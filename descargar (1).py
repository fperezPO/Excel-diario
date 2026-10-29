"""
Descarga automática de un Excel desde una página web y subida a SharePoint.

Qué hace, en orden:
  1. Abre la página indicada en config.yaml con un navegador invisible.
  2. Sigue los pasos de config.yaml (escribir el usuario, hacer clic, etc.).
  3. Hace clic en el botón de descarga y guarda el Excel.
  4. Sube el archivo a la carpeta de SharePoint indicada en config.yaml.

Las contraseñas nunca se escriben en este archivo ni en config.yaml:
se leen de los "secretos" del repositorio en GitHub.

Uso:  python descargar.py            (usa config.yaml)
      python descargar.py otra.yaml  (usa otro archivo de configuración)
"""

import os
import re
import sys
from datetime import datetime, timedelta
from functools import reduce
from pathlib import Path
from urllib.parse import quote, urlparse
from zoneinfo import ZoneInfo

import requests
import yaml
from playwright.sync_api import TimeoutError as TiempoAgotado
from playwright.sync_api import sync_playwright

CONFIG_PREDETERMINADA = Path(__file__).with_name("config.yaml")
CARPETA_DESCARGAS = Path("descargas")
CAPTURA_ERROR = Path("error.png")
ZONA_HORARIA = ZoneInfo("America/Montevideo")
ESPERA_PASO_MS = 30_000        # máximo de espera para que aparezca cada elemento
ESPERA_DESCARGA_MS = 180_000   # máximo de espera para que llegue el Excel
LOGIN = "https://login.microsoftonline.com"
GRAPH = "https://graph.microsoft.com/v1.0"

# Acción → dato obligatorio que indica sobre qué elemento actuar
ACCIONES = {
    "escribir": "campo",
    "elegir": "campo",
    "clic": "en",
    "descargar": "en",
    "ir_a": "url",
    "esperar": None,
}


class ErrorConfiguracion(Exception):
    """Un problema que se corrige en config.yaml o en los secretos de GitHub."""


# ─── Configuración ───────────────────────────────────────────────────────────

def cargar_config(ruta):
    try:
        config = yaml.safe_load(Path(ruta).read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise ErrorConfiguracion(f"No encuentro el archivo {ruta}.")
    except yaml.YAMLError as e:
        raise ErrorConfiguracion(f"{ruta} tiene un error de formato: {e}")

    if not isinstance(config, dict):
        raise ErrorConfiguracion(f"{ruta} está vacío o no tiene el formato esperado.")
    if "COMPLETAR" in str(config):  # solo mira los datos, no los comentarios
        raise ErrorConfiguracion(
            f"{ruta} todavía tiene datos marcados como COMPLETAR. "
            "Reemplazalos por los datos reales antes de ejecutar."
        )

    for clave in ("pagina_inicio", "pasos", "sharepoint"):
        if clave not in config:
            raise ErrorConfiguracion(f"Falta '{clave}' en {ruta}.")
    for clave in ("sitio", "carpeta", "nombre_archivo"):
        if clave not in config["sharepoint"]:
            raise ErrorConfiguracion(f"Falta '{clave}' dentro de 'sharepoint' en {ruta}.")

    pasos = config["pasos"]
    for n, paso in enumerate(pasos, start=1):
        accion = paso.get("accion")
        if accion not in ACCIONES:
            raise ErrorConfiguracion(
                f"Paso {n}: la acción '{accion}' no existe. "
                f"Opciones: {', '.join(ACCIONES)}."
            )
        dato = ACCIONES[accion]
        if dato and dato not in paso and "selector" not in paso:
            raise ErrorConfiguracion(f"Paso {n} ({accion}): falta '{dato}'.")
        if accion == "escribir" and "valor" not in paso and "valor_secreto" not in paso:
            raise ErrorConfiguracion(f"Paso {n} (escribir): falta 'valor' o 'valor_secreto'.")
        if accion == "elegir" and "opcion" not in paso:
            raise ErrorConfiguracion(f"Paso {n} (elegir): falta 'opcion'.")

    if not pasos or pasos[-1]["accion"] != "descargar":
        raise ErrorConfiguracion("El último paso tiene que ser 'descargar'.")
    if sum(p["accion"] == "descargar" for p in pasos) > 1:
        raise ErrorConfiguracion("Solo puede haber un paso 'descargar'.")
    return config


def verificar_secretos(config):
    """Comprueba al inicio que estén cargados todos los secretos que se van a usar."""
    necesarios = [p["valor_secreto"] for p in config["pasos"] if "valor_secreto" in p]
    necesarios += ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET"]
    faltan = [n for n in necesarios if not os.environ.get(n)]
    if faltan:
        raise ErrorConfiguracion(
            f"Faltan estos secretos: {', '.join(faltan)}. Cargalos en GitHub: Settings > "
            "Secrets and variables > Actions > New repository secret."
        )


def leer_secreto(nombre):
    valor = os.environ.get(nombre)
    if not valor:
        raise ErrorConfiguracion(
            f"Falta el secreto {nombre}. Cargalo en GitHub: Settings > "
            "Secrets and variables > Actions > New repository secret."
        )
    return valor


# ─── Navegador ───────────────────────────────────────────────────────────────

def describir(paso):
    """Texto corto del paso para el registro. Nunca muestra contraseñas."""
    objetivo = paso.get("selector") or paso.get("campo") or paso.get("en") or paso.get("url") or ""
    if paso["accion"] == "esperar":
        objetivo = f"{paso.get('segundos', 5)} segundos"
    return f"{paso['accion']} {objetivo}".strip()


def ubicar(page, paso, n):
    """Encuentra el elemento del paso buscando el texto que se ve en pantalla."""
    if "selector" in paso:  # opción avanzada: selector CSS exacto
        texto = paso["selector"]
        elemento = page.locator(texto).first
    else:
        if paso["accion"] in ("escribir", "elegir"):
            texto = str(paso["campo"])
            candidatos = [page.get_by_label(texto), page.get_by_placeholder(texto)]
        else:
            texto = str(paso["en"])
            candidatos = [
                page.get_by_role("button", name=texto),
                page.get_by_role("link", name=texto),
                page.get_by_role("menuitem", name=texto),
                page.get_by_role("tab", name=texto),
                page.get_by_text(texto, exact=True),
            ]
        elemento = reduce(lambda a, b: a.or_(b), candidatos).first

    try:
        elemento.wait_for(state="visible", timeout=ESPERA_PASO_MS)
    except TiempoAgotado:
        raise ErrorConfiguracion(
            f"Paso {n} ({paso['accion']}): no encontré en pantalla «{texto}». "
            "Revisá que el texto en config.yaml coincida con lo que se ve en la página. "
            "La captura de pantalla del error muestra dónde se detuvo."
        )
    return elemento


def ejecutar_pasos(page, pasos, fechas):
    """Sigue los pasos en orden y devuelve la descarga del último."""
    for n, paso in enumerate(pasos, start=1):
        accion = paso["accion"]
        print(f"Paso {n}: {describir(paso)}", flush=True)

        if accion == "ir_a":
            page.goto(str(paso["url"]))
        elif accion == "escribir":
            if "valor_secreto" in paso:
                valor = leer_secreto(paso["valor_secreto"])
            else:
                valor = str(paso["valor"]).format(**fechas)
            ubicar(page, paso, n).fill(valor)
        elif accion == "elegir":
            ubicar(page, paso, n).select_option(label=str(paso["opcion"]).format(**fechas))
        elif accion == "clic":
            ubicar(page, paso, n).click()
            page.wait_for_load_state()
        elif accion == "esperar":
            page.wait_for_timeout(float(paso.get("segundos", 5)) * 1000)
        elif accion == "descargar":
            boton = ubicar(page, paso, n)
            try:
                with page.expect_download(timeout=ESPERA_DESCARGA_MS) as info:
                    boton.click()
            except TiempoAgotado:
                raise ErrorConfiguracion(
                    f"Paso {n}: hice clic en «{describir(paso)}» pero no llegó ningún archivo "
                    f"en {ESPERA_DESCARGA_MS // 1000} segundos."
                )
            return info.value


def guardar(descarga, config, fechas):
    # El nombre original (con su serie aleatoria) se descarta; solo se conserva la extensión
    nombre = config["sharepoint"]["nombre_archivo"].format(**fechas)
    extension = Path(descarga.suggested_filename).suffix or ".xlsx"
    if not nombre.lower().endswith(extension.lower()):
        nombre += extension
    CARPETA_DESCARGAS.mkdir(exist_ok=True)
    destino = CARPETA_DESCARGAS / nombre
    descarga.save_as(destino)
    tamano = destino.stat().st_size
    if tamano == 0:
        raise ErrorConfiguracion("El archivo descargado está vacío.")
    print(f"Descargado: {destino.name} ({tamano:,} bytes)", flush=True)
    return destino


def descargar_excel(config, fechas):
    with sync_playwright() as p:
        navegador = p.chromium.launch()
        contexto = navegador.new_context(
            accept_downloads=True, locale="es-UY", timezone_id="America/Montevideo"
        )
        page = contexto.new_page()
        page.set_default_timeout(ESPERA_PASO_MS)
        try:
            print(f"Abriendo {config['pagina_inicio']}", flush=True)
            page.goto(str(config["pagina_inicio"]))
            descarga = ejecutar_pasos(page, config["pasos"], fechas)
            return guardar(descarga, config, fechas)
        except Exception:
            try:
                page.screenshot(path=CAPTURA_ERROR, full_page=True)
                print(f"Se guardó una captura de pantalla del error: {CAPTURA_ERROR}")
            except Exception:
                pass
            raise
        finally:
            navegador.close()


# ─── SharePoint (Microsoft Graph) ────────────────────────────────────────────

PISTAS = {
    400: "Revisá los datos de conexión con Microsoft y el nombre de la carpeta.",
    401: "Revisá MS_TENANT_ID, MS_CLIENT_ID y MS_CLIENT_SECRET. ¿Venció el secreto de Microsoft?",
    403: "La aplicación no tiene permiso sobre ese sitio. Pedile a TI que revise el permiso de escritura.",
    404: "No encontré el sitio, la biblioteca o la carpeta. Revisá la sección 'sharepoint' de config.yaml.",
}


def revisar(respuesta, que):
    if respuesta.ok:
        return
    try:
        datos = respuesta.json()
        error = datos.get("error")
        detalle = datos.get("error_description") or (
            error.get("message") if isinstance(error, dict) else error
        )
    except ValueError:
        detalle = respuesta.text[:300]
    raise ErrorConfiguracion(
        f"No se pudo {que} (código {respuesta.status_code}). "
        f"{PISTAS.get(respuesta.status_code, '')} Detalle: {detalle or 'sin detalle'}"
    )


def obtener_token():
    respuesta = requests.post(
        f"{LOGIN}/{leer_secreto('MS_TENANT_ID')}/oauth2/v2.0/token",
        data={
            "client_id": leer_secreto("MS_CLIENT_ID"),
            "client_secret": leer_secreto("MS_CLIENT_SECRET"),
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    revisar(respuesta, "conectar con Microsoft 365")
    return respuesta.json()["access_token"]


def normalizar_sitio(texto):
    """Devuelve (servidor, ruta del sitio) a partir de la dirección del sitio,
    de cualquier página dentro de él o de un enlace para compartir."""
    texto = str(texto).strip()
    url = urlparse(texto if "://" in texto else f"https://{texto}")
    partes = [p for p in url.path.split("/") if p]
    # Enlaces para compartir: /:f:/s/Sitio/... o /:x:/r/sites/Sitio/...
    if partes and re.fullmatch(r":[a-z]:", partes[0]):
        partes = partes[1:]
        if partes[:1] == ["s"]:
            partes = ["sites"] + partes[1:]
        elif partes[:1] == ["r"]:
            partes = partes[1:]
    if len(partes) >= 2 and partes[0] in ("sites", "teams"):
        return url.netloc, f"{partes[0]}/{partes[1]}"
    return url.netloc, ""


def nombres(respuesta, solo_carpetas=False):
    elementos = respuesta.json().get("value", []) if respuesta.ok else []
    return [e["name"] for e in elementos if not solo_carpetas or "folder" in e]


def verificar_carpeta(base, carpeta, cabeceras, sitio_id):
    """Evita que el archivo termine en una carpeta nueva creada por error."""
    respuesta = requests.get(f"{base}/root:/{quote(carpeta)}", headers=cabeceras, timeout=30)
    if respuesta.status_code == 404:
        raiz = requests.get(f"{base}/root/children", headers=cabeceras, timeout=30)
        bibliotecas = requests.get(f"{GRAPH}/sites/{sitio_id}/drives", headers=cabeceras, timeout=30)
        raise ErrorConfiguracion(
            f"No encontré la carpeta «{carpeta}» en SharePoint. "
            f"Carpetas en la biblioteca: {', '.join(nombres(raiz, True)) or 'ninguna'}. "
            f"Bibliotecas del sitio: {', '.join(nombres(bibliotecas)) or 'sin datos'}. "
            "Si está dentro de otra carpeta, escribí la ruta completa en 'carpeta' "
            "(por ejemplo «Stock/Foto stock WIS»); si está en otra biblioteca, completá 'biblioteca'."
        )
    revisar(respuesta, "encontrar la carpeta de SharePoint")
    if "folder" not in respuesta.json():
        raise ErrorConfiguracion(f"«{carpeta}» existe en SharePoint pero no es una carpeta.")


def subir_a_sharepoint(archivo, cfg):
    cabeceras = {"Authorization": f"Bearer {obtener_token()}"}

    host, ruta = normalizar_sitio(cfg["sitio"])
    url_sitio = f"{GRAPH}/sites/{host}:/{ruta}" if ruta else f"{GRAPH}/sites/{host}"
    respuesta = requests.get(url_sitio, headers=cabeceras, timeout=30)
    revisar(respuesta, "encontrar el sitio de SharePoint")
    sitio_id = respuesta.json()["id"]

    # Biblioteca de documentos: la principal ("Documentos"), salvo que se indique otra
    base = f"{GRAPH}/sites/{sitio_id}/drive"
    biblioteca = cfg.get("biblioteca")
    if biblioteca:
        respuesta = requests.get(f"{GRAPH}/sites/{sitio_id}/drives", headers=cabeceras, timeout=30)
        revisar(respuesta, "listar las bibliotecas del sitio")
        bibliotecas = {d["name"]: d["id"] for d in respuesta.json().get("value", [])}
        if biblioteca not in bibliotecas:
            raise ErrorConfiguracion(
                f"No existe la biblioteca «{biblioteca}». Disponibles: {', '.join(bibliotecas)}."
            )
        base = f"{GRAPH}/drives/{bibliotecas[biblioteca]}"

    carpeta = str(cfg["carpeta"]).strip("/")
    if carpeta:
        verificar_carpeta(base, carpeta, cabeceras, sitio_id)
    ruta_archivo = quote(f"{carpeta}/{archivo.name}" if carpeta else archivo.name)
    respuesta = requests.put(
        f"{base}/root:/{ruta_archivo}:/content",
        headers={**cabeceras, "Content-Type": "application/octet-stream"},
        data=archivo.read_bytes(),
        timeout=300,
    )
    revisar(respuesta, "subir el archivo a SharePoint")
    print(f"Subido a SharePoint: {respuesta.json().get('webUrl', archivo.name)}", flush=True)


# ─── Inicio ──────────────────────────────────────────────────────────────────

def main(ruta_config=CONFIG_PREDETERMINADA):
    config = cargar_config(ruta_config)
    verificar_secretos(config)
    ahora = datetime.now(ZONA_HORARIA)
    fechas = {"hoy": ahora, "ayer": ahora - timedelta(days=1)}
    archivo = descargar_excel(config, fechas)
    subir_a_sharepoint(archivo, config["sharepoint"])
    print("Listo.", flush=True)


if __name__ == "__main__":
    try:
        main(sys.argv[1] if len(sys.argv) > 1 else CONFIG_PREDETERMINADA)
    except ErrorConfiguracion as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        sys.exit(1)
