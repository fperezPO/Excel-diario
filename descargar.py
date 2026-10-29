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
import time
from datetime import datetime, timedelta
from functools import reduce
from pathlib import Path
from urllib.parse import quote, urlparse
from zoneinfo import ZoneInfo

import requests
import yaml
from playwright.sync_api import Error as ErrorPlaywright
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
    "tecla": None,
}

# Nombres de elementos HTML que se pueden usar solos en un selector (sin punto)
ETIQUETAS_HTML = {
    "a", "button", "div", "span", "input", "select", "textarea", "label", "img", "i",
    "li", "ul", "ol", "table", "thead", "tbody", "tr", "td", "th", "form", "nav", "p",
}

# Nombres de teclas que se pueden escribir en config.yaml
TECLAS = {
    "enter": "Enter", "intro": "Enter",
    "escape": "Escape", "esc": "Escape",
    "tab": "Tab", "tabulador": "Tab",
}


class ErrorConfiguracion(Exception):
    """Un problema que se corrige en config.yaml o en los secretos de GitHub."""


# ─── Configuración ───────────────────────────────────────────────────────────

class LectorSinRepetidos(yaml.SafeLoader):
    """Lee config.yaml avisando si un dato aparece dos veces en el mismo paso
    (normalmente se usaría el último sin decir nada)."""

    def construct_mapping(self, node, deep=False):
        vistos = set()
        for nodo_clave, _ in node.value:
            clave = self.construct_object(nodo_clave, deep=deep)
            if clave in vistos:
                raise ErrorConfiguracion(
                    f"'{clave}' aparece dos veces en el mismo paso "
                    f"(línea {nodo_clave.start_mark.line + 1}). Dejá solo uno."
                )
            vistos.add(clave)
        return super().construct_mapping(node, deep)


def cargar_config(ruta):
    try:
        config = yaml.load(Path(ruta).read_text(encoding="utf-8"), Loader=LectorSinRepetidos)
    except FileNotFoundError:
        raise ErrorConfiguracion(f"No encuentro el archivo {ruta}.")
    except yaml.YAMLError as e:
        marca = getattr(e, "problem_mark", None)
        donde = f" cerca de la línea {marca.line + 1}" if marca else ""
        raise ErrorConfiguracion(
            f"{ruta} tiene un error de formato{donde}. Revisá la sangría: cada paso empieza "
            "con 2 espacios y un guion («  - accion: ...») y sus datos van debajo con 4 espacios. "
            "Usá espacios, no tabulaciones."
        )

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
    inicio = str(config["pagina_inicio"]).strip()
    if not re.match(r"https?://", inicio) or inicio.count("://") > 1:
        raise ErrorConfiguracion(
            f"La dirección de 'pagina_inicio' no es válida: «{inicio}». "
            "Tiene que empezar con https:// una sola vez."
        )

    pasos = config["pasos"]
    for n, paso in enumerate(pasos, start=1):
        accion = paso.get("accion")
        if accion not in ACCIONES:
            raise ErrorConfiguracion(
                f"Paso {n}: la acción '{accion}' no existe. "
                f"Opciones: {', '.join(ACCIONES)}."
            )
        if "columna" in paso:
            if accion not in ("escribir", "elegir", "clic"):
                raise ErrorConfiguracion(f"Paso {n} ({accion}): 'columna' no se usa con esta acción.")
            if "selector" not in paso:
                raise ErrorConfiguracion(
                    f"Paso {n}: con 'columna' también va 'selector', la clase que comparten "
                    'los filtros de la grilla (por ejemplo selector: ".gr-filter").'
                )
        for clave in ("en", "campo"):
            texto = str(paso.get(clave, ""))
            if re.match(r"[.#][A-Za-z_]", texto):
                raise ErrorConfiguracion(
                    f"Paso {n}: «{texto}» es una clase o un id, no un texto de pantalla. "
                    f"Cambiá «{clave}:» por «selector:»."
                )
        dato = ACCIONES[accion]
        if dato and dato not in paso and "selector" not in paso:
            raise ErrorConfiguracion(f"Paso {n} ({accion}): falta '{dato}'.")
        if accion == "escribir" and "valor" not in paso and "valor_secreto" not in paso:
            raise ErrorConfiguracion(f"Paso {n} (escribir): falta 'valor' o 'valor_secreto'.")
        if accion == "elegir" and "opcion" not in paso:
            raise ErrorConfiguracion(f"Paso {n} (elegir): falta 'opcion'.")
        if accion == "tecla" and str(paso.get("tecla", "")).lower() not in TECLAS:
            raise ErrorConfiguracion(
                f"Paso {n} (tecla): las teclas disponibles son Enter, Escape y Tab."
            )
        if "selector" in paso:
            revisar_selector(paso["selector"], n)

    if not pasos or pasos[-1]["accion"] != "descargar":
        raise ErrorConfiguracion("El último paso tiene que ser 'descargar'.")
    if sum(p["accion"] == "descargar" for p in pasos) > 1:
        raise ErrorConfiguracion("Solo puede haber un paso 'descargar'.")
    return config


def revisar_selector(selector, n):
    """Detecta los errores más comunes al copiar un ícono desde 'Inspeccionar'."""
    if not selector:
        raise ErrorConfiguracion(
            f"Paso {n}: 'selector' está vacío. Si empieza con # tiene que ir entre comillas, "
            'por ejemplo selector: "#menu".'
        )
    selector = str(selector).strip()
    if selector.startswith("<"):
        etiqueta = selector.split(">", 1)[0]
        clases = re.search(r'class="([^"]*)"', etiqueta)
        ident = re.search(r'id="([^"]*)"', etiqueta)
        if ident and ident.group(1).strip():
            sugerencia = "#" + ident.group(1).strip()
        elif clases and clases.group(1).split():
            sugerencia = "." + ".".join(clases.group(1).split())
        else:
            sugerencia = None
        raise ErrorConfiguracion(
            f"Paso {n}: en 'selector' va la clase con un punto adelante, no el HTML completo."
            + (f' Para ese elemento sería: selector: "{sugerencia}"' if sugerencia else "")
        )
    # Una clase escrita sin el punto («gr-toolbar-btn filter» en vez de «.gr-toolbar-btn.filter»)
    palabras = selector.split()
    if not re.search(r"[.#\[\]:>*]", selector) and not all(p.lower() in ETIQUETAS_HTML for p in palabras):
        sugerencia = "." + ".".join(palabras)
        if len(palabras) == 1:
            consejo = f'Si es una clase: selector: "{sugerencia}". Si es un id: selector: "#{palabras[0]}".'
        else:
            consejo = (f'Si son las clases de un mismo elemento, van pegadas y cada una con '
                       f'su punto: selector: "{sugerencia}".')
        raise ErrorConfiguracion(f"Paso {n}: al selector «{selector}» le falta el punto adelante. {consejo}")


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

def objetivo(paso):
    """Sobre qué actúa el paso: el texto, el selector, la dirección o la tecla."""
    if paso.get("columna"):
        return f"columna {paso['columna']}"
    for clave in ("selector", "campo", "en", "url", "tecla"):
        if paso.get(clave):
            return str(paso[clave])
    return ""


def describir(paso):
    """Texto corto del paso para el registro. Nunca muestra contraseñas."""
    if paso["accion"] == "esperar":
        return f"esperar {paso.get('segundos', 5)} segundos"
    return f"{paso['accion']} {objetivo(paso)}".strip()


class Red:
    """Lleva la cuenta de los pedidos de datos de la página, para no avanzar
    mientras una búsqueda o una grilla todavía se está cargando."""

    TIPOS = ("document", "xhr", "fetch")

    def __init__(self, page):
        self.page = page
        self.pendientes = set()
        page.on("request", self._empieza)
        page.on("requestfinished", self._termina)
        page.on("requestfailed", self._termina)
        page.on("framenavigated", self._navega)

    def _empieza(self, pedido):
        if pedido.resource_type in self.TIPOS:
            self.pendientes.add(pedido)

    def _termina(self, pedido):
        self.pendientes.discard(pedido)

    def _navega(self, marco):
        if marco == self.page.main_frame:
            self.pendientes.clear()

    def esperar(self, quieto_ms=700, maximo_ms=15_000):
        """Espera a que la página pase un momento sin pedir datos (con un máximo)."""
        quieto = transcurrido = 0
        while quieto < quieto_ms and transcurrido < maximo_ms:
            self.page.wait_for_timeout(100)
            transcurrido += 100
            quieto = 0 if self.pendientes else quieto + 100


JS_FILTRO_DE_COLUMNA = """([titulo, selector]) => {
    const normalizar = (s) => (s || '').normalize('NFD').replace(/[\\u0300-\\u036f]/g, '')
        .replace(/\\s+/g, ' ').trim().toLowerCase();
    const visible = (el) => {
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== 'hidden';
    };

    // Títulos visibles con exactamente ese texto (sin importar tildes ni mayúsculas)
    const buscado = normalizar(titulo);
    const titulos = [];
    const recorrido = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    while (recorrido.nextNode()) {
        const padre = recorrido.currentNode.parentElement;
        if (padre && normalizar(recorrido.currentNode.nodeValue) === buscado && visible(padre)) {
            titulos.push(padre);
        }
    }
    if (!titulos.length) return 'sin-titulo';

    let filtros;
    try {
        filtros = [...document.querySelectorAll(selector)].filter(visible);
    } catch (e) {
        return 'selector-invalido';
    }
    if (!filtros.length) return 'sin-filtros';

    // El filtro que está en la misma columna que el título y más cerca de él
    let elegido = null, distancia = Infinity;
    for (const t of titulos) {
        const rt = t.getBoundingClientRect();
        const centro = rt.left + rt.width / 2;
        for (const f of filtros) {
            if (f.contains(t) || t.contains(f)) continue;
            const rf = f.getBoundingClientRect();
            if (centro < rf.left || centro > rf.right) continue;
            const d = Math.abs((rf.top + rf.bottom) / 2 - (rt.top + rt.bottom) / 2);
            if (d < distancia) { distancia = d; elegido = f; }
        }
    }
    if (!elegido) return 'no-alineado';
    if (elegido.matches('input, select, textarea')) return elegido;
    return elegido.querySelector('input, select, textarea') || elegido;
}"""


def ubicar_en_columna(page, paso, n):
    """Entre los elementos del selector, devuelve el que está debajo del título de la columna."""
    titulo, selector = str(paso["columna"]), str(paso["selector"])
    limite = time.monotonic() + ESPERA_PASO_MS / 1000
    while True:
        resultado = page.evaluate_handle(JS_FILTRO_DE_COLUMNA, [titulo, selector])
        elemento = resultado.as_element()
        if elemento:
            return elemento
        motivo = resultado.json_value()
        resultado.dispose()
        if motivo == "selector-invalido" or time.monotonic() > limite:
            break
        page.wait_for_timeout(300)

    mensajes = {
        "sin-titulo": f"no encontré en pantalla una columna titulada «{titulo}». "
                      "Revisá que el título esté escrito igual que en la grilla.",
        "selector-invalido": f"«{selector}» no es un selector válido.",
        "sin-filtros": f"no encontré en pantalla elementos «{selector}».",
        "no-alineado": f"encontré la columna «{titulo}», pero ningún «{selector}» debajo de ella.",
    }
    raise ErrorConfiguracion(
        f"Paso {n}: {mensajes[motivo]} La captura de pantalla del error muestra dónde se detuvo."
    )


def ubicar(page, paso, n):
    """Encuentra el elemento del paso buscando el texto que se ve en pantalla."""
    if "columna" in paso:  # filtros de una grilla: el que está debajo del título
        return ubicar_en_columna(page, paso, n)
    if "selector" in paso:  # elementos sin texto (íconos): clase o id del elemento
        texto = str(paso["selector"])
        elemento = page.locator(texto).filter(visible=True).first
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
                page.get_by_title(texto),  # íconos con texto al pasar el mouse
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


def hacer_clic(elemento, paso, n):
    try:
        elemento.click()
    except TiempoAgotado:
        raise ErrorConfiguracion(
            f"Paso {n}: encontré «{objetivo(paso)}» pero no pude hacer clic porque algo lo tapa "
            "(por ejemplo un calendario o un aviso abierto). Probá agregar antes un paso "
            "«accion: tecla» con «tecla: Escape». La captura de pantalla muestra qué lo tapa."
        )


JS_CARGAR_VALOR = """(campo, valor) => {
    campo.value = valor;
    campo.dispatchEvent(new Event('input', {bubbles: true}));
    campo.dispatchEvent(new Event('change', {bubbles: true}));
}"""


def a_formato_iso(valor):
    """Los campos de fecha propios del navegador solo aceptan año-mes-día."""
    for formato in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(valor.strip(), formato).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return valor


def escribir(campo, valor, paso, n):
    secreto = "valor_secreto" in paso
    try:
        if (campo.get_attribute("type") or "").lower() == "date":
            valor = a_formato_iso(valor)
        if not secreto:
            print(f"   valor: {valor}", flush=True)
        if campo.is_editable():
            campo.fill(valor)
        else:
            # Campos que solo se completan eligiendo en un calendario
            print("   El campo no deja escribir: cargo el valor directamente.", flush=True)
            campo.evaluate(JS_CARGAR_VALOR, valor)
    except ErrorPlaywright as e:
        detalle = "" if secreto else f" Detalle: {str(e).splitlines()[0]}"
        raise ErrorConfiguracion(
            f"Paso {n}: no pude escribir en «{objetivo(paso)}». Revisá que sea un campo para "
            f"escribir y que el formato coincida con el que muestra la página.{detalle}"
        )


def ejecutar_pasos(page, pasos, fechas):
    """Sigue los pasos en orden y devuelve la descarga del último."""
    red = Red(page)
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
            escribir(ubicar(page, paso, n), valor, paso, n)
        elif accion == "elegir":
            opcion = str(paso["opcion"]).format(**fechas)
            try:
                ubicar(page, paso, n).select_option(label=opcion)
            except TiempoAgotado:
                raise ErrorConfiguracion(
                    f"Paso {n}: la lista «{objetivo(paso)}» no tiene la opción «{opcion}»."
                )
        elif accion == "clic":
            hacer_clic(ubicar(page, paso, n), paso, n)
            page.wait_for_load_state()
            red.esperar()
        elif accion == "tecla":
            page.keyboard.press(TECLAS[str(paso["tecla"]).lower()])
            page.wait_for_load_state()
            red.esperar()
        elif accion == "esperar":
            page.wait_for_timeout(float(paso.get("segundos", 5)) * 1000)
        elif accion == "descargar":
            boton = ubicar(page, paso, n)
            try:
                with page.expect_download(timeout=ESPERA_DESCARGA_MS) as info:
                    hacer_clic(boton, paso, n)
            except TiempoAgotado:
                raise ErrorConfiguracion(
                    f"Paso {n}: hice clic en «{objetivo(paso)}» pero no llegó ningún archivo "
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
