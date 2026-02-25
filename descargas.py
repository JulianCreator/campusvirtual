import os
import re
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import unquote, urlparse, parse_qs

import flet
from flet import (
    Page,
    TextField,
    Dropdown,
    dropdown,
    ElevatedButton,
    Text,
    Column,
    alignment,
    Image,
    Icons,  # Asegúrate de usar 'Icons' en mayúsculas
    Container
)

# Intentamos importar pypdf para leer metadatos de PDF (opcional).
try:
    import pypdf  # pip install pypdf
    HAVE_PYPDF = True
except ImportError:
    HAVE_PYPDF = False

# Lista global para almacenar recursos antes de generar Excel
RECURSOS_ENCONTRADOS = []

# Credenciales únicas (mismas para las 3 plataformas)
MOODLE_USER = "extraccion_ustacv"
MOODLE_PASS = "000"

# Opciones de plataformas con nombre, URL, color y sufijo de carpeta
PLATAFORMAS = [
    {"name": "Pregrado", "url": "https://pregrado.ustabuca.edu.co", "color": "#00dba7", "folder_suffix": "Pregrado"},
    {"name": "Posgrado", "url": "https://posgrado.ustabuca.edu.co", "color": "#7700ca", "folder_suffix": "Posgrado"},
    {"name": "Educación Continua", "url": "https://educacioncontinua.ustabuca.edu.co", "color": "#06ff3b", "folder_suffix": "Educontinua"},
]

# -----------------------------------------------------------------------------
# FUNCIONES DE AYUDA
# -----------------------------------------------------------------------------
def limpiar_nombre(nombre):
    """
    Limpia el nombre eliminando caracteres no permitidos y limitando su longitud.
    """
    nombre_limpio = re.sub(r'[<>:"/\\|?*]', '_', nombre)
    return nombre_limpio[:100]

def remover_trailing_archivo(nombre):
    """
    Elimina la palabra 'Archivo' al final del nombre o antes de la extensión.
    """
    nombre = re.sub(r'\s*Archivo\s*$', '', nombre, flags=re.IGNORECASE)
    nombre = re.sub(r'(.*)(Archivo)(\.[^.]+)$', r'\1\3', nombre, flags=re.IGNORECASE)
    return nombre.strip()

def obtener_nombre_desde_url(url):
    """
    Extrae el nombre del archivo desde una URL.
    """
    return unquote(url.split('/')[-1].split('?')[0])

def extraer_url_onclick(onclick_value):
    """
    Extrae la URL de un atributo onclick que contiene window.open.
    """
    if not onclick_value:
        return None
    # Regex corregida: detecta window.open('...') y captura todo entre comillas simples
    match = re.search(r"window\.open\('([^']+)'", onclick_value)
    if match:
        return match.group(1)
    return None

# -----------------------------------------------------------------------------
# LÓGICA DE MOODLE
# -----------------------------------------------------------------------------
def iniciar_sesion(url_base, log_area):
    """
    Inicia sesión en la plataforma Moodle dada (url_base)
    con las credenciales MOODLE_USER / MOODLE_PASS.
    """
    login_url = f"{url_base}/login/index.php"
    session = requests.Session()
    try:
        r = session.get(login_url, timeout=10)
        r.raise_for_status()
    except requests.RequestException as e:
        log_area.value = "[ERROR] Al acceder a la página de login."
        log_area.color = "red"
        log_area.update()
        return None

    soup = BeautifulSoup(r.text, 'html.parser')
    token_input = soup.find("input", {"name": "logintoken"})
    if not token_input:
        log_area.value = "[ERROR] No se encontró logintoken en la página de login."
        log_area.color = "red"
        log_area.update()
        return None

    token_value = token_input["value"]
    login_data = {
        "username": MOODLE_USER,
        "password": MOODLE_PASS,
        "logintoken": token_value
    }

    try:
        lr = session.post(login_url, data=login_data, timeout=10)
        lr.raise_for_status()
    except requests.RequestException as e:
        log_area.value = "[ERROR] Al enviar credenciales."
        log_area.color = "red"
        log_area.update()
        return None

    if "loginerrormessage" not in lr.text.lower():
        log_area.value = "[INFO] Sesión iniciada correctamente."
        log_area.color = "green"
        log_area.update()
        return session
    else:
        log_area.value = "[ERROR] No se pudo iniciar sesión. Revisa credenciales."
        log_area.color = "red"
        log_area.update()
        return None

def obtener_soup(session, url, log_area):
    """
    Obtiene y parsea el contenido HTML de una URL utilizando la sesión dada.
    """
    try:
        resp = session.get(url, timeout=10)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, 'html.parser')
    except requests.RequestException as e:
        log_area.value = f"[ERROR] Al acceder a {url}."
        log_area.color = "red"
        log_area.update()
        return None

def obtener_num_secciones(session, url_base, id_curso, log_area):
    """
    Obtiene el número máximo de secciones en un curso dado.
    """
    url_curso = f"{url_base}/course/view.php?id={id_curso}"
    soup = obtener_soup(session, url_curso, log_area)
    if not soup:
        return 0

    max_seccion = 0
    anchors = soup.find_all('a', href=True)
    for a in anchors:
        href = a["href"]
        if "course/view.php" in href and "section=" in href:
            parsed = urlparse(href)
            q = parse_qs(parsed.query)
            if 'section' in q:
                try:
                    sec_num = int(q['section'][0])
                    if sec_num > max_seccion:
                        max_seccion = sec_num
                except ValueError:
                    pass
    return max_seccion

def obtener_nombre_curso(session, url_base, id_curso, log_area):
    """
    Obtiene el nombre del curso a partir de su ID.
    """
    url_curso = f"{url_base}/course/view.php?id={id_curso}"
    soup = obtener_soup(session, url_curso, log_area)
    if not soup:
        return f"Curso_{id_curso}"

    titulo = soup.find('h1') or soup.find('title') or soup.find('h2')
    if titulo:
        return limpiar_nombre(titulo.get_text(strip=True))
    return f"Curso_{id_curso}"

def obtener_tuplas_intermedias(session, url, log_area):
    """
    Obtiene recursos intermedios desde una URL específica.
    """
    soup = obtener_soup(session, url, log_area)
    if not soup:
        return []
    resultados = []

    div_res = soup.find('div', class_='resourceworkaround')
    if div_res:
        enlace = div_res.find('a', href=True)
        if enlace:
            url_final = enlace['href']
            texto = enlace.get_text(strip=True)
            texto = remover_trailing_archivo(texto)
            resultados.append((url_final, texto))

    a_plug = soup.find_all('a', href=lambda h: h and 'pluginfile.php' in h)
    for anc in a_plug:
        link_href = anc['href']
        text_anch = anc.get_text(strip=True) or ""
        text_anch = remover_trailing_archivo(text_anch)
        if not text_anch:
            text_anch = None
        resultados.append((link_href, text_anch))

    iframes = soup.find_all('iframe', src=lambda x: x and 'pluginfile.php' in x)
    embeds = soup.find_all('embed', src=lambda x: x and 'pluginfile.php' in x)
    objs = soup.find_all('object', data=lambda x: x and 'pluginfile.php' in x)
    for tag in iframes + embeds + objs:
        su = tag.get('src') or tag.get('data')
        resultados.append((su, None))

    # Quitar duplicados
    unique = []
    seen = set()
    for (u, t) in resultados:
        if u and u not in seen:
            seen.add(u)
            unique.append((u, t))

    return unique

def obtener_links_recursos(session, url_base, id_curso, seccion_num, log_area):
    """
    Retorna una lista de recursos con su URL, nombre y tipo.
    """
    url_seccion = f"{url_base}/course/view.php?id={id_curso}&section={seccion_num}"
    soup = obtener_soup(session, url_seccion, log_area)
    if not soup:
        return []

    recursos = []
    enlaces = soup.find_all('a', href=True)
    for enlace in enlaces:
        instancename = enlace.find('span', class_='instancename')
        if instancename:
            sub_ac = instancename.find('span', class_='accesshide')
            if sub_ac:
                sub_ac.decompose()
            nombre_visible = instancename.get_text(strip=True)
        else:
            nombre_visible = enlace.get_text(strip=True)

        nombre_visible = remover_trailing_archivo(nombre_visible)
        url_h = enlace['href']

        onclick_val = enlace.get('onclick', '')
        match_ = extraer_url_onclick(onclick_val)
        if match_ and 'redirect=1' in match_:
            url_h = match_

        if "pluginfile.php" in url_h:
            real_name = limpiar_nombre(obtener_nombre_desde_url(url_h))
            if len(real_name) > 70:
                real_name = real_name[:70]
            recursos.append((url_h, real_name, "file"))

        elif "mod/resource/view.php" in url_h:
            try:
                rh = session.head(url_h, allow_redirects=True, timeout=10)
                ctype = rh.headers.get('Content-Type', '')
                if ctype.startswith('application/'):
                    final_name = limpiar_nombre(nombre_visible)
                    final_name = remover_trailing_archivo(final_name)
                    if len(final_name) > 70:
                        fallback = limpiar_nombre(obtener_nombre_desde_url(url_h))
                        final_name = fallback or final_name
                    recursos.append((url_h, final_name, "file"))
                else:
                    tuplas = obtener_tuplas_intermedias(session, url_h, log_area)
                    for (fu, ft) in tuplas:
                        if ft:
                            nm = ft
                        else:
                            nm = nombre_visible
                        nm = limpiar_nombre(remover_trailing_archivo(nm))
                        if len(nm) > 70:
                            fb = limpiar_nombre(obtener_nombre_desde_url(fu))
                            nm = fb or nm
                        recursos.append((fu, nm, "file"))
            except requests.RequestException as e:
                log_area.value = "[ERROR] HEAD en " + url_h
                log_area.color = "red"
                log_area.update()

        elif "mod/url/view.php" in url_h:
            # Recurso URL
            recursos.append((url_h, nombre_visible, "url"))
        else:
            pass

    return recursos

def descargar_archivo(session, url, carpeta_destino, nombre_archivo, log_area):
    """
    Descarga un archivo desde una URL y lo guarda en la carpeta destino con el nombre especificado.
    """
    try:
        ruta_inicial = os.path.join(carpeta_destino, nombre_archivo)
        if os.path.exists(ruta_inicial):
            os.remove(ruta_inicial)

        r = session.get(url, stream=True, allow_redirects=True, timeout=10)
        if r.status_code != 200:
            log_area.value = "[ERROR] al descargar " + url
            log_area.color = "red"
            log_area.update()
            return False

        ctype = r.headers.get('Content-Type', '').lower()
        extension = None

        if '.rtf' in url.lower():
            extension = '.rtf'
        else:
            if 'application/pdf' in ctype:
                extension = '.pdf'
            elif 'application/msword' in ctype:
                extension = '.doc'
            elif 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' in ctype:
                extension = '.docx'
            elif 'application/vnd.ms-excel' in ctype:
                extension = '.xls'
            elif 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in ctype:
                extension = '.xlsx'
            elif 'application/rtf' in ctype or 'text/rtf' in ctype:
                extension = '.rtf'
            else:
                ext_fallback = os.path.splitext(url.split('?')[0])[-1]
                extension = ext_fallback

        if extension and not nombre_archivo.lower().endswith(extension.lower()):
            nombre_archivo += extension

        ruta_final = os.path.join(carpeta_destino, nombre_archivo)

        with open(ruta_final, 'wb') as f:
            for chunk in r.iter_content(chunk_size=1024):
                if chunk:
                    f.write(chunk)

        log_area.value = "[INFO] Archivo descargado: " + ruta_final
        log_area.color = "green"
        log_area.update()

        if extension == '.pdf' and HAVE_PYPDF:
            try:
                from pypdf import PdfReader
                with open(ruta_final, 'rb') as pdf_f:
                    reader = PdfReader(pdf_f)
                    info = reader.metadata
                    titulo_pdf = info.get('/Title', '') if info else ''
                time.sleep(0.2)
                if titulo_pdf:
                    titulo_pdf_limpio = limpiar_nombre(titulo_pdf.strip())[:70]
                    if titulo_pdf_limpio:
                        ruta_renombrada = os.path.join(carpeta_destino, titulo_pdf_limpio + '.pdf')
                        if not os.path.exists(ruta_renombrada):
                            os.rename(ruta_final, ruta_renombrada)
                            log_area.value = "[INFO] Renombrado PDF: " + ruta_renombrada
                            log_area.color = "green"
                            log_area.update()
            except Exception as e:
                log_area.value = "[WARN] No se pudo leer metadatos PDF."
                log_area.color = "orange"
                log_area.update()

        return True
    except Exception as e:
        log_area.value = "[ERROR] Al procesar " + url
        log_area.color = "red"
        log_area.update()
        return False

def recorrer_secciones_curso(session, url_base, id_curso, carpeta_curso, log_area, nombre_curso):
    """
    Recorre todas las secciones de un curso y descarga los recursos.
    """
    RECURSOS_ENCONTRADOS.clear()
    max_sec = obtener_num_secciones(session, url_base, id_curso, log_area)
    log_area.value = f"[INFO] El curso {id_curso} ({nombre_curso}) tiene secciones de 0 a {max_sec}."
    log_area.color = "blue"
    log_area.update()

    for sec in range(max_sec + 1):
        recs = obtener_links_recursos(session, url_base, id_curso, sec, log_area)
        log_area.value += f"\n[INFO] Sección {sec}: {len(recs)} recursos."
        log_area.update()
        if not recs:
            continue

        carpeta_secc = os.path.join(carpeta_curso, f"Seccion_{sec}")
        os.makedirs(carpeta_secc, exist_ok=True)

        for (ur, nm, tipo) in recs:
            if tipo == "url":
                # No se descarga => "presente"
                RECURSOS_ENCONTRADOS.append({
                    "ID_Curso": id_curso,
                    "Nombre_Curso": nombre_curso,
                    "Seccion": sec,
                    "Nombre": nm,
                    "Vinculo": ur,
                    "Estado": "presente"
                })
            else:
                # Archivos => se intenta descargar
                ok = descargar_archivo(session, ur, carpeta_secc, nm, log_area)
                estado = "descargada" if ok else "ausente"
                RECURSOS_ENCONTRADOS.append({
                    "ID_Curso": id_curso,
                    "Nombre_Curso": nombre_curso,
                    "Seccion": sec,
                    "Nombre": nm,
                    "Vinculo": ur,
                    "Estado": estado
                })

# -----------------------------------------------------------------------------
# INTERFAZ FLET
# -----------------------------------------------------------------------------
def main(page: Page):
    page.title = "Recursos Campus Virtual"
    # Configurar dimensiones de la ventana (actualizado a versiones recientes de Flet)
    page.window.width = 600  # Ventana más pequeña
    page.window.height = 500  # Ventana más pequeña

    # Mensaje de advertencia para el icono
    advertencia_icono = Text(
        value="",
        size=12,
        color="orange",
        text_align="center"
    )

    # Configurar el icono de la aplicación
    icon_path = "icono.ico"
    if not os.path.exists(icon_path):
        print(f"Advertencia: No se encontró el archivo de ícono en la ruta '{icon_path}'.")
        advertencia_icono.value = f"Advertencia: No se encontró el archivo de ícono en la ruta '{icon_path}'."
    else:
        page.window.icon = os.path.abspath(icon_path)
    page.update()

    # Establecer alineación centrada
    page.vertical_alignment = alignment.center
    page.horizontal_alignment = alignment.center

    # Logo
    try:
        logo = Image(src="logo.png", width=300, height=75)  # Logo más pequeño
    except Exception as e:
        logo = Text(value="[ERROR] No se encontró 'logo.png'", color="red")

    # Título con font personalizada
    titulo = Text(
        value="DESCARGA DE RECURSOS MOODLE",
        size=18,  # Tamaño de fuente reducido
        weight="bold",
        color="blue",
        text_align="center",
        font_family="Palette_Bold"  # Asegúrate de que la fuente 'Palette_Bold' esté instalada
    )

    # Dropdown de plataformas (drop_categoria)
    plataforma_dropdown = Dropdown(
        label="Plataforma Moodle",
        options=[
            dropdown.Option(text=platform["name"], data=platform["url"]) for platform in PLATAFORMAS
        ],
        value=PLATAFORMAS[0]["name"],  # Establecer el valor predeterminado usando el texto
        width=400  # Ancho reducido
    )

    # Campo para ID curso (input_rango)
    curso_id_field = TextField(label="ID del curso", width=400)  # Ancho reducido

    # Línea horizontal superior (Container)
    linea_superior = Container(
        height=10,
        bgcolor=PLATAFORMAS[0]["color"],  # Color por defecto basado en la primera plataforma
        width=500  # Ancho reducido
    )

    # Botón para descargar (btn_iniciar)
    descargar_btn = ElevatedButton(
        text="Descargar Recursos",
        icon=Icons.DOWNLOAD,  # Usar Icons en lugar de icons
        width=400,  # Ancho reducido
        bgcolor=PLATAFORMAS[0]["color"],  # Establecer color inicial
        color="white"  # Color del texto para buena legibilidad
    )

    # Mensaje de estado (status_text)
    estado_text = Text(
        value="",
        size=14,
        color="green",
        text_align="center"
    )

    # Función al cambiar la selección del dropdown para actualizar el color de la línea y del botón
    def on_plataforma_change(e):
        selected_text = e.control.value  # Obtener el texto seleccionado
        # Encontrar la plataforma que coincide con el texto seleccionado
        selected_platform = next((platform for platform in PLATAFORMAS if platform["name"] == selected_text), None)
        if selected_platform:
            # Cambiar el color de la línea horizontal
            linea_superior.bgcolor = selected_platform["color"]
            linea_superior.update()
            
            # Cambiar el color del botón de descarga
            descargar_btn.bgcolor = selected_platform["color"]
            descargar_btn.color = "white"  # Ajustar color del texto si es necesario
            descargar_btn.update()

    # Asignar la función al evento on_change del dropdown
    plataforma_dropdown.on_change = on_plataforma_change

    # Función al hacer clic en el botón
    def on_descargar_click(e):
        # Limpiar estado
        estado_text.value = "Iniciando proceso de descarga..."
        estado_text.color = "blue"
        estado_text.update()

        # Recuperar la plataforma seleccionada
        selected_text = plataforma_dropdown.value
        selected_platform = next((platform for platform in PLATAFORMAS if platform["name"] == selected_text), None)
        if selected_platform:
            base_url = selected_platform["url"]
        else:
            estado_text.value = "Por favor, selecciona una plataforma."
            estado_text.color = "red"
            estado_text.update()
            return

        curso_id = curso_id_field.value.strip()

        if not curso_id:
            estado_text.value = "Por favor, ingresa un ID de curso."
            estado_text.color = "red"
            estado_text.update()
            return

        # Iniciar sesión
        ses = iniciar_sesion(base_url, estado_text)
        page.update()
        if not ses:
            estado_text.value = "Fallo al iniciar sesión."
            estado_text.color = "red"
            estado_text.update()
            return

        # Obtener nombre del curso
        nombre_curso = obtener_nombre_curso(ses, base_url, curso_id, estado_text)

        # Determinar el sufijo de la carpeta basado en la plataforma
        folder_suffix = selected_platform["folder_suffix"]

        # Crear carpeta del curso con nombre dinámico basado en la plataforma
        base_dir = f"Descargas_{folder_suffix}"
        os.makedirs(base_dir, exist_ok=True)
        carpeta_curso = os.path.join(base_dir, nombre_curso_corto := (curso_id if selected_platform["name"].lower() in ["posgrado", "educación continua"] else nombre_curso[:10]))
        os.makedirs(carpeta_curso, exist_ok=True)

        # Recorrer secciones
        recorrer_secciones_curso(ses, base_url, curso_id, carpeta_curso, estado_text, nombre_curso)

        # Generar Excel
        df = pd.DataFrame(RECURSOS_ENCONTRADOS, columns=["ID_Curso", "Nombre_Curso", "Seccion", "Nombre", "Vinculo", "Estado"])
        excel_path = os.path.join(carpeta_curso, "recursos.xlsx")
        try:
            df.to_excel(excel_path, index=False)
            estado_text.value = f"Proceso completado.\nExcel generado en: {excel_path}"
            estado_text.color = "green"
        except Exception as e:
            estado_text.value = f"Fallo al generar Excel: {e}"
            estado_text.color = "red"

        estado_text.update()

    # Asignar la función al botón
    descargar_btn.on_click = on_descargar_click

    # Añadir todos los elementos a la página con el diseño solicitado
    page.add(
        Container(
            alignment=alignment.center,
            # Puedes ajustar el padding si es necesario
            content= Column(
                horizontal_alignment="center",
                alignment="center",
                spacing=15,  # Reducir espaciado para una apariencia más compacta
                controls=[
                    # Línea horizontal superior
                    linea_superior,
                    # Logo
                    logo,
                    # Título
                    titulo,
                    # Dropdown de plataformas
                    plataforma_dropdown,
                    # Campo para ID del curso
                    curso_id_field,
                    # Botón para descargar
                    descargar_btn,
                    # Mensaje de estado
                    estado_text,
                    # Mensaje de advertencia del icono (si aplica)
                    advertencia_icono
                ]
            )
        )
    )

# Ejecutar flet
if __name__ == "__main__":
    flet.app(target=main)

