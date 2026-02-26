import time #Para intriducir pausas controladas entre solicitudes HTTP y evitar saturar el servidor durante el scraping
import random #Generar invertavlos aleatorios entre solicitudes  para simular un comportamiento mas humano
import flet #Framework para interfaz grafica
from flet import Page, Column, Text, Dropdown, dropdown, TextField, ElevatedButton, Image, Container #se importan componentes especificos de flet
import requests #Para hacer solicitudes HTTP
from bs4 import BeautifulSoup #para parsear el HTML del sitio y extraer datos
import pandas as pd #Para crear datos estructurados 
import os #para interactuar con el sistema operativo

def calcular_inactividad(texto_tiempo):
    if "años" in texto_tiempo or "año" in texto_tiempo or "Nunca" in texto_tiempo:
        return True
    elif "días" in texto_tiempo or "día" in texto_tiempo:
        dias = int(texto_tiempo.split()[0])
        return dias > 56
    return False

def verificar_actividad_curso(session, id_curso):
    url_participantes = f"https://pregrado.ustabuca.edu.co/user/index.php?id={id_curso}"
    response = session.get(url_participantes)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        participantes_count = soup.find('p', {'data-region': 'participant-count'})
        numero_participantes = participantes_count.get_text(strip=True) if participantes_count else "Desconocido"

        participantes = soup.find_all('tr')
        total_estudiantes = 0
        estudiantes_inactivos = 0

        for participante in participantes:
            rol = participante.find('td', class_='cell c3')
            if rol and "Estudiante" in rol.get_text(strip=True):
                total_estudiantes += 1
                actividad = participante.find('td', class_='cell c5')
                if actividad:
                    tiempo_inactividad = actividad.get_text(strip=True)
                    if calcular_inactividad(tiempo_inactividad):
                        estudiantes_inactivos += 1

        if total_estudiantes == 0:
            return numero_participantes, "Inactivo"
        elif estudiantes_inactivos / total_estudiantes > 0.5:
            return numero_participantes, "Inactivo"
        else:
            return numero_participantes, "Activo"
    else:
        return "Desconocido", "Desconocido"

def iniciar_sesion_moodle():
    login_url = "https://pregrado.ustabuca.edu.co/login/index.php"
    username = "extraccion_ustacv"
    password = "000"

    try:
        session = requests.session()
        response = session.get(login_url)
        soup = BeautifulSoup(response.text, 'html.parser')
        token = soup.find("input", {"name": "logintoken"})["value"]

        login_data = {
            "username": username,
            "password": password,
            "logintoken": token
        }

        login_response = session.post(login_url, data=login_data)
        if login_response.status_code == 200 and "loginerrormessage" not in login_response.text:
            return session
        else:
            return None
    except:
        return None

def obtener_pagina_soup(session, url):
    response = session.get(url)
    if response.status_code == 200:
        return BeautifulSoup(response.text, 'html.parser')
    else:
        return None

def contar_usuarios_curso(session, course_id, numero_rango):
    contador_estudiantes = 0
    contador_profesores = 0
    nombres_docentes = []

    for page in range(numero_rango):
        url = f"https://pregrado.ustabuca.edu.co/user/index.php?id={course_id}&page={page}"
        response = session.get(url)
        if response.status_code != 200:
            break

        soup = BeautifulSoup(response.text, 'html.parser')
        span_elementos = soup.find_all('span', class_='inplaceeditable')

        if not span_elementos:
            break

        for span_elemento in span_elementos:
            a_element = span_elemento.find('a')
            if a_element:
                text = a_element.text.strip()
                title = a_element.get('title', '').strip()

                if "Profesor" in text or "Teacher" in text or "Non-editing teacher" in text:
                    nombre_usuario = title.replace("Tareas del rol", "").replace("Tareas De Rol", "").strip()
                    nombre_usuario = ' '.join(nombre_usuario.split())  # Elimina espacios extra
                    if nombre_usuario:
                        nombres_docentes.append(nombre_usuario)
                    contador_profesores += 1
                elif "Estudiante" in text or "Student" in text:
                    contador_estudiantes += 1

        time.sleep(random.uniform(0.5, 1.5))

    return contador_estudiantes, contador_profesores, nombres_docentes

def obtener_cursos_pagina(session, soup, division_nombre, subcategorias, numero_rango):
    cursos = []
    cursos_principales = soup.find_all('div', class_='card dashboard-card')
    for curso in cursos_principales:
        nombre_curso = curso.find('a', class_='aalink').get_text(strip=True)
        url_curso = curso.find('a', class_='aalink')['href']
        id_curso = url_curso.split('id=')[1]

        contador_estudiantes, contador_profesores, nombres_docentes = contar_usuarios_curso(session, id_curso, numero_rango)
        total_usuarios = contador_estudiantes + contador_profesores
        numero_participantes, estado_curso = verificar_actividad_curso(session, id_curso)

        subcategorias_dict = {
            f"Subcategoría {i+1}": subcategorias[i] if i < len(subcategorias) else ""
            for i in range(4)  # Máximo de 4 niveles de subcategorías, ajustable si es necesario
        }

        curso_data = {
            "División": division_nombre,
            "Nombre del curso": nombre_curso,
            "URL": url_curso,
            "Nombres de Docentes": ', '.join(nombres_docentes).title(),
            "Cantidad de Estudiantes": contador_estudiantes,
            "Cantidad de Profesores": contador_profesores,
            "Cantidad Total de Usuarios": total_usuarios,
            "Estado del Curso": estado_curso,
        }
        
        # Combinar con las subcategorías dinámicas
        curso_data.update(subcategorias_dict)
        cursos.append(curso_data)
    return cursos

def obtener_todos_los_cursos(session, id_categoria, division_nombre, nivel=0, subcategorias=[], numero_rango=50):
    data = []
    url = f"https://pregrado.ustabuca.edu.co/course/index.php?categoryid={id_categoria}"
    soup = obtener_pagina_soup(session, url)
    if not soup:
        return data

    cursos = obtener_cursos_pagina(session, soup, division_nombre, subcategorias, numero_rango)
    data.extend(cursos)

    subcategorias_html = soup.find_all('div', class_='category')
    for sub in subcategorias_html:
        nombre_sub = sub.find('h3').get_text(strip=True)
        enlace_sub = sub.find('a')
        sub_id = enlace_sub['href'].split('categoryid=')[1] if enlace_sub else None
        if sub_id:
            time.sleep(random.uniform(0.5, 1.5))
            sub_data = obtener_todos_los_cursos(
                session,
                sub_id,
                division_nombre,
                nivel + 1,
                subcategorias=subcategorias + [nombre_sub],
                numero_rango=numero_rango
            )
            data.extend(sub_data)

    return data

def guardar_a_excel(data, nombre_archivo="informe_moodle.xlsx"):
    columnas_deseadas = [
        "División",
        "Nombre del curso",
        "Subcategoría 1",
        "Subcategoría 2",
        "Subcategoría 3",
        "Subcategoría 4",
        "URL",
        "Nombres de Docentes",
        "Cantidad de Estudiantes",
        "Cantidad de Profesores",
        "Cantidad Total de Usuarios",
        "Estado del Curso"
    ]

    df = pd.DataFrame(data, columns=columnas_deseadas)
    df.to_excel(nombre_archivo, index=False)

def main(page: Page):
    
    page.fonts = { "Palette":"Palette_Bold.ttf"
        
    }

    page.title = "PREGRADO"
    page.horizontal_alignment = "center"
    page.vertical_alignment = "center"
    
    
    page.window.width = 500
    page.window.height = 400
    
    icon_path = "icono.ico"
    if not os.path.exists(icon_path):
        print(f"Advertencia: No se encontró el archivo de ícono en la ruta '{icon_path}'.")
    else:
        page.window.icon = os.path.abspath(icon_path)

    categorias = {
        4: "CILCE",
        3: "Humanidades",
        15: "División Ciencias Económicas, Administrativas y Contables",
        20: "Ciencias Básicas",
        27: "Campus Virtual",
        29: "División de Ingenierías y Arquitectura",
        31: "División Ciencias de la Salud",
        34: "División de Ciencias Jurídicas y Políticas"
    }

    status_text = Text(value="", size=14)
    drop_categoria = Dropdown(
        label="Selecciona la Categoría",
        options=[dropdown.Option(str(k), text=v) for k, v in categorias.items()],
        width=300
    )
    input_rango = TextField(label="Número de páginas a escanear por curso (ej: 50)", width=300)
    btn_iniciar = ElevatedButton(text="Iniciar Extracción",  bgcolor="#00dba7", color="#FFFFFF", on_click=lambda e: iniciar_extraccion(e))

    def iniciar_extraccion(e):
        if not drop_categoria.value:
            status_text.value = "Por favor selecciona una categoría."
            page.update()
            return
        if not input_rango.value.isdigit():
            status_text.value = "Por favor ingresa un número válido para el rango."
            page.update()
            return

        status_text.value = "Iniciando sesión en Pregrado..."
        page.update()
        session = iniciar_sesion_moodle()
        if not session:
            status_text.value = "No se pudo iniciar sesión. Revisa las credenciales."
            page.update()
            return

        id_categoria_usuario = int(drop_categoria.value)
        numero_rango = int(input_rango.value)

        if id_categoria_usuario not in categorias:
            status_text.value = "Categoría no válida."
            page.update()
            return

        division_nombre = categorias[id_categoria_usuario]
        status_text.value = f"Extrayendo información de: {division_nombre}..."
        page.update()

        data = obtener_todos_los_cursos(session, id_categoria_usuario, division_nombre, numero_rango=numero_rango)

        if data:
            # Guardar con el nombre de la categoría
            nombre_archivo = f"informe_{division_nombre.replace(' ', '_')}.xlsx"
            guardar_a_excel(data, nombre_archivo=nombre_archivo)
            status_text.value = f"Proceso completo. Se han guardado los datos en '{nombre_archivo}'."
        else:
          status_text.value = "No se encontraron cursos o no se pudo completar la extracción."
        page.update()

    page.add(
        Container(
            
            alignment=flet.alignment.center,

            # Ajusta el padding según tu preferencia (ej: padding=0)
            content= Column(
                horizontal_alignment= "center",
                alignment= "center",
                spacing=2.5, # Ajusta el espacio entre los controles, 0 = sin espacio
                controls=[
                    Container(height=10, bgcolor="#00dba7", width=600),
                    Image(src="logo.png", width=400, height=100),
                    Text("GENERADOR DE INFORMES PREGRADO", size=20, weight="bold", font_family="Palette"),
                    drop_categoria,
                    input_rango,
                    btn_iniciar,
                    status_text
                ]
            )
        )
    )

flet.app(target=main)


