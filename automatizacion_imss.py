import os
import glob
import unicodedata
import pandas as pd
from playwright.sync_api import sync_playwright
import time

def normalizar_texto(texto):
    """Limpia espacios, convierte a minúsculas y elimina acentos/diacríticos."""
    if pd.isna(texto) or not isinstance(texto, str):
        return ""
    # Convertir a minúsculas y quitar espacios
    t = texto.strip().lower()
    # Eliminar acentos
    t = "".join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')
    return t

# Mapeo de columna de rol (hoja administrativa) -> (valor_rol, nombre_rol) en MoCE
MAPEO_ROL_COLUMNAS = {
    'asistente medica': (1100, 'Asistente'),
    'asistente medico': (1100, 'Asistente'),
    'administrador de catalogo': (1500, 'Médico'),
    'director': (1500, 'Médico'),
    'jefes de servicios': (1500, 'Médico'),
    'jefe de servicio': (1500, 'Médico'),
    'agenda de citas': (1100, 'Asistente'),
    'agenda citas': (1100, 'Asistente'),
    'enfermera': (1400, 'Enfermera'),
    'enfermero': (1400, 'Enfermera'),
}

def es_hoja_administrativa(hoja):
    """Determina si la hoja corresponde a personal administrativo (u operativo)."""
    h = normalizar_texto(hoja)
    if 'operativ' in h:
        return False
    # 'admin' cubre variantes: 'administrativo', 'adminsitrativo' (typo), 'admin'
    if 'admin' in h:
        return True
    return False

def obtener_columnas_rol(df):
    """Identifica las columnas que marcan el rol en la hoja de personal administrativo."""
    palabras_clave = ['asistente', 'administrador', 'catalogo', 'director', 'jefe', 'agenda', 'enfermera', 'enfermero']
    return [c for c in df.columns.tolist() if any(p in normalizar_texto(str(c)) for p in palabras_clave)]

def detectar_rol_administrativo(row, columnas_rol):
    """Detecta el rol de un usuario administrativo según la columna marcada con 'X'.
    Devuelve (valor_rol, nombre_rol) o None si ningún rol viene marcado."""
    for col in columnas_rol:
        valor = row.get(col)
        if pd.notna(valor) and normalizar_texto(str(valor)) in ('x', 'si', 'true', '1', 's', 'y', 'yes'):
            clave = normalizar_texto(str(col))
            for patron, rol in MAPEO_ROL_COLUMNAS.items():
                if patron in clave:
                    return rol
    return None

def rol_para_usuario(row, es_administrativo, columnas_rol=None):
    """Devuelve (valor_rol, nombre_rol) según el tipo de hoja.
    - Operativo: siempre Médico (1500).
    - Administrativo: detecta la columna marcada con 'X'.
    Devuelve (None, None) si es administrativo y no hay rol marcado."""
    if es_administrativo:
        return detectar_rol_administrativo(row, columnas_rol or [])
    return (1500, 'Médico')

def esperar_alerta(page, texto, timeout=8):
    """Espera a que aparezca una alerta con el texto indicado en el contenedor <app-alert>.
    Devuelve True si la alerta se mostró, False si no apareció dentro del tiempo límite."""
    try:
        page.locator("app-alert .alert-container .alert").filter(has_text=texto).wait_for(
            state="visible", timeout=timeout * 1000
        )
        return True
    except Exception:
        return False

# CONFIGURACIÓN POR DEFECTO
URL_SISTEMA = 'https://ci-moceopd.imss.gob.mx/login'

def buscar_archivo_excel():
    """Busca el archivo de Excel en el directorio actual."""
    # Buscar todos los archivos .xlsx en la carpeta actual
    archivos = glob.glob("*.xlsx")
    # Filtrar archivos temporales de Excel y el archivo de salida
    archivos = [f for f in archivos if not f.startswith("~$") and f != "RESULTADO_CARGA_MOCE.xlsx"]
    
    if len(archivos) == 0:
        raise FileNotFoundError("No se encontró ningún archivo de Excel (.xlsx) en esta carpeta.")
    
    if len(archivos) == 1:
        print(f"\n[+] Archivo Excel detectado: {archivos[0]}")
        return archivos[0]
    
    print("\n--- SELECCIÓN DE EXCEL ---")
    print("Se encontraron múltiples archivos de Excel:")
    for idx, f in enumerate(archivos):
        print(f"  [{idx + 1}] {f}")
    while True:
        try:
            opcion = int(input("Selecciona el número del archivo que deseas usar: "))
            if 1 <= opcion <= len(archivos):
                return archivos[opcion - 1]
        except ValueError:
            pass
        print("[!] Opción inválida. Intenta de nuevo.")

def seleccionar_hoja(excel_file):
    """Permite seleccionar la hoja del Excel a procesar."""
    xl = pd.ExcelFile(excel_file)
    hojas = xl.sheet_names
    
    if len(hojas) == 1:
        print(f"[+] Procesando la única hoja: '{hojas[0]}'")
        return hojas[0]
        
    print("\n--- SELECCIÓN DE HOJA ---")
    print("El archivo contiene las siguientes hojas:")
    for idx, hoja in enumerate(hojas):
        print(f"  [{idx + 1}] {hoja}")
    while True:
        try:
            opcion = int(input("Selecciona el número de la hoja que deseas procesar: "))
            if 1 <= opcion <= len(hojas):
                return hojas[opcion - 1]
        except ValueError:
            pass
        print("[!] Opción inválida. Intenta de nuevo.")

def obtener_columnas(df):
    """Detecta y confirma las columnas de CURP, Nombre y Matrícula."""
    columnas = df.columns.tolist()
    print("\n--- DETECCIÓN DE COLUMNAS ---")
    
    # 1. Detectar CURP
    col_curp = None
    # Prioridad 1: Coincidencia exacta con CURP o Clave Empleado
    exactas_curp = [c for c in columnas if str(c).strip().lower() in ['curp', 'clave empleado', 'clave_empleado', 'clave de empleado', 'claveunica']]
    if exactas_curp:
        default_curp = exactas_curp[0]
    else:
        # Prioridad 2: Que contenga la palabra 'curp'
        contiene_curp = [c for c in columnas if 'curp' in str(c).lower()]
        if contiene_curp:
            default_curp = contiene_curp[0]
        else:
            # Prioridad 3: Que contenga 'clave' o 'empleado', excluyendo cve presupuestal
            contiene_clave = [c for c in columnas if any(x in str(c).lower() for x in ['clave', 'empleado']) and not any(x in str(c).lower() for x in ['presupuestal', 'cve'])]
            if contiene_clave:
                default_curp = contiene_clave[0]
            else:
                # Prioridad 4: Cualquier coincidencia parcial
                posibles_curp = [c for c in columnas if any(x in str(c).lower() for x in ['curp', 'clave', 'empleado'])]
                default_curp = posibles_curp[0] if posibles_curp else (columnas[0] if columnas else '')

    if default_curp:
        respuesta = input(f"¿Usar columna '{default_curp}' para las CURPs? (Presiona ENTER para SÍ, o escribe el nombre correcto): ").strip()
        col_curp = respuesta if respuesta else default_curp
    else:
        while col_curp not in columnas:
            col_curp = input("[!] No se detectó la columna CURP. Escribe el nombre exacto de la columna CURP: ").strip()

    # 2. Detectar Nombre
    col_nombre = None
    # Intentar coincidencia exacta o palabras clave directas
    exactas_nombre = [c for c in columnas if str(c).strip().lower() in ['nombre', 'nombre completo', 'nombres', 'nombre(s)', 'personal', 'médico', 'medico']]
    if exactas_nombre:
        default_nombre = exactas_nombre[0]
    else:
        # Coincidencias parciales excluyendo nombres de unidades/consultorios/etc.
        posibles_nombre = [c for c in columnas if any(x in str(c).lower() for x in ['nombre', 'medico', 'personal', 'completo'])]
        posibles_nombre = [c for c in posibles_nombre if not any(x in str(c).lower() for x in ['unidad', 'consultorio', 'clues', 'hospital', 'clue', 'presupuestal'])]
        default_nombre = posibles_nombre[0] if posibles_nombre else (columnas[0] if columnas else '')

    if default_nombre:
        respuesta = input(f"¿Usar columna '{default_nombre}' para los Nombres? (Presiona ENTER para SÍ, o escribe el nombre correcto): ").strip()
        col_nombre = respuesta if respuesta else default_nombre
    else:
        while col_nombre not in columnas:
            col_nombre = input("[!] No se detectó la columna de Nombre. Escribe el nombre exacto de la columna del Nombre: ").strip()

    # 3. Detectar columna de Matrícula existente o crear una
    col_matricula = None
    posibles_mat = [c for c in columnas if 'matri' in str(c).lower()]
    if posibles_mat:
        default_mat = posibles_mat[0]
        respuesta = input(f"¿Guardar la matrícula en la columna existente '{default_mat}'? (Presiona ENTER para SÍ, o escribe el nombre correcto): ").strip()
        col_matricula = respuesta if respuesta else default_mat
    else:
        respuesta = input("No se encontró columna de Matrícula. ¿Deseas crear la columna 'Matricula_Generada'? (ENTER para SÍ, o escribe un nombre personalizado): ").strip()
        col_matricula = respuesta if respuesta else 'Matricula_Generada'

    return col_curp, col_nombre, col_matricula

def ejecutar_automatizacion():
    print("=" * 60)
    print("        ASISTENTE DE AUTOMATIZACIÓN - IMSS MOCE")
    print("=" * 60)
    
    try:
        excel_file = buscar_archivo_excel()
        hoja = seleccionar_hoja(excel_file)
        
        # Leer el Excel
        print(f"\n[+] Cargando datos de '{excel_file}' en la hoja '{hoja}'...")
        df = pd.read_excel(excel_file, sheet_name=hoja)
        
        # Determinar si es la hoja de personal administrativo (cambia el rol por usuario)
        columnas_rol = obtener_columnas_rol(df)
        es_administrativo = es_hoja_administrativa(hoja) or bool(columnas_rol)
        if es_administrativo:
            print(f"[+] Hoja detectada como PERSONAL ADMINISTRATIVO. Columnas de rol: {columnas_rol}")
        else:
            print("[+] Hoja detectada como PERSONAL OPERATIVO. Todos los usuarios se darán de alta como Médico.")
        
        # Obtener columnas a usar
        col_curp, col_nombre, col_matricula = obtener_columnas(df)
        
        # Asegurarnos de que la columna de matrícula exista y sea tipo objeto (texto) para evitar errores de tipo en pandas
        if col_matricula not in df.columns:
            df[col_matricula] = ""
        df[col_matricula] = df[col_matricula].astype(object)
            
    except Exception as e:
        print(f"\n[!] Error inicial: {e}")
        return

    # Filtrar filas que tengan CURP
    df_medicos = df[df[col_curp].notna()].copy()
    
    # Filtrar filas que ya tengan una matrícula registrada para no procesar duplicados
    if col_matricula in df.columns:
        # Se consideran vacías las celdas nulas, vacías, o que tengan textos de error previo
        filtro_vacias = df_medicos[col_matricula].isna() | (df_medicos[col_matricula].astype(str).str.strip() == "") | (df_medicos[col_matricula].astype(str).str.contains("ERROR|NO GENERADA", case=False, na=True))
        df_medicos_a_procesar = df_medicos[filtro_vacias].copy()
        cant_omitidos = len(df_medicos) - len(df_medicos_a_procesar)
        if cant_omitidos > 0:
            print(f"[i] Se omitirán {cant_omitidos} registros que ya tienen una matrícula válida en la columna '{col_matricula}'.")
        df_medicos = df_medicos_a_procesar

    print(f"\n>>> Se encontraron {len(df_medicos)} registros pendientes por procesar.")
    if len(df_medicos) == 0:
        print("[!] No hay registros nuevos para procesar en esta hoja.")
        return

    with sync_playwright() as p:
        # Intentar abrir Edge (o Chromium si Edge no está instalado)
        print("\n>>> Iniciando navegador...")
        try:
            browser = p.chromium.launch(headless=False, channel="msedge")
        except Exception as e:
            print(f"[i] No se pudo abrir Edge mediante canal 'msedge' ({e}). Iniciando Chromium normal...")
            browser = p.chromium.launch(headless=False)
            
        context = browser.new_context()
        page = context.new_page()

        print(f">>> Navegando a {URL_SISTEMA}...")
        page.goto(URL_SISTEMA)

        print("\n" + "!" * 50)
        print("--- ACCIÓN MANUAL REQUERIDA ---")
        print("1. Inicia sesión en la ventana de Edge que se abrió.")
        print("2. Navega hasta estar dentro de la pestaña 'Alta de Usuarios'.")
        print("3. Cuando ya veas los campos de CURP y Matrícula listos:")
        print("   Regresa a esta ventana negra y presiona ENTER.")
        print("!" * 50 + "\n")
        
        input("Presiona ENTER aquí cuando ya estés en la pestaña de Alta de Usuarios...")

        exitosos = 0
        errores = 0

        for index, row in df_medicos.iterrows():
            curp_raw = str(row[col_curp]).strip()
            # Limpiar decimales flotantes como '963014927.0' -> '963014927'
            if curp_raw.endswith('.0'):
                curp = curp_raw[:-2]
            else:
                curp = curp_raw
                
            nombre = str(row[col_nombre]).strip()
            print(f"\n[*] [{exitosos+errores+1}/{len(df_medicos)}] Procesando: {nombre} | CURP: {curp}")

            # Determinar rol según el tipo de hoja
            rol_det = rol_para_usuario(row, es_administrativo, columnas_rol)
            if rol_det is None:
                print("    [!] No se detectó un rol marcado (X) para este usuario. Se omitirá.")
                df.at[index, col_matricula] = "ERROR: Sin rol marcado"
                errores += 1
                continue
            rol_value, rol_nombre = rol_det
            print(f"    [i] Rol asignado: {rol_nombre} ({rol_value})")

            try:
                # 1. Ingresar CURP de manera robusta
                campo_curp = None
                selectores_curp = [
                    "input[formcontrolname='curp']",
                    "input:near(label:text('CURP'))",
                    "input[placeholder*='CURP' i]",
                    "input[id*='curp' i]",
                    "input[name*='curp' i]",
                    "input[type='text']"
                ]
                
                # Buscamos el campo CURP de forma rápida (máximo 2s por intento)
                for selector in selectores_curp:
                    try:
                        locator = page.locator(selector).first
                        if locator.is_visible(timeout=2000):
                            campo_curp = locator
                            break
                    except Exception:
                        continue
                
                if campo_curp is None:
                    # Caída por defecto si no se detectó por visibilidad rápida
                    campo_curp = page.locator("input[formcontrolname='curp']").first
                
                # Llenar la CURP (Playwright limpia automáticamente el campo antes de escribir con .fill)
                campo_curp.click()
                campo_curp.press("Control+A")
                campo_curp.press("Backspace")
                campo_curp.fill(curp)
                
                # 2. Clic en Búsqueda de CURP de manera robusta
                btn_buscar = None
                selectores_buscar = [
                    "button:has-text('Búsqueda de CURP')",
                    "button:has(.bi-search)",
                    "button.btn-success",
                    "button:has-text('Buscar')"
                ]
                for selector in selectores_buscar:
                    try:
                        locator = page.locator(selector).first
                        if locator.is_visible(timeout=2000):
                            btn_buscar = locator
                            break
                    except Exception:
                        continue
                
                if btn_buscar is None:
                    btn_buscar = page.locator("button:has-text('Búsqueda de CURP')").first
                
                btn_buscar.click()
                
                # Esperar a que el sistema procese y genere la matrícula
                time.sleep(2.5) 
                
                # 3. Capturar Matrícula de manera robusta
                campo_matricula = None
                selectores_matricula = [
                    "input[formcontrolname='matricula']",
                    "input:near(label:text('Matrícula'))",
                    "input:near(label:text('Matricula'))",
                    "input[placeholder*='Matrícula' i]"
                ]
                
                for selector in selectores_matricula:
                    try:
                        locator = page.locator(selector).first
                        if locator.is_visible(timeout=2000):
                            campo_matricula = locator
                            break
                    except Exception:
                        continue
                
                if campo_matricula is None:
                    campo_matricula = page.locator("input[formcontrolname='matricula']").first
                
                matricula = campo_matricula.input_value()
                
                if matricula and matricula.strip():
                    print(f"    [+] Matrícula obtenida con éxito: {matricula}")
                    df.at[index, col_matricula] = matricula
                    exitosos += 1
                    
                    # 4. Seleccionar Rol
                    try:
                        # Probamos con el selector Angular formcontrolname
                        try:
                            page.select_option("select[formcontrolname='rol']", value=str(rol_value))
                        except Exception:
                            page.select_option("select:has-text('Seleccionar rol')", label=rol_nombre)
                        print(f"    [+] Rol seleccionado: {rol_nombre} ({rol_value})")
                    except Exception as e_rol:
                        print(f"    [i] Advertencia al seleccionar rol: {e_rol}")

                    # 5. Clic en Agregar (Solo si la matrícula se obtuvo con éxito)
                    btn_agregar = None
                    selectores_agregar = [
                        "button:has-text('Agregar')",
                        "button:has(.bi-plus-circle)",
                        "button.btn-secondary:has-text('Agregar')"
                    ]
                    for selector in selectores_agregar:
                        try:
                            locator = page.locator(selector).first
                            if locator.is_visible(timeout=2000):
                                btn_agregar = locator
                                break
                        except Exception:
                            continue
                    
                    if btn_agregar is None:
                        btn_agregar = page.locator("button:has-text('Agregar')").first
                    
                    # Verificar si el botón Agregar está deshabilitado
                    try:
                        if btn_agregar.is_disabled(timeout=1500):
                            print("    [!] El botón 'Agregar' está deshabilitado en el sistema. Omitiendo clic.")
                            df.at[index, col_matricula] = "ERROR: Botón Agregar deshabilitado"
                            errores += 1
                        else:
                            btn_agregar.click(timeout=2000)
                            print("    [+] Clic en Agregar ejecutado.")
                            
                            # 5.0 Validar mensaje de confirmación "Usuario agregado correctamente"
                            agregado_ok = esperar_alerta(page, "Usuario agregado correctamente")
                            if agregado_ok:
                                print("    [+] Confirmación detectada: 'Usuario agregado correctamente'.")
                            else:
                                print("    [!] NO se detectó la confirmación 'Usuario agregado correctamente'.")
                            time.sleep(1.5)
                            
                            # 5.1 Clic en Continuar (Confirmar la adición final)
                            btn_continuar = None
                            selectores_continuar = [
                                "button:has-text('Continuar')",
                                "button:has(.bi-arrow-right-circle)",
                                "button.btn-primary:has-text('Continuar')"
                            ]
                            for selector in selectores_continuar:
                                try:
                                    locator = page.locator(selector).first
                                    if locator.is_visible(timeout=2000):
                                        btn_continuar = locator
                                        break
                                except Exception:
                                    continue
                            
                            if btn_continuar is None:
                                btn_continuar = page.locator("button:has-text('Continuar')").first
                                
                            try:
                                btn_continuar.click(timeout=2000)
                                print("    [+] Clic en Continuar ejecutado.")
                                
                                # 5.1.1 Validar el popup "Personal operativo guardado exitosamente"
                                guardado_ok = esperar_alerta(page, "guardado exitosamente")
                                if guardado_ok:
                                    print("    [+] Confirmación detectada: 'Personal operativo guardado exitosamente'.")
                                else:
                                    print("    [!] NO se detectó la confirmación 'Personal operativo guardado exitosamente'.")
                                time.sleep(2)
                            except Exception as e_cont:
                                print(f"    [!] Advertencia al hacer clic en Continuar: {e_cont}")
                                time.sleep(1)
                                
                            # 5.2 Regresar a la pestaña de Alta de Usuarios para el siguiente médico
                            print("    [i] Regresando a la pestaña 'Alta de Usuarios'...")
                            regreso_exitoso = False
                            selectores_regreso = [
                                "button[ngbnavlink]:has-text('Alta de Usuarios')",
                                "button:has-text('Alta de Usuarios')",
                                "#ngb-nav-1",
                                "a:has-text('Alta de Usuarios')",
                                "span:has-text('Alta de Usuarios')",
                                "text=Alta de Usuarios"
                            ]
                            for selector in selectores_regreso:
                                try:
                                    locator = page.locator(selector).first
                                    if locator.is_visible(timeout=2000):
                                        locator.click()
                                        regreso_exitoso = True
                                        break
                                except Exception:
                                    continue
                            
                            if not regreso_exitoso:
                                print("    [!] No se pudo regresar automáticamente a 'Alta de Usuarios'.")
                                print("    Por favor, haz clic manualmente en la pestaña 'Alta de Usuarios' en el navegador...")
                                try:
                                    page.locator("input[formcontrolname='curp']").wait_for(state="visible", timeout=8000)
                                    print("    [+] Detectada pantalla de Alta de Usuarios de nuevo.")
                                except Exception:
                                    input("    Presiona ENTER en esta consola cuando ya estés de vuelta en la pestaña 'Alta de Usuarios'...")
                            else:
                                try:
                                    page.locator("input[formcontrolname='curp']").wait_for(state="visible", timeout=5000)
                                    time.sleep(1.5)
                                except Exception:
                                    time.sleep(2)

                            # 5.3 Eliminar el registro anterior (bote de basura) para poder capturar el siguiente usuario
                            try:
                                btn_trash = None
                                selectores_trash = [
                                    "button.btn-danger:has(.bi-trash)",
                                    "button:has(.bi-trash)",
                                    "i.bi-trash"
                                ]
                                for selector in selectores_trash:
                                    try:
                                        locator = page.locator(selector).first
                                        if locator.is_visible(timeout=1500):
                                            btn_trash = locator
                                            break
                                    except Exception:
                                        continue
                                if btn_trash is not None:
                                    btn_trash.click(timeout=2000)
                                    print("    [+] Registro anterior eliminado (clic en bote de basura).")
                                    time.sleep(1.5)
                                else:
                                    print("    [i] No hay registro anterior para eliminar en la pestaña.")
                            except Exception as e_trash:
                                print(f"    [i] Advertencia al eliminar el registro anterior: {e_trash}")
                    except Exception as e_btn:
                        print(f"    [!] Error al interactuar con el botón Agregar: {e_btn}")
                        df.at[index, col_matricula] = "ERROR: Falla botón Agregar"
                        errores += 1
                else:
                    print("    [!] El sistema no mostró matrícula para esta CURP.")
                    df.at[index, col_matricula] = "NO GENERADA"
                    errores += 1

            except Exception as e:
                print(f"    [!] Error al procesar a este médico: {e}")
                df.at[index, col_matricula] = f"ERROR: {str(e)[:50]}"
                errores += 1
                continue

        # 6. Guardar resultados
        output_file = 'RESULTADO_CARGA_MOCE.xlsx'
        try:
            df.to_excel(output_file, index=False, engine='openpyxl')

            # Marcar en rojo las CURPs con matrícula no generada o con error
            from openpyxl import load_workbook
            from openpyxl.styles import PatternFill
            wb = load_workbook(output_file)
            ws = wb.active
            encabezados = {cell.value: cell.column for cell in ws[1]}
            red_fill = PatternFill(start_color="00FF0000", end_color="00FF0000", fill_type="solid")
            if col_matricula in encabezados:
                col_mat_idx = encabezados[col_matricula]
                col_curp_idx = encabezados.get(col_curp)
                for fila in range(2, ws.max_row + 1):
                    valor = ws.cell(row=fila, column=col_mat_idx).value
                    es_error = (valor is None or not str(valor).strip()
                                or "ERROR" in str(valor).upper()
                                or "NO GENERADA" in str(valor).upper())
                    if es_error:
                        ws.cell(row=fila, column=col_mat_idx).fill = red_fill
                        if col_curp_idx:
                            ws.cell(row=fila, column=col_curp_idx).fill = red_fill
            wb.save(output_file)
            print("\n" + "=" * 60)
            print(f">>> PROCESO FINALIZADO <<<")
            print(f" - Registros procesados con éxito: {exitosos}")
            print(f" - Registros con error o vacíos: {errores}")
            print(f" - Archivo resultante guardado en: {output_file}")
            print("=" * 60)
        except Exception as e:
            print(f"\n[!] Error al escribir el archivo final {output_file}: {e}")
        
        input("\nPresiona ENTER aquí para cerrar el navegador y salir...")
        browser.close()

if __name__ == "__main__":
    ejecutar_automatizacion()
