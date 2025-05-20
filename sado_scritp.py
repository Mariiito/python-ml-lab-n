from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import pandas as pd
import glob
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()

USUARIO = os.getenv('NISSAN_USUARIO')
PASSWORD = os.getenv('NISSAN_PASSWORD')
URL = os.getenv('NISSAN_URL', 'https://10.90.8.11/RcnVentas/Security/Login')
OPCION_DROPDOWN = os.getenv('NISSAN_DROPDOWN_OPTION', '304 - SUCURSAL IZÚCAR DE MATAMOROS, SADO DE ORIENTE ATLIXCO')

# Configuración de período total
FECHA_INICIO_TOTAL = datetime(2021, 1, 1)
FECHA_FIN_TOTAL = datetime(2025, 5, 31)

def crear_carpeta_descargas():
    """Crea la carpeta 'gruposado' si no existe y retorna la ruta absoluta"""
    directorio_script = os.path.dirname(os.path.abspath(__file__))
    carpeta_descargas = os.path.join(directorio_script, "gruposado")
    
    if not os.path.exists(carpeta_descargas):
        os.makedirs(carpeta_descargas)
        print(f"Carpeta creada: {carpeta_descargas}")
    else:
        print(f"Carpeta ya existe: {carpeta_descargas}")
    
    return carpeta_descargas

def generar_periodos_trimestres(fecha_inicio, fecha_fin):
    """Genera periodos de 90 días desde fecha_inicio hasta fecha_fin"""
    periodos = []
    fecha_actual = fecha_inicio
    
    while fecha_actual <= fecha_fin:
        # Calcular la fecha final del período (90 días después)
        fecha_fin_periodo = fecha_actual + timedelta(days=89)
        
        # Asegurarse de no exceder la fecha final total
        if fecha_fin_periodo > fecha_fin:
            fecha_fin_periodo = fecha_fin
        
        periodos.append({
            'inicio': fecha_actual,
            'fin': fecha_fin_periodo,
            'inicio_str': fecha_actual.strftime("%d/%m/%Y"),
            'fin_str': fecha_fin_periodo.strftime("%d/%m/%Y"),
            'nombre': f"{fecha_actual.strftime('%Y%m%d')}_{fecha_fin_periodo.strftime('%Y%m%d')}"
        })
        
        # Avanzar al siguiente período (día siguiente al final del período actual)
        fecha_actual = fecha_fin_periodo + timedelta(days=1)
    
    return periodos

def configurar_driver():
    """Configura y retorna el driver de Chrome con carpeta de descarga personalizada"""
    carpeta_descargas = crear_carpeta_descargas()
    
    chrome_options = Options()
    
    # Configuraciones de seguridad
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--ignore-ssl-errors")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Configurar preferencias de descarga
    prefs = {
        "download.default_directory": carpeta_descargas,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    print(f"Driver configurado. Las descargas se guardarán en: {carpeta_descargas}")
    return driver

def esperar_elemento_con_multiples_estrategias(driver, wait, texto_elemento, descripcion="elemento"):
    """Intenta encontrar un elemento usando múltiples estrategias"""
    estrategias = [
        (By.XPATH, f"//a[contains(text(), '{texto_elemento}')]"),
        (By.XPATH, f"//a[normalize-space(text())='{texto_elemento}']"),
        (By.XPATH, f"//li//a[contains(text(), '{texto_elemento}')]"),
        (By.XPATH, f"//span[contains(text(), '{texto_elemento}')]/.."),
        (By.PARTIAL_LINK_TEXT, texto_elemento),
        (By.LINK_TEXT, texto_elemento)
    ]
    
    for i, (by, selector) in enumerate(estrategias):
        try:
            elemento = wait.until(EC.presence_of_element_located((by, selector)))
            print(f"✓ Encontrado {descripcion} con estrategia {i+1}")
            return elemento
        except Exception:
            continue
    
    print(f"✗ No se pudo encontrar {descripcion} con ninguna estrategia")
    return None

def verificar_descarga_completada(carpeta_descargas, timeout=120):
    """Verifica si la descarga se ha completado"""
    print(f"Verificando descarga en: {carpeta_descargas}")
    
    archivos_antes = set(os.listdir(carpeta_descargas)) if os.path.exists(carpeta_descargas) else set()
    
    tiempo_inicio = time.time()
    while time.time() - tiempo_inicio < timeout:
        try:
            archivos_actuales = set(os.listdir(carpeta_descargas)) if os.path.exists(carpeta_descargas) else set()
            archivos_nuevos = archivos_actuales - archivos_antes
            
            if archivos_nuevos:
                archivos_descarga_en_progreso = [f for f in archivos_nuevos if f.endswith('.crdownload')]
                
                if not archivos_descarga_en_progreso:
                    print(f"✓ Descarga completada. Archivo nuevo: {list(archivos_nuevos)[0]}")
                    return list(archivos_nuevos)[0]
                else:
                    print(f"⏳ Descarga en progreso... ({archivos_descarga_en_progreso})")
            
            time.sleep(3)
        except Exception as e:
            print(f"Error verificando descarga: {str(e)}")
            time.sleep(3)
    
    print(f"⚠️ Timeout de {timeout}s alcanzado. No se pudo verificar la descarga.")
    return None

def renombrar_archivo_descargado(carpeta_descargas, nombre_archivo_descargado, nuevo_nombre):
    """Renombra el archivo descargado con un nombre descriptivo"""
    if nombre_archivo_descargado:
        ruta_original = os.path.join(carpeta_descargas, nombre_archivo_descargado)
        extension = os.path.splitext(nombre_archivo_descargado)[1]
        ruta_nueva = os.path.join(carpeta_descargas, f"{nuevo_nombre}{extension}")
        
        try:
            os.rename(ruta_original, ruta_nueva)
            print(f"✓ Archivo renombrado: {nuevo_nombre}{extension}")
            return f"{nuevo_nombre}{extension}"
        except Exception as e:
            print(f"⚠️ Error renombrando archivo: {str(e)}")
            return nombre_archivo_descargado
    return None

def navegar_a_consulta_venta(driver, wait):
    """Navega hasta la sección de Consulta de Venta"""
    try:
        print("\n🧭 Iniciando navegación a Reportes > Ventas > Consulta de Venta...")
        
        # Esperar a que la página esté completamente cargada
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(5)
        
        # Paso 1: Buscar y hacer hover sobre "Reportes"
        print("Paso 1: Buscando menú 'Reportes'...")
        reportes_menu = esperar_elemento_con_multiples_estrategias(driver, wait, "Reportes", "menú Reportes")
        
        if not reportes_menu:
            variaciones_reportes = ["Reporte", "REPORTES", "Reports"]
            for variacion in variaciones_reportes:
                reportes_menu = esperar_elemento_con_multiples_estrategias(driver, wait, variacion, f"menú {variacion}")
                if reportes_menu:
                    break
        
        if not reportes_menu:
            raise Exception("No se pudo encontrar el menú 'Reportes'")
        
        # Hacer hover sobre Reportes
        driver.execute_script("arguments[0].scrollIntoView(true);", reportes_menu)
        time.sleep(1)
        actions = ActionChains(driver)
        actions.move_to_element(reportes_menu).perform()
        print("✓ Hover realizado sobre 'Reportes'")
        time.sleep(3)
        
        # Paso 2: Buscar "Ventas"
        print("Paso 2: Buscando menú 'Ventas'...")
        ventas_menu = esperar_elemento_con_multiples_estrategias(driver, wait, "Ventas", "menú Ventas")
        
        if not ventas_menu:
            variaciones_ventas = ["Venta", "VENTAS", "Sales"]
            for variacion in variaciones_ventas:
                ventas_menu = esperar_elemento_con_multiples_estrategias(driver, wait, variacion, f"menú {variacion}")
                if ventas_menu:
                    break
        
        if not ventas_menu:
            raise Exception("No se pudo encontrar el menú 'Ventas'")
        
        # Hacer hover sobre Ventas
        driver.execute_script("arguments[0].scrollIntoView(true);", ventas_menu)
        time.sleep(1)
        actions.move_to_element(ventas_menu).perform()
        print("✓ Hover realizado sobre 'Ventas'")
        time.sleep(3)
        
        # Paso 3: Buscar y hacer clic en "Consulta de Venta"
        print("Paso 3: Buscando 'Consulta de Venta'...")
        variaciones_consulta = [
            "Consulta de Venta",
            "Consulta de Ventas", 
            "Consulta Venta",
            "Consulta Ventas",
            "CONSULTA DE VENTA"
        ]
        
        consulta_venta = None
        for variacion in variaciones_consulta:
            consulta_venta = esperar_elemento_con_multiples_estrategias(driver, wait, variacion, f"opción {variacion}")
            if consulta_venta:
                break
        
        if not consulta_venta:
            raise Exception("No se pudo encontrar 'Consulta de Venta'")
        
        # Hacer clic en Consulta de Venta
        driver.execute_script("arguments[0].scrollIntoView(true);", consulta_venta)
        time.sleep(1)
        
        try:
            consulta_venta.click()
            print("✓ Clic realizado en 'Consulta de Venta' (método normal)")
        except Exception:
            driver.execute_script("arguments[0].click();", consulta_venta)
            print("✓ Clic realizado en 'Consulta de Venta' (JavaScript)")
        
        time.sleep(5)
        print(f"✓ Navegación completada. URL actual: {driver.current_url}")
        return True
        
    except Exception as e:
        print(f"✗ Error durante la navegación: {str(e)}")
        return False

def configurar_fechas_y_exportar(driver, wait, fecha_inicio_str, fecha_fin_str, nombre_periodo):
    """Configura las fechas específicas y realiza la exportación"""
    try:
        print(f"\n📅 Configurando período: {fecha_inicio_str} - {fecha_fin_str}")
        
        # Paso 1: Configurar fecha inicial
        try:
            fecha_inicial_input = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Fecha Inicial']"))
            )
            fecha_inicial_input.clear()
            fecha_inicial_input.send_keys(fecha_inicio_str)
            print(f"✓ Fecha inicial establecida: {fecha_inicio_str}")
            time.sleep(1)
            
        except Exception:
            try:
                fecha_inicial_input = driver.find_element(By.NAME, "fechaInicial")
                fecha_inicial_input.clear()
                fecha_inicial_input.send_keys(fecha_inicio_str)
                print(f"✓ Fecha inicial establecida (método alternativo): {fecha_inicio_str}")
            except Exception as e:
                print(f"✗ Error configurando fecha inicial: {str(e)}")
                return False
        
        # Paso 2: Configurar fecha final
        try:
            fecha_final_input = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Fecha Final']"))
            )
            fecha_final_input.clear()
            fecha_final_input.send_keys(fecha_fin_str)
            print(f"✓ Fecha final establecida: {fecha_fin_str}")
            time.sleep(1)
            
        except Exception:
            try:
                fecha_final_input = driver.find_element(By.NAME, "fechaFinal")
                fecha_final_input.clear()
                fecha_final_input.send_keys(fecha_fin_str)
                print(f"✓ Fecha final establecida (método alternativo): {fecha_fin_str}")
            except Exception as e:
                print(f"✗ Error configurando fecha final: {str(e)}")
                return False
        
        # Paso 3: Configurar tipo de fecha
        try:
            tipo_fecha_dropdown = wait.until(
                EC.presence_of_element_located((By.XPATH, "//select[contains(@name, 'tipoFecha') or contains(@id, 'tipoFecha') or contains(@class, 'tipoFecha')]"))
            )
            
            select_tipo_fecha = Select(tipo_fecha_dropdown)
            
            opciones_tipo_fecha = [
                "Fecha Captura Venta",
                "Fecha Factura Cliente",
                "Captura Venta"
            ]
            
            opcion_seleccionada = False
            for opcion in opciones_tipo_fecha:
                try:
                    select_tipo_fecha.select_by_visible_text(opcion)
                    print(f"✓ Tipo de fecha seleccionado: {opcion}")
                    opcion_seleccionada = True
                    break
                except Exception:
                    continue
            
            if not opcion_seleccionada:
                try:
                    opciones = select_tipo_fecha.options
                    if len(opciones) > 1:
                        select_tipo_fecha.select_by_index(1)
                        print(f"✓ Tipo de fecha seleccionado por índice: {opciones[1].text}")
                        opcion_seleccionada = True
                except Exception as e:
                    print(f"⚠️ No se pudo seleccionar tipo de fecha: {str(e)}")
            
            time.sleep(2)
            
        except Exception as e:
            print(f"⚠️ Error configurando tipo de fecha: {str(e)}")
        
        # Paso 4: Hacer clic en Exportar
        carpeta_descargas = crear_carpeta_descargas()
        
        try:
            exportar_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Exportar') or contains(@class, 'exportar')]"))
            )
            
            driver.execute_script("arguments[0].scrollIntoView(true);", exportar_button)
            time.sleep(1)
            
            try:
                exportar_button.click()
                print("✓ Clic realizado en 'Exportar'")
            except Exception:
                driver.execute_script("arguments[0].click();", exportar_button)
                print("✓ Clic realizado en 'Exportar' (JavaScript)")
            
            # Verificar descarga y renombrar archivo
            archivo_descargado = verificar_descarga_completada(carpeta_descargas)
            if archivo_descargado:
                archivo_renombrado = renombrar_archivo_descargado(carpeta_descargas, archivo_descargado, nombre_periodo)
                print(f"✓ Exportación completada para período {nombre_periodo}")
                return archivo_renombrado
            else:
                print(f"✗ No se pudo verificar la descarga para período {nombre_periodo}")
                return False
                
        except Exception as e:
            print(f"✗ Error haciendo clic en Exportar: {str(e)}")
            return False
            
    except Exception as e:
        print(f"✗ Error general en configuración de fechas: {str(e)}")
        return False

def combinar_archivos_excel(carpeta_descargas):
    """Combina todos los archivos Excel descargados en uno solo"""
    try:
        print("\n📊 Iniciando combinación de archivos Excel...")
        
        # Buscar todos los archivos Excel en la carpeta
        patron_excel = os.path.join(carpeta_descargas, "*.xlsx")
        archivos_excel = glob.glob(patron_excel)
        
        if not archivos_excel:
            # Buscar también archivos .xls
            patron_xls = os.path.join(carpeta_descargas, "*.xls")
            archivos_excel = glob.glob(patron_xls)
        
        if not archivos_excel:
            print("✗ No se encontraron archivos Excel para combinar")
            return False
        
        print(f"📁 Se encontraron {len(archivos_excel)} archivos para combinar:")
        for archivo in archivos_excel:
            print(f"  - {os.path.basename(archivo)}")
        
        # Lista para almacenar todos los DataFrames
        dataframes = []
        
        # Leer cada archivo Excel
        for archivo in archivos_excel:
            try:
                print(f"📖 Leyendo: {os.path.basename(archivo)}")
                df = pd.read_excel(archivo)
                
                # Añadir columna con información del período
                nombre_archivo = os.path.splitext(os.path.basename(archivo))[0]
                df['Periodo_Archivo'] = nombre_archivo
                
                dataframes.append(df)
                print(f"✓ Leído exitosamente: {len(df)} filas")
                
            except Exception as e:
                print(f"⚠️ Error leyendo {archivo}: {str(e)}")
                continue
        
        if not dataframes:
            print("✗ No se pudieron leer archivos Excel válidos")
            return False
        
        # Combinar todos los DataFrames
        print("🔄 Combinando todos los datos...")
        df_combinado = pd.concat(dataframes, ignore_index=True, sort=False)
        
        # Ordenar por fecha si existe una columna de fecha
        columnas_fecha_posibles = ['fecha', 'Fecha', 'FECHA', 'fecha_captura', 'Fecha_Captura', 'FECHA_CAPTURA']
        columna_fecha_encontrada = None
        
        for col in columnas_fecha_posibles:
            if col in df_combinado.columns:
                columna_fecha_encontrada = col
                break
        
        if columna_fecha_encontrada:
            try:
                df_combinado[columna_fecha_encontrada] = pd.to_datetime(df_combinado[columna_fecha_encontrada])
                df_combinado = df_combinado.sort_values(by=columna_fecha_encontrada)
                print(f"✓ Datos ordenados por {columna_fecha_encontrada}")
            except Exception as e:
                print(f"⚠️ No se pudo ordenar por fecha: {str(e)}")
        
        # Guardar el archivo combinado
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo_combinado = f"Ventas_Combinadas_2021_2025_{timestamp}.xlsx"
        ruta_archivo_combinado = os.path.join(carpeta_descargas, nombre_archivo_combinado)
        
        print(f"💾 Guardando archivo combinado: {nombre_archivo_combinado}")
        df_combinado.to_excel(ruta_archivo_combinado, index=False)
        
        print(f"✅ ¡Combinación completada!")
        print(f"📄 Archivo final: {nombre_archivo_combinado}")
        print(f"📊 Total de registros: {len(df_combinado):,}")
        print(f"📁 Ubicación: {ruta_archivo_combinado}")
        
        return ruta_archivo_combinado
        
    except Exception as e:
        print(f"✗ Error combinando archivos: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def login_nissan_rcn(usuario, password, opcion_dropdown):
    """Función principal mejorada para realizar el login y descarga por trimestres"""
    driver = None
    try:
        # Generar todos los períodos trimestres
        periodos = generar_periodos_trimestres(FECHA_INICIO_TOTAL, FECHA_FIN_TOTAL)
        print(f"\n📅 Se generaron {len(periodos)} períodos de descarga:")
        for i, periodo in enumerate(periodos, 1):
            print(f"  {i:2d}. {periodo['inicio_str']} - {periodo['fin_str']} ({periodo['nombre']})")
        
        driver = configurar_driver()
        print(f"\n🌐 Navegando a: {URL}")
        driver.get(URL)

        wait = WebDriverWait(driver, 15)

        # Realizar login
        print("\n🔐 Realizando login...")
        usuario_field = wait.until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Usuario']"))
        )
        password_field = driver.find_element(By.XPATH, "//input[@placeholder='Password']")
        usuario_field.clear()
        usuario_field.send_keys(usuario)
        password_field.clear()
        password_field.send_keys(password)

        login_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Ingresar')]")
        login_button.click()

        print("⏳ Esperando a que cargue la página después del login...")
        time.sleep(10)

        # Seleccionar dropdown
        try:
            dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//select")))
            select = Select(dropdown)
            select.select_by_visible_text(opcion_dropdown)
            print(f"✓ Opción seleccionada del dropdown: {opcion_dropdown}")
            time.sleep(3)
        except Exception as e:
            print(f"⚠️ No se pudo seleccionar la opción del dropdown: {str(e)}")

        # Navegar a la sección de consulta de venta
        if not navegar_a_consulta_venta(driver, wait):
            raise Exception("No se pudo navegar a la sección de Consulta de Venta")

        # Procesar cada período
        archivos_descargados = []
        print(f"\n🚀 Iniciando descarga de {len(periodos)} períodos...")
        
        for i, periodo in enumerate(periodos, 1):
            print(f"\n{'='*60}")
            print(f"📥 DESCARGANDO PERÍODO {i}/{len(periodos)}")
            print(f"📅 Período: {periodo['inicio_str']} - {periodo['fin_str']}")
            print(f"🏷️  Nombre: {periodo['nombre']}")
            print(f"{'='*60}")
            
            # Configurar fechas y exportar
            archivo_descargado = configurar_fechas_y_exportar(
                driver, wait, 
                periodo['inicio_str'], 
                periodo['fin_str'], 
                periodo['nombre']
            )
            
            if archivo_descargado:
                archivos_descargados.append(archivo_descargado)
                print(f"✅ Período {i}/{len(periodos)} completado exitosamente")
            else:
                print(f"❌ Error en el período {i}/{len(periodos)}")
            
            # Pausa entre descargas para evitar sobrecargar el servidor
            if i < len(periodos):  # No pausar después del último período
                print("⏸️  Pausa de 5 segundos antes del siguiente período...")
                time.sleep(5)
        
        # Combinar todos los archivos descargados
        print(f"\n{'='*60}")
        print("🔄 PROCESO DE COMBINACIÓN DE ARCHIVOS")
        print(f"{'='*60}")
        
        if archivos_descargados:
            carpeta_descargas = crear_carpeta_descargas()
            archivo_combinado = combinar_archivos_excel(carpeta_descargas)
            
            if archivo_combinado:
                print(f"\n🎉 ¡PROCESO COMPLETADO EXITOSAMENTE!")
                print(f"Total de períodos descargados: {len(archivos_descargados)}/{len(periodos)}")
                print(f"Archivo final combinado: {os.path.basename(archivo_combinado)}")
            else:
                print(f"\nDescargas completadas, pero hubo error en la combinación")
        else:
            print(f"\nNo se descargó ningún archivo exitosamente")

    except Exception as e:
        print(f"\nError durante el proceso: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # Verificar que pandas esté disponible
    try:
        import pandas as pd
    except ImportError:
        print("\nERROR: pandas no está instalado.")
        print("Instala pandas con: pip install pandas openpyxl")
        exit(1)
    
    if not USUARIO or not PASSWORD:
        print("\nERROR: Las credenciales no están configuradas.")
        print("Asegúrate de tener un archivo .env con:")
        print("NISSAN_USUARIO=tu_usuario")
        print("NISSAN_PASSWORD=tu_contraseña")
        print("NISSAN_URL=https://10.90.8.11/RcnVentas/Security/Login  # (opcional)")
        print("NISSAN_DROPDOWN_OPTION=307 - SADO DE ORIENTE  # (opcional)")
    else:
        print(f"\nINICIANDO PROCESO DE EXPORTACIÓN AUTOMÁTICA")
        print(f"Período total: {FECHA_INICIO_TOTAL.strftime('%d/%m/%Y')} - {FECHA_FIN_TOTAL.strftime('%d/%m/%Y')}")
        print(f"Las descargas se guardarán en la carpeta 'gruposado'")
        print(f"Al final se creará un archivo Excel combinado con todos los datos")
        
        login_nissan_rcn(USUARIO, PASSWORD, OPCION_DROPDOWN)