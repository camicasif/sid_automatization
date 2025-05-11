import os
import json
import time
import win32com.client as win32
import win32gui
import win32con
from PIL import Image, ImageGrab
import xlwings as xw
import openpyxl
import io

class ConfigManager:
    """Manejador de configuración desde archivo JSON"""
    def __init__(self, config_path='config.json'):
        self.config_path = config_path
        self._load_config()

    def _load_config(self):
        with open(self.config_path) as f:
            self._config = json.load(f)

    def get(self, *keys):
        result = self._config
        for key in keys:
            result = result[key]
        return result

    @property
    def tss_range(self):
        return self.get('celdas_tss', 'rango_llaves')

class TSSProcessor:
    """Procesador de archivos TSS con manejo mejorado de diálogos"""
    def __init__(self, tss_path, config):
        self.tss_path = tss_path
        self.config = config
        self._load_workbook()

    def _load_workbook(self):
        self.wb_tss = openpyxl.load_workbook(self.tss_path, data_only=True)
        info_sheet_index = self.config.get('hojas_tss', 'informacion')
        self.sheet_tss = self.wb_tss.worksheets[info_sheet_index]

    def obtener_valor(self, celda):
        """Obtener valor de celda limpiando espacios"""
        value = self.sheet_tss[celda].value
        return str(value).strip() if value else None

    def listar_ventanas_office(self):
        office_windows = []
        def callback(hwnd, _):
            title = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            if (win32gui.IsWindowVisible(hwnd) and title and
                    ("Excel" in title or "Office" in title or class_name in ['NUIDialog', '#32770'])):
                office_windows.append((hwnd, title, class_name))
        win32gui.EnumWindows(callback, None)
        return office_windows

    def cerrar_dialogos_office(self):
        dialogs_closed = 0
        windows = self.listar_ventanas_office()
        for hwnd, title, class_name in windows:
            try:
                if class_name in ['NUIDialog', '#32770'] and "Excel" not in title:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    time.sleep(2)
                    if win32gui.IsWindow(hwnd):
                        win32gui.SendMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_CLOSE, 0)
                        time.sleep(1)
                    if not win32gui.IsWindow(hwnd):
                        dialogs_closed += 1
            except:
                pass
        return dialogs_closed

    def capturar_multiples_rangos(self, rangos):
        """Captura múltiples rangos y guarda directamente en carpeta capturas"""
        os.makedirs("capturas", exist_ok=True)
        excel = None
        resultados = {}

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = True
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(os.path.abspath(self.tss_path))
            sheet = wb.Sheets(1)

            self.cerrar_dialogos_office()

            for nombre, rango in rangos.items():
                output_path = os.path.join("capturas", f"{nombre}.png")
                for intento in range(3):
                    try:
                        sheet.Range(rango).CopyPicture(Appearance=1, Format=2)
                        time.sleep(2)

                        img = ImageGrab.grabclipboard()
                        if img:
                            img.save(output_path)
                            resultados[nombre] = output_path
                            print(f"✅ {nombre} guardado en {output_path}")
                            break
                    except Exception as e:
                        print(f"⚠️ Intento {intento+1} para {nombre}: {str(e)}")
                        self.cerrar_dialogos_office()
                        time.sleep(1)
                else:
                    resultados[nombre] = None

            return resultados

        except Exception as e:
            print(f"❌ Error al capturar rangos: {str(e)}")
            return {}
        finally:
            if excel:
                try:
                    excel.DisplayAlerts = False
                    excel.Quit()
                except:
                    pass

    def extraer_imagen(self):
        """Extrae imagen incrustada y la guarda en capturas"""
        os.makedirs("capturas", exist_ok=True)
        foto_ubicacion = self.config.get('celdas_tss', 'foto_ubicacion')
        target_cell = self.sheet_tss[foto_ubicacion]
        merged_range = self._find_merged_range(target_cell)
        min_row, max_row, min_col, max_col = self._get_expanded_range(target_cell, merged_range)

        for img in self.sheet_tss._images:
            img_top = img.anchor._from.row + 1
            img_left = img.anchor._from.col + 1
            if (min_row <= img_top <= max_row) and (min_col <= img_left <= max_col):
                output_path = os.path.join("capturas", "ubicacion.png")
                self._save_image(img, output_path)
                return output_path
        return None

    def _find_merged_range(self, target_cell):
        """Encontrar rango combinado para la celda objetivo"""
        for merged_cell in self.sheet_tss.merged_cells.ranges:
            if target_cell.coordinate in merged_cell:
                print(f"✅ Celda combinada encontrada: {merged_cell.coord}")
                return merged_cell
        print(f"ℹ️ Celda no está combinada")
        return None

    def _get_expanded_range(self, target_cell, merged_range):
        """Obtener rango ampliado para búsqueda de imagen"""
        if merged_range:
            min_row, max_row = merged_range.min_row, merged_range.max_row
            min_col, max_col = merged_range.min_col, merged_range.max_col
        else:
            min_row = max_row = target_cell.row
            min_col = max_col = target_cell.column

        # Ampliar rango con márgenes
        return (
            max(1, min_row - 1),  # expanded_min_row
            max_row,  # expanded_max_row
            max(1, min_col - 1),  # expanded_min_col
            max_col  # expanded_max_col
        )


    def _save_image(self, img, output_path):
        """Guarda imagen en ruta específica"""
        try:
            img_data = img._data()
            img_pil = Image.open(io.BytesIO(img_data))
            img_pil.save(output_path)
            return True
        except Exception as e:
            print(f"❌ Error al guardar imagen: {str(e)}")
            return False

class SIDGenerator:
    """Generador de SID usando imágenes de carpeta capturas"""
    def __init__(self, plantilla_path, config):
        self.plantilla_path = plantilla_path
        self.config = config
        self.app = xw.App(visible=False)

    def crear_copia(self, output_path, datos_tss):
        """Versión con diagnóstico completo"""
        try:
            # 1. Configurar rutas de imágenes (con verificación)
            print("\n=== CONFIGURANDO RUTAS DE IMÁGENES ===")
            imagenes_requeridas = {
                'ubicacion': os.path.join("capturas", "ubicacion.png"),
                'llaves': os.path.join("capturas", "llaves.png"),
                'observaciones': os.path.join("capturas", "observaciones.png"),
                'ingreso': os.path.join("capturas", "ingreso.png")
            }

            # Diagnóstico: Mostrar rutas completas
            print("Rutas configuradas:")
            for nombre, path in imagenes_requeridas.items():
                print(f" - {nombre}: {os.path.abspath(path)}")

            # 2. Verificación de archivos
            print("\n=== INICIANDO VERIFICACIÓN ===")
            archivos_faltantes = []

            if not imagenes_requeridas:  # Verificar si el diccionario está vacío
                raise ValueError("El diccionario imagenes_requeridas está vacío")

            for nombre, path in imagenes_requeridas.items():
                abs_path = os.path.abspath(path)
                if os.path.exists(abs_path):
                    print(f"✅ {nombre}: {abs_path} (Existe)")
                else:
                    archivos_faltantes.append(abs_path)
                    print(f"❌ {nombre}: FALTANTE ({abs_path})")

            if archivos_faltantes:
                raise FileNotFoundError(
                    f"Archivos faltantes:\n" + "\n".join(archivos_faltantes))

            # Iniciar generación del SID
            print("\n=== INICIANDO GENERACIÓN DE SID ===")
            wb = self.app.books.open(self.plantilla_path)

            try:
                 # 1. Verificar plantilla SID
                if not os.path.exists(self.plantilla_path):
                    raise FileNotFoundError(f"Plantilla SID no encontrada en: {os.path.abspath(self.plantilla_path)}")

                print(f"✔ Plantilla SID encontrada: {os.path.abspath(self.plantilla_path)}")

                # 2. Verificar directorio de salida
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                # 1. Llenar portada
                print("✔ Llenando datos de portada...")
                self._fill_cover_page(wb, datos_tss)

                # 2. Insertar imágenes con verificación
                print("\n✔ Insertando imágenes:")
                for nombre, img_path in imagenes_requeridas.items():
                    try:
                        print(f" - {nombre}...", end=" ")
                        self._insert_image(wb, img_path, nombre)
                        print("✅")
                    except Exception as e:
                        print(f"❌ (Error: {str(e)})")
                        raise

                # 3. Guardar el SID
                print(f"\n✔ Guardando SID en: {output_path}")
                wb.save(output_path)
                return True

            except Exception as e:
                print(f"\n❌ Error durante generación: {str(e)}")
                return False
            finally:
                try:
                    wb.close()
                except:
                    pass

        except FileNotFoundError as e:
            print(f"\n❌ Error crítico: {str(e)}")
            return False
        except Exception as e:
            print(f"\n❌ Error inesperado: {str(e)}")
            return False

    def _fill_cover_page(self, wb, datos_tss):
        sheet_portada = wb.sheets[self.config.get('hojas_sid', 'portada')]
        sheet_portada[self.config.get('celdas_sid', 'codigo_portada')].value = datos_tss['id']


    def _insert_image(self, wb, img_path, tipo):
        try:
            # 1. Mapeo de configuración (usando índices numéricos)
            config_map = {
                'ubicacion': (self.config.get('hojas_sid', 'ubicacion_sitio'), 'foto_ubicacion'),
                'llaves': (self.config.get('hojas_sid', 'datos_generales'), 'llaves_datos'),
                'observaciones': (self.config.get('hojas_sid', 'datos_generales'), 'observaciones_generales'),
                'ingreso': (self.config.get('hojas_sid', 'datos_generales'), 'ingreso')
            }

            # 2. Obtener índice de hoja y celda
            sheet_index, celda_key = config_map[tipo]

            # 3. Acceder a la hoja por índice
            try:
                sheet = wb.sheets[sheet_index]  # Usar índice numérico directamente
            except Exception as e:
                # Diagnóstico de hojas disponibles
                available_sheets = [(i, s.name) for i, s in enumerate(wb.sheets)]
                raise ValueError(
                    f"No se encontró la hoja con índice {sheet_index}. "
                    f"Hojas disponibles (índice, nombre): {available_sheets}"
                ) from e

            # 4. Obtener celda de destino
            celda = self.config.get('celdas_sid', celda_key)

            # 5. Insertar imagen
            sheet.pictures.add(
                img_path,
                left=sheet.range(celda).left,
                top=sheet.range(celda).top,
                width=sheet.range(celda).width,
                height=sheet.range(celda).height
            )

        except Exception as e:
            print(f"\n❌ ERROR insertando {tipo} en hoja {sheet_index} (celda {celda}):")
            print(f"Ruta imagen: {img_path}")
            print(f"Error completo: {str(e)}")
            raise
    def _limpiar_capturas(self):
        """Elimina archivos temporales después de usarlos"""
        try:
            for filename in os.listdir("capturas"):
                file_path = os.path.join("capturas", filename)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    print(f"⚠️ Error al eliminar {file_path}: {e}")
        except:
            pass

def main():
    config = ConfigManager()
    output_folder = "SIDs"
    os.makedirs(output_folder, exist_ok=True)

    # Buscar archivo TSS
    tss_files = [f for f in os.listdir("TSS_PRUEBA") if f.endswith(('.xls', '.xlsx', '.xlsm'))]
    if not tss_files:
        print("❌ No se encontraron archivos TSS")
        return

    # Procesar TSS
    tss_path = os.path.join("TSS_PRUEBA", tss_files[0])
    tss_processor = TSSProcessor(tss_path, config)

    # Obtener datos básicos
    datos = {
        'id': tss_processor.obtener_valor(config.get('celdas_tss', 'id')),
        'name': tss_processor.obtener_valor(config.get('celdas_tss', 'name'))
    }

    if not all(datos.values()):
        print("❌ Faltan datos requeridos en el TSS")
        return

    # Procesar imágenes
    print("=== Extrayendo imagen de ubicación ===")
    tss_processor.extraer_imagen()

    print("\n=== Capturando múltiples rangos ===")
    rangos_a_capturar = {
        'llaves': config.get('celdas_tss', 'rango_llaves'),
        'observaciones': config.get('celdas_tss', 'rango_observaciones_generales'),
        'ingreso': config.get('celdas_tss', 'rango_ingreso')
    }
    tss_processor.capturar_multiples_rangos(rangos_a_capturar)

    # Generar SID
    sid_generator = SIDGenerator("SID MIC BO 3YPLAN 2024_Name_ID_RevP.xlsx", config)
    nuevo_nombre = f"SID MIC BO 3YPLAN 2024_{datos['name']}_{datos['id']}_RevP.xlsx"
    output_path = os.path.join(output_folder, nuevo_nombre)

    if sid_generator.crear_copia(output_path, datos):
        print(f"\n✅ SID generado exitosamente: {output_path}")
    else:
        print("\n❌ Fallo al generar SID")

if __name__ == "__main__":
    main()
