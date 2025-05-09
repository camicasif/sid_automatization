import os
import json
import time
import tempfile
import win32com.client as win32
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
        """Obtener valor anidado de configuración"""
        result = self._config
        for key in keys:
            result = result[key]
        return result

    @property
    def tss_range(self):
        """Obtener rango configurado para captura de imagen"""
        return self.get('celdas_tss', 'rango_llaves')


class TSSProcessor:
    """Procesador de archivos TSS (Template Site Sheet)"""

    def __init__(self, tss_path, config):
        self.tss_path = tss_path
        self.config = config
        self._load_workbook()

    def _load_workbook(self):
        """Cargar el workbook de Excel"""
        self.wb_tss = openpyxl.load_workbook(self.tss_path, data_only=True)
        info_sheet_index = self.config.get('hojas_tss', 'informacion')
        self.sheet_tss = self.wb_tss.worksheets[info_sheet_index]

    def obtener_valor(self, celda):
        """Obtener valor de celda limpiando espacios"""
        value = self.sheet_tss[celda].value
        return str(value).strip() if value else None

    def extraer_imagen(self):
        """Extrae imagen de ubicación del sitio"""
        foto_ubicacion = self.config.get('celdas_tss', 'foto_ubicacion')
        target_cell = self.sheet_tss[foto_ubicacion]

        # Buscar rango combinado
        merged_range = self._find_merged_range(target_cell)

        # Definir rango de búsqueda ampliado
        min_row, max_row, min_col, max_col = self._get_expanded_range(target_cell, merged_range)
        print(f"🔍 Rango de búsqueda: filas {min_row}-{max_row}, columnas {min_col}-{max_col}")

        # Buscar y procesar imagen
        return self._process_image_in_range(min_row, max_row, min_col, max_col)

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

    def _process_image_in_range(self, min_row, max_row, min_col, max_col):
        """Procesar imagen encontrada en el rango especificado"""
        for img in self.sheet_tss._images:
            img_top = img.anchor._from.row + 1
            img_left = img.anchor._from.col + 1

            if (min_row <= img_top <= max_row) and (min_col <= img_left <= max_col):
                print(f"🖼️ Imagen encontrada en fila={img_top}, columna={img_left}")
                return self._save_temp_image(img)

        print("⚠️ No se encontraron imágenes en el rango")
        return None

    def _save_temp_image(self, img):
        """Guardar imagen en archivo temporal"""
        try:
            img_data = img._data()
            img_pil = Image.open(io.BytesIO(img_data))

            temp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            img_pil.save(temp_img.name)
            temp_img.close()
            return temp_img.name
        except Exception as e:
            print(f"❌ Error al procesar imagen: {str(e)}")
            return None

    def capturar_rango_como_imagen(self, rango=None):
        """Capturar rango de Excel como imagen"""
        rango = rango or self.config.tss_range
        excel = win32.gencache.EnsureDispatch('Excel.Application')

        try:
            excel.Visible = True  # Necesario para operaciones de clipboard
            wb = excel.Workbooks.Open(
                r'C:\Users\VICTUS\PycharmProjects\sid_automatization\TSS_PRUEBA\TSS_MIGUILLAS_rev.0_29.04.2025.xlsm')
            sheet = wb.Sheets(1)  # Primera hoja

            print(f"📷 Capturando rango {rango} como imagen...")
            sheet.Range(rango).CopyPicture(Appearance=1, Format=2)
            time.sleep(1)  # Esperar para operación de copiado

            return self._capture_clipboard_image()

        except Exception as e:
            print(f"❌ Error al capturar rango: {e}")
            return None
        finally:
            excel.Quit()

    def _capture_clipboard_image(self):
        """Capturar imagen del portapapeles y guardar en temporal"""
        try:
            img = ImageGrab.grabclipboard()
            if not img:
                print("⚠️ No se encontró imagen en el portapapeles")
                return None

            temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            temp_path = temp_file.name
            img.save(temp_path)
            temp_file.close()

            print(f"✅ Imagen temporal guardada en: {temp_path}")
            return temp_path

        except Exception as e:
            print(f"❌ Error al capturar imagen: {e}")
            return None


class SIDGenerator:
    """Generador de archivos SID (Site Information Document)"""

    def __init__(self, plantilla_path, config):
        self.plantilla_path = plantilla_path
        self.config = config
        self.app = xw.App(visible=False)

    def crear_copia(self, output_path, datos_tss, imagen_ubicacion=None, imagen_rango=None):
        """Crear copia del SID con datos procesados"""
        try:
            wb = self.app.books.open(self.plantilla_path)

            self._fill_cover_page(wb, datos_tss)
            self._insert_location_image(wb, imagen_ubicacion)
            self._insert_data_range_image(wb, imagen_rango)

            wb.save(output_path)
            wb.close()
            return True

        except Exception as e:
            print(f"❌ Error generando SID: {str(e)}")
            return False
        finally:
            self.app.quit()

    def _fill_cover_page(self, wb, datos_tss):
        """Rellenar datos en la portada"""
        sheet_portada = wb.sheets[self.config.get('hojas_sid', 'portada')]
        sheet_portada[self.config.get('celdas_sid', 'codigo_portada')].value = datos_tss['id']

    def _insert_location_image(self, wb, imagen_ubicacion):
        """Insertar imagen de ubicación si existe"""
        if not imagen_ubicacion or not os.path.exists(imagen_ubicacion):
            return

        sheet = wb.sheets[self.config.get('hojas_sid', 'ubicacion_sitio')]
        celda = self.config.get('celdas_sid', 'foto_ubicacion')

        sheet.pictures.add(
            imagen_ubicacion,
            left=sheet.range(celda).left,
            top=sheet.range(celda).top,
            width=None,
            height=None
        )
        os.unlink(imagen_ubicacion)

    def _insert_data_range_image(self, wb, imagen_rango):
        """Insertar imagen de rango de datos si existe"""
        if not imagen_rango or not os.path.exists(imagen_rango):
            return

        sheet = wb.sheets[self.config.get('hojas_sid', 'datos_generales')]
        celda_llaves = self.config.get('celdas_sid', 'llaves_datos')

        sheet.pictures.add(
            imagen_rango,
            left=sheet.range(celda_llaves).left,
            top=sheet.range(celda_llaves).top,
            width=None,
            height=None
        )
        os.unlink(imagen_rango)

def main():
    # Inicializar configuración
    config = ConfigManager()

    # Preparar directorio de salida
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
    imagen_ubicacion = tss_processor.extraer_imagen()

    print("\n=== Capturando rango como imagen ===")
    imagen_rango = tss_processor.capturar_rango_como_imagen()

    if not imagen_rango:
        print("❌ No se pudo capturar el rango como imagen")
        return

    # Generar SID
    sid_generator = SIDGenerator("SID MIC BO 3YPLAN 2024_Name_ID_RevP.xlsx", config)
    nuevo_nombre = f"SID MIC BO 3YPLAN 2024_{datos['name']}_{datos['id']}_RevP.xlsx"
    output_path = os.path.join(output_folder, nuevo_nombre)

    if sid_generator.crear_copia(output_path, datos, imagen_ubicacion, imagen_rango):
        print(f"\n✅ SID generado exitosamente: {output_path}")
    else:
        print("\n❌ Fallo al generar SID")


if __name__ == "__main__":
    main()