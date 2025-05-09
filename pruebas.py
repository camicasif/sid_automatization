import win32com.client as win32


def verificar_excel():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        wb = excel.Workbooks.Open(r'C:\Users\VICTUS\PycharmProjects\sid_automatization\TSS_PRUEBA\TSS_MIGUILLAS_rev.0_29.04.2025.xlsm')
        sheet = wb.Sheets(1)  # Primera hoja

        print("Probando CopyPicture...")
        sheet.Range("A84:AM98").CopyPicture(Appearance=1, Format=2)

        print("Probando Paste...")
        sheet.Paste()

        print("Operación completada con éxito")
    except Exception as e:
        print(f"Error COM detallado: {e}")
        if hasattr(e, 'excepinfo'):
            print(f"Info adicional: {e.excepinfo}")
    finally:
        excel.Quit()


verificar_excel()



import datetime
import os
import json
import time

import xlwings as xw
import openpyxl
from PIL import Image, ImageGrab
import io
import tempfile
import win32com.client as win32
import tempfile

# Cargar configuración
with open('config.json') as f:
    config = json.load(f)


class TSSProcessor:
    def __init__(self, tss_path):
        self.tss_path = tss_path
        self.wb_tss = openpyxl.load_workbook(tss_path, data_only=True)
        self.sheet_tss = self.wb_tss.worksheets[config['hojas_tss']['informacion']]

    def obtener_valor(self, celda):
        return str(self.sheet_tss[celda].value).strip() if self.sheet_tss[celda].value else None

    def extraer_imagen(self):
        """Busca imágenes que intersecten con el rango combinado, incluyendo un margen ampliado"""
        foto_ubicacion = config['celdas_tss']['foto_ubicacion']
        target_cell = self.sheet_tss[foto_ubicacion]

        # 1. Encontrar el rango combinado
        merged_range = None
        for merged_cell in self.sheet_tss.merged_cells.ranges:
            if target_cell.coordinate in merged_cell:
                merged_range = merged_cell
                break

        if not merged_range:
            print(f"ℹ️ Celda {foto_ubicacion} NO está combinada. Usando solo esta celda.")
            min_row = max_row = target_cell.row
            min_col = max_col = target_cell.column
        else:
            print(f"✅ Celda {foto_ubicacion} está en rango combinado: {merged_range.coord}")
            min_row, max_row = merged_range.min_row, merged_range.max_row
            min_col, max_col = merged_range.min_col, merged_range.max_col

        # 2. Ampliar el rango (filas: -1/+0, columnas: -1/+0)
        expanded_min_row = max(1, min_row - 1)  # Asegurar que no sea menor a 1
        expanded_min_col = max(1, min_col - 1)  # Asegurar que no sea menor a 1
        expanded_max_row = max_row  # No ampliamos hacia abajo (ajustar si necesario)
        expanded_max_col = max_col  # No ampliamos hacia la derecha (ajustar si necesario)

        print(
            f"🔍 Rango de búsqueda ampliado: filas {expanded_min_row}-{expanded_max_row}, columnas {expanded_min_col}-{expanded_max_col}")

        # 3. Buscar imágenes cuya posición inicial esté en el rango ampliado
        for img in self.sheet_tss._images:
            img_top = img.anchor._from.row + 1
            img_left = img.anchor._from.col + 1

            print(f"  📌 Imagen en: fila={img_top}, columna={img_left}")

            # Verificar si la esquina superior izquierda está en el rango ampliado
            if (expanded_min_row <= img_top <= expanded_max_row) and (expanded_min_col <= img_left <= expanded_max_col):
                print("  🖼️ ¡Imagen encontrada dentro del rango ampliado!")
                try:
                    img_data = img._data()
                    img_pil = Image.open(io.BytesIO(img_data))

                    temp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                    img_pil.save(temp_img.name)
                    temp_img.close()
                    return temp_img.name
                except Exception as e:
                    print(f"  ❌ Error al procesar imagen: {str(e)}")

        print("⚠️ No se encontraron imágenes dentro del rango ampliado.")
        return None

    def capturar_rango_como_imagen(self):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            wb = excel.Workbooks.Open(
                r'C:\Users\VICTUS\PycharmProjects\sid_automatization\TSS_PRUEBA\TSS_MIGUILLAS_rev.0_29.04.2025.xlsm')
            sheet = wb.Sheets(1)  # Primera hoja

            # Definir el rango a capturar
            rango = "A84:AM98"
            print(f"Capturando rango {rango} como imagen...")

            # Copiar como imagen
            sheet.Range(rango).CopyPicture(Appearance=1, Format=2)

            # Esperar un momento para que se complete la operación de copiado
            time.sleep(1)

            # Intentar capturar directamente del portapapeles
            try:
                img = ImageGrab.grabclipboard()
                if img:
                    # Crear archivo temporal
                    temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                    temp_path = temp_file.name
                    temp_file.close()

                    # Guardar la imagen
                    img.save(temp_path)
                    print(f"✅ Imagen guardada temporalmente en: {temp_path}")
                    return temp_path
                else:
                    print("⚠️ No se encontró imagen en el portapapeles")
                    return None

            except Exception as clipboard_error:
                print(f"❌ Error al capturar del portapapeles: {clipboard_error}")
                return None

        except Exception as e:
            print(f"❌ Error general al capturar rango como imagen: {e}")
            return None
        finally:
            excel.Quit()

class SIDGenerator:
    def __init__(self, plantilla_path):
        self.plantilla_path = plantilla_path
        self.app = xw.App(visible=False)

    def crear_copia(self, output_path, datos_tss, imagen_ubicacion=None, imagen_rango=None):
        """Crea el SID usando datos del TSS ya procesados"""
        try:
            # Crear copia del SID
            wb = self.app.books.open(self.plantilla_path)

            # Portada
            sheet_portada = wb.sheets[config['hojas_sid']['portada']]
            sheet_portada[config['celdas_sid']['codigo_portada']].value = datos_tss['id']

            # Ubicación (imagen)
            if imagen_ubicacion and os.path.exists(imagen_ubicacion):
                sheet_ubicacion = wb.sheets[config['hojas_sid']['ubicacion_sitio']]
                celda = config['celdas_sid']['foto_ubicacion']
                sheet_ubicacion.pictures.add(
                    imagen_ubicacion,
                    left=sheet_ubicacion.range(celda).left,
                    top=sheet_ubicacion.range(celda).top,
                    width=None,
                    height=None
                )
                os.unlink(imagen_ubicacion)

            # Datos generales (rango)
            if imagen_rango and os.path.exists(imagen_rango):
                sheet_datos = wb.sheets[config['hojas_sid']['datos_generales']]
                sheet_datos.pictures.add(
                    imagen_rango,
                    left=sheet_datos.range("B480").left,
                    top=sheet_datos.range("B480").top,
                    width=None,
                    height=None
                )
                os.unlink(imagen_rango)
            # Guardar y cerrar
            wb.save(output_path)
            wb.close()
            return True

        except Exception as e:
            print(f"Error generando SID: {str(e)}")
            return False
        finally:
            self.app.quit()

def main():
    # Configuración de paths
    output_folder = "SIDs"
    os.makedirs(output_folder, exist_ok=True)

    # Buscar archivo TSS
    tss_files = [f for f in os.listdir("TSS_PRUEBA") if f.endswith(('.xls', '.xlsx', '.xlsm'))]
    if not tss_files:
        print("No se encontraron archivos TSS")
        return

    tss_path = os.path.join("TSS_PRUEBA", tss_files[0])

    # Procesar TSS
    tss_processor = TSSProcessor(tss_path)

    datos = {
        'id': tss_processor.obtener_valor(config['celdas_tss']['id']),
        'name': tss_processor.obtener_valor(config['celdas_tss']['name'])
    }

    if not all(datos.values()):
        print("Faltan datos requeridos en el TSS")
        return

    imagen_path = tss_processor.extraer_imagen()

    print("=== Iniciando captura de rango como imagen ===")
    imagen_rango = tss_processor.capturar_rango_como_imagen()

    if not imagen_rango:
        print("❌ Error al capturar rango como imagen. Debug info:")
        return
    else:
        print("✅ Rango capturado exitosamente")


    # Generar SID
    sid_generator = SIDGenerator("SID MIC BO 3YPLAN 2024_Name_ID_RevP.xlsx")
    nuevo_nombre = f"SID MIC BO 3YPLAN 2024_{datos['name']}_{datos['id']}_RevP.xlsx"
    output_path = os.path.join(output_folder, nuevo_nombre)

    if sid_generator.crear_copia(output_path, datos, imagen_path, imagen_rango):
        print(f"✅ SID generado: {output_path}")
    else:
        print("❌ Error al generar SID")


if __name__ == "__main__":
    main()