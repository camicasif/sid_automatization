import io
import os
import json
import warnings
from datetime import timedelta

import win32com
from PIL import Image, ImageGrab
import openpyxl
import xlwings as xw
import win32com.client as win32
import win32gui
import win32con
import time
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl.reader.drawings')

# -*- coding: utf-8 -*-
def listar_ventanas_office():
    office_windows = []

    def callback(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        class_name = win32gui.GetClassName(hwnd)
        if (win32gui.IsWindowVisible(hwnd) and title and
                ("Excel" in title or "Office" in title or class_name in ['NUIDialog', '#32770'])):
            office_windows.append((hwnd, title, class_name))

    win32gui.EnumWindows(callback, None)
    return office_windows


def cerrar_dialogos_office():
    dialogs_closed = 0
    windows = listar_ventanas_office()
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

class TSSInstance:
    """Representa un archivo TSS individual con sus metadatos"""
    def __init__(self, file_path,config_path='config.json'):

        self.file_path = file_path
        self.name = "DEFAULT_NAME"
        self.id = "DEFAULT_ID"
        self.data = {'textos': {}, 'imagenes': {}}
        self.resultados_dir = ""
        self.config = _cargar_configuracion(config_path)
        self._extraer_metadatos()


    def _extraer_metadatos(self):
        """Extrae name/id al inicializar cada instancia usando la configuración"""
        try:
            wb = openpyxl.load_workbook(self.file_path, data_only=True)
            # Usar nombres de configuración en lugar de nombres directos de hojas
            self.name = self._leer_celda(wb, "informacion", "H8")
            self.id = self._leer_celda(wb, "informacion", "H7")
            wb.close()
        except Exception as e:
            print(f"⚠️ Error extrayendo metadatos de {self.file_path}: {str(e)}")
            self.name = f"ERROR_{os.path.basename(self.file_path)}"
            self.id = time.strftime('%Y%m%d%H%M%S')

    def _leer_celda(self, wb, sheet_config_name, celda):

        try:
            # Obtener el índice de la hoja desde la configuración
            sheet_index = self._obtener_hoja_indice('tss', sheet_config_name)

            # Obtener la hoja por índice
            sheet = wb.worksheets[sheet_index]

            # Leer y limpiar el valor de la celda
            valor = sheet[celda].value
            return str(valor).strip() if valor is not None else ""

        except KeyError as e:
            print(f"⚠️ Error: No se encontró la hoja '{sheet_config_name}' en la configuración")
            return ""
        except IndexError as e:
            print(f"⚠️ Error: Índice {sheet_index} no existe en el workbook para hoja '{sheet_config_name}'")
            return ""
        except Exception as e:
            print(f"⚠️ Error leyendo celda {celda} de hoja '{sheet_config_name}': {str(e)}")
            return ""

    def _obtener_hoja_indice(self, workbook_type, sheet_name):
        return self.config['hojas'][workbook_type][sheet_name]

def _limpiar_texto(texto):
    """Limpio texto para usar en nombres de archivos"""
    return ''.join(c for c in texto if c not in '\\/:*?"<>|').replace(" ", "_")


def _cargar_configuracion(config_path):
    with open(config_path, encoding='utf-8') as f:
        return json.load(f)


class TSSBatchProcessor:
    """Procesa múltiples archivos TSS en lote"""

    def __init__(self, config_path='config.json'):
        self.config = _cargar_configuracion(config_path)
        self.tss_instances = []  # Lista de objetos TSSInstance
        self.total_time = 0

    def procesar_lote(self, tss_folder="TSS"):

        """Procesa todos los TSS en un directorio con medición de tiempo"""
        print("\n" + "=" * 50)
        print(" INICIANDO PROCESAMIENTO POR LOTES ")
        print("=" * 50 + "\n")
        """Procesa todos los TSS en un directorio"""
        tss_files = self._encontrar_archivos_tss(tss_folder)
        total_files = len(tss_files)
        start_time_total = time.monotonic()

        for i, tss_path in enumerate(tss_files, 1):
            print(f"\n📂 Procesando archivo {i} de {total_files}")
            file_start_time = time.monotonic()

            tss_instance = TSSInstance(tss_path)
            self.tss_instances.append(tss_instance)
            self._procesar_individual(tss_instance)

            file_time = time.monotonic() - file_start_time
            self.total_time += file_time
            print(f"⏱️ Tiempo archivo: {timedelta(seconds=file_time)}")

            # Estimación del tiempo restante
            remaining_files = total_files - i
            avg_time = self.total_time / i
            estimated_remaining = avg_time * remaining_files
            print(f"⏳ Estimado restante: {timedelta(seconds=estimated_remaining)}")

        total_elapsed = time.monotonic() - start_time_total
        print("\n" + "=" * 50)
        print(" RESUMEN DE TIEMPOS ")
        print("=" * 50)
        print(f"📊 Total archivos procesados: {total_files}")
        print(f"⏱️ Tiempo total: {timedelta(seconds=total_elapsed)}")
        print(f"⏱️ Tiempo promedio por archivo: {timedelta(seconds=total_elapsed / total_files if total_files else 0)}")
        print("=" * 50 + "\n")

    def _encontrar_archivos_tss(self, folder_path):
        """Encuentra todos los archivos Excel (.xlsx) en el directorio especificado"""
        tss_files = []
        try:
            if not os.path.exists(folder_path):
                print(f"⚠️ El directorio {folder_path} no existe")
                return tss_files

            for filename in os.listdir(folder_path):
                if filename.lower().endswith(('.xls', '.xlsx', '.xlsm')) and not filename.startswith('~$'):
                    full_path = os.path.join(folder_path, filename)
                    tss_files.append(full_path)

            print(f"📁 Encontrados {len(tss_files)} archivos TSS en {folder_path}")
            return tss_files
        except Exception as e:
            print(f"❌ Error buscando archivos TSS: {str(e)}")
            return []

    def _procesar_individual(self, tss_instance):
        """Procesamiento completo para un TSS"""
        print(f"\n🔁 Procesando {tss_instance.name}_{tss_instance.id}")

        # 1. Configurar rutas
        output_dir = os.path.join("resultados", f"{tss_instance.name}_{tss_instance.id}")
        os.makedirs(output_dir, exist_ok=True)

        valores = {
            'name': tss_instance.name,
            'id': tss_instance.id
        }
        nombre_archivo = self.config['nombre_sid']['formato'].format(**valores)
        nombre_archivo = _limpiar_texto(nombre_archivo)

        output_folder = "SIDs"
        os.makedirs(output_folder, exist_ok=True)
        output_path = os.path.join(output_folder, nombre_archivo)

        # 2. Procesar contenido (adaptar tus métodos actuales)
        self._extraer_datos(tss_instance)

        self._generar_sid(
            tss_instance,
            self.config['nombre_sid']['plantilla'],
            output_path
        )
        print(f"✅ Proceso completado para {tss_instance.file_path}")

    # Configuración y helpers básicos

    def _obtener_hoja_indice(self, workbook_type, sheet_name):
        return self.config['hojas'][workbook_type][sheet_name]



    #Capturar informacion del tss

    def _extraer_datos(self, tss_instance):
        """Procesa el TSS agrupando elementos por tipo para optimización"""
        print(f"\n=== EXTRAYENDO DATOS DE {tss_instance.name}_{tss_instance.id} ===")
        wb_tss = None
        try:
            wb_tss = openpyxl.load_workbook(tss_instance.file_path, data_only=True)
            tss_instance.resultados_dir = os.path.join("resultados", f"{tss_instance.name}_{tss_instance.id}")
            os.makedirs(tss_instance.resultados_dir, exist_ok=True)

            # Organizar elementos por tipo para procesamiento eficiente
            elementos_por_tipo = {
                'rango': [],
                'imagen': [],
                'texto': []
            }

            for elemento in self.config['elementos']:
                elementos_por_tipo[elemento['tipo']].append(elemento)

            # Procesar textos primero (más rápido)
            for elemento in elementos_por_tipo['texto']:
                self._procesar_texto(wb_tss, tss_instance, elemento)

            # Procesar rangos (requiere Excel COM)
            if elementos_por_tipo['rango']:
                self._procesar_rangos_agrupados(wb_tss, tss_instance, elementos_por_tipo['rango'])

            # Procesar imágenes
            for elemento in elementos_por_tipo['imagen']:
                self._procesar_imagen(wb_tss, tss_instance, elemento)

            print(f"✅ Extracción completada para {tss_instance.name}_{tss_instance.id}")
            return True

        except Exception as e:
            print(f"❌ Error en extracción de datos: {str(e)}")
            return False
        finally:
            wb_tss.close()
    # Procesamiento interno del tss

    def _procesar_texto(self, wb_tss, tss_instance, elemento):
        """Procesa un elemento de texto y lo almacena en la instancia"""
        try:
            sheet_index = self._obtener_hoja_indice('tss', elemento['origen']['hoja'])
            sheet = wb_tss.worksheets[sheet_index]
            valor = sheet[elemento['origen']['celda']].value
            tss_instance.data['textos'][elemento['nombre']] = str(valor).strip() if valor else ""
            print(f"Texto '{elemento['nombre']}' extraído: {tss_instance.data['textos'][elemento['nombre']][:50]}...")
        except Exception as e:
            print(f"⚠️ Error procesando texto {elemento['nombre']}: {str(e)}")

    def _procesar_rangos_agrupados(self, wb_tss, tss_instance, elementos_rango):
        """Procesa múltiples rangos usando Excel COM"""
        try:

            rangos_dict = {
                elem['nombre']: {
                    'rango': elem['origen']['rango'],
                    'hoja': elem['origen']['hoja']
                }
                for elem in elementos_rango
            }

            resultados = self.capturar_multiples_rangos(tss_instance, rangos_dict)


                # 4. Almacenar rutas de imágenes válidas
            for nombre, ruta_imagen in resultados.items():
                if ruta_imagen and os.path.exists(ruta_imagen):
                    tss_instance.data['imagenes'][nombre] = ruta_imagen
                    print(f"✅ Rango '{nombre}' guardado en {ruta_imagen}")
                else:
                    print(f"⚠️ No se pudo capturar el rango '{nombre}' o la imagen no existe")

            return True

        except Exception as e:
            print(f"❌ Error en procesamiento de rangos agrupados: {str(e)}")
            return False

    def capturar_multiples_rangos(self, tss_instance, rangos_dict):
        """Captura rangos usando Excel COM y guarda en la carpeta de la instancia"""
        excel = None
        resultados = {}
        wb = None

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = True
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(os.path.abspath(tss_instance.file_path))
            cerrar_dialogos_office()

            os.makedirs(tss_instance.resultados_dir, exist_ok=True)

            for nombre, config in rangos_dict.items():
                output_path = os.path.join(tss_instance.resultados_dir, f"{nombre}.png")
                print("outputpath resultados dir", output_path)
                sheet_index = self._obtener_hoja_indice('tss', config['hoja']) + 1
                sheet = wb.Sheets(sheet_index)

                for intento in range(3):
                    try:
                        sheet.Range(config['rango']).CopyPicture(Appearance=1, Format=2)
                        time.sleep(2)

                        img = ImageGrab.grabclipboard()
                        if img:
                            img.save(output_path)
                            resultados[nombre] = output_path
                            print(f"✅ {nombre} guardado en {output_path}")
                            break
                    except Exception as e:
                        print(f"⚠️ Intento {intento + 1} para {nombre}: {str(e)}")
                        time.sleep(1)
                else:
                    resultados[nombre] = None
                    print(f"❌ No se pudo capturar el rango '{nombre}'")

            return resultados


        except Exception as e:
            print(f"❌ Error crítico en captura de rangos: {str(e)}")
            return {}  # Retornar diccionario vacío en caso de error crítico

        finally:
            # Cerrar todo correctamente
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception as e:
                print(f"⚠️ Error cerrando libro: {str(e)}")
            try:
                if excel is not None:
                    excel.DisplayAlerts = False
                    excel.Quit()
            except Exception as e:
                print(f"⚠️ Error cerrando Excel: {str(e)}")
            # Liberar recursos COM
            del wb
            del excel

    def _procesar_imagen(self, wb_tss, tss_instance, elemento):
        """Busca imágenes mostrando el rango de celdas de búsqueda"""
        try:
            sheet_index = self._obtener_hoja_indice('tss', elemento['origen']['hoja'])
            sheet = wb_tss.worksheets[sheet_index]
            celda = sheet[elemento['origen']['celda']]

            # Determinar coordenadas de búsqueda
            merged_range = self._encontrar_rango_combinado(celda, sheet)  # Pasar sheet como parámetro
            min_row, max_row, min_col, max_col = self._obtener_rango_expandido(celda, merged_range)

            # Convertir coordenadas numéricas a formato de letra de columna (A, B, C...)
            col_letter_start = openpyxl.utils.get_column_letter(min_col)
            col_letter_end = openpyxl.utils.get_column_letter(max_col)

            print(f"🔍 Buscando imagen {elemento['nombre']} en rango: "
                  f"{col_letter_start}{min_row}:{col_letter_end}{max_row} "
                  f"(Columnas {min_col}-{max_col}, Filas {min_row}-{max_row})")

            # Verificar si la hoja tiene imágenes antes de intentar acceder
            if not hasattr(sheet, '_images'):
                print(f"⚠️ Hoja {sheet.title} no contiene imágenes")
                return None

            # Buscar imagen en el rango
            for img in sheet._images:
                img_top = img.anchor._from.row + 1
                img_left = img.anchor._from.col + 1

                if (min_row <= img_top <= max_row) and (min_col <= img_left <= max_col):
                    img_path = os.path.join(tss_instance.resultados_dir, f"{elemento['nombre']}.png")
                    os.makedirs(os.path.dirname(img_path), exist_ok=True)  # Asegurar que el directorio existe

                    image_bytes = img._data()
                    image = Image.open(io.BytesIO(image_bytes))
                    image.save(img_path)
                    tss_instance.data['imagenes'][elemento['nombre']] = img_path
                    print(f"✅ Imagen '{elemento['nombre']}' encontrada en posición: "
                          f"Columna {img_left}, Fila {img_top}")
                    return img_path

            print(f"⚠️ Imagen {elemento['nombre']} no encontrada en el rango especificado")
            return None

        except Exception as e:
            print(f"❌ Error al buscar imagen: {str(e)}")
            return None

    def _encontrar_rango_combinado(self, target_cell, sheet):
        """Encontrar rango combinado para la celda objetivo"""
        for merged_cell in sheet.merged_cells.ranges:  # Usar sheet en lugar de self.sheet_tss
            if target_cell.coordinate in merged_cell:
                print(f"\n ✅ Celda combinada encontrada: {merged_cell.coord}")
                return merged_cell
        print(f"\n ℹ️ Celda no está combinada")
        return None

    def _obtener_rango_expandido(self, target_cell, merged_range):
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
            max_row,              # expanded_max_row
            max(1, min_col - 1),  # expanded_min_col
            max_col               # expanded_max_col
        )

    #Generacion de sid

    def _generar_sid(self, tss_instance, plantilla_path, output_path):
        """Genera el SID con los datos extraídos, soportando múltiples celdas destino"""
        print("\n=== GENERANDO SID ===")
        app = xw.App(visible=False)

        try:
            wb_sid = app.books.open(plantilla_path)

            # 1. Insertar textos (ahora soporta múltiples celdas destino)
            for elemento in self.config['elementos']:
                if elemento['tipo'] == 'texto' and elemento['nombre'] in tss_instance.data['textos']:
                    sheet_index = self._obtener_hoja_indice('sid', elemento['destino']['hoja'])
                    sheet = wb_sid.sheets[sheet_index]
                    valor = tss_instance.data['textos'][elemento['nombre']]

                    # Insertar el mismo valor en todas las celdas especificadas
                    for celda in elemento['destino']['celdas']:
                        sheet[celda].value = valor
                        print(f"Texto '{elemento['nombre']}' insertado en {celda}")

            # 2. Insertar imágenes/rangos (ya soporta múltiples celdas via _insertar_imagen)
            for elemento in self.config['elementos']:
                if elemento['tipo'] in ['imagen', 'rango'] and elemento['nombre'] in tss_instance.data['imagenes']:
                    self._insertar_imagen(wb_sid,tss_instance, elemento)
            # Guardar el resultado
            wb_sid.save(output_path)
            print(f"\n✅ SID generado correctamente en: {os.path.abspath(output_path)}")

        except Exception as e:
            print(f"\n❌ Error generando SID: {str(e)}")
            raise
        finally:
            app.quit()

    def _obtener_hoja(self, wb, sheet_identifier, book_type='sid'):
        """
        Obtiene una hoja por nombre o índice, con manejo de errores mejorado
        :param wb: Libro de trabajo (xlwings)
        :param sheet_identifier: Nombre o índice de la hoja
        :param book_type: 'sid' o 'tss' (para el mapeo de config)
        :return: Objeto hoja
        """
        try:
            # Si es string, buscar en la configuración
            sheet_index = self._obtener_hoja_indice(book_type, sheet_identifier)
            return wb.sheets[sheet_index]

        except Exception as e:
            available_sheets = "\n".join([f"- {s.name} (índice {i})" for i, s in enumerate(wb.sheets)])
            raise ValueError(
                f"No se pudo encontrar la hoja '{sheet_identifier}'.\n"
                f"Hojas disponibles:\n{available_sheets}"
            ) from e

    def _insertar_imagen(self, wb_sid, tss_instance, elemento):
        """Versión que soporta tamaño específico para imágenes y centrado en celda"""

        nombre = elemento['nombre']

        try:
            print(f"\n=== Insertando imagen '{nombre}' ===")

            # 1. Verificar existencia de la imagen
            img_path = os.path.abspath(tss_instance.data['imagenes'].get(nombre))
            if not os.path.exists(img_path):
                raise FileNotFoundError(
                    f"Imagen no encontrada.\nBuscada: {img_path}")

            # 2. Obtener hoja destino
            sheet_index = self._obtener_hoja_indice('sid', elemento['destino']['hoja'])
            sheet = wb_sid.sheets[sheet_index]
            print(f"Hoja destino: {sheet.name} (índice {sheet.index})")

            # 3. Obtener y convertir dimensiones (cm a puntos)
            width_cm = elemento['destino'].get('ancho')  # En cm
            height_cm = elemento['destino'].get('alto')  # En cm

            # Convertir cm a puntos (1 cm = 28.35 puntos)
            width = width_cm * 28.35 if width_cm is not None else None
            height = height_cm * 28.35 if height_cm is not None else None

            print(f"Configuración de tamaño - Ancho: {width_cm}cm ({width}pt), Alto: {height_cm}cm ({height}pt)")

            # 4. Procesar TODAS las celdas destino
            for celda in elemento['destino']['celdas']:
                try:
                    rango = sheet.range(celda)
                    print(f"Insertando en celda: {rango.address}")

                    # Calcular posición centrada
                    left = rango.left + 5
                    top = rango.top + 5
                    # Insertar imagen
                    picture = sheet.pictures.add(
                        img_path,
                        left=left,
                        top=top,
                        width=width,
                        height=height
                    )

                    # picture.api.ShapeRange.ZOrder(win32com.client.constants.msoSendToBack)
                    #
                    # # Opcional: Bloquear posición y tamaño
                    # picture.api.Placement = 1

                    # # Mantener relación de aspecto si solo se especifica una dimensión
                    # if width is not None and height is None:
                    #     # Mantener relación de aspecto basado en el ancho
                    #     img = Image.open(img_path)
                    #     aspect_ratio = img.height / img.width
                    #     picture.height = width * aspect_ratio
                    #     # Recalcular posición vertical después de ajustar altura
                    #     picture.top = rango.top + (rango.height - picture.height) / 2
                    # elif height is not None and width is None:
                    #     # Mantener relación de aspecto basado en el alto
                    #     img = Image.open(img_path)
                    #     aspect_ratio = img.width / img.height
                    #     picture.width = height * aspect_ratio
                    #     # Recalcular posición horizontal después de ajustar ancho
                    #     picture.left = rango.left + (rango.width - picture.width) / 2

                    print(f"✅ Imagen insertada en {celda} - Tamaño: {width or 'auto'}x{height or 'auto'}")

                except Exception as e:
                    print(f"⚠️ Error insertando en {celda}: {type(e).__name__} - {str(e)}")

            return True

        except Exception as e:
            print(f"\n❌ ERROR insertando '{nombre}': {type(e).__name__}")
            print(f"Mensaje: {str(e)}")
            return False


# Uso del sistema
if __name__ == "__main__":
    processor = TSSBatchProcessor('config.json')

    # Procesar todos los TSS encontrados
    processor.procesar_lote("TSS")

    # Alternativa para procesar uno específico
    # tss_instance = TSSInstance("ruta/especifica.xlsx")
    # processor._procesar_individual(tss_instance)