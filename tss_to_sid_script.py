import os
import xlwings as xw
from shutil import copyfile
import openpyxl


def debug_hojas(workbook):
    """Función para debuggear todas las hojas del archivo"""
    print("\n=== DEBUG DE HOJAS ===")
    for i, sheet in enumerate(workbook.sheets):
        print(f"\nHoja {i}: '{sheet.name}'")
        #print(f"  - Protegida: {'Sí' if sheet.protect_contents else 'No'}")

        # Verificar celdas específicas (A3 y A45)
        for celda in ['A3', 'A45']:
            try:
                valor = sheet.range(celda).value
                print(f"  - {celda}: {valor if valor else '<vacío>'}")
            except Exception as e:
                print(f"  - {celda}: Error al leer ({str(e)})")


def crear_copia_sid():
    # Configuración inicial
    plantilla_sid = "SID MIC BO 3YPLAN 2024_Name_ID_RevP.xlsx"
    output_folder = "SIDs"
    os.makedirs(output_folder, exist_ok=True)

    # Buscar archivo TSS
    tss_files = [f for f in os.listdir("TSS_PRUEBA") if f.lower().endswith(('.xls', '.xlsx', '.xlsm'))]
    if not tss_files:
        print("No se encontraron archivos TSS en TSS_PRUEBA")
        return

    tss_path = os.path.join("TSS_PRUEBA", tss_files[0])

    try:
        # Obtener datos del TSS
        wb_tss = openpyxl.load_workbook(tss_path, data_only=True)
        name = wb_tss.worksheets[0]['H8'].value
        id_val = str(wb_tss.worksheets[0]['H7'].value)

        if not name or not id_val:
            print("No se encontraron Name o ID en H8/H7")
            return

        # Preparar archivo de salida
        nuevo_nombre = f"SID MIC BO 3YPLAN 2024_{name}_{id_val}_RevP.xlsx"
        output_path = os.path.join(output_folder, nuevo_nombre)

        # Proceso con xlwings
        app = xw.App(visible=False)

        try:
            # 1. Abrir y debuggear plantilla original
            print("\n=== PLANTILLA ORIGINAL ===")
            wb_template = app.books.open(plantilla_sid)
            wb_template.save(output_path)
            wb_template.close()

            # 2. Procesar archivo de salida
            print("\n=== ARCHIVO DE SALIDA ===")
            wb_sid = app.books.open(output_path)

            target_sheet = wb_sid.sheets[4]

            print(f"\nHoja seleccionada: '{target_sheet.name}'")

            # 3. Modificar A45
            print(f"\n[ANTES] A45 = {target_sheet.range('A45').value}")
            target_sheet.range('A45').value = id_val
            print(f"[DESPUÉS] A45 = {target_sheet.range('A45').value}")

            # 4. Verificar A3
            print(f"\nValor en A3: {target_sheet.range('A3').value}")

            wb_sid.save()
            wb_sid.close()

            print(f"\n✅ Archivo generado: {output_path}")

        except Exception as e:
            print(f"\n❌ Error durante el procesamiento: {str(e)}")
            raise
        finally:
            app.quit()

    except Exception as e:
        print(f"\n❌ Error general: {str(e)}")


if __name__ == "__main__":
    crear_copia_sid()