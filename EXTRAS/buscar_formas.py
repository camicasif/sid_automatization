import glob
import os
import re

import xlwings as xw


def analizar_grupos_formas(archivo_excel, hoja_num=10):
    """
    Analiza grupos de formas y sus contenidos en una hoja Excel

    Args:
        archivo_excel (str): Ruta del archivo Excel
        hoja_num (int): Número de hoja (base 1)
    """
    app = xw.App(visible=False)
    try:
        # Abrir el libro
        libro = app.books.open(archivo_excel)
        hoja = libro.sheets[hoja_num - 1]

        print(f"\nAnalizando hoja: '{hoja.name}'")

        # Procesar todas las formas
        for shape in hoja.shapes:
            print(f"\nForma encontrada: {shape.name} (Tipo: {shape.type})")

            # Manejar grupos
            if shape.type == 'group':
                print("  └ Este es un GRUPO que contiene:")

                # Acceder a las formas dentro del grupo
                try:
                    group_items = shape.api.GroupItems
                    for i in range(group_items.Count):
                        sub_shape = group_items.Item(i + 1)  # Índices base 1

                        # Obtener propiedades de la subforma
                        shape_name = sub_shape.Name
                        shape_type = sub_shape.Type

                        print(f"    ├ Subforma {i + 1}: {shape_name} (Tipo: {shape_type})")

                        # Obtener texto si existe
                        try:
                            if hasattr(sub_shape, 'TextFrame2'):
                                texto = sub_shape.TextFrame2.TextRange.Text
                                if texto and texto.strip():
                                    print(f"    │   Texto: '{texto.strip()}'")
                                else:
                                    print("    │   Texto: [vacío]")
                        except Exception as text_error:
                            print(f"    │   Error al leer texto: {str(text_error)}")

                        # Obtener posición y tamaño
                        try:
                            print(f"    │   Posición: Top={sub_shape.Top}, Left={sub_shape.Left}")
                            print(f"    │   Tamaño: Width={sub_shape.Width}, Height={sub_shape.Height}")
                        except Exception as pos_error:
                            print(f"    │   Error al obtener posición: {str(pos_error)}")

                except Exception as group_error:
                    print(f"    └ Error al procesar grupo: {str(group_error)}")

            # Manejar formas individuales
            else:
                try:
                    if hasattr(shape.api, 'TextFrame2'):
                        texto = shape.api.TextFrame2.TextRange.Text
                        if texto and texto.strip():
                            print(f"  └ Texto: '{texto.strip()}'")
                except Exception as text_error:
                    print(f"  └ Error al leer texto: {str(text_error)}")

    except Exception as e:
        print(f"\nError general: {str(e)}")
    finally:
        libro.close()
        app.quit()

    def _actualizar_sectores_con_tecnologias(self, wb_sid, tss_instance):
        """Actualiza los TextBox tipo 1 (sectores) con tecnologías correspondientes"""
        try:
            print("\n=== ACTUALIZANDO SECTORES CON TECNOLOGÍAS ===")
            sheet = wb_sid.sheets[self._obtener_hoja_indice('sid', 'antenas')]

            # Procesar cada grupo en la hoja
            for shape in sheet.shapes:
                if shape.type == 'group':
                    try:
                        # Buscar TextBox de sectores (tipo 1) en el grupo
                        sectores = []
                        for sub_shape in shape.api.GroupItems:
                            if sub_shape.Type == 1:  # TextBox tipo 1
                                texto = sub_shape.TextFrame2.TextRange.Text.strip()
                                if texto.startswith("SECTOR"):
                                    sectores.append(sub_shape)

                        # Si encontramos los 3 sectores
                        if len(sectores) == 3:
                            # Extraer número de antena del grupo (ej: "Group 10" -> antena 1)
                            try:
                                antena_num = int(re.search(r'\d+', shape.name).group()) % 10
                                if antena_num == 0:
                                    antena_num = 10
                            except:
                                continue

                            # Procesar cada sector
                            for i, sector_shape in enumerate(sectores, 1):
                                sector_folder = os.path.join(tss_instance.resultados_dir, f"Antena_{antena_num}")
                                patron = os.path.join(sector_folder, f"Antena_{antena_num}_Sector_{i}*.png")

                                # Extraer tecnologías de los archivos
                                tecnologias = set()
                                for img_path in glob.glob(patron):
                                    if '(' in img_path and ')' in img_path:
                                        tech_part = img_path.split('(')[1].split(')')[0]
                                        tecnologias.update(
                                            t.strip() for t in tech_part.replace('-', ',').split(',') if t.strip())

                                # Actualizar texto del sector
                                if tecnologias:
                                    texto_original = sector_shape.TextFrame2.TextRange.Text.strip()
                                    nuevo_texto = f"{texto_original}\n({' + '.join(sorted(tecnologias))})"
                                    sector_shape.TextFrame2.TextRange.Text = nuevo_texto
                                    print(f"✅ Antena {antena_num} - Sector {i}: {nuevo_texto}")

                    except Exception as e:
                        print(f"⚠️ Error procesando grupo {shape.name}: {str(e)}")

        except Exception as e:
            print(f"❌ Error general: {str(e)}")

# Ejemplo de uso
if __name__ == "__main__":
    analizar_grupos_formas("../SID MIC BO 3YPLAN 2024_ERICSON_NAME_ID.xlsx", hoja_num=9)