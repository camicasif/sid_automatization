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


# Ejemplo de uso
if __name__ == "__main__":
    analizar_grupos_formas("SID MIC BO 3YPLAN 2024_Name_ID_RevP.xlsx", hoja_num=10)