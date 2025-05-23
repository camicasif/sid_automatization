import psutil
import os


def cerrar_archivo_excel_bloqueado(nombre_archivo):
    # Obtener la ruta absoluta del archivo (opcional, pero útil para comparaciones)
    ruta_archivo = os.path.abspath(nombre_archivo)

    for proceso in psutil.process_iter(['pid', 'name', 'open_files']):
        try:
            # Verificar si el proceso tiene archivos abiertos
            if proceso.info['open_files'] is not None:
                for archivo in proceso.info['open_files']:
                    # Comprobar si el archivo bloqueado coincide (ignorando el prefijo '~$')
                    if (nombre_archivo in archivo.path or
                            os.path.basename(ruta_archivo) in archivo.path or
                            archivo.path.endswith(nombre_archivo.replace('~$', ''))):
                        print(
                            f"Cerrando proceso {proceso.info['name']} (PID: {proceso.info['pid']}) que bloquea el archivo.")
                        proceso.terminate()  # Terminar el proceso
                        return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    print("No se encontró ningún proceso bloqueando el archivo.")
    return False


# Nombre del archivo bloqueado (puedes ajustarlo)
nombre_archivo = "~$SID MIC BO 3YPLAN 2024_Name_ID_RevP.xlsx"
cerrar_archivo_excel_bloqueado(nombre_archivo)