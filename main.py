import pandas as pd
from tkinter import Tk, filedialog
import os

# Función para cargar el archivo Excel
def cargar_archivo_excel():
    Tk().withdraw()  # Ocultar la ventana principal de Tkinter
    archivo_excel = filedialog.askopenfilename(
        title="Selecciona el archivo .xlsx",
        filetypes=[("Archivos Excel con macros", "*.xlsx")]
    )
    
    if archivo_excel:
        # Leer el archivo Excel sin asumir encabezados
        df = pd.read_excel(archivo_excel, header=None, engine='openpyxl')
        return df, archivo_excel
    else:
        print("No se seleccionó ningún archivo.")
        return None, None

# Función para reorganizar las columnas
def reorganizar_columnas(df):
    if 0 in df.columns:
        # Separar la columna en múltiples columnas utilizando la coma como delimitador
        df_reorganizado = df[0].str.split(',', expand=True)

        # Verificar el número de columnas en el DataFrame
        num_columnas = df_reorganizado.shape[1]
        print(f"Número de columnas en el DataFrame: {num_columnas}")

        # Definir nombres de columnas esperados
        nombres_columnas = [
            "ID", "IP", "Fecha", "Comba", "Método", "Data-type", "URL",
            "HTTP", "Err", "Navegador", "Acción"
        ]

        # Asignar nombres de columnas solo si el número de columnas coincide
        if num_columnas <= len(nombres_columnas):
            df_reorganizado.columns = nombres_columnas[:num_columnas]
        else:
            print(f"Advertencia: El número de columnas ({num_columnas}) es mayor que el esperado ({len(nombres_columnas)}).")
            # Crear nombres de columna genéricos adicionales
            df_reorganizado.columns = nombres_columnas + [f"Columna_{i}" for i in range(num_columnas - len(nombres_columnas))]
        
        return df_reorganizado
    else:
        print("El archivo no tiene datos en la primera columna.")
        return None



# Función para eliminar columnas no deseadas
def eliminar_columnas_no_deseadas(df_reorganizado):
    if df_reorganizado is not None:
        # Eliminar columnas: 'Comba', 'Data-type', y 'Err'
        columnas_a_eliminar = ['Comba', 'Data-type', 'Err']
        df_limpio = df_reorganizado.drop(columns=columnas_a_eliminar, errors='ignore')

        # Renombrar la columna 'Acción' a 'Response'
        df_limpio = df_limpio.rename(columns={'Acción': 'Response'})
        
        return df_limpio
    else:
        print("No se puede eliminar columnas, el DataFrame está vacío o no es válido.")
        return None

# Función para guardar el archivo reorganizado
def guardar_archivo(df_reorganizado, ruta_original):
    if df_reorganizado is not None:
        # Crear una nueva ruta para guardar el archivo
        nueva_ruta = os.path.splitext(ruta_original)[0] + '_reorganizado.xlsx'
        
        # Guardar el archivo reorganizado
        df_reorganizado.to_excel(nueva_ruta, index=False, engine='openpyxl')
        print(f"Archivo reorganizado guardado como: {nueva_ruta}")
        return nueva_ruta
    else:
        print("No se puede guardar el archivo, el DataFrame está vacío o no es válido.")
        return None

# Función principal del sistema
def sistema():
    # Cargar el archivo Excel
    df, archivo_excel = cargar_archivo_excel()
    
    if df is not None:
        # Reorganizar las columnas
        df_reorganizado = reorganizar_columnas(df)
        
        if df_reorganizado is not None:
            # Eliminar columnas no deseadas
            df_limpio = eliminar_columnas_no_deseadas(df_reorganizado)
            
            if df_limpio is not None:
                # Mostrar los datos reorganizados sin las columnas eliminadas
                print("Datos reorganizados y limpios:")
                print(df_limpio)
                
                # Guardar el archivo reorganizado
                ruta_guardada = guardar_archivo(df_limpio, archivo_excel)
                
                return ruta_guardada

# Ejecutar el sistema
if __name__ == "__main__":
    sistema()
