import pandas as pd
import glob

def procesar_archivos_top_pelicula(ruta_archivos, salida_txt):
    archivos = glob.glob(f"{ruta_archivos}/*.xlsx")  # Buscar todos los archivos .xlsx
    
    if not archivos:
        print("No se encontraron archivos .xlsx en la carpeta.")
        return
    
    data = {}  # Diccionario para almacenar la suma total por cadena de cine
    total=0
    # Primera parte: Acumular los datos de todos los archivos
    for archivo in archivos:
        df = pd.read_excel(archivo, usecols=[1, 11], dtype=str)  # Columnas F (cine) y K (valor)
        df.dropna(inplace=True)  # Elimina filas vacías
        
        for _, row in df.iterrows():
            cine = row.iloc[0].strip()  # Nombre del cine sin espacios extras
            try:
                valor = int(row.iloc[1])  # Convertir a entero
            except ValueError:
                continue  # Omitir filas con valores inválidos
            
            if cine in data:
                data[cine] += valor  # Sumar el valor
            else:
                data[cine] = valor  # Inicializar

    # Ordenar los datos globales por mayor asistencia
    data_ordenada = sorted(data.items(), key=lambda item: item[1], reverse=True)
    
         # Calcular el total
    for cine, suma in data_ordenada:
        total += suma
    print(f"Total acumulado: {total}")

    # Definir anchos de columna
    max_cine_length = max(len(cine) for cine in data_ordenada)  # Ancho dinámico basado en el nombre más largo
    ancho_cine = max(max_cine_length, 50)  # Mínimo 40 caracteres
    ancho_asistencia = 20  # Espacio para el número alineado a la derecha
    ancho_porcentaje = 15  # Espacio para el porcentaje alineado a la derecha
    
    # Escribir la primera parte en el archivo de salida
    with open(salida_txt, "w", encoding="utf-8") as f:
        f.write("\n--- REGIONAL: Resultado TOP cine ---\n")
        f.write(f"{'Funcion'.ljust(ancho_cine)} {'Asistencia'.rjust(ancho_asistencia)} {'% Total'.rjust(ancho_porcentaje)}\n")
        f.write(f"{'-' * ancho_cine} {'-' * ancho_asistencia} {'-' * ancho_porcentaje}\n")
        # Escribir los primeros 5 cines con su asistencia y el porcentaje
        for cine, suma in data_ordenada[:5]:
            porcentaje = (suma / total) * 100  # Calcular el porcentaje
            f.write(f"{cine.ljust(ancho_cine)} {f'{suma:,.0f}'.rjust(ancho_asistencia)} {f'{porcentaje:.4f}%'.rjust(ancho_porcentaje)}\n")  # Mostrar hasta 4 decimales
        # Escribir el total acumulado
        f.write("---------------------------------------------------------------------------------------\n")
        f.write(f"{'Total'.ljust(ancho_cine)} {f'{total:,.0f}'.rjust(ancho_asistencia)} {'100.00%'.rjust(ancho_porcentaje)}\n")
    #segunda parte
    with open(salida_txt, "a", encoding="utf-8") as f:  # Abrir el archivo de salida para escribir todo
        for archivo in archivos:
            print(f"Procesando archivo: {archivo}")  # Verificar cuál archivo se está procesando
            df = pd.read_excel(archivo, usecols=[1, 11], dtype=str, skiprows=2)  # Lee columnas B y K (11)
            
            df.dropna(inplace=True)  # Elimina filas con valores nulos
            
            archivo_data = {}  # Almacenar los resultados de este archivo de forma separada
            
            for _, row in df.iterrows():
                clave = row.iloc[0]  # Columna 1 (nombre de la película)
                try:
                    valor = int(row.iloc[1])  # Columna 11 (valor a sumar)
                except ValueError:
                    continue  # Si no se puede convertir a número, se omite la fila
                
                if clave in archivo_data:
                    archivo_data[clave] += valor  # Sumar los valores
                else:
                    archivo_data[clave] = valor  # Inicializar la suma para esa película

            # Ordenar los resultados por la suma de manera descendente (de mayor a menor)
            archivo_data_ordenado = sorted(archivo_data.items(), key=lambda item: item[1], reverse=True)

             # Calcular el total
            total=0 #limpiamos
            for clave, suma2 in archivo_data_ordenado:
                total += suma2
            print(f"Total acumulado: {total}")
            # Definir anchos de columna
            max_cine_length = max(len(clave) for clave in archivo_data_ordenado)  # Ancho dinámico basado en el nombre más largo
            ancho_cine = max(max_cine_length, 50)  # Mínimo 30 caracteres
            ancho_asistencia = 20  # Espacio para el número alineado a la derecha
            ancho_porcentaje = 15
            
            # Escribir la primera parte en el archivo de salida
            with open(salida_txt, "a", encoding="utf-8") as f:
                f.write(f"\n--- Top pelicula de: {archivo} ---\n")
                f.write(f"{'Funcion'.ljust(ancho_cine)} {'Asistencia'.rjust(ancho_asistencia)} {'% Total'.rjust(ancho_porcentaje)}\n")
                f.write(f"{'-' * ancho_cine} {'-' * ancho_asistencia} {'-' * ancho_porcentaje}\n")
                for clave, suma in archivo_data_ordenado[:5]:
                    porcentaje = (suma / total) * 100  # Calcular el porcentaje
                    f.write(f"{clave.ljust(ancho_cine)} {f'{suma:,.0f}'.rjust(ancho_asistencia)} {f'{porcentaje:.4f}%'.rjust(ancho_porcentaje)}\n")  # Mostrar hasta 4 decimales
            # Escribir el total acumulado
                f.write("---------------------------------------------------------------------------------------\n")
                f.write(f"{'Total'.ljust(ancho_cine)} {f'{total:,.0f}'.rjust(ancho_asistencia)} {'100.00%'.rjust(ancho_porcentaje)}\n")
    print(f"Archivo {salida_txt} generado exitosamente.")


def procesar_archivos_cadena_cine(ruta_archivos, salida_txt):
    archivos = glob.glob(f"{ruta_archivos}/*.xlsx")  # Buscar todos los archivos .xlsx
    
    if not archivos:
        print("No se encontraron archivos .xlsx en la carpeta.")
        return
    
    data = {}  # Diccionario para almacenar la suma total por cadena de cine
    total=0
    # Primera parte: Acumular los datos de todos los archivos
    for archivo in archivos:
        df = pd.read_excel(archivo, usecols=[5, 11], dtype=str)  # Columnas F (cine) y K (valor)
        df.dropna(inplace=True)  # Elimina filas vacías
        
        for _, row in df.iterrows():
            cine = row.iloc[0].strip()  # Nombre del cine sin espacios extras
            try:
                valor = int(row.iloc[1])  # Convertir a entero
            except ValueError:
                continue  # Omitir filas con valores inválidos
            
            if cine in data:
                data[cine] += valor  # Sumar el valor
            else:
                data[cine] = valor  # Inicializar

    # Ordenar los datos globales por mayor asistencia
    data_ordenada = sorted(data.items(), key=lambda item: item[1], reverse=True)

    for cine, suma in data_ordenada:
        total += suma
    print(f"Total acumulado: {total}")

    # Definir anchos de columna
    max_cine_length = max(len(cine) for cine in data_ordenada)  # Ancho dinámico basado en el nombre más largo
    ancho_cine = max(max_cine_length, 50)  # Mínimo 30 caracteres
    ancho_asistencia = 20  # Espacio para el número alineado a la derecha
    ancho_porcentaje = 15  # Espacio para el porcentaje alineado a la derecha

    # Escribir la primera parte en el archivo de salida
    with open(salida_txt, "w", encoding="utf-8") as f:
        f.write("\n--- REGIONAL: Resultado por cadena de cine ---\n")
        f.write(f"{'Cadena'.ljust(ancho_cine)} {'Asistencia'.rjust(ancho_asistencia)} {'% Visita Total'.rjust(ancho_porcentaje)}\n")
        f.write(f"{'-' * ancho_cine} {'-' * ancho_asistencia} {'-' * ancho_porcentaje}\n")

        for cine, suma in data_ordenada:
            porcentaje = (suma / total) * 100  # Calcular el porcentaje
            f.write(f"{cine.ljust(ancho_cine)} {f'{suma:,.0f}'.rjust(ancho_asistencia)} {f'{porcentaje:.4f}%'.rjust(ancho_porcentaje)}\n")  # Mostrar hasta 4 decimales
        f.write("---------------------------------------------------------------------------------------\n")
        f.write(f"{'Total'.ljust(ancho_cine)} {f'{total:,.0f}'.rjust(ancho_asistencia)} {'100.00%'.rjust(ancho_porcentaje)}\n")

    # Segunda parte: Procesar cada archivo individualmente
    with open(salida_txt, "a", encoding="utf-8") as f:  # "a" para agregar datos sin sobrescribir
        for archivo in archivos:
            df = pd.read_excel(archivo, usecols=[5, 11], dtype=str, skiprows=2)  
            df.dropna(inplace=True)

            archivo_data = {}  # Diccionario para este archivo
            
            for _, row in df.iterrows():
                cine = row.iloc[0].strip()
                try:
                    valor = int(row.iloc[1])  
                except ValueError:
                    continue  
                
                if cine in archivo_data:
                    archivo_data[cine] += valor  
                else:
                    archivo_data[cine] = valor  

            # Ordenar los datos del archivo individual
            archivo_data_ordenado = sorted(archivo_data.items(), key=lambda item: item[1], reverse=True)
             # Calcular el total
            total=0 #limpiamos
            for cine, suma2 in archivo_data_ordenado:
                total += suma2
            print(f"Total acumulado: {total}")

            # Escribir los resultados de este archivo
            f.write(f"\n--- Resultado de cadena de cine de: {archivo} ---\n")
            f.write(f"{'Cadena'.ljust(ancho_cine)} {'Asistencia'.rjust(ancho_asistencia)} {'% Visita Total'.rjust(ancho_porcentaje)}\n")
            f.write(f"{'-' * ancho_cine} {'-' * ancho_asistencia} {'-' * ancho_porcentaje}\n")
            for cine, suma in archivo_data_ordenado:
                porcentaje = (suma / total) * 100  # Calcular el porcentaje
                f.write(f"{cine.ljust(ancho_cine)} {f'{suma:,.0f}'.rjust(ancho_asistencia)} {f'{porcentaje:.4f}%'.rjust(ancho_porcentaje)}\n")  # Mostrar hasta 4 decimales
            # Escribir el total acumulado
            f.write("---------------------------------------------------------------------------------------\n")
            f.write(f"{'Total'.ljust(ancho_cine)} {f'{total:,.0f}'.rjust(ancho_asistencia)} {'100.00%'.rjust(ancho_porcentaje)}\n")
    print(f"Archivo {salida_txt} generado exitosamente.")

# Uso de la función
procesar_archivos_top_pelicula("C:\\Users\\Usuario\\Documents\\python ues\\Archivos prueba", "topPelicula.txt")
procesar_archivos_cadena_cine("C:\\Users\\Usuario\\Documents\\python ues\\Archivos prueba", "cadenaCine.txt")
