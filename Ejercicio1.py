import pandas as pd
import glob
import re
import os

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



def procesar_archivos_asistencia_cadena_cine(ruta_archivos, salida_txt):
    archivos = glob.glob(f"{ruta_archivos}/*.xlsx")

    if not archivos:
        print("No se encontraron archivos .xlsx en la carpeta.")
        return
    
    data = {}  # Diccionario { Título: { Cine: Asistencia } }
    total_asistencia = 0  # Total global de asistencia
    
    for archivo in archivos:
        df = pd.read_excel(archivo, usecols=[1, 5, 11], dtype=str, skiprows=2)  
        print(f"Leyendo archivo: {archivo}")
        
        df.dropna(inplace=True)  # Eliminar filas vacías
        
        for _, row in df.iterrows():
            titulo = row.iloc[0].strip()  # Columna 1 (Título de la película)
            cine = row.iloc[1].strip()  # Columna 5 (Cine)
            
            try:
                asistencia = int(row.iloc[2])  # Columna 11 (Asistencia)
            except ValueError:
                print(f"Valor inválido en archivo {archivo}: {row}")
                continue  # Omitir filas con valores no numéricos
            
            if titulo not in data:
                data[titulo] = {}

            if cine in data[titulo]:
                data[titulo][cine] += asistencia
            else:
                data[titulo][cine] = asistencia

    # Si no hay datos, se detiene el programa antes de continuar
    if not data:
        print("No se encontraron datos válidos en los archivos.")
        return

    # Ordenar los títulos por mayor asistencia total
    data_ordenada = sorted(data.items(), key=lambda item: sum(item[1].values()), reverse=True)

    # Calcular total general
    total_asistencia = sum(sum(cines.values()) for _, cines in data_ordenada)
    #print(f"Total Global de Asistencia: {total_asistencia:,}")

    with open(salida_txt, "w", encoding="utf-8") as f:
        f.write("\n--- REGIONAL: Asistencia por Película ---\n\n")

        for titulo, cines in data_ordenada:
            total_pelicula = sum(cines.values())  # Suma total de asistencia por película
            porcentaje_pelicula = (total_pelicula / total_asistencia) * 100  # Porcentaje del total global

            f.write(f"\n{titulo.ljust(40)} Asistencia: {f'{total_pelicula:,.0f}'.ljust(10)}   %Asistencia Por Cadena\n")
           # f.write(f"{titulo.ljust(40)} Asistencia: {total_pelicula:,.0f} ({porcentaje_pelicula:.2f}%)\n")
            f.write("---------------------------------------------------------------------------------------\n")

            # Ordenar cines por mayor asistencia
            cines_ordenados = sorted(cines.items(), key=lambda item: item[1], reverse=True)

            for cine, suma in cines_ordenados:
                porcentaje_cine = (suma / total_pelicula) * 100  # Porcentaje individual dentro de la película
                f.write(f"{cine.ljust(40)} {f'{suma:,.0f}'.rjust(15)} {f'{porcentaje_cine:.2f}%'.rjust(15)}\n")
                
            f.write("---------------------------------------------------------------------------------------\n")
            f.write(f"{'Total'.ljust(63)}{'100.00%'.rjust(10)}\n")

        f.write(f" **Total Global:** {total_asistencia:,} visualizaciones\n")
        f.write(f"\n")

    print(f"Archivo {salida_txt} generado exitosamente.")


def procesar_archivos_asistencia_estrenos(ruta_archivos, salida_txt):
    archivos = glob.glob(f"{ruta_archivos}/*.xlsx")
    
    if not archivos:
        print("No se encontraron archivos .xlsx en la carpeta.")
        return
    
    data = {}  # Diccionario { (Semana inicial, Año, Película, Semana Película): [Suma Col 17, Suma Col 21, Suma Col 25] }
    total_col_25 = {}  # Diccionario { Película: Total de Columna 25 }
    #restante = {}
    porcentajeSemana1=0
    porcentajeSemana2=0
    for archivo in archivos:
        nombre_archivo = os.path.basename(archivo)
        match = re.match(r"(\d+)-(\d+)", nombre_archivo)  # Extraer Semana y Año del nombre
        
        if not match:
            print(f"Nombre de archivo no coincide con el formato esperado: {nombre_archivo}")
            continue
        
        semana_inicial, anio = match.groups()
        semana_inicial, anio = int(semana_inicial), int(anio)
        
        df = pd.read_excel(archivo, usecols=[1, 3, 17, 21, 25], dtype=str, skiprows=2)
        df.dropna(inplace=True)
        
        for _, row in df.iterrows():
            pelicula = row.iloc[0].strip()
            semana_pelicula = row.iloc[1].strip()
            
            try:
                suma_col_17 = int(row.iloc[2].replace(",", "")) if row.iloc[2] else 0
                suma_col_21 = int(row.iloc[3].replace(",", "")) if row.iloc[3] else 0
                suma_col_25 = int(row.iloc[4].replace(",", "")) if row.iloc[4] else 0
            except ValueError:
                print(f"Valor inválido en archivo {archivo}: {row}")
                continue
            
            clave = (semana_inicial, anio, pelicula, semana_pelicula)  # Agrupar por Semana Inicial, Año, Película y Semana Película
            
            if clave not in data:
                data[clave] = [0, 0, 0]
            
            data[clave][0] += suma_col_17
            data[clave][1] += suma_col_21
            data[clave][2] += suma_col_25

            # Sumar columna 25 por película
            if pelicula not in total_col_25:
                total_col_25[pelicula] = 0
            total_col_25[pelicula] += suma_col_25  # 🔹 Aquí se suman correctamente TODOS los valores de col. 25 por película.

    with open(salida_txt, "w", encoding="utf-8") as f:
        f.write("--- Reporte de Asistencia por Semana Inicial y Año ---\n\n")

        # Encabezado de la tabla
        header = f"{'Película':<50}|{'Jueves día inicial':<20}|{'Fin de semana':<15}|{'Total Semana 1':<15}|{'% SEMANA 1':<15}"
        header += f"|{'Total Semana 2':<15}|{'% SEMANA 2':<15}|{'RESTANTE':<15}|{'TOTAL Recaudado':<15}|\n"
        f.write(header)
        f.write("=" * len(header) + "\n")

        peliculas_procesadas = set()

        # Ordenar por Semana Inicial y Año antes de escribir
        for (semana_inicial, anio, pelicula, semana_pelicula), sumas in sorted(data.items()):
            if (semana_inicial, anio, pelicula) in peliculas_procesadas:
                continue  # Evita repetir la película en múltiples filas
            
            semana1 = data.get((semana_inicial, anio, pelicula, "1"), [0, 0, 0])
            semana2 = data.get((semana_inicial, anio, pelicula, "2"), [0, 0, 0])

            total_pelicula = total_col_25.get(pelicula, 0)  # 🔹 Total real de la columna 25 por película
            restante = total_pelicula-(semana1[2]+semana2[2])
            porcentajeSemana1 = (semana1[2]/total_pelicula)*100
            porcentajeSemana2 = (semana2[2]/total_pelicula)*100
            linea = f"{pelicula:<50}|{semana1[0]:<20,}|{semana1[1]:<15,}|{semana1[2]:<15,}|{f'{porcentajeSemana1:.2f}%':<15}|"
            linea += f"{semana2[2]:<15,}|{f'{porcentajeSemana2:.2f}%':<15}|{restante:<15,}|{total_pelicula:<15,}|\n"

            
            f.write(linea)
            peliculas_procesadas.add((semana_inicial, anio, pelicula))
    
    print(f"Archivo {salida_txt} generado exitosamente.")

# Uso de la función
procesar_archivos_top_pelicula("C:\\Users\\Usuario\\Documents\\python ues\\Archivos prueba", "topPelicula.txt")
procesar_archivos_cadena_cine("C:\\Users\\Usuario\\Documents\\python ues\\Archivos prueba", "cadenaCine.txt")
procesar_archivos_asistencia_cadena_cine("C:\\Users\\Usuario\\Documents\\python ues\\Archivos prueba", "peliculaCadenaCine.txt")
procesar_archivos_asistencia_estrenos("C:\\Users\\Usuario\\Documents\\python ues\\Reportes", "asistenciaEstreno.txt")