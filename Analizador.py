import pandas as pd
import re
from collections import Counter
import os

class AnalizadorSentimientos:
    def __init__(self, ruta_positivas, ruta_negativas):
        """
        Inicializa el analizador cargando las listas de palabras
        """
        self.palabras_positivas = self.cargar_palabras(ruta_positivas)
        self.palabras_negativas = self.cargar_palabras(ruta_negativas)
        
    def cargar_palabras(self, ruta_archivo):
        """
        Carga las palabras desde un archivo de texto
        """
        try:
            with open(ruta_archivo, 'r', encoding='utf-8') as file:
                # Limpiar espacios y convertir a minúsculas
                palabras = [linea.strip().lower() for linea in file if linea.strip()]
            return palabras
        except FileNotFoundError:
            print(f"Error: No se encontró el archivo {ruta_archivo}")
            return []
    
    def limpiar_texto(self, texto):
        """
        Limpia y normaliza el texto para análisis
        """
        if pd.isna(texto):
            return ""
        
        # Convertir a minúsculas
        texto = str(texto).lower()
        
        # Eliminar caracteres especiales y números, mantener letras y espacios
        texto = re.sub(r'[^a-záéíóúüñ\s]', ' ', texto)
        
        # Eliminar espacios múltiples
        texto = re.sub(r'\s+', ' ', texto).strip()
        
        return texto
    
    def tokenizar(self, texto):
        """
        Divide el texto en palabras individuales (tokens)
        """
        return texto.split()
    
    def analizar_comentario(self, comentario):
        """
        Analiza un comentario individual
        """
        # Limpiar y tokenizar
        texto_limpio = self.limpiar_texto(comentario)
        palabras = self.tokenizar(texto_limpio)
        
        # Contar palabras positivas y negativas
        conteo_positivo = 0
        conteo_negativo = 0
        palabras_encontradas = {
            'positivas': [],
            'negativas': []
        }
        
        for palabra in palabras:
            if palabra in self.palabras_positivas:
                conteo_positivo += 1
                palabras_encontradas['positivas'].append(palabra)
            elif palabra in self.palabras_negativas:
                conteo_negativo += 1
                palabras_encontradas['negativas'].append(palabra)
        
        # Determinar sentimiento
        if conteo_positivo > conteo_negativo:
            sentimiento = "Positivo"
        elif conteo_negativo > conteo_positivo:
            sentimiento = "Negativo"
        else:
            sentimiento = "Neutral"
        
        return {
            'sentimiento': sentimiento,
            'conteo_positivo': conteo_positivo,
            'conteo_negativo': conteo_negativo,
            'palabras_encontradas': palabras_encontradas
        }
    
    def analizar_dataframe(self, df, columna_comentarios):
        """
        Analiza todos los comentarios en un DataFrame
        """
        resultados = []
        resumen_palabras = {
            'positivas': Counter(),
            'negativas': Counter()
        }
        
        for idx, row in df.iterrows():
            comentario = row[columna_comentarios]
            analisis = self.analizar_comentario(comentario)
            
            # Actualizar resumen de palabras
            for palabra in analisis['palabras_encontradas']['positivas']:
                resumen_palabras['positivas'][palabra] += 1
            for palabra in analisis['palabras_encontradas']['negativas']:
                resumen_palabras['negativas'][palabra] += 1
            
            # Agregar resultados
            resultados.append({
                'indice': idx,
                'comentario_original': comentario,
                'sentimiento': analisis['sentimiento'],
                'palabras_positivas': ', '.join(analisis['palabras_encontradas']['positivas']),
                'palabras_negativas': ', '.join(analisis['palabras_encontradas']['negativas']),
                'conteo_positivo': analisis['conteo_positivo'],
                'conteo_negativo': analisis['conteo_negativo']
            })
        
        return resultados, resumen_palabras
    
    def generar_reporte(self, df_original, resultados, resumen_palabras, archivo_salida):
        """
        Genera un reporte completo con los resultados
        """
        # Crear DataFrame con resultados
        df_resultados = pd.DataFrame(resultados)
        
        # Combinar con datos originales si es necesario
        df_final = pd.concat([df_original, df_resultados.iloc[:, 1:]], axis=1)
        
        # Guardar resultados detallados
        df_final.to_excel(archivo_salida, index=False)
        
        # Mostrar estadísticas en consola
        print("\n=== ESTADÍSTICAS GENERALES ===")
        print(f"Total comentarios analizados: {len(resultados)}")
        
        conteo_sentimientos = df_final['sentimiento'].value_counts()
        for sentimiento, conteo in conteo_sentimientos.items():
            porcentaje = (conteo / len(resultados)) * 100
            print(f"{sentimiento}: {conteo} ({porcentaje:.1f}%)")
        
        print("\n=== PALABRAS MÁS FRECUENTES ===")
        print("\nTop 5 palabras positivas:")
        for palabra, conteo in resumen_palabras['positivas'].most_common(5):
            print(f"  - {palabra}: {conteo} veces")
        
        print("\nTop 5 palabras negativas:")
        for palabra, conteo in resumen_palabras['negativas'].most_common(5):
            print(f"  - {palabra}: {conteo} veces")
        
        print(f"\n✅ Reporte guardado en: {archivo_salida}")
        
        return df_final

def main():
    """
    Función principal del programa
    """
    # Configurar rutas de archivos
    ruta_positivas = "datos/palabras_positivas.txt"
    ruta_negativas = "datos/palabras_negativas.txt"
    ruta_excel = "datos/comentarios.xlsx"
    ruta_salida = "datos/resultados_analisis.xlsx"
    
    # Verificar que existen los archivos
    if not os.path.exists(ruta_excel):
        print(f"Error: No se encuentra el archivo {ruta_excel}")
        print("Creando archivo de ejemplo...")
        
        # Crear archivo de ejemplo con los datos proporcionados
        datos_ejemplo = {
            'Fecha de publicación': ['04/10/2021', '05/10/2021', '06/10/2021'],
            'Autor': ['Eli Rocha', 'Ale Avila', 'Cristian Moya'],
            'Comentario': [
                'Ultimamente que he subido fotos de estado y agrego una musica de la misma historia, la foto original al publicarse sale super borrosa pierde su calidad, ya hecho reclamos en la opcion ayuda pero no me dan solucion',
                'Desde hace tiempo no me deja iniciar sesion, me aparece un mensaje que dice: Se produjo un error inesperado. Intenta iniciar sesion de nuevo.',
                'cada vez que veo un video en la seccion de videos, siempre se queda congelado. sigue el audio escuchandose, pero el video se congela y es muy cansado y fastidioso'
            ],
            'Número de estrellas': [1, 1, 1],
            'Nombre de la aplicación': ['Facebook', 'Facebook', 'Facebook']
        }
        
        df_ejemplo = pd.DataFrame(datos_ejemplo)
        df_ejemplo.to_excel(ruta_excel, index=False)
        print(f"✅ Archivo de ejemplo creado en: {ruta_excel}")
    
    # Cargar datos
    print("Cargando datos...")
    df = pd.read_excel(ruta_excel)
    
    # Crear analizador
    analizador = AnalizadorSentimientos(ruta_positivas, ruta_negativas)
    
    # Verificar que se cargaron las palabras
    print(f"Palabras positivas cargadas: {len(analizador.palabras_positivas)}")
    print(f"Palabras negativas cargadas: {len(analizador.palabras_negativas)}")
    
    # Analizar comentarios
    print("\nAnalizando comentarios...")
    resultados, resumen_palabras = analizador.analizar_dataframe(df, 'Comentario')
    
    # Generar reporte
    analizador.generar_reporte(df, resultados, resumen_palabras, ruta_salida)

if __name__ == "__main__":
    main()