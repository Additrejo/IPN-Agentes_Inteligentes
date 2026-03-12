from pathlib import Path
import pandas as pd
import re
from collections import Counter
import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser
import os

class VentanaResultados:
    def __init__(self, estadisticas, top_positivas, top_negativas, ruta_reporte):
        self.ventana = tk.Tk()
        self.ventana.title("Resultados del Analisis de Sentimientos")
        self.ventana.geometry("800x600")
        
        # Configurar estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        # Frame principal
        main_frame = ttk.Frame(self.ventana)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Titulo
        titulo = ttk.Label(main_frame, text="ANALISIS DE SENTIMIENTOS COMPLETADO", 
                          font=('Arial', 16, 'bold'))
        titulo.pack(pady=10)
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Estadisticas
        stats_frame = ttk.LabelFrame(main_frame, text="Estadisticas Generales", padding=10)
        stats_frame.pack(fill='x', pady=5)
        
        for key, value in estadisticas.items():
            frame = ttk.Frame(stats_frame)
            frame.pack(fill='x', pady=2)
            ttk.Label(frame, text=f"{key}:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
            ttk.Label(frame, text=value, font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
        
        # Frame para palabras
        palabras_frame = ttk.Frame(main_frame)
        palabras_frame.pack(fill='both', expand=True, pady=10)
        
        # Positivas
        pos_frame = ttk.LabelFrame(palabras_frame, text="Top 10 Palabras Positivas", padding=10)
        pos_frame.pack(side=tk.LEFT, fill='both', expand=True, padx=5)
        
        pos_text = tk.Text(pos_frame, height=12, width=30)
        pos_text.pack(fill='both', expand=True)
        
        for palabra, conteo in top_positivas:
            pos_text.insert(tk.END, f"- {palabra}: {conteo} veces\n")
        pos_text.config(state=tk.DISABLED)
        
        # Negativas
        neg_frame = ttk.LabelFrame(palabras_frame, text="Top 10 Palabras Negativas", padding=10)
        neg_frame.pack(side=tk.RIGHT, fill='both', expand=True, padx=5)
        
        neg_text = tk.Text(neg_frame, height=12, width=30)
        neg_text.pack(fill='both', expand=True)
        
        for palabra, conteo in top_negativas:
            neg_text.insert(tk.END, f"- {palabra}: {conteo} veces\n")
        neg_text.config(state=tk.DISABLED)
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Botones
        botones_frame = ttk.Frame(main_frame)
        botones_frame.pack(fill='x', pady=10)
        
        ttk.Button(botones_frame, text="Abrir Archivo", 
                  command=lambda: self.abrir_archivo(ruta_reporte)).pack(side=tk.LEFT, padx=5)
        ttk.Button(botones_frame, text="Cerrar", 
                  command=self.ventana.destroy).pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(main_frame, text=f"Archivo guardado en: {ruta_reporte}", 
                 font=('Arial', 8), wraplength=700).pack(pady=5)
    
    def abrir_archivo(self, ruta):
        try:
            webbrowser.open(str(ruta))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")
    
    def mostrar(self):
        self.ventana.mainloop()


class AnalizadorSentimientos:
    def __init__(self, ruta_base):
        self.ruta_base = Path(ruta_base)
        self.ruta_palabras = self.ruta_base / "Palabras de entrenamiento"
        
        self.archivo_positivas = self.ruta_palabras / "palabras-positivas.txt"
        self.archivo_negativas = self.ruta_palabras / "palabras-negativas.txt"
        self.archivo_excel = self.buscar_archivo_excel()
        
        # --- CORRECCION DEFINITIVA PARA EL ESCRITORIO ---
        # Opcion 1: Usar la variable de entorno USERPROFILE
        perfil_usuario = os.environ.get('USERPROFILE', '')
        self.escritorio = Path(perfil_usuario) / "Desktop"
        
        # Opcion 2: Si la anterior falla, usar una ruta fija
        if not self.escritorio.exists():
            self.escritorio = Path("C:/Users/addi_/Desktop")
        
        self.archivo_salida = self.escritorio / "Resultados Analisis Sentimientos.xlsx"
        # ------------------------------------------------
        
        print("\n" + "="*60)
        print("VERIFICANDO ARCHIVOS")
        print("="*60)
        print(f"Escritorio: {self.escritorio}")
        print(f"Archivo salida: {self.archivo_salida}")
        
        self.palabras_positivas = self.cargar_palabras(self.archivo_positivas)
        self.palabras_negativas = self.cargar_palabras(self.archivo_negativas)
    
    def buscar_archivo_excel(self):
        """Busca el archivo Excel"""
        archivo_esperado = self.ruta_base / "Datos opiniones Facebook.xlsx"
        if archivo_esperado.exists():
            print(f"Excel encontrado: {archivo_esperado.name}")
            return archivo_esperado
        
        # Buscar cualquier Excel
        archivos = list(self.ruta_base.glob("*.xlsx"))
        if archivos:
            print(f"Usando: {archivos[0].name}")
            return archivos[0]
        
        return None
    
    def cargar_palabras(self, ruta_archivo):
        """Carga las palabras"""
        try:
            if not ruta_archivo.exists():
                print(f"Error: No se encuentra {ruta_archivo.name}")
                return []
            
            codificaciones = ['utf-8', 'latin-1', 'cp1252']
            for cod in codificaciones:
                try:
                    with open(ruta_archivo, 'r', encoding=cod) as file:
                        palabras = [linea.strip().lower() for linea in file if linea.strip()]
                    print(f"Cargadas {len(palabras)} palabras de {ruta_archivo.name}")
                    return palabras
                except UnicodeDecodeError:
                    continue
            
            print(f"Error: No se pudo leer {ruta_archivo.name}")
            return []
        except Exception as e:
            print(f"Error: {e}")
            return []
    
    def limpiar_texto(self, texto):
        """Limpia el texto"""
        if pd.isna(texto):
            return ""
        texto = str(texto).lower()
        texto = re.sub(r'[^a-záéíóúüñ0-9\s]', ' ', texto)
        texto = re.sub(r'\s+', ' ', texto).strip()
        return texto
    
    def analizar_comentario(self, comentario):
        """Analiza un comentario"""
        texto = self.limpiar_texto(comentario)
        palabras = texto.split()
        
        pos = 0
        neg = 0
        palabras_pos = []
        palabras_neg = []
        
        for palabra in palabras:
            if palabra in self.palabras_positivas:
                pos += 1
                palabras_pos.append(palabra)
            elif palabra in self.palabras_negativas:
                neg += 1
                palabras_neg.append(palabra)
        
        if pos > neg:
            sentimiento = "Positivo"
        elif neg > pos:
            sentimiento = "Negativo"
        else:
            sentimiento = "Neutral"
        
        return {
            'sentimiento': sentimiento,
            'palabras_pos': ', '.join(palabras_pos),
            'palabras_neg': ', '.join(palabras_neg),
            'conteo_pos': pos,
            'conteo_neg': neg
        }
    
    def analizar_dataframe(self, df, columna):
        """Analiza el dataframe"""
        resultados = []
        resumen_pos = Counter()
        resumen_neg = Counter()
        
        total = len(df)
        print(f"\nAnalizando {total} comentarios...")
        
        for idx, row in df.iterrows():
            analisis = self.analizar_comentario(row[columna])
            
            if analisis['palabras_pos']:
                for p in analisis['palabras_pos'].split(', '):
                    if p:
                        resumen_pos[p] += 1
            
            if analisis['palabras_neg']:
                for p in analisis['palabras_neg'].split(', '):
                    if p:
                        resumen_neg[p] += 1
            
            resultados.append({
                'sentimiento': analisis['sentimiento'],
                'palabras_positivas': analisis['palabras_pos'],
                'palabras_negativas': analisis['palabras_neg'],
                'conteo_positivo': analisis['conteo_pos'],
                'conteo_negativo': analisis['conteo_neg']
            })
            
            if (idx + 1) % 50 == 0:
                print(f"   Procesados {idx + 1}/{total}")
        
        return resultados, resumen_pos, resumen_neg
    
    def ejecutar(self):
        """Ejecuta el analisis"""
        
        if not self.archivo_excel:
            messagebox.showerror("Error", "No se encontro el archivo Excel")
            return
        
        if not self.palabras_positivas or not self.palabras_negativas:
            messagebox.showerror("Error", "No se pudieron cargar las palabras necesarias")
            return
        
        try:
            # Cargar Excel
            df = pd.read_excel(self.archivo_excel)
            print(f"\nColumnas: {list(df.columns)}")
            print(f"Filas: {len(df)}")
            
            # Encontrar columna de comentarios
            columna = None
            for col in df.columns:
                if 'comentario' in str(col).lower():
                    columna = col
                    break
            
            if not columna:
                messagebox.showerror("Error", "No se encontro columna de comentarios")
                return
            
            print(f"Usando columna: {columna}")
            
            # Analizar
            resultados, resumen_pos, resumen_neg = self.analizar_dataframe(df, columna)
            
            # Crear DataFrame final
            df_resultados = pd.DataFrame(resultados)
            df_final = pd.concat([df, df_resultados], axis=1)
            
            # Verificar que el directorio existe antes de guardar
            if not self.escritorio.exists():
                print(f"Creando directorio: {self.escritorio}")
                self.escritorio.mkdir(parents=True, exist_ok=True)
            
            # Guardar
            print(f"Guardando en: {self.archivo_salida}")
            df_final.to_excel(self.archivo_salida, index=False)
            print(f"Reporte guardado exitosamente")
            
            # Preparar estadisticas
            conteo = df_final['sentimiento'].value_counts()
            total = len(resultados)
            
            estadisticas = {
                "Total comentarios": str(total),
                "Comentarios Positivos": f"{conteo.get('Positivo', 0)} ({conteo.get('Positivo', 0)/total*100:.1f}%)",
                "Comentarios Negativos": f"{conteo.get('Negativo', 0)} ({conteo.get('Negativo', 0)/total*100:.1f}%)",
                "Comentarios Neutrales": f"{conteo.get('Neutral', 0)} ({conteo.get('Neutral', 0)/total*100:.1f}%)",
                "Palabras positivas unicas": str(len(resumen_pos)),
                "Palabras negativas unicas": str(len(resumen_neg))
            }
            
            # Top 10
            top_pos = resumen_pos.most_common(10)
            top_neg = resumen_neg.most_common(10)
            
            # Mostrar ventana
            ventana = VentanaResultados(estadisticas, top_pos, top_neg, self.archivo_salida)
            ventana.mostrar()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error durante el analisis: {e}")
            import traceback
            traceback.print_exc()


def main():
    print("\n" + "="*50)
    print("ANALIZADOR DE SENTIMIENTOS")
    print("="*50)
    
    ruta_base = r"C:\Users\addi_\OneDrive\Documentos\IPN\9no Semestre\Agentes inteligentes expertos\1er. Bloque\Analizador de sentimientos\IPN-Agentes_Inteligentes"
    
    analizador = AnalizadorSentimientos(ruta_base)
    analizador.ejecutar()


if __name__ == "__main__":
    main()