import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from reportlab.pdfgen import canvas
import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from datetime import datetime

def crear_pdf_articulo(nombre_archivo, datos_por_articulo, carpeta_destino="pdfs"):
    # Crear la carpeta si no existe
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)
    
    # Ruta completa del archivo
    ruta_completa = os.path.join(carpeta_destino, f"{nombre_archivo}.pdf")
    
    # Crear el PDF
    c = canvas.Canvas(ruta_completa, pagesize=letter)
    width, height = letter
    
    # Configurar posiciones iniciales
    y_position = height - 50
    line_height = 20
    
    # Título principal
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y_position, f"REPORTE DE ARTÍCULO: {nombre_archivo}")
    y_position -= 30
    
    # Fecha y hora de generación
    c.setFont("Helvetica", 10)
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    c.drawString(50, y_position, f"Generado el: {fecha_actual}")
    y_position -= 40
    
    # Verificar si hay datos
    if not datos_por_articulo:
        c.setFont("Helvetica", 12)
        c.drawString(50, y_position, "No se encontraron datos para este artículo.")
    else:
        # Información del artículo (tomar datos del primer resultado)
        primer_resultado = datos_por_articulo[0]
        articulo_codigo = primer_resultado[0]
        descripcion = primer_resultado[1]
        
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, y_position, f"Artículo: {articulo_codigo}")
        y_position -= line_height
        
        c.setFont("Helvetica", 12)
        c.drawString(50, y_position, f"Descripción: {descripcion}")
        y_position -= 30
        
        # Encabezados de la tabla
        c.setFont("Helvetica-Bold", 11)
        c.drawString(50, y_position, "UBICACIONES Y STOCK:")
        y_position -= 25
        
        # Headers de la tabla
        c.setFont("Helvetica-Bold", 10)
        c.drawString(50, y_position, "Localizador")
        c.drawString(200, y_position, "En Mano")
        c.drawString(300, y_position, "LPN")
        y_position -= 5
        
        # Línea separadora
        c.line(50, y_position, 550, y_position)
        y_position -= 15
        
        # Datos de cada ubicación
        c.setFont("Helvetica", 10)
        total_stock = 0
        
        for i, (art, desc, en_mano, localizador, lpn) in enumerate(datos_por_articulo):
            # Verificar si necesitamos una nueva página
            if y_position < 100:
                c.showPage()
                y_position = height - 50
                c.setFont("Helvetica", 10)
            
            # Convertir en_mano a número para sumar
            try:
                stock_numerico = float(en_mano) if en_mano != '' else 0
                total_stock += stock_numerico
                stock_texto = str(en_mano) if en_mano != '' else '0'
            except (ValueError, TypeError):
                stock_texto = str(en_mano)

            # Manejar el LPN (evitar mostrar nan)
            if pd.isna(lpn) or str(lpn).lower() == 'nan':
                lpn_texto = "-"
            else:
                lpn_texto = str(lpn)
            
            # Escribir los datos
            c.drawString(50, y_position, str(localizador))
            c.drawString(200, y_position, stock_texto)
            c.drawString(300, y_position, lpn_texto)

            y_position -= line_height
        
        # Línea separadora antes del total
        y_position -= 10
        c.line(50, y_position, 550, y_position)
        y_position -= 20
        
        # Total de stock
        c.setFont("Helvetica-Bold", 11)
        c.drawString(50, y_position, f"TOTAL EN STOCK: {total_stock}")
        y_position -= 20
        
        # Resumen
        c.setFont("Helvetica", 10)
        c.drawString(50, y_position, f"Total de ubicaciones encontradas: {len(datos_por_articulo)}")
    
    # Guardar el PDF
    c.save()
    return ruta_completa

def mostrar_mensaje_exito(mensaje, label_mensaje):
    label_mensaje.configure(text=mensaje, font=('Arial', 10, 'bold'))
    
    # Limpiar después de 3 segundos
    label_mensaje.after(3000, lambda: label_mensaje.configure(text="", font=('Arial', 10)))

def obtener_datos_por_articulo(articulo, archivo_excel):
    resultados = []
    
    try:
        # Si es un StringVar de tkinter, extraer el valor
        if hasattr(archivo_excel, 'get'):
            ruta_archivo = archivo_excel.get()
        else:
            ruta_archivo = archivo_excel
        
        # Verificar que la ruta no esté vacía
        if not ruta_archivo:
            print("Error: No se ha seleccionado ningún archivo")
            return []
        
        # CORRECCIÓN: Usar ruta_archivo en lugar de archivo_excel
        todas_las_hojas = pd.read_excel(ruta_archivo, sheet_name=None)
        
        # Iterar sobre cada hoja
        for nombre_hoja, df in todas_las_hojas.items():
            # Verificar que la hoja tenga las columnas necesarias
            if 'Artículo' in df.columns:
                # Convertir la columna Artículo a string para evitar problemas con floats
                df['Artículo'] = df['Artículo'].astype(str)
                # Limpiar decimales innecesarios (ej: "426367.0" -> "426367")
                df['Artículo'] = df['Artículo'].str.replace(r'\.0$', '', regex=True)
                
                # Convertir el parámetro articulo a string también
                articulo_str = str(articulo)
                
                # Buscar filas donde la columna 'Artículo' coincida con el parámetro
                coincidencias = df[df['Artículo'] == articulo_str]
                
                # Para cada coincidencia, extraer los datos en el orden solicitado
                for _, fila in coincidencias.iterrows():
                    resultado = (
                        fila.get('Artículo', ''),           # Columna D
                        fila.get('Desc Artículo', ''),      # Columna E  
                        fila.get('En Mano', ''),            # Columna H
                        fila.get('Localizador', ''),        # Columna C
                        fila.get('LPN', '')                 # Columna O
                    )
                    resultados.append(resultado)
        
        return resultados
        
    except FileNotFoundError:
        print(f"Error: No se pudo encontrar el archivo {ruta_archivo}")
        return []
    except Exception as e:
        print(f"Error al procesar el archivo: {str(e)}")
        return []

def crear_pdf_pasillo(pasillo, datos_por_pasillo, carpeta_destino="pdfs"):
    """Crea un PDF con los datos del pasillo encontrados en el Excel"""
    # Crear la carpeta si no existe
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)
    
    nombre_archivo = f"pasillo{pasillo}"
    # Ruta completa del archivo
    ruta_completa = os.path.join(carpeta_destino, f"{nombre_archivo}.pdf")
    
    # Crear el PDF
    c = canvas.Canvas(ruta_completa, pagesize=letter)
    width, height = letter
    
    # Configurar posiciones iniciales
    y_position = height - 50
    line_height = 18
    
    # Título principal
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y_position, f"REPORTE DE PASILLO: {pasillo}")
    y_position -= 30
    
    # Fecha y hora de generación
    c.setFont("Helvetica", 10)
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    c.drawString(50, y_position, f"Generado el: {fecha_actual}")
    y_position -= 40
    
    # Verificar si hay datos
    if not datos_por_pasillo:
        c.setFont("Helvetica", 12)
        c.drawString(50, y_position, "No se encontraron datos para este pasillo.")
    else:
        # Información del pasillo
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, y_position, f"Pasillo: P{int(pasillo):02d}")
        y_position -= 25
        
        c.setFont("Helvetica", 11)
        c.drawString(50, y_position, "Artículos ordenados por altura y posición")
        y_position -= 35
        
        # Encabezados de la tabla
        c.setFont("Helvetica-Bold", 10)
        c.drawString(50, y_position, "Localizador")
        c.drawString(150, y_position, "Artículo")
        c.drawString(230, y_position, "Descripción")
        c.drawString(380, y_position, "En Mano")
        c.drawString(450, y_position, "LPN")
        y_position -= 5
        
        # Línea separadora
        c.line(50, y_position, 500, y_position)
        y_position -= 15
        
        # Datos de cada ubicación
        c.setFont("Helvetica", 9)
        total_articulos = 0
        altura_actual = None
        
        for i, (localizador, articulo, desc, en_mano, lpn) in enumerate(datos_por_pasillo):
            # Verificar si necesitamos una nueva página
            if y_position < 80:
                c.showPage()
                y_position = height - 50
                # Repetir encabezados en nueva página
                c.setFont("Helvetica-Bold", 14)
                c.drawString(50, y_position, f"Pasillo: P{int(pasillo):02d} (continuación)")
                y_position -= 30
                c.setFont("Helvetica-Bold", 10)
                c.drawString(50, y_position, "Localizador")
                c.drawString(150, y_position, "Artículo")
                c.drawString(230, y_position, "Descripción")
                c.drawString(380, y_position, "En Mano")
                c.drawString(450, y_position, "LPN")
                y_position -= 5
                c.line(50, y_position, 500, y_position)
                y_position -= 15
                c.setFont("Helvetica", 9)
            
            # Extraer altura del localizador para agrupar visualmente
            try:
                partes = localizador.split('.')
                altura = int(partes[2]) if len(partes) >= 3 else 0
                
                # Si cambió la altura, agregar separador visual
                if altura_actual is not None and altura != altura_actual:
                    y_position -= 5
                    c.setFont("Helvetica", 8)
                    c.setFillColorRGB(0.7, 0.7, 0.7)
                    c.drawString(50, y_position, f"--- Altura {altura} ---")
                    c.setFillColorRGB(0, 0, 0)  # Volver a negro
                    y_position -= 10
                    c.setFont("Helvetica", 9)
                elif altura_actual is None:
                    # Primera altura
                    c.setFont("Helvetica", 8)
                    c.setFillColorRGB(0.7, 0.7, 0.7)
                    c.drawString(50, y_position, f"--- Altura {altura} ---")
                    c.setFillColorRGB(0, 0, 0)
                    y_position -= 10
                    c.setFont("Helvetica", 9)
                
                altura_actual = altura
            except (ValueError, IndexError):
                pass
            
            # Manejar stock
            try:
                stock_texto = str(en_mano) if en_mano != '' and not pd.isna(en_mano) else '0'
            except:
                stock_texto = '0'
            
            # Manejar el LPN (evitar mostrar nan)
            if pd.isna(lpn) or str(lpn).lower() == 'nan':
                lpn_texto = "-"
            else:
                lpn_texto = str(lpn)
            
            # Manejar artículo (convertir float a string limpio si es necesario)
            if pd.isna(articulo):
                articulo_texto = ""
            else:
                articulo_str = str(articulo)
                # Limpiar .0 del final si es un float
                articulo_texto = articulo_str.replace('.0', '') if articulo_str.endswith('.0') else articulo_str
            
            # Truncar textos largos para que entren
            desc_truncada = str(desc)[:18] + "..." if len(str(desc)) > 18 else str(desc)
            
            # Escribir los datos
            c.drawString(50, y_position, str(localizador))
            c.drawString(150, y_position, articulo_texto)
            c.drawString(230, y_position, desc_truncada)
            c.drawString(380, y_position, stock_texto)
            c.drawString(450, y_position, lpn_texto)
            y_position -= line_height
            
            total_articulos += 1
        
        # Línea separadora antes del total
        y_position -= 10
        c.line(50, y_position, 500, y_position)
        y_position -= 20
        
        # Resumen
        c.setFont("Helvetica-Bold", 11)
        c.drawString(50, y_position, f"TOTAL DE ARTÍCULOS EN EL PASILLO: {total_articulos}")
    
    # Guardar el PDF
    c.save()
    return ruta_completa

def obtener_datos_por_pasillo(pasillo, archivo_excel):
    resultados = []
    
    try:
        # Si es un StringVar de tkinter, extraer el valor
        if hasattr(archivo_excel, 'get'):
            ruta_archivo = archivo_excel.get()
        else:
            ruta_archivo = archivo_excel
        
        # Verificar que la ruta no esté vacía
        if not ruta_archivo:
            print("Error: No se ha seleccionado ningún archivo")
            return []
        
        # Formatear el pasillo con ceros a la izquierda (ej: "2" -> "P02")
        pasillo_formateado = f"P{int(pasillo):02d}"
        
        # Leer todas las hojas del archivo Excel
        todas_las_hojas = pd.read_excel(ruta_archivo, sheet_name=None)
        
        # Lista temporal para poder ordenar después
        datos_temporales = []
        
        # Iterar sobre cada hoja
        for nombre_hoja, df in todas_las_hojas.items():
            # Verificar que la hoja tenga las columnas necesarias
            if 'Localizador' in df.columns:
                # Convertir localizador a string para poder hacer comparaciones
                df['Localizador'] = df['Localizador'].astype(str)
                
                # Filtrar filas que pertenezcan al pasillo
                filas_pasillo = df[df['Localizador'].str.startswith(pasillo_formateado, na=False)]
                
                # Para cada fila del pasillo, extraer los datos
                for _, fila in filas_pasillo.iterrows():
                    localizador = fila.get('Localizador', '')
                    
                    # Extraer altura y posición para ordenar
                    try:
                        # Formato: P02.041.3.1 -> dividir en partes
                        partes = localizador.split('.')
                        if len(partes) >= 4:
                            posicion = int(partes[1])  # 041
                            altura = int(partes[2])    # 3
                        else:
                            posicion = 999  # Para casos raros, ponerlos al final
                            altura = 999
                    except (ValueError, IndexError):
                        posicion = 999
                        altura = 999
                    
                    resultado = (
                        localizador,                        # Columna C
                        fila.get('Artículo', ''),           # Columna D
                        fila.get('Desc Artículo', ''),      # Columna E  
                        fila.get('En Mano', ''),            # Columna H
                        fila.get('LPN', ''),                # Columna O
                        altura,                             # Para ordenar
                        posicion                            # Para ordenar
                    )
                    datos_temporales.append(resultado)
        
        # Ordenar por altura y después por posición
        datos_temporales.sort(key=lambda x: (x[5], x[6]))  # altura, posición
        
        # Quitar las columnas de ordenamiento y retornar solo los datos necesarios
        resultados = [(loc, art, desc, stock, lpn) for loc, art, desc, stock, lpn, _, _ in datos_temporales]
        
        return resultados
    
    except FileNotFoundError:
        print(f"Error: No se pudo encontrar el archivo {ruta_archivo}")
        return []
    except Exception as e:
        print(f"Error al procesar el archivo: {str(e)}")
        return []

def buscar_por_articulo(entry_articulo, archivo_excel, label_mensaje):
    """Crea PDF para artículo específico"""
    articulo = entry_articulo.get().strip()
    
    if not articulo:
        messagebox.showwarning("Advertencia", "Por favor ingresa el nombre del artículo")
        return
    
    try:
        nombre_archivo = f"Articulo{articulo}"
        datos_por_articulo = obtener_datos_por_articulo(articulo, archivo_excel)
        ruta_pdf = crear_pdf_articulo(nombre_archivo, datos_por_articulo)
        mensaje = f"✓ PDF creado exitosamente: {nombre_archivo}.pdf"
        mostrar_mensaje_exito(mensaje, label_mensaje)
    except Exception as e:
        messagebox.showerror("Error", f"Error al crear el PDF: {str(e)}")

def buscar_por_pasillo(entry_pasillo, archivo_excel, label_mensaje):
    """Crea PDF para pasillo específico"""
    pasillo = entry_pasillo.get().strip()
    
    if not pasillo:
        messagebox.showwarning("Advertencia", "Por favor ingresa el número de pasillo")
        return
    
    try:
        nombre_archivo = f"Pasillo{pasillo}"
        datos_por_pasillo = obtener_datos_por_pasillo(pasillo, archivo_excel)
        ruta_pdf = crear_pdf_pasillo(pasillo, datos_por_pasillo)
        mensaje = f"✓ PDF creado exitosamente: {nombre_archivo}.pdf"
        mostrar_mensaje_exito(mensaje, label_mensaje)
    except Exception as e:
        messagebox.showerror("Error", f"Error al crear el PDF: {str(e)}")

# --- Función para buscar archivo Excel
def seleccionar_archivo(archivo_excel, lbl_archivo):
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if archivo:
        archivo_excel.set(archivo)
        lbl_archivo.config(text=f"Archivo: {archivo}")

# --- Función para mostrar el formulario correcto
def mostrar_formulario(form_container, modo_var, form_buscar_por_articulo, form_buscar_por_pasillo):
    for child in form_container.winfo_children():
        child.pack_forget()
    if modo_var.get() == "Buscar por artículo":
        form_buscar_por_articulo.pack(fill="x")
    elif modo_var.get() == "Buscar por pasillo":
        form_buscar_por_pasillo.pack(fill="x")

def main():
    root = tk.Tk()
    root.title("Gestionador de artículos")
    root.geometry("600x400")

    # Variable para guardar el archivo seleccionado
    archivo_excel = tk.StringVar(value="")

    # --- Logo en la esquina superior derecha
    logo = tk.PhotoImage(file="marolio_logo.png").subsample(5, 5)
    lbl_logo = ttk.Label(root, image=logo)
    lbl_logo.image = logo 
    lbl_logo.place(relx=1.0, y=0, anchor="ne", x=-10)  

    # Label para mostrar archivo seleccionado
    lbl_archivo = ttk.Label(root, text="Archivo: (ninguno seleccionado)")

    # --- Botón para seleccionar archivo
    btn_archivo = ttk.Button(
        root,
        text="Buscar archivo Excel",
        command=lambda: seleccionar_archivo(archivo_excel, lbl_archivo)
    )

    btn_archivo.pack(pady=10)
    lbl_archivo.pack()

    # Label para mensajes de estado
    label_mensaje = ttk.Label(root, text="", font=('Arial', 10))
    label_mensaje.pack(side="bottom",pady=30)

    # --- Combobox para elegir el modo
    ttk.Label(root, text="Elegí el modo:").pack(pady=(20, 5))

    modo_var = tk.StringVar()
    combo_modo = ttk.Combobox(root, textvariable=modo_var, state="readonly")
    combo_modo["values"] = ("Buscar por artículo", "Buscar por pasillo")
    combo_modo.pack()

    # Contenedor para los formularios
    form_container = ttk.Frame(root, padding=10)
    form_container.pack(fill="both", expand=True, pady=15)

    # --- Formulario por Buscar por articulo
    form_buscar_por_articulo = ttk.Frame(form_container, padding=10)
    ttk.Label(form_buscar_por_articulo, text="Artículo:").grid(row=0, column=0, sticky="w", pady=5)
    entry_articulo = ttk.Entry(form_buscar_por_articulo)
    entry_articulo.grid(row=0, column=1, pady=5, sticky="ew")

    ttk.Button(
        form_buscar_por_articulo, 
        text="Obtener PDF", 
        command=lambda: buscar_por_articulo(entry_articulo, archivo_excel, label_mensaje)
    ).grid(row=2, column=0, columnspan=2, pady=10)

    # --- Formulario por Buscar por pasillo
    form_buscar_por_pasillo = ttk.Frame(form_container, padding=10)
    ttk.Label(form_buscar_por_pasillo, text="Pasillo:").grid(row=0, column=0, sticky="w", pady=5)
    entry_pasillo = ttk.Entry(form_buscar_por_pasillo)
    entry_pasillo.grid(row=0, column=1, pady=5, sticky="ew")

    ttk.Button(
        form_buscar_por_pasillo, 
        text="Obtener PDF", 
        command=lambda: buscar_por_pasillo(entry_pasillo, archivo_excel, label_mensaje)
    ).grid(row=2, column=0, columnspan=2, pady=10)

    combo_modo.bind(
        "<<ComboboxSelected>>",
        lambda event: mostrar_formulario(form_container, modo_var, form_buscar_por_articulo, form_buscar_por_pasillo)
    )

    root.mainloop()

if __name__ == "__main__":
    main()