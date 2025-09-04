# Instructivo de Instalación y Uso

## Gestionador de Artículos - Generador de PDFs

### 📋 Descripción

Esta aplicación permite generar reportes en PDF a partir de archivos Excel, organizando la información por artículo específico, por pasillo completo, o registrando diferencias encontradas.

---

## 🔧 Instalación

### Paso 1: Instalar Python

1. Descargar Python desde [python.org](https://www.python.org/downloads/)
2. Durante la instalación, **marcar la casilla "Add Python to PATH"**
3. Verificar la instalación abriendo CMD y escribiendo: `python --version`

### Paso 2: Instalar las librerías necesarias

Abrir la terminal/CMD y ejecutar los siguientes comandos uno por uno:

```bash
pip install tkinter
pip install pandas
pip install reportlab
pip install openpyxl
```

### Paso 3: Preparar los archivos

1. Crear una carpeta para el proyecto (ej: "GestionadorArticulos")
2. Guardar el archivo `main.py` en esa carpeta
3. Colocar la imagen `marolio_logo.png` en la misma carpeta

### Paso 4: Ejecutar la aplicación

1. Abrir terminal/CMD en la carpeta del proyecto
2. Ejecutar: `python main.py`

---

## 📖 Cómo usar la aplicación

### Interfaz Principal

Al ejecutar la aplicación verás:

* Botón "Buscar archivo Excel" en la parte superior
* Selector de modo (Buscar por artículo / Buscar por pasillo / Diferencias)
* Formularios dinámicos según el modo seleccionado
* Área de mensajes en la parte inferior

### Paso a Paso de Uso

#### 1. Seleccionar archivo Excel

* Hacer clic en **"Buscar archivo Excel"**
* Navegar hasta tu archivo Excel (.xlsx o .xls)
* Seleccionar el archivo
* Verificar que aparezca la ruta del archivo seleccionado

#### 2. Elegir el modo de búsqueda

##### **Modo: Buscar por artículo**

* Seleccionar "Buscar por artículo" del menú desplegable
* Ingresar el código del artículo en el campo "Artículo"
* Hacer clic en "Obtener PDF"

**¿Qué genera?**

* PDF con todas las ubicaciones donde se encuentra ese artículo
* Stock total del artículo
* Detalle por ubicación (Localizador, En Mano, LPN)

##### **Modo: Buscar por pasillo**

* Seleccionar "Buscar por pasillo" del menú desplegable
* Ingresar el número de pasillo (ej: 2, 15, 34)
* Hacer clic en "Obtener PDF"

**¿Qué genera?**

* PDF con todos los artículos del pasillo especificado
* Organizado por altura y posición
* Total de artículos en el pasillo

##### **Modo: Diferencias**

* Seleccionar "Diferencias" del menú desplegable
* Ingresar el **Localizador** a revisar
* Seleccionar el **Estado** (FALTANTE o SOBRANTE)
* Hacer clic en **"Agregar diferencia"**

**¿Qué hace?**

* Registra en una hoja adicional llamada **"Diferencias"** dentro del mismo Excel:

  * Localizador
  * Artículo
  * Descripción
  * Stock (En Mano)
  * LPN
  * Estado seleccionado (FALTANTE / SOBRANTE)

⚠️ **Advertencia importante:**
Al usar esta funcionalidad, **NO debes tener el archivo Excel abierto**, de lo contrario los cambios no podrán guardarse.

#### 3. Encontrar los PDFs generados

* Los PDFs se guardan automáticamente en la carpeta **"pdfs"**
* Esta carpeta se crea automáticamente en la misma ubicación que `main.py`
* Nombres de archivo:

  * Artículos: `Articulo[CODIGO].pdf`
  * Pasillos: `Pasillo[NUMERO].pdf`

---

## 📊 Formato requerido del Excel

### Columnas necesarias en el archivo Excel:

| Columna | Nombre        | Descripción                        |
| ------- | ------------- | ---------------------------------- |
| C       | Localizador   | Ubicación física (ej: P02.041.3.1) |
| D       | Artículo      | Código del artículo                |
| E       | Desc Artículo | Descripción del producto           |
| H       | En Mano       | Cantidad disponible                |
| O       | LPN           | Número de lote/pallet              |

### Formato del Localizador

* **Estructura**: `P[Pasillo].[Posición].[Altura].[Subdivision]`
* **Ejemplo**: `P02.041.3.1`

  * P02 = Pasillo 2
  * 041 = Posición 41
  * 3 = Altura 3
  * 1 = Subdivisión 1

---

## ⚠️ Solución de Problemas

### Error: "No module named 'tkinter'"

* **Windows**: `pip install tk`
* **Linux**: `sudo apt-get install python3-tk`
* **Mac**: `brew install python-tk`

### Error: "No se pudo encontrar el archivo"

* Verificar que se haya seleccionado un archivo Excel
* Confirmar que el archivo no esté abierto en Excel
* Verificar permisos de lectura del archivo

### Error: "No se encontraron datos"

* Revisar que el archivo Excel tenga las columnas correctas
* Verificar que el código de artículo existe en el Excel
* Para pasillos, confirmar que el formato del Localizador sea correcto

### La aplicación no abre

* Verificar que Python esté correctamente instalado
* Confirmar que `marolio_logo.png` esté en la misma carpeta
* Ejecutar desde terminal para ver mensajes de error

### PDFs se ven mal o incompletos

* Verificar que los datos del Excel no tengan caracteres especiales
* Confirmar que las columnas tengan los nombres exactos requeridos

---

## 💡 Consejos de Uso

### Para mejores resultados:

* **Mantener el archivo Excel cerrado** mientras usas la aplicación
* **Usar códigos de artículo exactos** (sin espacios adicionales)
* **Verificar el formato del Excel** antes de procesar
* **Revisar los PDFs generados** en la carpeta "pdfs"

### Funcionalidades adicionales:

* La aplicación procesa **múltiples hojas** del mismo archivo Excel
* Los datos se **ordenan automáticamente** por ubicación
* Se muestra **fecha y hora** de generación en cada PDF
* Los **totales se calculan automáticamente**
* Ahora puedes **registrar diferencias** directamente en el mismo Excel

---

## 📁 Estructura de archivos necesaria

```
GestionadorArticulos/
├── main.py
├── marolio_logo.png
└── pdfs/ (se crea automáticamente)
    ├── Articulo123456.pdf
    ├── Pasillo02.pdf
    └── ...
```

---

## 🆘 Soporte

Si encuentras problemas:

1. Verificar que todas las librerías están instaladas
2. Confirmar que el archivo Excel tiene el formato correcto
3. Revisar que `marolio_logo.png` esté presente
4. Ejecutar desde terminal para ver mensajes de error detallados

---

*Última actualización: Septiembre 2025*
