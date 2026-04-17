# Calificador OMR

Aplicación de escritorio para generar, escanear y calificar hojas de respuestas OMR. Procesa PDFs escaneados con múltiples estudiantes y exporta los resultados a Excel.

## Requisitos del sistema

- Python 3.10 o superior
- poppler (requerido por pdf2image)
  - macOS: `brew install poppler`
  - Windows: descargar desde https://github.com/oschwartz10612/poppler-windows y agregar la carpeta `bin/` al PATH
  - Linux: `sudo apt install poppler-utils`

## Instalación

```bash
git clone https://github.com/DavidCyberDuck/OMRGraderEDUCA.git
cd OMRGraderEDUCA
python install.py
```

El script `install.py` instala poppler y las dependencias de Python automáticamente.

Para instalar solo las dependencias de Python:

```bash
pip install -r omr_app/requirements.txt
```

## Uso

```bash
python omr_app/main.py
```

### Generar hojas OMR

1. Configurar el nombre del examen y el número de preguntas de la Sección 1
2. Hacer clic en "Generar Hoja OMR" y guardar el PDF
3. Imprimir las hojas para los estudiantes

### Calificar exámenes

1. Los estudiantes rellenan sus hojas con lápiz o bolígrafo negro
2. Escanear todas las hojas en un solo PDF (una página por estudiante, mínimo 200 dpi)
3. En la aplicación, seleccionar el PDF escaneado y configurar la clave de respuestas
4. Hacer clic en "Calificar PDF" y elegir la carpeta de destino para el Excel

## Estructura del proyecto

```
omr_app/
├── main.py              # Interfaz gráfica (Tkinter)
├── layout.py            # Coordenadas de la hoja (fuente única de verdad)
├── sheet_generator.py   # Generación de PDFs
├── omr_scanner.py       # Detección de burbujas con OpenCV
├── omr_grader.py        # Calificación de respuestas
├── omr_exporter.py      # Exportación a Excel
└── requirements.txt
```

## Lista de alumnos (opcional)

Carga un archivo Excel con dos columnas — `folio` y `nombre` — para incluir automáticamente el nombre del estudiante en los resultados. Los folios no encontrados aparecen resaltados en amarillo en el Excel exportado.

## Salida Excel

El archivo generado contiene cinco hojas:

- **Resumen** — folio, nombre, puntaje y promedio de autoconocimiento por estudiante
- **Detalle Preguntas** — respuesta por pregunta con indicación de acierto o error
- **Autoconocimiento** — valores de la Sección 2 por estudiante
- **Gráficas** — gráfica de barras de puntajes y gráfica de línea de autoconocimiento
- **Clave de Respuestas** — clave configurada para el examen

## Recomendaciones de escaneo

- Rellenar los círculos completamente con lápiz o bolígrafo negro
- Escanear a 200 dpi como mínimo, en blanco y negro o escala de grises
- Colocar la hoja recta en el escáner, sin dobleces ni arrugas
