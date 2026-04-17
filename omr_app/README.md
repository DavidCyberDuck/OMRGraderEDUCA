# OMR Grader

Aplicación de escritorio para calificar exámenes automáticamente usando hojas OMR escaneadas en PDF.

## Características

- 📋 **Genera hojas OMR** listas para imprimir (Sección 1: opción múltiple, Sección 2: autoconocimiento)
- 📂 **Procesa PDFs escaneados** con múltiples estudiantes por archivo
- 🗝 **Clave de respuestas configurable** y número de preguntas ajustable (hasta 40)
- 📊 **Exporta a Excel** con tabla de resultados, detalle por pregunta, autoconocimiento y gráficas

## Requisitos

```bash
pip install opencv-python numpy Pillow pdf2image reportlab openpyxl
```

También necesitas **poppler** para pdf2image:
- **Windows**: Descargar desde https://github.com/oschwartz10612/poppler-windows y agregar al PATH
- **macOS**: `brew install poppler`
- **Linux**: `sudo apt install poppler-utils`

## Instalación

```bash
git clone <repo>
cd omr_grader
pip install -r requirements.txt
python main.py
```

## Uso

### 1. Generar hojas OMR
1. Abre la app: `python main.py`
2. Configura el nombre del examen y número de preguntas (Sección 1)
3. Clic en **"Generar Hoja OMR"** → guarda el PDF
4. Imprime las hojas para los estudiantes

### 2. Calificar exámenes
1. Los estudiantes rellenan sus hojas con lápiz o bolígrafo negro
2. Escanea todas las hojas en un solo PDF (una página por estudiante)
3. En la app, selecciona el PDF escaneado
4. Configura la clave de respuestas (menús desplegables A/B/C/D)
5. Clic en **"Calificar PDF"** y elige dónde guardar el Excel
6. ¡Listo! El Excel tendrá resultados, detalle y gráficas

## Estructura del proyecto

```
omr_app/
├── main.py              # GUI principal (Tkinter)
├── sheet_generator.py   # Genera PDFs con hojas OMR
├── omr_scanner.py       # Detecta burbujas en imágenes
├── omr_grader.py        # Califica respuestas
├── omr_exporter.py      # Exporta resultados a Excel
├── omr_config.json      # Configuración guardada automáticamente
└── README.md
```

## Hoja OMR — Estructura

```
┌─────────────────────────────────────┐
│  [■] Nombre: _______  Fecha: _____  [■]
│                              ID Est.
│  Sección 1 (opción múltiple)  [D][U]
│  1. ( )A ( )B ( )C ( )D      [0][0]
│  2. ( )A ( )B ( )C ( )D      [1][1]
│  ...                          ...
│  ─────────────────────────────────  │
│  Sección 2 (autoconocimiento)       │
│  1. ( )1 ( )2 ( )3 ( )4 ( )5       │
│  ...                                │
│  [■]                            [■] │
└─────────────────────────────────────┘
```

## Consejos para mejores resultados

- Usar lápiz o bolígrafo negro, rellenar completamente el círculo
- Escanear a 200 dpi mínimo, en blanco y negro o escala de grises
- Asegurarse de que la hoja esté recta en el escáner
- Evitar arrugar o doblar las hojas antes de escanear

## Salida Excel

El archivo Excel generado contiene 4 hojas:
1. **Resumen** — tabla de todos los estudiantes con puntaje, calificación letra y promedio de autoconocimiento
2. **Detalle Preguntas** — respuesta por pregunta de cada estudiante, con colores verde/rojo
3. **Autoconocimiento** — valores de la Sección 2 por estudiante
4. **Gráficas** — gráfica de barras de puntajes y gráfica de línea de autoconocimiento
