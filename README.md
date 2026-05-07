# Style Mapper para DITA

Herramienta web para preparar documentos Word antes de convertirlos a DITA
con OxygenAuthor. Detecta estilos no reconocidos, permite mapearlos a estilos
estándar y genera un documento nuevo listo para el Batch Converter.

## Estructura del proyecto

```
style-mapper/
├── app.py                  ← Interfaz Streamlit (punto de entrada)
├── database.py             ← Base de datos SQLite
├── style_processor.py      ← Leer y modificar estilos en .docx
├── xml_parser.py           ← Leer el XML de configuración de Oxygen
├── requirements.txt        ← Dependencias Python
├── render.yaml             ← Configuración de despliegue en Render
└── data/
    └── stylesWordToDita.xml ← Configuración de estilos de OxygenAuthor
```

## Cómo ejecutar localmente

### 1. Crear entorno virtual
```bash
python -m venv venv
```

### 2. Activar el entorno virtual
- **Windows:**  `venv\Scripts\activate`
- **Mac/Linux:** `source venv/bin/activate`

### 3. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 4. Ejecutar la app
```bash
streamlit run app.py
```

Se abrirá automáticamente en `http://localhost:8501`

## Notas importantes

- La base de datos (`style_mapper.db`) se crea automáticamente al iniciar.
- En Render (plan gratuito), las reglas guardadas se pierden al redesplegar.
  Para persistencia permanente, añadir una base de datos PostgreSQL en Render.
- El archivo `data/stylesWordToDita.xml` contiene la configuración de Oxygen
  y se carga automáticamente en cada inicio.
