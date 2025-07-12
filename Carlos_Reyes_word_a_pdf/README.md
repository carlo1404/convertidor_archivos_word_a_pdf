# Convertidor Word a PDF

Este programa permite convertir archivos de Microsoft Word (.docx) a PDF mediante una interfaz gráfica sencilla creada con Tkinter.

## Características

- ✅ Interfaz gráfica intuitiva y moderna
- ✅ Conversión de archivos .docx a PDF
- ✅ Barra de progreso durante la conversión
- ✅ Selección de carpeta de destino
- ✅ Manejo de errores
- ✅ Interfaz en español

## Requisitos

- Python 3.10 o superior
- Microsoft Word instalado en el sistema (requisito de la librería `docx2pdf` en Windows)
- Windows (por compatibilidad con docx2pdf)

## Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/carlo1404/convertidor_archivos_word_a_pdf.git
cd convertidor_archivos_word_a_pdf
```

### 2. Instalar dependencias

```bash
pip install -r requirements.txt
```

O instalar manualmente:

```bash
pip install docx2pdf
```

Si tienes varias versiones de Python y el comando anterior no funciona, prueba:

```bash
python -m pip install docx2pdf
```

O bien:

```bash
py -m pip install docx2pdf
```

> **Nota:** Tkinter viene incluido con la mayoría de las instalaciones de Python en Windows.

## Uso

1. Ejecuta el programa:
   
   ```bash
   python convertidor_word_pdf.py
   ```

2. Haz clic en "Seleccionar archivo" y elige un archivo `.docx` de tu computadora.
3. Haz clic en "Convertir a PDF".
4. Selecciona la carpeta donde quieres guardar el PDF convertido.
5. El programa mostrará un mensaje indicando si la conversión fue exitosa o si ocurrió algún error.

## Funcionamiento

- El programa utiliza la librería `docx2pdf` para realizar la conversión, lo que garantiza que el formato del documento se mantenga igual al original.
- La interfaz gráfica está hecha con Tkinter y es muy intuitiva.
- La conversión se realiza en un hilo separado para no congelar la interfaz.
- Incluye una barra de progreso durante la conversión.

## Estructura del Proyecto

```
convertidor_archivos_word_a_pdf/
├── convertidor_word_pdf.py    # Archivo principal del programa
├── requirements.txt           # Dependencias del proyecto
├── README.md                 # Este archivo
└── .gitignore               # Archivos a ignorar por Git
```

## Documentación en el código

El código fuente está ampliamente comentado para facilitar su comprensión y mantenimiento. Cada función incluye docstrings explicando su propósito y funcionamiento.

## Autor

**Carlos Reyes** - Desarrollador del convertidor Word a PDF

## Solución de problemas

### Error: `ModuleNotFoundError: No module named 'docx2pdf'`

Esto significa que la librería `docx2pdf` no está instalada en tu entorno de Python. Para solucionarlo, ejecuta uno de los siguientes comandos en la terminal:

```bash
pip install docx2pdf
```

O si tienes varias versiones de Python:

```bash
python -m pip install docx2pdf
```

O bien:

```bash
py -m pip install docx2pdf
```

Luego, vuelve a ejecutar el programa.

### Error: "Microsoft Word no está instalado"

La librería `docx2pdf` requiere que Microsoft Word esté instalado en el sistema. Si no lo tienes instalado, puedes:

1. Instalar Microsoft Word desde la Microsoft Store o descargarlo desde el sitio oficial de Microsoft
2. O usar una alternativa como LibreOffice (aunque puede requerir configuración adicional)

### La conversión no funciona correctamente

- Asegúrate de que el archivo .docx no esté corrupto
- Verifica que tienes permisos de escritura en la carpeta de destino
- Cierra el archivo Word si está abierto en otra aplicación

## Contribuciones

Las contribuciones son bienvenidas. Para contribuir:

1. Haz un fork del proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo `LICENSE` para más detalles.

## Versión

Versión actual: 1.0

## Fecha de creación

11/07/2025 