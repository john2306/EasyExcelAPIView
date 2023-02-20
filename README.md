# EasyExcelAPIView
Este módulo permite exportar fácilmente datos de una modelo de Django a un archivo de Excel 2007 o superior.

## Requerimientos
* Python 3.6 o superior
* Django 2.2 o superior
* openpyxl 3.0.3 o superior
* django-rest-framework 3.11.0 o superior

## Dependencias
* pip install openpyxl==3.1.1
* pip install djangorestframework==3.14.0

## Instalación
Descargar o copiar el archivo `EasyExcelAPIView.py` y agregarlo a su proyecto.

## Uso
Para usar el módulo, se debe crear una clase que herede de `EasyExcelAPIView` y agregar el atributo model con el modelo de Django que se desea exportar.

```python
from EasyExcelAPIView import EasyExcelAPIView

class ExportarModeloExcel(EasyExcelAPIView):
    model = Modelo
```
 
### Atributos   
* `model`: Modelo de Django que se desea exportar (necesario).
* `filename`: Nombre del archivo de Excel que se generará (opcional).
* `header`: Lista de encabezados que se desean mostrar en el archivo de Excel. 
            Por defecto se usan los nombres o el `verbose_name` de los campos del modelo (opcional).
* `sheet_name`: Nombre de la hoja de Excel que se generará (opcional).
* `fields`: Lista de campos que se desean exportar. Por defecto se exportan todos los campos del modelo (opcional).
* `with_relacion`: Booleano que indica si se desea exportar los campos relacionados. Por defecto es `False` (opcinal).