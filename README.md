# EasyExcelAPIView
Este módulo permite exportar fácilmente datos de una modelo de Django a un archivo de Excel 2007 o superior.

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
* `model`: Modelo de Django que se desea exportar.
* `filename`: Nombre del archivo de Excel que se generará. Por defecto es `export.xlsx`.
* `sheet_name`: Nombre de la hoja de Excel que se generará. Por defecto es `Hoja 1`.
* `fields`: Lista de campos que se desean exportar. Por defecto se exportan todos los campos del modelo.
* `with_relacion`: Booleano que indica si se desea exportar los campos relacionados. Por defecto es `False`.