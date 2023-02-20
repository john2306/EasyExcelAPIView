from .EasyExcelAPIView import EasyExcelAPIView
"""
1. Forma más fácil de usar el módulo
"""

class EjemploUsarEasyExcelAPIView(EasyExcelAPIView):
    """
    Extiende la clase EasyExcelAPIView para usarla y personalizarla.
    : param model: Nombre del modelo de Django a usar
    : param fields: Por defecto son todos los campos del modelo, excluyendo los campos de relaciones.
    : param with_relation: Por defecto es False, si es True, incluye los campos de relaciones.
    : param headers: Por defecto son los nombres de los campos del modelo o el verbose_name de los campos.
    """
    model = 'Clientes' # Aquí debe importar el modelo de Django