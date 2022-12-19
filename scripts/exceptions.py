class Error(Exception):
    """ Clase base para otras excepciones"""
    pass

class AssetStatusError(Error):
    """ Se invoca cuando un activo fijo posee valores True para las columnas 'is_active' 
    y 'is_capex """