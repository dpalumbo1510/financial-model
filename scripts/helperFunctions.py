import numpy as np
from datetime import datetime
import calendar

def change_to_last_day(date):
    # Los dias de las fechas en el encabezado del modelo son los 31. 
    # En algunos dataframes del script, estos son el 1, entonces es necesario 
    # cambiar el día para que algunos loops puedan buscar la fecha del
    # encabezado.

    new_date = datetime(
        year = date.year, 
        month = date.month, 
        day = calendar.monthrange(
            year = date.year,
            month = date.month)[1])

    return new_date
