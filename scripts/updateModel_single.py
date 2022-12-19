from datetime import datetime

import pandas as pd
import xlwings as xw

with xw.App() as app:

    last_actual_date = datetime.fromisoformat('2023-04-30')

    # Cargar el libro

    model_wb = xw.Book(r'Projects/FinancialModel/output/Ekorec_update_test.xlsx')
    monthly_ws = model_wb.sheets['Mensual']
    annual_ws = model_wb.sheets['Anual']

    # Parametros generales #

    label_mapping = {
        'Achieved pool price': 'MD_Captured' ,
        'TTF': 'TTF',
        'Nat gas variable term': 'Tv_gas',
        'Steam price calcculations': 'Heat_price',
        'Steam discount': 'Heat_discount_%',
        'Operating incentive, Ro': 'RO',
        'Efficiency HHV': 'Elec_eff',
        'Steam demand': 'Heat_demand',
        'Running hours': 'Running_hours', 
        'Gross generation': 'Generation_MWh',
        'Electricity to the grid': 'Export_MWh',
        'Steam to client CHP': 'Heat_CHP_MWh',
        'Steam to client Boiler': 'Heat_boiler_MWh',
        'Market': 'Income_MD_€', 
        'Ro revenue': 'Income_RO_€',
        'Heat to industry CHP': 'Income_steam_CHP_€',
        'Heat to industry boiler': 'Income_steam_boiler_€',
        'Nat gas variable term CHP': 'Variable_gas_CHP_cost_€',
        'Nat gas variable term Boiler': 'Variable_gas_boiler_cost_€',
        'CO2 adjustment': 'CO2_cost_€', 
        'CHP Maintenance ': 'Manto_cost_€',
        'Electricity tax': 'Variable_tax_7%',
        'Nat gas fixed term': 'Fix_gas_cost_€',
        'Operation (Personnel expenses)': 'Personnel_€',
        'Insurance': 'Insurance_€',
        'Land Rent': 'Land_rent_€',
        'Administr. Legal cost / Supplies': 'Adm&legal_cost_€',
        'Management fee': 'Managem_fee_€',
        'IBI, IAE, etc': 'Corporate_tax_€',
    }
    exception_labels = [
        'Gross margin',
        'EBITDA',
        'Depreciation on own assets',
        'Depreciation on acquired assets',
        'EBIT',
        'Profit before tax',
        'Profit after tax'
    ]
    
    model_inputs_df = pd.read_excel(
        r'Projects/FinancialModel/input/monthlyInputs_update.xlsx',
        index_col='DiaLocal' 
    )

    # ---- Actualización hoja mensual ---- #

    # Construir el dataframe de fechas

    app.status_bar = 'Actualizando encabezados...'
    dates_range = monthly_ws.range('E5:XFD5')

    months = []
    column_numbers = []
    status = []
    column = 0

    for cell in dates_range:
        if cell.value != None:
            month = dates_range[0, column].value

            if month <= last_actual_date:
                status.append('Actual')
            else:
                status.append('Forecast')
            
            months.append(month)
            column_numbers.append(column)
            column += 1
        else:
            break

    date_column_map = {
        'date': pd.Series(months),
        'columnNumber': pd.Series(column_numbers),
        'monthStatus': pd.Series(status)
    }

    model_dates = pd.DataFrame(data=date_column_map).set_index('date')
    last_month_column = int(model_dates.tail(1)['columnNumber'])
    months = model_dates.index
    dates_range = monthly_ws.range(
        (5, 5),
        (5, 4 + column) # Se suman 4 porque se deben considerar las 5 primeras columnas (zero-indexed)
    )

    del(column)

    # Actualizacion del encabezado de fechas

    for month in months:
        if month <= last_actual_date:
            column = int(model_dates['columnNumber'][month])
            date_cell = dates_range[0, column]
            date_status_cell = date_cell.offset(row_offset = -1)

            date_status_cell.value = 'Actual'
            date_status_cell.color = (226, 239, 218)
            date_status_cell.font.italic = True
            date_cell.color = (226, 239, 218)
            date_cell.font.italic = True
        else:
            break

    del(month)
    del(column)

    actual_periods = model_dates[model_dates['monthStatus'] == 'Actual']
    actual_months = actual_periods.index
    forecast_periods = model_dates[model_dates['monthStatus'] == 'Forecast']
    forecast_months = forecast_periods.index

    # ---- P&L ---- #

    app.status_bar = "Actualizando P&L..."

    pnl_labels_range = monthly_ws.range((53, 2), (104, 2))
    pnl_values_range = monthly_ws.range(
        (53, 5), 
        (104, 5 + last_month_column)
    )
    pnl_rows = range(0, pnl_labels_range.rows.count)

    # Actual #

    for month in actual_months:
        column = int(actual_periods['columnNumber'][month])

        for row in pnl_rows:
            row_label = pnl_labels_range[row, 0].value
            input_data_column = label_mapping.get(row_label)
            value_cell = pnl_values_range[row, column]
            
            if input_data_column == None:
                if row_label in exception_labels:
                    value_cell.font.italic = True
                else:
                    value_cell.color = (244, 249, 241)
                    value_cell.font.italic = True
            else:    
                amount = model_inputs_df[input_data_column][month]
                value_cell.value = amount
                value_cell.color = (244, 249, 241)
                value_cell.font.italic = True

    x = 0
    