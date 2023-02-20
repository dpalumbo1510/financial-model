from datetime import date
from copy import copy
from time import strftime

import pandas as pd
import xlwings as xw
import xlsxwriter

from helperFunctions import *
from fixedAssets_conso import *

# Configuración previa de Pandas:
pd.options.mode.chained_assignment = None  # default='warn'
# ----

model_start_date = input(
    '¿Cuál es la fecha de inicio del modelo? (AAAA-MM-DD): ')

with xw.App() as app:

# Configuración del entorno de Excel para optimizar la velocidad del script:

    app.screen_updating = False
    app.visible = False
    app.interactive = False
    app.calculation = 'manual'

# Fechas para la generacion del modelo:

    print('Inicializando...')
    model_start_date = date.fromisoformat(model_start_date)
    model_start_date = pd.Timestamp(model_start_date)
    all_periods = []

# Creación y configuración del libro:

    new_wb = xw.Book()
    new_wb.activate()
    new_wb_first_sheet = new_wb.sheets[0]

# Carga del template:

    templates_wb = xw.Book(
        r'C:\VSCODE-Local\Projects\FinancialModel\templates\modelTemplate.xlsx')

# Selección de las plantas:

    plants_information = pd.read_excel(
        r'C:\VSCODE-Local\Projects\FinancialModel\input\plants.xlsx')
    plant_list = plants_information.index

# Parametros del template:

    plant_template_ws = templates_wb.sheets['PlantTemplate']
    plant_template_last_row = str(plant_template_ws.used_range.last_cell.row)
    conso_a_template_ws = templates_wb.sheets['Conso_A_Template']
    conso_a_ws_name = 'Conso - Annual'
    conso_m_template_ws = templates_wb.sheets['Conso_M_Template']
    conso_m_ws_name = 'Conso - Monthly'
    holding_template_ws = templates_wb.sheets['Holding']
    conso_adj_temp_ws = templates_wb.sheets['Conso Adj']

    month_status_row = 4
    dates_row = 5
    content_start_row = 3

# Carga de data externa:

    fixed_assets = pd.read_excel(
        r'C:\VSCODE-Local\Projects\FinancialModel\input\fixed_assets_conso.xlsx'
    )

# Generación de la hoja de cada planta:

    ws_counter = 0
    all_capex_monthly = pd.DataFrame(
        columns=['plant_name', 'start_date', 'historic_cost']
    )

    for plant in plant_list:

        last_ws_generated = new_wb.sheets[ws_counter]
        plant_name = plants_information.loc[plant, 'name']
        plant_code = plants_information.loc[plant, 'it_code']
        plant_EOL_year = str((plants_information.loc[plant, 'reg_per_end']).year)

        plant_template_ws.copy(after = last_ws_generated, name = plant_name)
        
        current_ws = new_wb.sheets[plant_name]
        current_ws.range('B2').value = plant_name
        current_ws.range('B3').value = plant_code
        current_ws.range('B4').value = 'End of regulatory life: ' + plant_EOL_year

        reg_period_end = plants_information.loc[plant, 'reg_per_end']
        
        model_timeframe = pd.date_range(
            start=model_start_date, 
            end=reg_period_end, 
            freq='M'
        ).to_series().index
        
        plant_months = []
        plant_col_numbers = []
        
        current_column = 6
        column_number = 0

        print("Generando modelo para " + plant_name)

        for month in model_timeframe:
            
            current_ws.range('D3:' + 'D' + plant_template_last_row).copy(
                current_ws.range(
                    content_start_row, 
                    current_column
                )
            )
            
            # Encabezado de fechas:
            date_cell = current_ws.range((dates_row, current_column))
            date_cell.value = month
            date_cell.number_format = "mm-aaaa"
            date_cell.font.bold = True
            date_cell.api.Borders.LineStyle = 1
            date_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
            date_cell.color = (221, 235, 247)
            
            # Encabezado del estatus del mes:
            month_status_cell = current_ws.range((month_status_row, current_column))
            month_status_cell.font.bold = True
            month_status_cell.api.Borders.LineStyle = 1
            month_status_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
            month_status_cell.value = "Forecast"
            month_status_cell.color = (221, 235, 247)

            date_cell.column_width = 13
            
            plant_months.append(month)
            plant_col_numbers.append(column_number)
            
            current_column += 1
            column_number +=1
            
            all_periods.append(month)

        date_column_map = {
            'date': pd.Series(plant_months),
            'columnNumber': pd.Series(plant_col_numbers)
        }

        plant_model_dates = pd.DataFrame(data=date_column_map).set_index('date')
        plant_model_months = plant_model_dates.index
        last_period_column = int(plant_model_dates.tail(1).values)
        last_period_date = plant_model_dates.tail(1).index.item()
        
        current_ws.range('D:D').delete()

        # Depreciación de activos existentes:

        plant_fx_assets = fixed_assets[fixed_assets['Plant code'] == plant_code]
        
        capex, curr_assets_amort, capex_amortization = calculate_fixed_assets(
            plant_fx_assets, 
            model_start_date
        )

        dep_mapping = {
            'Plant_Active_Assets_Dep': curr_assets_amort,
            'Plant_CAPEX_Assets_Dep': capex_amortization
        }

        for range_name, data_source in dep_mapping.items():

            try:
                # Cortar el dataframe hasta la fecha final del modelo
                dep_last_period = data_source.tail(1).index.item()

                if dep_last_period > last_period_date:
                    data_source = data_source[:last_period_date]

                else:
                    pass

                dep_periods = data_source.index
                labels = current_ws.range(range_name)
                labels_lastrow = (labels.row + (labels.rows.count) - 1)
                values = current_ws.range(
                    (labels.row, 5),
                    (labels_lastrow, 5 + last_period_column)
                )
                range_rows = range(0, labels.rows.count)

                for row in range_rows:
                    for period in dep_periods:
                        amount = data_source[period]
                        new_period = change_to_last_day(period)
                        month_column = int(plant_model_dates['columnNumber'][new_period])
                        values[0, month_column].value = amount * (-1)

            except ValueError:
                pass

            app.calculate()

        # ---- Actualización resultado acumulado al cierre ---- #
        
        ret_earnings_label = current_ws.range('Plant_Results_Accrued')
        ret_earnings_values = current_ws.range(
            (ret_earnings_label.row, 5), 
            (ret_earnings_label.row, 5 + last_period_column)
        )

        py_ret_earn_value = ret_earnings_values.offset(column_offset = -1)        

        py_result_range = ret_earnings_values.offset(
            row_offset = 1, 
            column_offset = -1
        )

        for month in plant_model_months:
            month_number = month.month
            if month_number == 1:
                month_column = int(plant_model_dates['columnNumber'][month])
                py_ret_earnings = py_ret_earn_value[0, month_column].address
                py_result = py_result_range[0, month_column].address
                ret_earnings_values[0, month_column].formula = \
                    '=' + py_ret_earnings + '+' + py_result + ''
            else:
                pass

            app.calculate()

        # ---- Acumulación del resultado del año en balance ---- #

        pat_label = current_ws.range('Plant_PAT')
        pat_values = current_ws.range(
            (pat_label.row, 5),
            (pat_label.row, 5 + last_period_column)
        )

        cy_result_label = current_ws.range('Plant_Results_CurrentYear')
        cy_result_values = current_ws.range(
            (cy_result_label.row, 5), 
            (cy_result_label.row, 5 + last_period_column)
        )
        
        for month in plant_model_months:
            month_column = int(plant_model_dates['columnNumber'][month])
            current_cell = cy_result_values[0, month_column]
            month_number = month.month
            if month_number == 1:
                current_cell.formula = \
                    '=' + pat_values[0, month_column].address + ''
            else:
                accrued_result = current_cell.offset(column_offset = -1)
                current_cell.formula = \
                    '=' + accrued_result.address + '+' + \
                        pat_values[0, month_column].address + ''

            app.calculate()

        ws_counter += 1
        print('Pestaña ' + '"' + plant_name + '"' + ' completada.')

# Generación de la hoja del Holding

    holding_template_ws.copy(after = new_wb_first_sheet, name = 'Holding')

# Generación de la hoja de los ajustes de consolidación

    conso_adj_temp_ws.copy(after = new_wb_first_sheet, name = 'Conso Adjs')

# Generación de la hoja con el consolidado mensual:

    print('Generando consolidado mensual...')
    
    period_to_model = pd.Series(
        all_periods
    ).drop_duplicates().sort_values(ascending=True).reset_index(drop=True)

    conso_m_template_ws.copy(
        after=new_wb_first_sheet, 
        name=conso_m_ws_name
    )

    conso_m_template_last_row = str(conso_m_template_ws.used_range.last_cell.row)
    conso_m_ws = new_wb.sheets[conso_m_ws_name]

    base_column = conso_m_ws.range('D3:' + 'D' + conso_m_template_last_row)
    first_plant_ws = new_wb.sheets[2].name
    last_plant_ws = new_wb.sheets[-1].name
    
    for row in range(0, base_column.rows.count):
            
        current_cell = base_column[row, 0]
        offset_cell = conso_m_ws.range(current_cell.row, 3)
        formula_address = offset_cell.get_address(
            row_absolute=False, 
            column_absolute=False
        )
            
        if current_cell.value == 'F':
                current_cell.formula = '=SUM(' + "'" + first_plant_ws + ':' + last_plant_ws + "'" + '!' + formula_address + ')'
        else:
            pass

    current_column = 6
    column_number = 0

    months = []
    column_numbers = []
    column_addresses = []

    for month in period_to_model:
        
        base_column.copy(conso_m_ws.range(content_start_row, current_column))
        
        # Date headers creation
        date_cell = conso_m_ws.range((dates_row, current_column))
        date_cell.value = month
        date_cell.number_format = "mm-aaaa"
        date_cell.font.bold = True
        date_cell.api.Borders.LineStyle = 1
        date_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        date_cell.color = (221, 235, 247)
        
        # Month status header creation
        month_status_cell = conso_m_ws.range((month_status_row, current_column))
        month_status_cell.font.bold = True
        month_status_cell.api.Borders.LineStyle = 1
        month_status_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        month_status_cell.value = "Forecast"
        month_status_cell.color = (221, 235, 247)
        
        date_cell.column_width = 13
        
        months.append(month)
        column_numbers.append(column_number)
        column_address = xlsxwriter.utility.xl_col_to_name(current_column - 2) #//TODO EXPLICAR EL POR QUÉ DEL -2
        column_addresses.append(column_address)
        
        current_column += 1
        column_number += 1

    date_column_map = {
        'date': pd.Series(months),
        'columnAddress': pd.Series(column_addresses),
        'columnNumber': pd.Series(column_numbers)
    }

    model_dates = pd.DataFrame(data=date_column_map).set_index('date')
    model_months = model_dates.index
    last_period_column = int(model_dates['columnNumber'].tail(1).values)
    last_period_date = model_dates.tail(1).index.item()
    
    conso_m_ws.range('D:D').delete()

    # ---- Acumulación del resultado del año en balance ---- #

    cy_result_label_m = conso_m_ws.range('Conso_M_Results_CurrentYear')
    cy_result_values_m = conso_m_ws.range(
        (cy_result_label_m.row, 5),
        (cy_result_label_m.row, 5 + last_period_column)
    )

    pat_label_m = conso_m_ws.range('Conso_M_PAT')
    pat_values_m = conso_m_ws.range(
        (pat_label_m.row, 5),
        (pat_label_m.row, 5 + last_period_column)
    )

    for month in period_to_model:
        month_column = int(model_dates['columnNumber'][month])
        current_cell = cy_result_values_m[0, month_column]
        month_number = month.month
        if month_number == 1:
            current_cell.formula = '=' + pat_values_m[0, month_column].address
        else:
            pass
    
    # ---- Actualización resultado acumulado al cierre ---- #

    ret_earnings_label_m = conso_m_ws.range('Conso_M_Results_Accrued')
    ret_earnings_values_m = conso_m_ws.range(
            (ret_earnings_label_m.row, 5), 
            (ret_earnings_label_m.row, 5 + last_period_column)
        )

    py_result_range = ret_earnings_values_m.offset(
        row_offset = 1, 
        column_offset = -1
    )

    py_ret_earn_value = ret_earnings_values_m.offset(column_offset = -1)
        
    for month in period_to_model:
        month_number = month.month
        if month_number == 1:
            month_column = int(model_dates['columnNumber'][month])
            py_ret_earnings = py_ret_earn_value[0, month_column].address
            py_result = py_result_range[0, month_column].address
            ret_earnings_values_m[0, month_column].formula = \
                '=' + py_ret_earnings + '+' + py_result + ''
        else:
            pass

    app.calculate()
    print('Consolidado mensual completado.')

    # Generación del consolidado anual:

    print('Generando consolidado anual...')

    conso_a_template_ws.copy(
        after=new_wb_first_sheet, 
        name=conso_a_ws_name
    )    
    
    conso_a_template_last_row = str(conso_a_template_ws.used_range.last_cell.row)
    conso_a_ws = new_wb.sheets[conso_a_ws_name]

    model_start_year = model_start_date.year
    model_end_year = period_to_model.tail(1).item().year
    model_dates_monthly = pd.DataFrame(data=date_column_map).set_index('date')
    model_dates_annual = pd.DataFrame()
    conso_ws_column = 0

    for year in range(model_start_year, (model_end_year + 1)):
        data_list = []
        annual_df = model_dates_monthly.loc[str(year)]
        year_first_column = annual_df['columnAddress'].head(1).iloc[0]
        year_last_column = annual_df['columnAddress'].tail(1).iloc[0]
        
        data_list.append([
            year, 
            year_first_column, 
            year_last_column, 
            conso_ws_column
            ]
        )
        
        list_bridge_df = pd.DataFrame(
            data=data_list, 
            columns=[
                'year', 
                'year_first_column', 
                'year_last_column',
                'current_sheet_column'
            ]
        )
        
        model_dates_annual = pd.concat([model_dates_annual, list_bridge_df])
        conso_ws_column += 1

    model_dates_annual = model_dates_annual.set_index('year')
    model_years = model_dates_annual.index

    current_column = 6
    column_number = 0

    for year in model_years:
        
        conso_a_ws.range('D3:' + 'D' + conso_a_template_last_row).copy(conso_a_ws.range(3, current_column))
        
        # Date headers creation
        date_cell = conso_a_ws.range((dates_row, current_column))
        date_cell.value = str(year)
        date_cell.font.bold = True
        date_cell.api.Borders.LineStyle = 1
        date_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        date_cell.color = (221, 235, 247)

        # Year status header creation
        month_status_cell = conso_a_ws.range((month_status_row, current_column))
        month_status_cell.font.bold = True
        month_status_cell.api.Borders.LineStyle = 1
        month_status_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        month_status_cell.value = "Forecast"
        month_status_cell.color = (221, 235, 247)
        
        date_cell.column_width = 13
        
        current_column += 1
        column_number += 1

    conso_a_ws.range('D:D').delete()

    model_last_year = model_dates_annual.tail(1).index.item()
    last_year_column = int(model_dates_annual.tail(1)['current_sheet_column'])

    # Carga de data 

    ret_earnings_label_a = conso_a_ws.range('Conso_A_Results_Accrued')
    ret_earnings_values_a = conso_a_ws.range(
            (ret_earnings_label_a.row, 5), 
            (ret_earnings_label_a.row, 5 + last_period_column)
        )

    py_result_range = ret_earnings_values_a.offset(
        row_offset = 1, 
        column_offset = -1
    )

    py_ret_earn_value = ret_earnings_values_a.offset(column_offset = -1)

    # Generación de la sumatoria

    data_mapping = {
        'Conso_A_Revenues': 'Conso_M_Revenues',
        'Conso_A_DirectExp': 'Conso_M_DirectExp',
        'Conso_A_Other_operating_income': 'Conso_M_Other_operating_income',
        'Conso_A_Bad_debt_expenses': 'Conso_M_Bad_debt_expenses',
        'Conso_A_PnL': 'Conso_M_PnL',
        'Conso_A_Financing_Income': 'Conso_M_Financing_Income',
        'Conso_A_Financing_Cost_TP': 'Conso_M_Financing_Cost_TP',
        'Conso_A_Financing_Cost_Group': 'Conso_M_Financing_Cost_Group',
        'Conso_A_Financing_Cost_Holding': 'Conso_M_Financing_Cost_Holding',
        'Conso_A_Depreciation': 'Conso_M_Depreciation',
        'Conso_A_Share_of_profit': 'Conso_M_Share_of_profit',
        'Conso_A_Tax': 'Conso_M_Tax',
        'Conso_A_NCI': 'Conso_M_NCI',
        'Conso_A_Operating_CF': 'Conso_M_Operating_CF',
        'Conso_A_Investing_CF': 'Conso_M_Investing_CF',
        'Conso_A_Financing_CF': 'Conso_M_Financing_CF',
        'Conso_A_Current_Assets_Var': 'Conso_M_Current_Assets_Var',
        'Conso_A_Current_Liab_Var': 'Conso_M_Current_Liab_Var'
    }

    sum_items = [
       'Personnel cost',
       'Plant operation',
       'Maintenance',
       'Facility management',
       'Corporate services',
       'Taxes',
       'Non-personnel costs',
       'OPEX I',
       'Allocated cost (Division support) - Services received',
       'Allocated cost (Division support) - Services provided',
       'Allocated cost (Group support) - Services received',
       'Allocated cost (Group support) - Services provided',
       'OPEX II',
       'Bonus',
       'OPEX III',
       'OPEX IV',
       'IFRS 16 related reclassification (-)',
       'OPEX',
       'Third party',
       'Group',
       'Holding financing',
       'Depreciation on own assets',
       'Depreciation on aquired assets',
       'Inventories',
       'Trade receivables (incl accrued revenue) - 3rd party',
       'Trade receivables',
       'Short-term loans',
       'Trade payables to third parties',
       'Trade payables',
    ]

    for target, source in data_mapping.items():
        conso_a_labels = conso_a_ws.range(target)
        
        # El label de los rangos con una sola linea son almacenados como string.
        # Python posteriormente convierte este string en una lista donde cada
        # letra es un elemento. Para estos casos se crea previamente la lista
        # y se añade el string completo como un solo elemento.
        if conso_a_labels.count == 1:
            conso_a_items = []
            conso_a_items.append(conso_a_labels.value)
        else:
            conso_a_items = conso_a_labels.value

        conso_a_values = conso_a_ws.range(
            (conso_a_labels.row, 5),
            (conso_a_labels.row + len(conso_a_items), 5 + last_year_column)
        )
        conso_m_labels = conso_m_ws.range(source)

        # IDEM comentario linea 523.
        if conso_m_labels.count == 1: 
            conso_m_items = []
            conso_m_items.append(conso_m_labels.value)
        else:
            conso_m_items = conso_m_labels.value

        for item in conso_a_items:
            if item in conso_m_items:
                if item in sum_items:
                    continue
                else:
                    lookup_row = conso_m_items.index(item)
                    lookup_row_num = str(conso_m_labels[lookup_row, 0].row)
                    target_row = conso_a_items.index(item)
                    for year in model_years:
                        first_column = model_dates_annual['year_first_column'][year]
                        last_column = model_dates_annual['year_last_column'][year]
                        target_column = int(model_dates_annual['current_sheet_column'][year])
                        target_cell = conso_a_values[target_row, target_column]
                        target_cell.formula = '=SUM(' + "'" + conso_m_ws_name + "'!" + first_column + lookup_row_num + ":" + last_column + lookup_row_num + ')'
                    app.calculate()
            else:
                pass

    # ---- Actualización resultado acumulado al cierre ---- #
        
    for year in model_years:
        year_column = int(model_dates_annual['current_sheet_column'][year])
        py_ret_earnings = py_ret_earn_value[0, year_column].address
        py_result = py_result_range[0, year_column].address
        ret_earnings_values_a[0, year_column].formula = \
            '=' + py_ret_earnings + '+' + py_result + ''

    print('Consolidado anual completado.')
    
    #! CLOSING

    app.screen_updating = True
    new_wb.sheets[0].delete()
    new_wb.save(r'C:\VSCODE-Local\Projects\FinancialModel\output\mod_test.xlsx')
    print('Modelo completado.')
    templates_wb.close()
    new_wb.close()