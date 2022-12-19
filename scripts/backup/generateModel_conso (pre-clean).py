from datetime import datetime
from datetime import date
import calendar
from copy import copy
from time import strftime

import pandas as pd
import xlwings as xw

from helperFunctions import *
from workingCapital import *
from fixedAssets_conso import *

with xw.App() as app:

# Configuración del entorno de Excel para optimizar la velocidad del script:

    app.screen_updating = False
    app.visible = False
    app.interactive = False
    app.calculation = 'manual'

# Fechas para la generacion del modelo:

    model_start_date = date.fromisoformat('2023-01-01')
    all_periods = []

# Parametros financieros y fiscales:

    income_tax_rate = 0.28
    max_bin_deduction_rate = 0.70

# ---- Construcción del modelo ---- #

# Creación y configuración del libro:

    new_wb = xw.Book()
    new_wb.activate()
    new_wb_first_sheet = new_wb.sheets[0]

# Carga de la información sobre las plantas:

    print("Cargando información sobre las plantas...")
    plants_information = pd.read_excel(r'Projects\FinancialModel\input\plants.xlsx')
    plant_list = plants_information.index

# Carga del template y parámetros generales del mismo:

    print("Cargando y configurando el template...")
    templates_wb = xw.Book(r'Projects\FinancialModel\templates\modelTemplate.xlsx')
    plant_template_ws = templates_wb.sheets['PlantTemplate']
    plant_template_last_row = str(plant_template_ws.used_range.last_cell.row)

    conso_m_template_ws = templates_wb.sheets['Conso_M_Template']
    conso_m_ws_name = 'Conso - Monthly'
    conso_a_template_ws = templates_wb.sheets['Conso_A_Template']
    conso_a_ws_name = 'Conso - Annual'
    
# Cálculo de la última linea de cada rango:

    comm_inputs_labels = plant_template_ws.range('Plant_Commercial_Inputs')
    comm_inputs_lastrow = (
        comm_inputs_labels.row + (comm_inputs_labels.rows.count) - 1
    )
    
    tech_inputs_labels = plant_template_ws.range('Plant_Technical_Inputs')
    tech_inputs_lastrow = (
        tech_inputs_labels.row + (tech_inputs_labels.rows.count) - 1
    )

    reg_inputs_labels = plant_template_ws.range('Plant_Regulatory_Inputs')
    reg_inputs_lastrow = (
        reg_inputs_labels.row + (reg_inputs_labels.rows.count) - 1
    )

    opex_labels = plant_template_ws.range('Plant_OPEX')
    opex_lastrow = (opex_labels.row + (opex_labels.rows.count) - 1)

    current_fa_dep_labels = plant_template_ws.range('Plant_Active_Assets_Dep')
    current_fa_dep_lastrow = (
        current_fa_dep_labels.row + (current_fa_dep_labels.rows.count) - 1
    )

    capex_assets_dep_labels = plant_template_ws.range('Plant_CAPEX_Assets_Dep')
    capex_assets_dep_lastrow = (
        capex_assets_dep_labels.row + (capex_assets_dep_labels.rows.count) - 1
    )

    current_int_dep_labels = plant_template_ws.range('Plant_Active_Int_Dep')
    capex_int_dep_labels = plant_template_ws.range('Plant_CAPEX_Int_Dep')
    
    month_status_row = 4
    dates_row = 5
    content_start_row = 6

# Carga de data externa:

    print("Cargando datos externos...")

    comm_inputs = pd.read_excel(
        'Projects/FinancialModel/input/commercial_inputs.xlsx',
        parse_dates=['Mes'],
        index_col='Mes'
    )

    tech_inputs = pd.read_excel(
        'Projects/FinancialModel/input/technical_inputs.xlsx',
        parse_dates=['Mes'],
        index_col='Mes'
    )

    reg_inputs = pd.read_excel(
        'Projects/FinancialModel/input/reg_inputs.xlsx',
        parse_dates=['Mes'],
        index_col='Mes'
    )

    fin_inputs = pd.read_excel(
        'Projects/FinancialModel/input/fin_inputs.xlsx',
        parse_dates=['Mes'],
        index_col='Mes'
    )

    opex = pd.read_excel(
        'Projects/FinancialModel/input/opex.xlsx',
        parse_dates=['Mes'],
        index_col='Mes'
    )

    fixed_assets = pd.read_excel(
        'Projects\FinancialModel/input/fixed_assets_conso.xlsx'
    )

    # pool_weekly_sales = 'Projects\FinancialModel/input/Instrinsic_value_weekly_EKOREC_2023_2033_EnergySales (test).xlsx'

    prev_year_closing = pd.read_excel(
    "Projects\FinancialModel/input/closingPrevYear.xlsx", 
    index_col = 'description'
    )

    other_current_transactions = pd.read_excel(
    "Projects\FinancialModel/input/otherCurrentTransactions.xlsx", 
    index_col= 'date'
    )

    other_noncurrent_transactions = pd.read_excel(
    "Projects\FinancialModel/input/otherNonCurrTransactions.xlsx", 
    index_col= 'date'
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
        print('Generando columnas mensuales...')

        for month in model_timeframe:
            
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
            
            current_ws.range('D6:' + 'D' + plant_template_last_row).copy(
                current_ws.range(
                    content_start_row, 
                    current_column
                )
            )

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

        # Definición de los rangos de la hoja:

        comm_inputs_labels = current_ws.range('Plant_Commercial_Inputs')
        comm_inputs_values = current_ws.range(
            (comm_inputs_labels.row, 5), 
            (comm_inputs_lastrow, 5 + last_period_column)
        )
        comm_inputs_rows = range(0, comm_inputs_labels.rows.count)

        tech_inputs_labels = current_ws.range('Plant_Technical_Inputs')
        tech_inputs_values = current_ws.range(
            (tech_inputs_labels.row, 5), 
            (tech_inputs_lastrow, 5 + last_period_column)
        )
        tech_inputs_rows = range(0, tech_inputs_labels.rows.count)

        reg_inputs_labels = current_ws.range('Plant_Regulatory_Inputs')
        reg_inputs_values = current_ws.range(
            (reg_inputs_labels.row, 5), 
            (reg_inputs_lastrow, 5 + last_period_column)
        )
        reg_inputs_rows = range(0, reg_inputs_labels.rows.count)

        opex_labels = current_ws.range('Plant_OPEX')
        opex_values = current_ws.range(
            (opex_labels.row, 5),
            (opex_lastrow, 5 + last_period_column)
        )
        opex_rows = range(0, opex_labels.rows.count)

        current_fa_dep_labels = current_ws.range('Plant_Active_Assets_Dep')
        current_fa_dep_values = current_ws.range(
            (current_fa_dep_labels.row, 5),
            (current_fa_dep_lastrow, 5 + last_period_column)
        )
        current_fa_dep_rows = range(0, current_fa_dep_labels.rows.count)

        capex_range_start = current_ws.range('Plant_CAPEX_Start')
        
        capex_assets_dep_labels = current_ws.range('Plant_CAPEX_Assets_Dep')
        capex_assets_dep_values = current_ws.range(
            (capex_assets_dep_labels.row, 5),
            (capex_assets_dep_lastrow, 5 + last_period_column)
        )
        capex_assets_dep_rows = range(0, capex_assets_dep_labels.rows.count)

        ret_earnings_label = current_ws.range('Plant_Results_Accrued')
        ret_earnings_values = current_ws.range(
            (ret_earnings_label.row, 5), 
            (ret_earnings_label.row, 5 + last_period_column)
        )
        
        py_result_range = ret_earnings_values.offset(
            row_offset = 1, 
            column_offset = -1
        )
        
        py_ret_earn_value = ret_earnings_values.offset(column_offset = -1)

        income_tax_label = current_ws.range('Plant_TAX_IncomeTax')
        income_tax_values = current_ws.range(
            (income_tax_label.row, 5),
            (income_tax_label.row, 5 + last_period_column)
        )
        
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

        def_tax_asset_label = current_ws.range('Plant_TAX_DeferredAsset')
        opening_def_tax = current_ws.range((def_tax_asset_label.row, 4)).value
        def_tax_asset_values = current_ws.range(
            (def_tax_asset_label.row, 5), 
            (def_tax_asset_label.row, 5 + last_period_column)
        )

        dtlc_label = current_ws.range('Plant_TAX_DTLC')
        dtlc_values = current_ws.range(
            (dtlc_label.row, 5), 
            (dtlc_label.row, 5 + last_period_column))

        legal_reserve_label = current_ws.range('Plant_Equity_Reserves')
        legal_reserve_values = current_ws.range(
            (legal_reserve_label.row, 5),
            (legal_reserve_label.row, 5 + last_period_column)
        )

        share_capital_label = current_ws.range('Plant_Equity_ShareCapital')
        share_capital_values = current_ws.range(
            (share_capital_label.row, 5),
            (share_capital_label.row, 5 + last_period_column)
        )

        # //TODO Incorporar la amortización de los intangibles.
        
        current_int_dep_labels = current_ws.range('Plant_Active_Int_Dep')
        
        capex_int_dep_labels = current_ws.range('Plant_CAPEX_Int_Dep')

        # //TODO Agregar el resto de los rangos

        # Dataframes de la planta

        plant_comm_inputs = comm_inputs[comm_inputs['Plant code'] == plant_code]
        plant_tech_inputs = tech_inputs[tech_inputs['Plant code'] == plant_code]
        plant_reg_inputs = reg_inputs[reg_inputs['Plant code'] == plant_code]
        plant_fin_inputs = fin_inputs[fin_inputs['Plant code'] == plant_code]
        plant_opex = opex[opex['Plant code'] == plant_code]
        plant_fx_assets = fixed_assets[fixed_assets['Plant code'] == plant_code]

        # //TODO Agregar el resto de los dataframes

        # Funciones

        capex, curr_assets_amort, capex_amortization = calculate_fixed_assets(
            plant_fx_assets, 
            model_start_date
        )

        capex_bridge_list = []
        
        for asset in capex.index:
            start_date = capex['start_date'][asset]
            historic_cost = capex['historic_cost'][asset]
            capex_bridge_list.append([
                plant_name,
                start_date,
                historic_cost
                ]
            )
           
        capex_bridge_df = pd.DataFrame(
            capex_bridge_list,
            columns = ['plant_name', 'start_date', 'historic_cost']
        )

        all_capex_monthly = pd.concat(
            [all_capex_monthly, capex_bridge_df], 
            ignore_index=True
        )

        # //TODO Aquí van las funciones de los cálculos
        
        # Carga de la data en la hoja

        
        #//TODO CARGAR LAS CIFRAS DEL AÑO ANTERIOR AQUI (antes de bin_balance).

        bin_balance = opening_def_tax / income_tax_rate
        
        # Inputs

        print('Cargando inputs comerciales...')

        for row in comm_inputs_rows:
            try:
                row_label = comm_inputs_labels[row, 0].value
                for month in plant_model_months:
                    month_column = int(plant_model_dates['columnNumber'][month])
                    amount = plant_comm_inputs[row_label][month]
                    comm_inputs_values[row, month_column].value = amount
            
            except KeyError:
                pass

        print('Cargando inputs tecnicos...')

        for row in tech_inputs_rows:
            try:
                row_label = tech_inputs_labels[row, 0].value
                for month in plant_model_months:
                    month_column = int(plant_model_dates['columnNumber'][month])
                    amount = plant_tech_inputs[row_label][month]
                    tech_inputs_values[row, month_column].value = amount
            
            except KeyError:
                pass

        print('Cargando inputs regulatorios...')

        for row in reg_inputs_rows:
            try:
                row_label = reg_inputs_labels[row, 0].value
                for month in plant_model_months:
                    month_column = int(plant_model_dates['columnNumber'][month])
                    amount = plant_reg_inputs[row_label][month]
                    reg_inputs_values[row, month_column].value = amount
            
            except KeyError:
                pass

        print('Cargando OPEX...')

        for row in opex_rows:
            try:
                row_label = opex_labels[row, 0].value
                for month in plant_model_months:
                    month_column = int(plant_model_dates['columnNumber'][month])
                    amount = plant_opex[row_label][month]
                    opex_values[row, month_column].value = amount

            except KeyError:
                pass

        print('Cargando CAPEX...')

        number_of_assets = len(capex)
        capex_start_row = capex_range_start.row
        
        # Se resta uno porque la primera fila también tendrá contenido.
        capex_range_end = capex_start_row + number_of_assets - 1

        capex_range = current_ws.range('' 
            + str(capex_start_row) 
            + ':' 
            + str(capex_range_end) 
            + ''
        )

        capex_range.insert()

        # Rango de celdas donde van los nombres de los activos.
        capex_labels = current_ws.range(
            (capex_start_row, 2), 
            (capex_range_end, 3)
        ) 
        
        capex_values_range = current_ws.range(
            (capex_start_row, 5), 
            (capex_range_end, last_period_column)
        )
        
        row = 0 # Para el control del loop que sigue a continuación

        # Este loop inserta los CAPEX en la hoja
        
        for asset in capex.index:
            capex_labels[row, 0].value = capex['name'][asset]
            capex_labels[row, 1].value = "€"
            purchase_date = capex['start_date'][asset]
            purchase_date = change_to_last_day(purchase_date)
            month_column = int(plant_model_dates['columnNumber'][purchase_date])
            amount = capex['historic_cost'][asset]
            capex_values_range[row, month_column].value = amount
            
            row += 1

        # Las siguientes dos declaraciones son para ajustar visualmente la 
        # tabla de CAPEX al nuevo contenido recien agregado.

        top_blank_row = current_ws.range('' 
            + str(capex_start_row) 
            + ':' 
            + str(capex_start_row) 
            + ''
        ).offset(row_offset = -1).delete()
        
        bottom_blank_row = current_ws.range('' 
            + str(capex_range_end) 
            + ':' 
            + str(capex_range_end) 
            + ''
        ).delete()

        print('Cargando depreciación/amortización...')

        # Activos existentes

        try: 
        #//! Este bloque se empleó para ignorar las plantas que no tienen
        #//! activos fijos. Es mejor trasladarlo como una condición a la 
        #//! ejecución de la función para que nisiquiera la ejecute.
            # Cortar el dataframe hasta la fecha final del modelo
            curr_asset_last_amort_period = curr_assets_amort.tail(1).index.item()
            
            if curr_asset_last_amort_period > last_period_date:
                curr_assets_amort = curr_assets_amort[:last_period_date]

            else:
                pass
            
            amort_periods = curr_assets_amort.index

            for row in current_fa_dep_rows:
                for period in amort_periods:
                    amount = curr_assets_amort[period]
                    new_period = change_to_last_day(period)
                    month_column = int(plant_model_dates['columnNumber'][new_period])
                    current_fa_dep_values[0, month_column].value = amount * (-1)

        except ValueError:
            pass

        # CAPEX

        try:
        #//! Este bloque se empleó para ignorar las plantas que no tienen
        #//! activos fijos. Es mejor trasladarlo como una condición a la 
        #//! ejecución de la función para que nisiquiera la ejecute.
            # Cortar el dataframe hasta la fecha final del modelo

            capex_amort_last_period = capex_amortization.tail(1).index.item()

            if capex_amort_last_period > last_period_date:
                capex_amortization = capex_amortization[:last_period_date]

            else:
                pass

            amort_periods = capex_amortization.index
            
            for row in capex_assets_dep_rows:
                for period in amort_periods:
                    amount = capex_amortization[period]
                    new_period = change_to_last_day(period)
                    month_column = int(plant_model_dates['columnNumber'][new_period])
                    capex_assets_dep_values[0, month_column].value = amount * (-1)

        except ValueError:
            pass

        # ---- Actualización resultado acumulado al cierre ---- #

        print('Actualizando resultados acumulados...')
        
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

        # ---- Acumulación del resultado del año en balance ---- #

        print('Actualizando resultados del periodo...')
        
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

        # ---- Income Tax & BINs ---- #
        
        app.calculate() # Para que se actualicen los valores de las celdas de resultado.

        print('Calculando impuesto renta y BINs...')
        
        for month in plant_model_months:
            month_number = month.month
            if month_number == 12:
                month_column = int(plant_model_dates['columnNumber'][month])
                accrued_result = cy_result_values[0, month_column].value
                pnl_tax_cell = income_tax_values[0, month_column]
                def_tax_asset_cell = def_tax_asset_values[0, month_column]
                dtlc_cell = dtlc_values[0, month_column]

                if accrued_result < 0:
                    bin_balance += (accrued_result) * - 1
                    def_tax_amount = bin_balance * income_tax_rate
                    def_tax_asset_cell.value = def_tax_amount
                    dtlc_cell.value = def_tax_amount

                elif accrued_result > 0:
                    if bin_balance > 0:
                        deduction_value = accrued_result * max_bin_deduction_rate
                        
                        if deduction_value < bin_balance:
                            pass    
                        else:
                            deduction_value = bin_balance

                        deferred_tax = deduction_value * income_tax_rate
                        pnl_tax_cell.value = \
                            ((accrued_result - deduction_value) * income_tax_rate) * (-1)
                        def_tax_asset_cell.value = \
                            (def_tax_asset_cell.value) - deferred_tax
                        dtlc_cell.value = (dtlc_cell.value) - deferred_tax
                        bin_balance -= deduction_value

                    else:
                        pnl_tax_cell.value = ((accrued_result) * income_tax_rate) * (-1)    
                else:
                    pass
            else:
                pass

        # ---- Actualización reserva legal ---- #

        print('Actualizando reserva legal...')
        #//TODO Cambiar modificar la variable legal_reserve para que tenga el .value al final.

        for month in plant_model_months:
            month_column = int(plant_model_dates['columnNumber'][month])
            month_number = month.month
            if month_column == 0:
                pass
            else:
                if month_number == 1:
                    year_result = cy_result_values[0, month_column].offset(column_offset = -1).value
                    legal_reserve = legal_reserve_values[0, month_column].offset(column_offset = -1)
                    share_capital = \
                        share_capital_values[0, month_column].offset(column_offset = -1).value
                    if year_result > 0:
                        if legal_reserve.value < (int(share_capital) * 0.20):
                            add_to_reserve = (py_result_range[0, month_column].value) * 0.10
                            new_reserve = legal_reserve.value + add_to_reserve
                            if new_reserve > (int(share_capital) * 0.20):
                                delta = (int(share_capital) * 0.20) - new_reserve
                                if delta < 0:
                                    add_to_reserve += delta
                                    legal_reserve.value += add_to_reserve
                                    py_result_range[0, month_column].value -= add_to_reserve
                                else:
                                    pass
                            else:
                                legal_reserve.value += add_to_reserve
                                py_result_range[0, month_column].value -= add_to_reserve
                        else:
                            pass
                    else:
                        pass
                else:
                    pass
        
        # //! FINAL LOOP PLANTA
        ws_counter += 1
        print('Pestaña ' + '"' + plant_name + '"' + ' completada.')

# ----------------------------------------------------------------------------#
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

    base_column = conso_m_ws.range('D6:' + 'D' + conso_m_template_last_row)
    first_plant_ws = new_wb.sheets[2].name
    last_plant_ws = new_wb.sheets[ws_counter + 1].name # Se debe sumar uno porque este contador comienza en cero. Buscar una solucion mas elegante.
    
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

    for month in period_to_model:
        
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

        base_column.copy(conso_m_ws.range(content_start_row, current_column))
        
        date_cell.column_width = 13
        
        months.append(month)
        column_numbers.append(column_number)
        
        current_column += 1
        column_number += 1

    date_column_map = {
        'date': pd.Series(months), 
        'columnNumber':pd.Series(column_numbers)
    }

    model_dates = pd.DataFrame(data=date_column_map).set_index('date')
    model_months = model_dates.index
    last_period_column = int(model_dates.tail(1).values)
    last_period_date = model_dates.tail(1).index.item()
    
    conso_m_ws.range('D:D').delete()

    # Agrupacion de CAPEX y determinación de numero de plantas

    all_capex_grouped = all_capex_monthly.groupby(
        ['plant_name', 'start_date']
        ).sum()

    plants_with_capex = all_capex_monthly['plant_name'].drop_duplicates().reset_index()
    num_of_plants_w_capex = len(plants_with_capex)

    # Definición de rangos

    capex_start_row = conso_m_ws.range('Conso_M_CAPEX_Start')
    capex_end_row = capex_start_row.row + (num_of_plants_w_capex - 1) # El - 1 es porque la primera fila tendrá contenido.

    capex_range = conso_m_ws.range(''
        + str(capex_start_row.row)
        + ':'
        + str(capex_end_row)
        + ''
    )

    capex_range.insert()

    capex_labels = conso_m_ws.range(
        (capex_start_row.row, capex_start_row.column),
        (capex_end_row, capex_start_row.column) 
    )

    capex_values_m = conso_m_ws.range(
        (capex_start_row.row, 5),
        (capex_end_row, 5 + last_period_column)
    )

    pat_label_m = conso_m_ws.range('Conso_M_PAT')
    pat_values_m = conso_m_ws.range(
        (pat_label_m.row, 5),
        (pat_label_m.row, 5 + last_period_column)
    )
    
    cy_result_label_m = conso_m_ws.range('Conso_M_Results_CurrentYear')
    cy_result_values_m = conso_m_ws.range(
        (cy_result_label_m.row, 5),
        (cy_result_label_m.row, 5 + last_period_column)
    )
    
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

    # Inserción CAPEX por planta:

    capex_labels_area = range(0, capex_labels.rows.count)

    try:
        for row in capex_labels_area:
            plant_name = plants_with_capex['plant_name'][row]
            current_row = capex_labels[row, 0]
            current_row.value = plant_name
            capex_by_month = all_capex_grouped.loc[plant_name]
            capex_months = capex_by_month.index
            for month in capex_months:
                month_number = int(capex_by_month.tail(1).index.item().month)
                column_number = int(model_dates['columnNumber'][month_number])
                value_cell = capex_values_m[row, column_number]
                value_cell.value = capex_by_month['historic_cost'][month]
    except KeyError:
        pass

    top_blank_row = conso_m_ws.range('' 
        + str(capex_labels.row) 
        + ':' 
        + str(capex_labels.row) 
        + ''
    ).offset(row_offset = -1).delete()
    
    bottom_blank_row = conso_m_ws.range('' 
        + str(capex_end_row) 
        + ':' 
        + str(capex_end_row) 
        + ''
    ).delete()

    # ---- Acumulación del resultado del año en balance ---- #

    print('Actualizando resultados del periodo...')
        
    for month in period_to_model:
        month_column = int(model_dates['columnNumber'][month])
        current_cell = cy_result_values_m[0, month_column]
        month_number = month.month
        if month_number == 1:
            current_cell.formula = '=' + pat_values_m[0, month_column].address
        else:
            pass
    
    # ---- Actualización resultado acumulado al cierre ---- #

    print('Actualizando resultados acumulados...')
        
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

# ----------------------------------------------------------------------------#
# Generación de la hoja con el consolidado anual:

    print('Generando consolidado anual...')

    app.status_bar = "Creando consolidado anual..."

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
        annual_df = model_dates_monthly[str(year)]
        year_first_column = int(annual_df.head(1).iloc[0])
        year_last_column = int(annual_df.tail(1).iloc[0])
        
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
        
        conso_a_ws.range('D6:' + 'D' + conso_a_template_last_row).copy(conso_a_ws.range(6, current_column))
        date_cell.column_width = 13
        
        current_column += 1
        column_number += 1

    conso_a_ws.range('D:D').delete()

    model_last_year = model_dates_annual.tail(1).index.item()
    last_year_column = int(model_dates_annual.tail(1)['current_sheet_column'])

    # Carga de data 

    # Definicion de rangos

    conso_m_rev_labels = conso_m_ws.range('Conso_M_Revenues')
    conso_m_rev_items = conso_m_rev_labels.value
    
    conso_a_rev_labels = conso_a_ws.range('Conso_A_Revenues')
    conso_a_rev_items = conso_a_rev_labels.value
    conso_a_rev_values = conso_a_ws.range(
        (conso_a_rev_labels.row, 5),
        (conso_a_rev_labels.row + len(conso_a_rev_items), 5 + last_year_column)
    )

    conso_m_dex_labels = conso_m_ws.range('Conso_M_DirectExp')
    conso_m_dex_items = conso_m_dex_labels.value
    
    conso_a_dex_labels = conso_a_ws.range('Conso_A_DirectExp')
    conso_a_dex_items = conso_a_dex_labels.value
    conso_a_dex_values = conso_a_ws.range(
        (conso_a_dex_labels.row, 5),
        (conso_a_dex_labels.row + len(conso_a_dex_items), 5 + last_year_column)
    )

    conso_m_pnl_labels = conso_m_ws.range('Conso_M_PnL')
    conso_m_pnl_items = conso_m_pnl_labels.value
    
    conso_a_pnl_labels = conso_a_ws.range('Conso_A_PnL')
    conso_a_pnl_items = conso_a_pnl_labels.value
    conso_a_pnl_values = conso_a_ws.range(
        (conso_a_pnl_labels.row, 5),
        (conso_a_pnl_labels.row + len(conso_a_pnl_items), 5 + last_year_column)
    )

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

    for item in conso_a_rev_items:
        if item in conso_m_rev_items:
            lookup_row = conso_m_rev_items.index(item)
            target_row = conso_a_rev_items.index(item)
            for year in model_years:
                first_column = int(model_dates_annual['year_first_column'][year])
                last_column = int(model_dates_annual['year_last_column'][year])
                target_column = int(model_dates_annual['current_sheet_column'][year])
                sum_range = conso_m_ws.range(
                    (conso_m_rev_labels[lookup_row, 0].row, 5 + first_column),
                    (conso_m_rev_labels[lookup_row, 0].row, 5 + last_column)
                )
                amount = sum(sum_range.value)
                target_cell = conso_a_rev_values[target_row, target_column]
                target_cell.value = amount
        else:
            pass

    for item in conso_a_dex_items:
        if item in conso_m_dex_items:
            lookup_row = conso_m_dex_items.index(item)
            target_row = conso_a_dex_items.index(item)
            for year in model_years:
                first_column = int(model_dates_annual['year_first_column'][year])
                last_column = int(model_dates_annual['year_last_column'][year])
                target_column = int(model_dates_annual['current_sheet_column'][year])
                sum_range = conso_m_ws.range(
                    (conso_m_dex_labels[lookup_row, 0].row, 5 + first_column),
                    (conso_m_dex_labels[lookup_row, 0].row, 5 + last_column)
                )
                amount = sum(sum_range.value)
                target_cell = conso_a_dex_values[target_row, target_column]
                target_cell.value = amount
        else:
            pass

    for item in conso_a_pnl_items:
        if item in conso_m_pnl_items:
            lookup_row = conso_m_pnl_items.index(item)
            target_row = conso_a_pnl_items.index(item)
            for year in model_years:
                first_column = int(model_dates_annual['year_first_column'][year])
                last_column = int(model_dates_annual['year_last_column'][year])
                target_column = int(model_dates_annual['current_sheet_column'][year])
                sum_range = conso_m_ws.range(
                    (conso_m_pnl_labels[lookup_row, 0].row, 5 + first_column),
                    (conso_m_pnl_labels[lookup_row, 0].row, 5 + last_column)
                )
                amount = sum(sum_range.value)
                target_cell = conso_a_pnl_values[target_row, target_column]
                target_cell.value = amount
        else:
            pass

    
    all_capex_monthly['start_date'] = all_capex_monthly['start_date'].apply(lambda x: x.year)
    all_capex_grouped = all_capex_monthly.groupby(['plant_name', 'start_date']).sum()

    # CAPEX
    
    capex_start_row = conso_a_ws.range('Conso_A_CAPEX_Start')
    capex_end_row = capex_start_row.row + (num_of_plants_w_capex - 1) # El - 1 es porque la primera fila tendrá contenido.

    capex_range = conso_a_ws.range(''
        + str(capex_start_row.row)
        + ':'
        + str(capex_end_row)
        + ''
    )

    capex_range.insert()

    capex_labels = conso_a_ws.range(
        (capex_start_row.row, capex_start_row.column),
        (capex_end_row, capex_start_row.column) 
    )

    capex_values_a = conso_a_ws.range(
        (capex_start_row.row, 5),
        (capex_end_row, 5 + last_period_column)
    )

    capex_labels_area = range(0, capex_labels.rows.count)

    try:
        for row in capex_labels_area:
            plant_name = plants_with_capex['plant_name'][row]
            current_row = capex_labels[row, 0]
            current_row.value = plant_name
            capex_by_year = all_capex_grouped.loc[plant_name]
            capex_years = capex_by_year.index
            for year in capex_years:
                # year = int(capex_by_year.tail(1).index.item().year)
                column_number = int(model_dates_annual['current_sheet_column'][year])
                value_cell = capex_values_a[row, column_number]
                value_cell.value = capex_by_year['historic_cost'][year]
    except KeyError:
        pass

    top_blank_row = conso_a_ws.range('' 
        + str(capex_labels.row) 
        + ':' 
        + str(capex_labels.row) 
        + ''
    ).offset(row_offset = -1).delete()
    
    bottom_blank_row = conso_a_ws.range('' 
        + str(capex_end_row) 
        + ':' 
        + str(capex_end_row) 
        + ''
    ).delete()
    
    # Carga depreciacion/amortizacion

    

    # ---- Actualización resultado acumulado al cierre ---- #

    print('Actualizando resultados acumulados...')
        
    for year in model_years:
        year_column = int(model_dates_annual['current_sheet_column'][year])
        py_ret_earnings = py_ret_earn_value[0, year_column].address
        py_result = py_result_range[0, year_column].address
        ret_earnings_values_a[0, year_column].formula = \
            '=' + py_ret_earnings + '+' + py_result + ''

    print('Consolidado anual completado.')

# ----------------------------------------------------------------------------#
# Guardar libro y cierre de la instancia de Excel:

    app.screen_updating = True
    new_wb.sheets[0].delete()
    new_wb.save(r'Projects\FinancialModel\output\modelo_consolidado.xlsx')
    print('Modelo completado.')
    templates_wb.close()
    new_wb.close()