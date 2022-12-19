from datetime import date

import pandas as pd
import xlwings as xw

from helperFunctions import *
from workingCapital import *
from fixedAssets import *

with xw.App() as app:

    # ---- Parametros de la aplicación ---- #
    app.interactive = False
    # app.screen_updating = False
    
    # ---- Parametros generales ---- #

    monthly_ws_name = 'Mensual'
    annual_ws_name = 'Anual'
    month_status_row = 4
    dates_row = 5
    values_first_row = 6
    income_tax_rate = 0.28
    max_bin_deduction_rate = 0.70

    # ---- Fecha de la generacion del modelo ---- #

    model_start_date = date.fromisoformat('2023-01-01')

    # ---- Pestaña de la planta ---- #

    new_wb = xw.Book()

    app.status_bar = "Cargando template..."
    templates_wb = xw.Book(
        r'Projects\FinancialModel\templates\simpleTemplate.xlsm'
    )

    app.status_bar = "Cargando macros..."
    complete_alert_macro = templates_wb.macro('completeAlert')
    
    app.status_bar = "Cargando datos de la planta..."
    plant_information = pd.read_excel(
        r'Projects\FinancialModel\input\plants_simple.xlsx'
    )
    
    monthly_template_ws = templates_wb.sheets['monthlyTemplate']
    template_last_row = str(monthly_template_ws.used_range.last_cell.row)

    new_wb_first_sheet = new_wb.sheets[0]
    plant_name = plant_information.loc[0, 'name']
    plant_code = plant_information.loc[0, 'it_code']
    plant_eol_year = str((plant_information.loc[0, 'reg_per_end']).year)

    monthly_template_ws.copy(
        after=new_wb_first_sheet, 
        name=monthly_ws_name
    )

    monthly_ws = new_wb.sheets[monthly_ws_name]
    monthly_ws.range('B2').value = plant_name
    monthly_ws.range('B3').value = plant_code
    monthly_ws.range('B4').value = 'End of regulatory life: ' + plant_eol_year

    app.status_bar = "Calculando período regulatorio restante..."
    
    reg_period_end = plant_information.loc[0, 'reg_per_end']
    model_timeframe = pd.date_range(
        start=model_start_date, 
        end=reg_period_end, 
        freq='M'
    ).to_series().index

    current_column = 6
    column_number = 0
    
    months = []
    column_numbers = []

    app.status_bar = "Creando columnas mensuales..."
    
    for month in model_timeframe:
        
        # Date headers creation
        date_cell = monthly_ws.range((dates_row, current_column))
        date_cell.value = month
        date_cell.number_format = "mm-aaaa"
        date_cell.font.bold = True
        date_cell.api.Borders.LineStyle = 1
        date_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        date_cell.color = (221, 235, 247)

        # Month status header creation
        month_status_cell = monthly_ws.range((month_status_row, current_column))
        month_status_cell.font.bold = True
        month_status_cell.api.Borders.LineStyle = 1
        month_status_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        month_status_cell.value = "Forecast"
        month_status_cell.color = (221, 235, 247)
        
        monthly_ws.range('D6:' + 'D' + template_last_row).copy(
            monthly_ws.range(
                values_first_row, 
                current_column
            )
        )
        
        date_cell.column_width = 11
        
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

    monthly_ws.range('D:D').delete()

    # ---- Carga de la data ---- #

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

    app.status_bar = "Cargando informacion financiera..."
    
    model_inputs_df = pd.read_excel(
        'Projects\FinancialModel/input/monthlyInputs.xlsx', 
        index_col='DiaLocal', 
        parse_dates=True
    )
    
    pool_weekly_sales = 'Projects\FinancialModel/input/Instrinsic_value_weekly_EKOREC_2023_2033_EnergySales (test).xlsx'
    
    fixed_assets = 'Projects\FinancialModel/input/fixed_assets.xlsx'
    
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
    
    app.status_bar = "Calculando transacciones..."
    pool_incomes = calculate_pool_income(pool_weekly_sales, model_timeframe)
    heat_incomes = calculate_heat_income(model_inputs_df, model_timeframe)
    ro_incomes = calculate_ro_income(model_inputs_df, model_timeframe)
    gas_cost = calculate_30_day_items(model_inputs_df, 'Gas', model_timeframe)
    capex, curr_assets_amort, capex_amortization = calculate_fixed_assets(
        fixed_assets, 
        model_start_date
    )
    
    oth_curr_items_m, oth_curr_items_a = calculate_other_items(
        other_current_transactions
    )
    
    oth_non_curr_items_m, oth_non_curr_items_a = calculate_other_items(
        other_noncurrent_transactions
    )
    
    # ---- PnL ---- #
    app.status_bar = "Cargando P&L..."
    
    pnl_labels_range = monthly_ws.range((53, 2), (104, 2))
    pnl_values_range = monthly_ws.range(
        (53, 5), 
        (104, 5 + last_period_column)
    )
    pnl_rows = range(0, pnl_labels_range.rows.count)

    # Cifras del año anterior
    for row in pnl_rows:
        try:
            row_label = pnl_labels_range[row].value
            row_number = pnl_labels_range[row].row
            amount = prev_year_closing['amount'][row_label]
            monthly_ws.range((row_number, 4)).value = amount
        
        except KeyError:
            pass

    # Forecast
    for row in pnl_rows:
        row_label = pnl_labels_range[row, 0].value
        input_data_column = label_mapping.get(row_label)
        
        if input_data_column == None:
            pass
        else:
            for month in model_months:
                month_column = int(model_dates['columnNumber'][month])
                amount = model_inputs_df[input_data_column][month]
                pnl_values_range[row, month_column].value = amount
                    
    # ---- Cashflow ---- #
    app.status_bar = "Cargando Cashflow..."
    
    cf_labels_range = monthly_ws.range((108, 2), (138, 2))
    cf_values_range = monthly_ws.range(
        (108, 5), 
        (138, 5 + last_period_column)
    )

    cashflow_rows = range(0, cf_labels_range.rows.count)

    for row in cashflow_rows:
        row_label = cf_labels_range[row, 0].value
        
        for month in model_months:
                try:
                    month_column = int(model_dates['columnNumber'][month])
                    amount = int(
                        oth_non_curr_items_m.loc[month, row_label]
                    )
                    cf_values_range[row, month_column].value = amount
                
                except KeyError:
                    pass
    
    # ---- Carga de variacion de capital de trabajo ---- #
    app.status_bar = "Cargando variaciones de WC..."
    
    wc_labels_range = monthly_ws.range((144, 2), (177, 2))
    wc_values_range = monthly_ws.range(
        (144, 5), 
        (177, 5 + last_period_column)
    )

    working_capital_rows = range(0, wc_labels_range.rows.count)

    for row in working_capital_rows:
        row_label = wc_labels_range[row, 0].value

        if row_label == 'Pool':
            for month in model_months:
                month_column = int(model_dates['columnNumber'][month])
                amount = pool_incomes['poolVariation'][month]
                wc_values_range[row, month_column].value = amount
        
        elif row_label == 'Heat':
            for month in model_months:
                month_column = int(model_dates['columnNumber'][month])
                amount = heat_incomes['heatVariation'][month]
                wc_values_range[row, month_column].value = amount
        
        elif row_label == 'RO':
            for month in model_months:
                month_column = int(model_dates['columnNumber'][month])
                amount = ro_incomes['roVariation'][month]
                wc_values_range[row, month_column].value = amount
        
        elif row_label == 'Gas':
            for month in model_months:
                month_column = int(model_dates['columnNumber'][month])
                amount = gas_cost['variation'][month]
                wc_values_range[row, month_column].value = amount
        else:
            for month in model_months:
                try:
                    month_column = int(model_dates['columnNumber'][month])
                    amount = int(oth_curr_items_m.loc[month, row_label])
                    wc_values_range[row, month_column].value = amount
                
                except KeyError:
                    pass
        
    # ---- Activos fijos ---- #

    # Carga de CAPEX
    app.status_bar = "Cargando CAPEX..."
    
    # Se parte por la 187 para evitar que se extienda el fondo gris del tìtulo.
    capex_range_start = 187 
    number_of_assets = len(capex)
    
    # Se resta uno porque la 187 es la primera fila y también tendrá contenido.
    capex_range_end = capex_range_start + number_of_assets - 1 
    
    capex_range = monthly_ws.range('' 
        + str(capex_range_start) 
        + ':' 
        + str(capex_range_end) 
        + ''
    )
    
    capex_range.insert()
    
    # Rango de celdas donde van los nombres de los activos.
    capex_name_range = monthly_ws.range(
        (capex_range_start, 2), 
        (capex_range_end, 3)
    ) 
    
    capex_values_range = monthly_ws.range(
        (capex_range_start, 5), 
        (capex_range_end, last_period_column)
    )
    
    row = 0 # Para el control del loop que sigue a continuación

    # Este loop inserta los CAPEX en la hoja
    
    for asset in capex.index:
        capex_name_range[row, 0].value = capex['name'][asset]
        capex_name_range[row, 1].value = "€"
        purchase_date = capex['start_date'][asset]
        purchase_date = change_to_last_day(purchase_date)
        month_column = int(model_dates['columnNumber'][purchase_date])
        amount = capex['historic_cost'][asset]
        capex_values_range[row, month_column].value = amount
        
        row += 1

    # Las siguientes dos declaraciones son para ajustar visualmente la tabla 
    # de CAPEX al nuevo contenido recien agregado.

    top_blank_row = monthly_ws.range('' 
        + str(capex_range_start) 
        + ':' 
        + str(capex_range_start) 
        + ''
    ).offset(row_offset = -1).delete()
    
    bottom_blank_row = monthly_ws.range('' 
        + str(capex_range_end) 
        + ':' 
        + str(capex_range_end) 
        + ''
    ).delete()
    
    # Carga de depreciacion de activos existentes
    app.status_bar = "Cargando depreciación/amortizacion..."
    
    curr_asset_amort_range = monthly_ws.range(
        (85, 5), 
        (85, 5 + last_period_column)
    )

    # Cortar el dataframe hasta la fecha final del modelo

    curr_asset_last_amort_period = curr_assets_amort.tail(1).index.item()
    
    if curr_asset_last_amort_period > last_period_date:
        curr_assets_amort = curr_assets_amort[:last_period_date]

    else:
        pass

    amort_periods = curr_assets_amort.index
    
    for period in amort_periods:
        amount = curr_assets_amort[period]
        new_period = change_to_last_day(period)
        month_column = int(model_dates['columnNumber'][new_period])
        curr_asset_amort_range[0, month_column].value = amount * (-1)

    # Carga de depreciacion del CAPEX

    capex_amort_range = monthly_ws.range(
        (91, 5), 
        (91, 5 + last_period_column)
    )

    # Cortar el dataframe hasta la fecha final del modelo

    capex_amort_last_period = capex_amortization.tail(1).index.item()

    if capex_amort_last_period > last_period_date:
        capex_amortization = capex_amortization[:last_period_date]

    else:
        pass

    amort_periods = capex_amortization.index

    for period in amort_periods:
        amount = capex_amortization[period]
        new_period = change_to_last_day(period)
        month_column = int(model_dates['columnNumber'][new_period])
        capex_amort_range[0, month_column].value = amount * (-1)

    # ---- Balance ---- #
    app.status_bar = "Cargando balance..."
    
    lines_by_capex = (number_of_assets) - 2 
    bs_labels_range = monthly_ws.range(
        (192 + lines_by_capex, 2), 
        (259 + lines_by_capex, 2)
    )

    balance_sheet_rows = range(0, bs_labels_range.rows.count)

    # Cifras del año anterior
    for row in balance_sheet_rows:
        try:
            row_label = bs_labels_range[row].value
            row_number = bs_labels_range[row].row
            amount = prev_year_closing['amount'][row_label]
            monthly_ws.range((row_number, 4)).value = amount
        
        except KeyError:
            pass

    # ---- Actualización resultado acumulado al cierre ---- #
    app.status_bar = "Actualizando resultados acumulados..."
    
    ret_earnings_range = monthly_ws.range(
        ((227 + lines_by_capex), 5), 
        ((227 + lines_by_capex), 5 + last_period_column)
    )
    
    py_result_range = ret_earnings_range.offset(
        row_offset = 1, 
        column_offset = -1
    )
    
    py_ret_earn_range = ret_earnings_range.offset(column_offset = -1)
    
    for month in model_months:
        month_number = month.month
        if month_number == 1:
            month_column = int(model_dates['columnNumber'][month])
            py_ret_earnings = py_ret_earn_range[0, month_column].address
            py_result = py_result_range[0, month_column].address
            ret_earnings_range[0, month_column].formula = \
                '=' + py_ret_earnings + '+' + py_result + ''
        else:
            pass

    # ---- Acumulación del resultado del año en balance ---- #
    app.status_bar = "Actualizando resultados del periodo..."
    
    cy_result_range = monthly_ws.range(
        ((228 + lines_by_capex), 5), 
        ((228 + lines_by_capex), 5 + last_period_column)
    )

    for month in model_months:
        month_column = int(model_dates['columnNumber'][month])
        current_cell = cy_result_range[0, month_column]
        month_number = month.month
        if month_number == 1:
            current_cell.formula = \
                '=' + pnl_values_range[51, month_column].address + ''
        else:
            accrued_result = current_cell.offset(column_offset = -1)
            current_cell.formula = \
                '=' + accrued_result.address + '+' + \
                    pnl_values_range[51, month_column].address + ''

    # ---- Income Tax & BINs ---- #

    app.status_bar = "Calculando impuesto renta y actualizando BINs..."
    
    opening_def_tax = monthly_ws.range(((199 + lines_by_capex), 4)).value
    def_tax_asset_range = monthly_ws.range(
        ((199 + lines_by_capex), 5), 
        ((199 + lines_by_capex), 5 + last_period_column)
    )
    dtlc_range = monthly_ws.range(
        ((230 + lines_by_capex, 5)), 
        (230 + lines_by_capex, 5 + last_period_column))
    
    bin_balance = opening_def_tax / income_tax_rate
    
    for month in model_months:
        month_number = month.month
        if month_number == 12:
            month_column = int(model_dates['columnNumber'][month])
            accrued_result = cy_result_range[0, month_column].value
            pnl_tax_cell = pnl_values_range[48, month_column]
            if bin_balance > 0:
                def_tax_asset_cell = def_tax_asset_range[0, month_column]
                dtlc_cell = dtlc_range[0, month_column]
                
                if accrued_result > 0:
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
                    bin_balance += accrued_result
            else:
                pnl_tax_cell.value = ((accrued_result) * income_tax_rate) * (-1)
        else:
            pass
    
    # ---- Actualización reserva legal ---- #
    app.status_bar = "Actualizando reserva legal..."
    legal_reserve_range = monthly_ws.range(
        ((226 + lines_by_capex), 5), 
        ((226 + lines_by_capex), 5 + last_period_column)
    )
    
    share_capital_range = monthly_ws.range(
        ((225 + lines_by_capex), 5), 
        ((225 + lines_by_capex), 5 + last_period_column)
    )

    for month in model_months:
        month_column = int(model_dates['columnNumber'][month])
        month_number = month.month
        if month_column == 0:
            pass
        else:
            if month_number == 1:
                legal_reserve = legal_reserve_range[0, month_column].offset(column_offset = -1)
                share_capital = \
                    share_capital_range[0, month_column].offset(column_offset = -1).value
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

    # ---- Pestaña anualizada ---- #
    app.status_bar = "Creando pestaña anual..."

    annual_template_ws = templates_wb.sheets['annualTemplate']

    annual_template_ws.copy(
        before = monthly_ws, 
        name = annual_ws_name
    )

    summary_ws = new_wb.sheets[annual_ws_name]
    summary_ws.range('B2').value = plant_name
    summary_ws.range('B3').value = plant_code
    summary_ws.range('B4').value = 'End of regulatory life: ' + plant_eol_year

    model_start_year = model_start_date.year
    model_end_year = (plant_information.loc[0, 'reg_per_end'].year) + 1
    model_dates_monthly = pd.DataFrame(data=date_column_map).set_index('date')
    model_dates_annual = pd.DataFrame()
    current_ws_column = 0

    for year in range(model_start_year, model_end_year):
        data_list = []
        annual_df = model_dates_monthly[str(year)]
        year_first_column = int(annual_df.head(1).iloc[0])
        year_last_column = int(annual_df.tail(1).iloc[0])
        
        data_list.append([
            year, 
            year_first_column, 
            year_last_column, 
            current_ws_column
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
        current_ws_column += 1

    model_dates_annual = model_dates_annual.set_index('year')
    model_years = model_dates_annual.index

    app.status_bar = "Creando columnas anuales..."

    current_column = 6
    column_number = 0

    for year in model_years:
        
        # Date headers creation
        date_cell = summary_ws.range((5, current_column))
        date_cell.value = str(year)
        date_cell.font.bold = True
        date_cell.api.Borders.LineStyle = 1
        date_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        date_cell.color = (221, 235, 247)

        # Year status header creation
        month_status_cell = summary_ws.range((4, current_column))
        month_status_cell.font.bold = True
        month_status_cell.api.Borders.LineStyle = 1
        month_status_cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        month_status_cell.value = "Forecast"
        month_status_cell.color = (221, 235, 247)
        
        summary_ws.range('D6:' + 'D' + template_last_row).copy(summary_ws.range(6, current_column))
        date_cell.column_width = 11
        
        current_column += 1
        column_number += 1

    summary_ws.range('D:D').delete()

    model_last_year = model_dates_annual.tail(1).index.item()
    last_year_column = int(model_dates_annual.tail(1)['current_sheet_column'])
    
    # ---- Data anualizada ---- #

    model_inputs_annual = model_inputs_df.groupby(by=lambda date: date.year).sum()
    pool_incomes_annual = pool_incomes.groupby(by=lambda date: date.year).sum()
    heat_incomes_annual = heat_incomes.groupby(by=lambda date: date.year).sum()
    ro_incomes_annual = ro_incomes.groupby(by=lambda date: date.year).sum()
    gas_cost_annual = gas_cost.groupby(by=lambda date: date.year).sum()
    
    # ---- P&L ---- #
    app.status_bar = "Cargando P&L..."
    pnl_values_range = summary_ws.range(
        (53, 5), 
        (104, 5 + last_year_column)
    )

    # Cifras del año anterior #

    for row in pnl_rows:
        try:
            row_label = pnl_labels_range[row].value
            row_number = pnl_labels_range[row].row
            amount = prev_year_closing['amount'][row_label]
            summary_ws.range((row_number, 4)).value = amount
        
        except KeyError:
            pass

    # Forecast
    for row in pnl_rows:
        row_label = pnl_labels_range[row, 0].value
        input_data_column = label_mapping.get(row_label)
        
        if input_data_column == None:
            pass
        
        else:
            for year in model_years:
                year_column = int(model_dates_annual['current_sheet_column'][year])
                amount = model_inputs_annual[input_data_column][year]
                pnl_values_range[row, year_column].value = amount
    
    # ---- Cashflow ---- #
    app.status_bar = "Cargando Cashflow..."
    cf_values_range = summary_ws.range(
        (108, 5),
        (138, 5 + last_year_column)
    )

    for row in cashflow_rows:
        row_label = cf_labels_range[row, 0].value
        
        for year in model_years:
            try:
                year_column = int(model_dates_annual['current_sheet_column'][year])
                amount = int(oth_non_curr_items_a.loc[year, row_label])
                cf_values_range[row, year_column].value = amount
            except KeyError:
                pass

    # ---- Carga de variacion de capital de trabajo ---- #
    app.status_bar = "Cargando variaciones de WC..."
    wc_values_range = summary_ws.range(
        (144, 5), 
        (177, 5 + last_year_column)
    )

    for row in working_capital_rows:
        row_label = wc_labels_range[row, 0].value
        
        if row_label == 'Pool':
            for year in model_years:
                year_column = int(model_dates_annual['current_sheet_column'][year])
                amount = pool_incomes_annual['poolVariation'][year]
                wc_values_range[row, year_column].value = amount
        
        elif row_label == 'Heat':
            for year in model_years:
                year_column = int(model_dates_annual['current_sheet_column'][year])
                amount =  heat_incomes_annual['heatVariation'][year]
                wc_values_range[row, year_column].value = amount
        
        elif row_label == 'RO':
            for year in model_years:
                year_column = int(model_dates_annual['current_sheet_column'][year])
                amount =  ro_incomes_annual['roVariation'][year]
                wc_values_range[row, year_column].value = amount
        
        elif row_label == 'Gas':
            for year in model_years:
                year_column = int(model_dates_annual['current_sheet_column'][year])
                amount =  gas_cost_annual['variation'][year]
                wc_values_range[row, year_column].value = amount
        else:
            for year in model_years:
                try:
                    year_column = int(model_dates_annual['current_sheet_column'][year])
                    amount = int(oth_curr_items_a.loc[year, row_label])
                    wc_values_range[row, year_column].value = amount
                
                except KeyError:
                    pass
    
    # ---- Activos fijos ---- #

    # Carga de CAPEX
    app.status_bar = "Cargando CAPEX..."

    capex_range = summary_ws.range('' 
        + str(capex_range_start) 
        + ':' 
        + str(capex_range_end) 
        + ''
    )
    
    capex_range.insert()

    capex_name_range = summary_ws.range(
        (capex_range_start, 2), 
        (capex_range_end, 3)
    ) 
    
    capex_values_range = summary_ws.range(
        (capex_range_start, 5), 
        (capex_range_end, 5 + last_year_column)
    )

    row = 0 # Para el control del loop que sigue a continuación

    all_capex = capex.index

    for asset in all_capex:
        capex_name_range[row, 0].value = capex['name'][asset]
        capex_name_range[row, 1].value = "€"
        purchase_date = capex['start_date'][asset]
        purchase_year = (change_to_last_day(purchase_date)).year
        year_column = int(model_dates_annual['current_sheet_column'][purchase_year])
        amount = capex['historic_cost'][asset]
        capex_values_range[row, year_column].value = amount
        
        row += 1   

    # Las siguientes dos declaraciones son para ajustar visualmente la tabla de 
    # CAPEX al nuevo contenido recien agregado.

    top_blank_row = summary_ws.range('' 
        + str(capex_range_start) 
        + ':' 
        + str(capex_range_start) 
        + ''
    ).offset(row_offset = -1).delete()
    
    bottom_blank_row = summary_ws.range('' 
        + str(capex_range_end) 
        + ':' 
        + str(capex_range_end) 
        + ''
    ).delete()
    
    # Carga de depreciacion de activos existentes
    app.status_bar = "Cargando depreciación/amortizacion..."
    
    curr_assets_amort = curr_assets_amort.groupby(lambda x: x.year).sum()

    curr_asset_amort_range = summary_ws.range(
        (85, 5), 
        (85, 5 + last_year_column)
    )

    # Cortar el dataframe hasta la fecha final del modelo

    curr_asset_last_amort_period = curr_assets_amort.tail(1).index.item()
    
    if curr_asset_last_amort_period > model_last_year:
        curr_assets_amort = curr_assets_amort[:model_last_year]

    else:
        pass

    amort_periods = curr_assets_amort.index

    for year in amort_periods:
        amount = curr_assets_amort[year]
        year_column = int(model_dates_annual['current_sheet_column'][year])
        curr_asset_amort_range[0, year_column].value = amount * (-1)

    # Carga de depreciacion del CAPEX

    capex_amortization = capex_amortization.groupby(lambda x: x.year).sum()
    
    capex_amort_range = summary_ws.range(
        (91, 5), 
        (91, 5 + last_year_column)
    )

    # Cortar el dataframe hasta la fecha final del modelo

    capex_amort_last_period = capex_amortization.tail(1).index.item()

    if capex_amort_last_period > model_last_year:
        capex_amortization = capex_amortization[:model_last_year]

    else:
        pass

    amort_periods = capex_amortization.index

    for year in amort_periods:
        amount = capex_amortization[year]
        year_column = int(model_dates_annual['current_sheet_column'][year])
        capex_amort_range[0, year_column].value = amount * (-1)

    # ---- Balance ---- #
    app.status_bar = "Cargando balance..."
    bs_labels_range = summary_ws.range(
        (192 + lines_by_capex, 2), 
        (259 + lines_by_capex, 2)
    )

    # Cifras del año anterior
    for row in balance_sheet_rows:
        try:
            row_label = bs_labels_range[row].value
            row_number = bs_labels_range[row].row
            amount = prev_year_closing['amount'][row_label]
            summary_ws.range((row_number, 4)).value = amount
        
        except KeyError:
            pass

    # ---- Actualización resultados al cierre, Income TAX, BINs y 
    # reserva legal ---- #

    # Las columnas de los meses 12 de la pestaña mensual ya poseen los montos
    # de balance. Solo se traen a la pestaña anual.
    app.status_bar = "Actualizando resultados..."

    ret_earnings_range_a = summary_ws.range(
        (227 + lines_by_capex, 5), 
        (227 + lines_by_capex, 5 + last_year_column)
    )

    cy_result_range_a = summary_ws.range(
        (228 + lines_by_capex, 5),
        (228 + lines_by_capex, 5 + last_period_column)
    )

    def_tax_asset_range_a = summary_ws.range(
        (199 + lines_by_capex, 5), 
        (199 + lines_by_capex, 5 + last_year_column)
    )

    dtlc_range_a = summary_ws.range(
        (230 + lines_by_capex, 5), 
        (230 + lines_by_capex, 5 + last_year_column))

    legal_reserve_range_a = summary_ws.range(
        (226 + lines_by_capex, 5), 
        (226 + lines_by_capex, 5 + last_year_column)
    )

    pnl_tax_range_m = monthly_ws.range(
        (103, 5), 
        (103, 5 + last_period_column)
    )
    
    pnl_tax_range_a = summary_ws.range(
        (103, 5), 
        (103, 5 + last_year_column)
    )

    for year in model_years:
        year_column = int(model_dates_annual['current_sheet_column'][year])
        monthly_column = int(model_dates_annual['year_last_column'][year])
        
        ret_earnings_cell = ret_earnings_range_a[0, year_column]
        cy_result_cell = cy_result_range_a[0, year_column]
        def_tax_asset_cell = def_tax_asset_range_a[0, year_column]
        dtlc_cell = dtlc_range_a[0, year_column]
        pnl_tax_cell = pnl_tax_range_a[0, year_column]
        legal_reserve_cell = legal_reserve_range_a[0, year_column]

        ret_earnings_amount = ret_earnings_range[0, monthly_column].value
        cy_result_amount = cy_result_range[0, monthly_column].value
        def_tax_asset_amount = def_tax_asset_range[0, monthly_column].value
        dtlc_amount = dtlc_range[0, monthly_column].value
        pnl_tax_amount = pnl_tax_range_m[0, monthly_column].value
        legal_reserve_amount = legal_reserve_range[0, monthly_column].value
        
        ret_earnings_cell.value = ret_earnings_amount
        cy_result_cell.value = cy_result_amount
        def_tax_asset_cell.value = def_tax_asset_amount
        dtlc_cell.value = dtlc_amount
        pnl_tax_cell.value = pnl_tax_amount
        legal_reserve_cell.value = legal_reserve_amount

    # ---- Guardar y cerrar el archivo ---- #
    app.status_bar = "Finalizando..."
    new_wb.sheets[0].delete()
    time_string = datetime.today().strftime('%d-%m-%Y-%H%M')
    wb_name = 'Projects\FinancialModel\\' + 'output\\' + plant_name + '_' + time_string + '.xlsx'
    app.status_bar = "Completo."
    complete_alert_macro()
    templates_wb.close()
    new_wb.save(wb_name)
    # app.screen_updating = True
