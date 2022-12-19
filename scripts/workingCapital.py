from datetime import datetime
from datetime import date
import calendar
from copy import copy
from time import strftime

import pandas as pd
from dateutil.relativedelta import relativedelta

def group_by_month(week):
    # Esta funcion se definió para poder agrupar las fechas semanales en meses (ampliar)
    date = datetime(
        week.year,
        week.month,
        1
    )
    
    return date

def calculate_pool_income(pool_weekly_sales, model_periods):
    
    pool_sales_df = pd.read_excel(pool_weekly_sales, index_col = 'DiaLocal')
    num_of_weeks_per_month = pool_sales_df.groupby(group_by_month).count().rename(columns={'Income_MD_€':'numOfWeeks'})
    pool_income_df = pd.DataFrame(columns=['date', 'poolVariation', 'cashAmount'])
    
    for month in model_periods:
    
        # ---- Creacion de los rangos de fechas ---- #
         
        previous_month = month + relativedelta(months = -1)
        previous_month_last_day = calendar.monthrange(
            previous_month.year, 
            previous_month.month
        )[1]
        previous_month = datetime(
            previous_month.year,
            previous_month.month,
            previous_month_last_day
        )

        current_month = datetime(
            month.year,
            month.month,
            1
        )
        
        last_two_weeks_prev_month = pd.date_range(
            end=previous_month, 
            periods=2, 
            freq = 'W'
        ).to_series()
        
        num_of_weeks = int(num_of_weeks_per_month.loc[current_month])
        
        if num_of_weeks == 4:
            first_two_weeks_curr_month = pd.date_range(
                start=current_month, 
                periods=2, 
                freq='W'
            ).to_series()
        
        else:
            first_two_weeks_curr_month = pd.date_range(
                start=current_month, 
                periods=3, 
                freq='W'
            ).to_series()
        
        last_two_weeks_curr_month = pd.date_range(
            end=month, 
            periods=2, 
            freq='W'
        ).to_series()
        
        # ---- Calculo del saldo de la CxC al inicio del periodo ----
        
        receivables_opening_balance = 0
        
        for week in last_two_weeks_prev_month:
            receivables_opening_balance = receivables_opening_balance + pool_sales_df['Income_MD_€'][last_two_weeks_prev_month[week]]

        # ---- Calculo del ingreso a caja del mes ----
        
        # El monto pagado durante el mes corresponde a las dos semanas del mes anterior y las dos primeras del mes actual
        
        cash_periods = pd.concat([last_two_weeks_prev_month, first_two_weeks_curr_month]) 
        cash_amount = 0 
        
        for week in cash_periods:
            cash_amount = cash_amount + pool_sales_df['Income_MD_€'][cash_periods[week]]

        # ---- Calculo del saldo de la CxC al cierre del periodo

        receivables_closing_balance = 0
        
        for week in last_two_weeks_curr_month:
            receivables_closing_balance = receivables_closing_balance + pool_sales_df['Income_MD_€'][last_two_weeks_curr_month[week]]    

        # ---- Calculo de la variación ----
        
        receivables_variation = receivables_closing_balance - receivables_opening_balance

        values = []
        values.append(month)
        values.append(int(receivables_variation))
        values.append(int(cash_amount))

        df_last_item = len(pool_income_df)
        pool_income_df.loc[df_last_item] = values
    
    pool_income_df = pool_income_df.set_index('date')
    
    return pool_income_df

def calculate_heat_income(model_inputs, model_periods):

    heat_sales = model_inputs['Heat_CHP_MWh']
    heat_income_df = pd.DataFrame(columns=['date', 'heatVariation', 'cashAmount'])
    
    for month in model_periods:
    
        values = []

        # ---- Calculo del saldo de la CxC al inicio del período

        month_start = datetime(
            month.year,
            month.month,
            1
        )
        
        opening_periods = pd.date_range(
            end=month_start, 
            periods=2, 
            freq='M'
        ).to_series()

        receivables_opening_balance = 0
        
        for period in opening_periods:
            receivables_opening_balance = receivables_opening_balance + heat_sales[opening_periods[period]]

        # ---- Calculo del ingreso a caja del mes ----

        # El monto que ingresa a caja es la venta de dos meses atrás:
       
        sales_month = opening_periods[0]

        cash_amount = 0
        cash_amount = heat_sales[sales_month]

        # ---- Calculo del saldo de la CxC al cierre del período

        receivables_closing_balance =  heat_sales[opening_periods[1]] + heat_sales[month]
        
        receivables_variation = receivables_closing_balance - receivables_opening_balance

        values.append(month)
        values.append(receivables_variation)
        values.append(cash_amount)

        df_last_item = len(heat_income_df)
        heat_income_df.loc[df_last_item] = values

    heat_income_df = heat_income_df.set_index('date')
    
    return heat_income_df

def calculate_ro_income(model_inputs, model_periods):

    ro_income = model_inputs['Income_RO_€']
    ro_incomes_df = pd.DataFrame(columns=['date','roVariation','cashCurrMonth'])
    
    bop_receivables = 0
    eop_receivables = 0
    curr_year_paid_amount = 0
    curr_year_pay_period = 0
    cash_curr_month = 0
    past_year_pay_period = 0
    past_year_paid_amount = 0
    dates_prev_year = pd.Series()
    dates_curr_year = pd.Series()
    past_year_next_ro_month = 0
    past_year_ro_accrued = 0
    cash_from_past_year = 0
    period_ro_amount = 0

    payment_calendar = {
        1: 0.29,
        2: 0.63,
        3: 0.73,
        4: 0.78,
        5: 0.79,
        6: 0.80,
        7: 0.82,
        8: 0.83,
        9: 0.84,
        10: 0.88,
        11: 0.94,
        12: 0.95,
        13: 0.98,
        14: 1
    }
    
    for month in model_periods:
    
        values = []
        month_number = month.month
        
        # ---- Calculo de caja por RO del año anterior ----

        if past_year_pay_period == 11:

            period_ro_amount = dates_prev_year[past_year_next_ro_month]
            past_year_ro_accrued = (past_year_ro_accrued + period_ro_amount)
            cash_from_past_year = (past_year_ro_accrued * payment_calendar.get(past_year_pay_period)) - past_year_paid_amount
            past_year_paid_amount = past_year_paid_amount + cash_from_past_year
            past_year_pay_period += 1

        
        elif past_year_pay_period == 12:
            
            index_date = datetime(
                (month.year - 1),
                12,
                calendar.monthrange((month.year) - 1, 12)[1]
            )
            
            past_year_ro_accrued = (past_year_ro_accrued + dates_prev_year[index_date])
            cash_from_past_year = (past_year_ro_accrued * payment_calendar.get(past_year_pay_period)) - past_year_paid_amount
            past_year_paid_amount = past_year_paid_amount + cash_from_past_year
            past_year_pay_period += 1

        elif past_year_pay_period >= 13:
            
            cash_from_past_year = (past_year_ro_accrued * payment_calendar.get(past_year_pay_period)) - past_year_paid_amount
            past_year_paid_amount = past_year_paid_amount + cash_from_past_year
            
            if past_year_pay_period == 13:
                past_year_pay_period += 1
            
            else:
                dates_prev_year = pd.Series(dtype='int64')
                past_year_next_ro_month = 0
                past_year_paid_amount = 0
                past_year_ro_accrued = 0
                past_year_pay_period = 0

        else:
            pass

        # ---- Calculo de caja y CxC por RO del año actual ----
        
        ro_amount = 0

        if month_number  == 1:
            cash_curr_year = 0
            curr_year_pay_period = 0
            curr_year_paid_amount = 0
            dates_curr_year = ro_income[str(month.year)]
        
        elif month_number == 2:
            pass
        
        else:
            curr_year_pay_period += 1
            
            for period in range(1, (month_number - 1)):
                index_date = datetime(
                    month.year,
                    period,
                    calendar.monthrange(month.year, period)[1]
                )
                
                ro_amount = ro_amount + dates_curr_year[index_date]
            
            ro_accrued_for_pay = ro_amount * (payment_calendar.get(curr_year_pay_period))
            cash_curr_year = ro_accrued_for_pay - curr_year_paid_amount
            curr_year_paid_amount = curr_year_paid_amount + cash_curr_year

            if month_number == 5:
                cash_from_past_year = 0
            
            elif month_number == 12:
                past_year_pay_period = 11
                past_year_next_ro_month = datetime(
                    month.year, 
                    (month_number -1),
                    calendar.monthrange(month.year, (month_number - 1))[1]
                )
                past_year_paid_amount = copy(curr_year_paid_amount)
                past_year_ro_accrued = copy(ro_amount)
                dates_prev_year = dates_curr_year[10:12]
                
            else:
                pass
        
        cash_curr_month = cash_curr_year + cash_from_past_year
        eop_receivables += dates_curr_year[month] - cash_curr_month
        receivables_variation = eop_receivables - bop_receivables

        values.append(month)
        values.append(receivables_variation)
        values.append(cash_curr_month)

        df_lastItem = len(ro_incomes_df)
        ro_incomes_df.loc[df_lastItem] = values

        bop_receivables = eop_receivables

    ro_incomes_df = ro_incomes_df.set_index('date')
    
    return ro_incomes_df

def calculate_30_day_items(model_inputs, label, model_periods):
    
    label_item_mapping = {
        'Gas': 'Variable_gas_CHP_cost_€'
    }
    
    selected_item = label_item_mapping.get(label)
    item_values = model_inputs[selected_item]

    results_df = pd.DataFrame(columns=['date','variation','cashMov'])

    for month in model_periods:

        values = []

        past_month_date = pd.date_range(end=month, periods=2, freq='M').to_series(index=[0,1])
        
        cash_movement = item_values[past_month_date[0]]
        receivables_variation = item_values[past_month_date[1]] - item_values[past_month_date[0]]
        
        values.append(month)
        values.append(receivables_variation)
        values.append(cash_movement)

        df_last_item = len(results_df)
        results_df.loc[df_last_item] = values

    results_df = results_df.set_index('date')
    
    return results_df

def calculate_other_items(other_curr_transactions):

    monthly_df = other_curr_transactions.groupby(by=['date','label']).sum()
    annual_df = other_curr_transactions.groupby(by=[lambda date: date.year, 'label']).sum()

    return monthly_df, annual_df


## PARA EKOREC ##

model_start_date = date.fromisoformat('2023-01-01')

plant_information = pd.read_excel(
        r'Projects\FinancialModel\input\plants.xlsx'
    )

plant_name = plant_information.loc[0, 'name']
plant_code = plant_information.loc[0, 'it_code']
plant_eol_year = str((plant_information.loc[0, 'reg_per_end']).year)
reg_period_end = plant_information.loc[0, 'reg_per_end']

model_timeframe = pd.date_range(
        start=model_start_date, 
        end=reg_period_end, 
        freq='M'
    ).to_series().index

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

model_inputs_df = pd.read_excel(
        'Projects/FinancialModel/input/monthlyInputs.xlsx', 
        index_col='DiaLocal', 
        parse_dates=True
    )

pool_weekly_sales = 'Projects/FinancialModel/input/pool_weekly_sales.xlsx'

pool_incomes = calculate_pool_income(pool_weekly_sales, model_timeframe)
heat_incomes = calculate_heat_income(model_inputs_df, model_timeframe)
ro_incomes = calculate_ro_income(model_inputs_df, model_timeframe)
gas_cost = calculate_30_day_items(model_inputs_df, 'Gas', model_timeframe)

# to_excel

pool_incomes.to_excel('Projects/FinancialModel/output/pool_cash.xlsx')
heat_incomes.to_excel('Projects/FinancialModel/output/heat_cash.xlsx')
ro_incomes.to_excel('Projects/FinancialModel/output/ro_cash.xlsx')
gas_cost.to_excel('Projects/FinancialModel/output/gas_cash.xlsx')