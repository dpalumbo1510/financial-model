from __future__ import division
from tkinter.filedialog import test
import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta
from exceptions import AssetStatusError

def calculate_fixed_assets(fixed_assets_source, model_generation_date):

    # ---- Preprocesamiento listado de activos fijos ---- #

    fixed_assets_df = pd.read_excel(fixed_assets_source)

    fixed_assets_df = fixed_assets_df.assign(
        end_date = pd.Series(),
        periods_remaining = pd.Series(),
        is_active = pd.Series(),
        is_capex = pd.Series()
    )
    
    for asset in fixed_assets_df.index:
        
        asset_start_date = fixed_assets_df['start_date'][asset]
        useful_life_in_months = (fixed_assets_df['useful_life'][asset]) * 12
        months_delta = relativedelta(months = useful_life_in_months)
        fixed_assets_df['end_date'][asset] = asset_start_date + months_delta
        
        asset_end_date = fixed_assets_df['end_date'][asset]

        if asset_end_date < model_generation_date:
            fixed_assets_df['is_active'][asset] = False
            fixed_assets_df['is_capex'][asset] = False
            fixed_assets_df['periods_remaining'][asset] = 0
        
        elif asset_start_date < model_generation_date:
            fixed_assets_df['is_active'][asset] = True
            fixed_assets_df['is_capex'][asset] = False
            
            remaining_life_period = relativedelta(
                asset_end_date, 
                model_generation_date
            )
            
            months_remaining = (remaining_life_period.years * 12) + remaining_life_period.months
            fixed_assets_df['periods_remaining'][asset] = months_remaining
        
        else:
            fixed_assets_df['is_active'][asset] = False
            fixed_assets_df['is_capex'][asset] = True
            fixed_assets_df['periods_remaining'][asset] = useful_life_in_months

        asset_is_active = fixed_assets_df['is_active'][asset]
        asset_is_capex = fixed_assets_df['is_capex'][asset]

        try:

            # True - True: error lógico. No puede estar activo y ser CAPEX a la vez.#
            if asset_is_active is True and asset_is_capex is True:
                raise AssetStatusError
            
            # El motivo por el cual no se evaluan las otras posiblidades del IF:

            # 1) False - False: el activo no estaría operativo y no sería CAPEX,
            #    es decir, el activo en la práctica no existe.
            # 2) False - True: es CAPEX y se trata en uno de los bloque IF 
            #    anteriores.
    
            else:
                pass

        except AssetStatusError:
            print("Existe un activo con valor 'True' para 'is_active' y 'is_capex'")
            break
    
    # Las series agregadas luego de crear el dataframe poseen dtype "Objeto",
    # por lo que se deben inferir los dtypes correctos en base a su contenido
    # para el procesamiento que se hará mas adelante.
    
    fixed_assets_df = fixed_assets_df.infer_objects()

    # ---- CAPEX ---- #
    
    capex_df = fixed_assets_df[fixed_assets_df["is_capex"] == True]

    # Creacion de tabla de amortización de los CAPEX.

    all_amort_periods_df = pd.DataFrame({
        'asset_id': pd.Series(),
        'amort_period': pd.Series(dtype='datetime64[ns]'),
        'amort_amount': pd.Series()
    })
    
    capex_amortization = pd.DataFrame() # Inicializacion de la variable.
    amortization_list = []
    
    for asset in capex_df.index:

        asset_id = capex_df['asset_id'][asset]
        historic_cost = capex_df['historic_cost'][asset]
        amortization_periods = int(capex_df['periods_remaining'][asset])
        amortization_amount = historic_cost / amortization_periods
        month_delta = relativedelta(months=0)

        for period in range(0, amortization_periods):
            amortization_date = capex_df['start_date'][asset] + month_delta
            amortization_list.append([
                asset_id, 
                amortization_date, 
                amortization_amount
                ]
            )
            month_delta = month_delta + relativedelta(months =+ 1)

    list_bridge_df = pd.DataFrame(
        amortization_list, 
        columns = ['asset_id', 'amort_period', 'amort_amount']
    )
        
    all_amort_periods_df = pd.concat([all_amort_periods_df, list_bridge_df], ignore_index=True)

    capex_amortization = all_amort_periods_df.groupby('amort_period')['amort_amount'].agg(np.sum)
        
    # ---- Activos existentes ---- #

    current_assets_df = fixed_assets_df[
        (fixed_assets_df['is_capex'] == False) & (fixed_assets_df['is_active'] == True)
    ]
    
    # Creacion tabla de amortización de activos en uso.
    
    curr_assets_amort_df = pd.DataFrame({
        'asset_id': pd.Series(),
        'amort_period': pd.Series(dtype='datetime64[ns]'),
        'amort_amount': pd.Series()
        }
    )

    curr_assets_amort = pd.DataFrame()
    amortization_list = []   

    for asset in current_assets_df.index:

        asset_id = current_assets_df['asset_id'][asset]
        historic_cost = current_assets_df['historic_cost'][asset]
        amortization_periods = int(current_assets_df['periods_remaining'][asset])
        useful_life_in_months = current_assets_df['useful_life'][asset] * 12
        amortization_amount = historic_cost / useful_life_in_months
        month_delta = relativedelta(months=0)

        for period in range(0, amortization_periods):
            amortization_date = model_generation_date + month_delta
            amortization_list.append([
                asset_id, 
                amortization_date, 
                amortization_amount
                ]
            )
            month_delta = month_delta + relativedelta(months =+ 1)

    list_bridge_df = pd.DataFrame(
        amortization_list, 
        columns = ['asset_id', 'amort_period', 'amort_amount']
    )
    
    curr_assets_amort_df = pd.concat([curr_assets_amort_df, list_bridge_df], ignore_index=True)
    
    curr_assets_amort = curr_assets_amort_df.groupby('amort_period')['amort_amount'].agg(np.sum)

    return capex_df, curr_assets_amort, capex_amortization