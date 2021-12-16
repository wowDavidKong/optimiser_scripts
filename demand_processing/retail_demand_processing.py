"""
This script filters the demand on the species in question and writes the output to a new
Excel file

"""

import numpy as np
import pandas as pd
import logging

log = logging.getLogger()

# logging to print to console
console = logging.StreamHandler()
log.addHandler(console)
log.setLevel(logging.INFO)

retail_demand_loc = input('Please specify the file patch, including, file name: ')
write_loc = input('Please specify the location where the excel file is to be written out to: ')
species = input('Please specify the species: BEEF, PORK or LAMB: ')

if species not in (['BEEF', 'PORK', 'LAMB']):
    raise ValueError('Incorrect species selection. Please specify either BEEF, PORK or LAMB')


def read_retail_demand(sheet_name: str):
    """
    A function that reads in the individual Excel demand tabs and returns a DataFrame with just the primals associated
    with the species selected.

    Args:
        sheet_name (str): The name of the Excel sheet Can be either HW, TRUG or BUN

    Returns:
        pd.DataFrame: a DataFrame containing the formatted demand by species selected
    """

    retail_demand = pd.read_excel(f'{retail_demand_loc}.xlsx', sheet_name=sheet_name)
    retail_demand = retail_demand[retail_demand['Family'] == species]
    retail_demand['Date'] = pd.to_datetime(retail_demand['Date'], format='yyyy-mm-dd').dt.date


    return retail_demand


# We apply the function to each CRM site and write out DataFrames to a new Excel spreadsheet.
hw_retail_demand = read_retail_demand(sheet_name="HW")
log.info('Adding Heathwood Demand')
trug_retail_demand = read_retail_demand(sheet_name="TRUG")
log.info('Adding Truganina Demand')
bun_retail_demand = read_retail_demand(sheet_name="BUN")
log.info('Adding Bunbury Demand')

with pd.ExcelWriter(f'{write_loc}/retail_demand_vol.xlsx') as writer:
    hw_retail_demand.to_excel(writer, sheet_name='HW', index=False)
    trug_retail_demand.to_excel(writer, sheet_name='TRUG', index=False)
    bun_retail_demand.to_excel(writer, sheet_name='BUN', index=False)


