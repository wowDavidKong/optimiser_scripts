"""

This script is used to convert the vendorline and butchershop crate volume demand into their kg equivalents

"""

import pandas as pd
import numpy as np

vendor_demand_loc = input('Please specify the file patch, including, file name: ')
write_loc = input('Please specify the location where the excel file is to be written out to: ')


def read_vendor_demand(sheet_name=None):
    """
    A function that reads in the the individual tabs in excel and returns a dataframe with crate-to-kg demand conversions
    and pairs down entries to those just used

    Args:
        sheet_name (str): The name of the excel sheet. Can be either HW, TRUG or BUN

    Returns:
        pd.DataFrame: a DataFrame containing the adjusted demand and only the primals of interest
    """

    # We read in the excel file sheet and clean up missing WoW code values
    raw_vendor_demand = pd.read_excel(f'{vendor_demand_loc}.xlsx', sheet_name=sheet_name)
    raw_vendor_demand['Date'] = pd.to_datetime(raw_vendor_demand['Date'], format='yyyy-mm-dd').dt.date
    missing_wow_codes = [99000, 99001, 99002]
    for value in missing_wow_codes:
        raw_vendor_demand.loc[raw_vendor_demand['PrimalID'] == value, 'WOW code'] = value

    # We create a dictionary of the WoW code and crate-to-kg multiplier, adjust the demand for the primals in
    # question and generate a DataFrame with only the primals required.
    crate_to_kg_val = dict([(99000, 4),
                            (99001, 4),
                            (99002, 4),
                            (82767, 13.92),
                            (82766, 15.065),
                            (82765, 13.48),
                            (82764, 14.62),
                            (81638, 13.8),
                            (81637, 13.6),
                            (81636, 13.6)])

    for k, v in crate_to_kg_val.items():
        raw_vendor_demand.loc[raw_vendor_demand['WOW code'] == k, 'Proposed Purchases'] = raw_vendor_demand[
                                                                                              'Proposed Purchases'] * v

    req_wow_codes = list(crate_to_kg_val.keys())
    updated_df = raw_vendor_demand[raw_vendor_demand['WOW code'].isin(req_wow_codes)]

    return updated_df

# We apply the function to each CRM site and write out DataFrames to a new Excel spreadsheet.
hw_vendor_demand = read_vendor_demand(sheet_name="HW")
trug_vendor_demand = read_vendor_demand(sheet_name="TRUG")
bun_vendor_demand = read_vendor_demand(sheet_name="BUN")

with pd.ExcelWriter(f'{write_loc}/vendorline_demand_vol.xlsx') as writer:
    hw_vendor_demand.to_excel(writer, sheet_name='HW', index=False)
    trug_vendor_demand.to_excel(writer, sheet_name='TRUG', index=False)
    bun_vendor_demand.to_excel(writer, sheet_name='BUN', index=False)





