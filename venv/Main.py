import os
import pandas as pd
import xlrd
import numpy as np

# Location of the New Add Paramter Tables
root = r'V:\APLA\jobs\AP\Petrochina\024076 - PetroChina 2020YE\10 Reports\zzz_Parameter Tables Work\1. NewAdd PUD Parameter Sheets'

'''
# Create two mother dataframes (oil and gas)
gas_mother_df = pd.DataFrame(columns=['District', 'Field', 'Block', 'WellBlock', 'Reservoir', 'DatumDepth', 'Porosity',
        'WaterSat', 'Permeabilty', 'TempC', 'InitPressure', 'SpGravity', 'PDP_EvalMethod', 'PDP_Area', 'PDP_Thickness',
        'PDP_NetRockVol', 'PDP_OGIP', 'PDP_OGIPSep','PDNP_Area', 'PDNP_Thickness', 'PDNP_NetRockVol', 'PDNP_OGIP',
        'PDNP_OGIPSep', 'PUD_Area', 'PUD_Thickness', 'PUD_NetRockVol', 'PUD_OGIP', 'PUD_OGIPSep', 'RecoveryMech'])

oil_mother_df = pd.DataFrame(columns=['District', 'Field', 'Block', 'WellBlock', 'Reservoir', 'DatumDepth', 'Porosity',
        'WaterSat', 'Permeabilty', 'TempC', 'InitPressure', 'OilDensity', 'OilViscosity', 'OilFVF', 'PDP_EvalMethod',
        'PDP_Area', 'PDP_Thickness', 'PDP_NetRockVol', 'PDP_OOIP', 'PDNP_Area', 'PDNP_Thickness',
        'PDNP_NetRockVol', 'PDNP_OOIP', 'PUD_Area', 'PUD_Thickness', 'PUD_NetRockVol', 'PUD_OOIP', 'RecoveryMech'])
'''

# Create empty dfs
mother_oil = pd.DataFrame()
mother_gas = pd.DataFrame()
district = ""
field = ""
hdrs = ""


# Make sure the excel files are of .xlsx extensions ****
for dir, subdir, files in os.walk(root):

    # If files is not empty
    if files:
        for f in files:
            loc = dir + '\\' + f

            # Open file
            wb = xlrd.open_workbook(loc)
            sheets = wb.sheet_names()
            for sheet in sheets:
                if 'Gas' in sheet:
                    print(sheet)
                    district = wb.sheet_by_name(sheet).cell_value(0, 2)
                    field = wb.sheet_by_name(sheet).cell_value(1, 2)

                    # Read sheet, delete 2nd and 3rd columns (chinese translation and total columns)
                    y = pd.read_excel(loc, sheet_name=sheet, header=None, skiprows=4, nrows=117)
                    y = y.drop([1, 2], axis=1)

                    # if header contains 'Unnamed,' then delete
                    y.columns = y.iloc[0]
                    y = y[[col for col in y.columns if col is not np.nan]]
                    y.columns = range(y.shape[1])

                    # Make first column of data become the headers
                    hdrs_gas = np.array(y.iloc[:, 0])
                    hdrs_gas = np.insert(hdrs_gas, 117, 'District')
                    hdrs_gas = np.insert(hdrs_gas, 118, 'Field')

                    # Transpose df, drop 1st row
                    y_T = y.T
                    y_T = y_T.drop(y_T.index[0])

                    # add district and field to end of df
                    y_T['District'] = district
                    y_T['Field'] = field

                    mother_gas = mother_gas.append(other=y_T, ignore_index=False)


                elif 'Oil' in sheet:
                    print(sheet)
                    district = wb.sheet_by_name(sheet).cell_value(0, 2)
                    field = wb.sheet_by_name(sheet).cell_value(1, 2)

                    # Read sheet, delete 2nd and 3rd columns (chinese translation and total columns)
                    x = pd.read_excel(loc, sheet_name=sheet, header = None, skiprows=4, nrows=103)
                    x = x.drop([1, 2], axis=1)

                    # if header contains 'Unnamed,' then delete
                    x.columns = x.iloc[0]
                    x = x[[col for col in x.columns if col is not np.nan]]
                    x.columns = range(x.shape[1])
                    
                    # Make first column of data become the headers
                    hdrs_oil = np.array(x.iloc[:,0])
                    hdrs_oil = np.insert(hdrs_oil, 103, 'District')
                    hdrs_oil = np.insert(hdrs_oil, 104, 'Field')

                    # Transpose df, drop 1st row
                    x_T = x.T
                    x_T = x_T.drop(x_T.index[0])

                    # add district and field to end of df
                    x_T['District'] = district
                    x_T['Field'] = field

                    mother_oil = mother_oil.append(other = x_T, ignore_index= False)

                else:
                    next


mother_oil.columns = hdrs_oil
mother_oil.to_csv(r'V:\APLA\jobs\AP\Petrochina\024076 - PetroChina 2020YE\10 Reports\zzz_Parameter Tables Work\NewAdds_Oil_output.csv')

mother_gas.columns = hdrs_gas
mother_gas.to_csv(r'V:\APLA\jobs\AP\Petrochina\024076 - PetroChina 2020YE\10 Reports\zzz_Parameter Tables Work\NewAdds_Gas_output.csv')