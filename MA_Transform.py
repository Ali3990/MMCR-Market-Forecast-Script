# This process will use the Moody's Analytics API to retrieve data from baskets under the MA account. Refer to Mnemonics Compiler for list of compiled data baskets.
import os
import openpyxl
import pandas as pd
import datetime
from pathlib import Path
from Moodys_API_script.Moodys_API import download_basket
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from dotenv import load_dotenv

load_dotenv()
acckey=str(os.getenv("acc_key"))
enckey=str(os.getenv("enc_key"))

# create folder director. Refer to README.md for more details on structure.
root = Path(__file__).parent.parent
folder_paths = [
    root/"Data"/"MA Forecast Data",
    root/"Data"/"REIS Data",
    root/"Forecast Workbooks"
]

for folder in folder_paths:
    folder.mkdir(parents=True, exist_ok=True)

databuffet_dir = root / "Data" / "MA Forecast Data"

# See api script for instructions on setting variables
BASKET_NAME = "TM Forecast - Data Buffet"
filename = BASKET_NAME + ".xlsx"

# runs the api call and stores dataframe in memory
df = download_basket(BASKET_NAME, databuffet_dir, filename, acckey, enckey, engine="openpyxl")

# See 'Mnemonic_compiler.xlsx' file >> 'Dict' tab for dictionary compiler. Add more markets as necessary. Paste the entire dictionary here.
# e.g. 'market': ['MMC abbreviation', 'REIS abbreviation', 'geocode']
tm_dict = {
    'US Metro Total': ['US', 'US', 'IUSA'],
    'Denver':  ['DEN',  'DE',  'IUSA_MDEN'],
    'Los Angeles':  ['LA',  'LA',  'IUSA_DMLOS'],
    'Oakland-East Bay':  ['OAK',  'OA',  'IUSA_DMOAK'],
    'Orange County':  ['OC',  'OC',  'IUSA_DMANA'],
    'Portland':  ['POR',  'PO',  'IUSA_MPOT'],
    'San Jose':  ['SJ',  'SJ',  'IUSA_MSAJ'],
    'San Diego':  ['SD',  'SD',  'IUSA_MSAN'],
    'Ventura County':  ['VC',  'VN',  'IUSA_MOXN'],
    'Fairfield County':  ['FC',  'FC',  'IUSA_MBSD'],
    'Boston':  ['BOS',  'BO',  'IUSA_MBOS'],
    'New York Metro':  ['NYM',  'NY',  'IUSA_DMNWY'],
    'Northern New Jersey':  ['NNJ',  'NJ',  'IUSA_DMNEK'],
    'Long Island':  ['LI',  'LI',  'IUSA_DMNAS'],
    'District of Columbia':  ['DC',  'DC',  'IUSA_MWAA'],
    'Atlanta':  ['ATL',  'AT',  'IUSA_MATS'],
    'Austin':  ['AUS',  'AU',  'IUSA_MAUS'],
    'Central New Jersey':  ['CNJ',  'CJ',  'IUSA_DMLNB'],
    'Charleston':  ['CHS',  'CN',  'IUSA_MCHS'],
    'Charlotte':  ['CHR',  'CR',  'IUSA_MCLT'],
    'Dallas':  ['DAL',  'DA',  'IUSA_DMDAL'],
    'Fort Lauderdale':  ['FL',  'FL',  'IUSA_DMFOT'],
    'Fort Worth':  ['FW',  'FW',  'IUSA_DMDLL'],
    'Miami':  ['MIA',  'MI',  'IUSA_DMMIA'],
    'Orlando':  ['ORL',  'OR',  'IUSA_MORL'],
    'San Bernardino-Riverside':  ['SB',  'SB',  'IUSA_MRIV'],
    'Tampa-St. Petersburg':  ['TAM',  'TA',  'IUSA_MTAM'],
    'US Metro Total':  ['US',  'US',  'IUSA'],
    'Seattle 1':  ['SEA',  'SE',  'IUSA_DMEVE'],
    'Seattle 2':  ['EVE',  'SE',  'IUSA_DMSEB'],
    'San Francisco 1':  ['SF',  'SF',  'IUSA_DMSAF'],
    'San Francisco 2':  ['SR',  'SF',  'IUSA_DMSRF'],
    'Raleigh-Durham 1':  ['RAL',  'RD',  'IUSA_MRAL'],
    'Raleigh-Durham 2':  ['DUR',  'RD',  'IUSA_MDUR']
}

def moody_data_transform(tm_dict, df, save_dir):
    output_path = os.path.join(save_dir, "Mnemonic_transformed_data.xlsx")
    
    wb = Workbook()
    ws = wb.active
    first_block = True  # used to track first iteration for row trimming

    for market, values in tm_dict.items():
        geocode = values[2]

        # find columns matching the geocode
        matching_cols = [col for col in df.columns if "." in col and col.split(".")[-1] == geocode]
        if not matching_cols:
            continue

        # ensure first column is always included and no duplicates
        sub_df = df[matching_cols].copy()

        # trim first 5 rows
        if not first_block:
            sub_df = sub_df.iloc[5:]
        else:
            first_block = False

        # add Market column as the second column
        sub_df.insert(0, "Market", market)

        # write to Excel
        for r in dataframe_to_rows(sub_df, index=True, header=True):
            ws.append(r)

        # add a blank row separator between blocks
        ws.append([])

    wb.save(output_path)
    print(f"✅ Saved vertically stacked Excel to {output_path}")
    
moody_data_transform(tm_dict, df, databuffet_dir)






