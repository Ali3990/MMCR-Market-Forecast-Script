# This process will use the Moody's Analytics API to retrieve data from baskets under the MA account. Refer to Mnemonics Compiler for list of compiled data baskets.
import os
import openpyxl
import pandas as pd
import datetime
import sys

from Moodys_API_script.Moodys_API import download_basket
from dotenv import load_dotenv

load_dotenv()
acckey=str(os.getenv("acc_key"))
enckey=str(os.getenv("enc_key"))

# See api script for instructions on setting variables
BASKET_NAME = "TM Forecast - Data Buffet"
target_dir = r'C:\Users\ALi\OneDrive - MMC\Desktop\MMCR\Apt Forecasts\Forecast process 2025\Data\MA Forecast Data'
filename = BASKET_NAME + ".xlsx"

# runs the api call and stores dataframe in memory
df = download_basket(BASKET_NAME, target_dir, filename, acckey, enckey)

# See Mnemonic_compiler.xlsx file >> 'Dict' for dictionary compiler. Add more markets as necessary. Paste the entire dictionary here.
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

def moody_data_splitter(dict, save_dir):
    for market, values in tm_dict.items():
        #grabs 3rd element in tm_dict
        geocode = values[2] 

        # create list of exact geocode matches in column headers
        matching_cols = [col for col in df.columns 
                        if "." in col and col.split(".")[-1] == geocode]

        if not matching_cols:
            continue   # handle the columns by skipping over ones that do not have matching geocode.
        
        # matching columns with the correct geocode in header
        # Column A from the dataframe is auto-included when saving to file when argument index=True
        sub_df = df.loc[:, matching_cols]

        file_path = os.path.join(save_dir, f"{market}.xlsx")
        sub_df.to_excel(file_path, index=True)

        print(f"Saved {market} to output directory.")


moody_data_splitter(tm_dict, target_dir)
