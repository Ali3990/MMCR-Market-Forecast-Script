# MMCR-Market-Forecast-Script
Forecasting process for the first half of the year.

Instructions:
Note: Copy 'MMCR-Market-Forecast-Script' to a dedicated folder. The script will create multiple directories to store your files.

1. Use the Mnemonic_compiler to add any new markets or mnemonics.
2. For markets that have multiple geocodes, refer to instructions within MA_Transform.py to account for aggregation/weighted averaging functions.
3. Create a .env file with your Moody's API credentials (no quotations on the access or encryption key), save within the Moodys_API_script folder. If you do not have one, visit and generate keys at: https://www.economy.com/myeconomy/api-key-info 

    Example:
    acc_key=dsfnl3950-290******
    enc_key=sf097fdss-23jh*****

4. MA_Transform.py will call the save basket on data buffet, transform the data in a vertical stack and apply an identifier for later use. Folder directories will also be automatically created at the parent folder level.

*Created directories:

(Saved-folder-name)
    |_MMCR-Market-Forecast-Script
    |_*Forecast Workbooks
    |_*Data
        |_*REIS Data
        |_*MA Forecast Data
