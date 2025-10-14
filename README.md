# MMCR-Market-Forecast-Script
Forecasting process for the first half of the year.

Instructions:
Note: Copy 'MMCR-Market-Forecast-Script' to a dedicated folder. The script will create multiple directories to store your files.

1. Use the Mnemonic_compiler.xlsx to add any new markets or mnemonics. This is a dynamic process, so users can add as many markets or mnemonics. The data will be formatted to fit alongside REIS data, however, it is advisable to keep to essential variables to reduce testing runtime.
2. For markets that have multiple geocodes, refer to notes in Mnemonic_compiler.xlsx. The MA data for divisional metros will be merged (either sum or weighted average) in the main forecast file.
3. Create a .env file with your Moody's API credentials (no quotations on the access or encryption key), save within the Moodys_API_script folder. If you do not have one, visit and generate keys at: https://www.economy.com/myeconomy/api-key-info 

    Example:
    acc_key=dsf50-290*****
    enc_key=sf09s-23j*****

4. MA_Transform.py will call the save basket on data buffet, transform the data in a vertical stack and apply an identifier for later use. Folder directories will also be automatically created at the parent folder level.

*Created directories:

(Saved-folder-name)
    |_MMCR-Market-Forecast-Script
    |_*Forecast Workbooks
    |_*Data
        |_*REIS Data
        |_*MA Forecast Data

5. Use the reg_test.py file to run the main forecast file under "./Forecast Workbooks" directory. Select the market you wish to test.