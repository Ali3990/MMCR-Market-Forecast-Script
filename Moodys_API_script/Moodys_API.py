# Refer to documentation

import requests
import pandas as pd
import datetime
import json
from time import sleep
import hashlib
import hmac
import io
from dotenv import load_dotenv
import os
import openpyxl


#####
# Setup:
# 1. Store your access key, encryption key, and basket name.
# Get your keys at:
# https://www.economy.com/myeconomy/api-key-info
load_dotenv()
acckey=str(os.getenv("acc_key"))
enckey=str(os.getenv("enc_key"))

# The id can be found in the URL of the page of the basket of mnemonics, but basket name will work by itself:
# BASKET_NAME = "TM Forecast - Double geocodes - Baseline"
# basket_id = "E3BCC362-65B1-4970-9E5D-4407D4E28831"
# !! Remember to set the start date of the basket to '1985M01' ("Jan 1985")


# # provide directory to save the output file:
# target_dir = r'C:\Users\Alan Li\Desktop\MMCR\Apt Forecasts\Forecast Process 2025\Data\MA Forecast Data'
# filename = BASKET_NAME + ".xlsx"

#####
# Function: Make API request, including a freshly generated signature.
#
# Arguments:
# 1. Part of the endpoint, i.e., the URL after "https://api.economy.com/data/v1/"
# 2. Your access key.
# 3. Your personal encryption key.
# 4. Optional: default GET, but specify POST when requesting action from the API.
#
# Returns:
# HTTP response object.

ENC_KEY = enckey
ACC_KEY = acckey

def api_call(apiCommand, accKey, encKey, call_type="GET"):
    url = "https://api.economy.com/data/v1/" + apiCommand
    timeStamp = datetime.datetime.strftime(datetime.datetime.utcnow(), "%Y-%m-%dT%H:%M:%SZ")
    payload = bytes(accKey + timeStamp, "utf-8")
    signature = hmac.new(bytes(encKey, "utf-8"), payload, digestmod=hashlib.sha256)
    head = {"AccessKeyId":accKey,
            "Signature":signature.hexdigest(),
            "TimeStamp":timeStamp}
    sleep(1)
    if call_type == "POST":
        response = requests.post(url, headers=head)
    elif call_type =="DELETE":
        response = requests.delete(url, headers=head)
    else:
        response = requests.get(url, headers=head)
    return(response)

def download_basket(BASKET_NAME, target_dir, filename, ACC_KEY, ENC_KEY, engine):
    #####
    # 1. Set Basket name at the beginning.
    # 2. Get list of baskets.
    # 3. Extract the basket with a given name, and save its ID for later.
    baskets = pd.DataFrame(json.loads(api_call("baskets/", ACC_KEY, ENC_KEY).text))
    basketId = baskets.loc[baskets["name"]==BASKET_NAME, "basketId"].item()
    print("Basket ID: " + basketId)
    print("Basket Name: " + BASKET_NAME)

    # 4. Execute a particular basket using its ID.
    # This requires that the optional argument "type" be set to "POST".
    call = ("orders?type=baskets&action=run&id=" + basketId)
    order = api_call(call, ACC_KEY, ENC_KEY, call_type="POST")
    orderId = order.text[12:48]
    print("Order ID: " + orderId)

    #####
    # Download the output:
    # 5. Periodically check if the order has completed.
    if order.status_code != 200:
        sleep(3)
        print("Failed! Status Code: "+ str(order.status_code))
    else:
        sleep(3)
        print("Successful Order! Status Code: " + str(order.status_code))

    # 6. Download completed output.
    new_call = ("orders?type=baskets&id=" + basketId)
    get_basket = api_call(new_call, ACC_KEY, ENC_KEY)

    # 7. Load Excel file directly from memory (assuming API returns .xlsx)
    data_df = pd.read_excel(io.BytesIO(get_basket.content))

    # Optional: Set index, clean data if needed
    data_df = data_df.set_index(data_df.columns[0])
    data_df.dropna(how='all', axis=1, inplace=True)
    data_df = data_df.loc[:, (data_df != "").any(axis=0)]

    data_df.to_excel(os.path.join(target_dir,filename), engine=engine)

    return data_df

# #test
# download_basket(BASKET_NAME, target_dir, filename, acckey, enckey, engine="openpyxl")