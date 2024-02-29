import pandas as pd
import json
import requests
from datetime import datetime, timedelta
import gspread
import time
import os
import webbrowser
import zipfile
import shutil
import math
import warnings
warnings.filterwarnings('ignore')
def get_wms_report(cookies, payloads, download_folder):

    url = "https://wms.ssc.shopee.vn/api/v2/apps/basic/reportcenter/create_export_task"

    payload = payloads
    headers = {
    'authority': 'wms.ssc.shopee.vn',
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'en-US,en;q=0.9',
    'content-type': 'application/json;charset=UTF-8',
    'cookie': cookies,
    'origin': 'https://wms.ssc.shopee.vn',
    'referer': 'https://wms.ssc.shopee.vn/v2/reportcenter/reportcenter',
    'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'x-csrftoken': 'dJgldqNiOjHyOPNTuqRIpZNOjcpeKW5F'
    }

    response = requests.request("POST", url, headers=headers, data=payload)

    print(response.text)

    time.sleep(5)

    url = "https://wms.ssc.shopee.vn/api/v2/apps/basic/reportcenter/search_export_task?is_myself=1&pageno=1&count=1"

    payload = {}
    headers = {
    'authority': 'wms.ssc.shopee.vn',
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'en-US,en;q=0.9',
    'cookie': cookies,
    'referer': 'https://wms.ssc.shopee.vn/v2/reportcenter/reportcenter',
    'sec-ch-ua': '"Google Chrome";v="113", "Chromium";v="113", "Not-A.Brand";v="24"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36',
    'x-csrftoken': 'F7DoD3f39shh5vtoyxZT5dAA7BHUjgN5'
    }
    print("Wating for WMS report to be done")
    checking = True
    x=0
    
    while checking == True:
        response = requests.request("GET", url, headers=headers, data=payload)
        status = json.loads(response.text)['data']['list'][0]['task_status']
        
        if status != 2:
            time.sleep(3)
            x += 3
            print(f'Waiting time: {x}s')
        else:
            print('Status Done. Begin download report')
            time.sleep(1)
            response = requests.request("GET", url, headers=headers, data=payload)
            download_link = json.loads(response.text)['data']['list'][0]['download_link']
            file_downloaded = json.loads(response.text)['data']['list'][0]['export_file_name']
            webbrowser.register('chrome', None, webbrowser.BackgroundBrowser("C:/Program Files/Google/Chrome/Application/chrome.exe"))
            webbrowser.get('chrome').open(download_link)
            checking = False
    time.sleep(3)

    #check download files
    file_path = os.path.join(download_folder, file_downloaded)
    while True:
        if os.path.exists(file_path):
            print("Download Completed")
            time.sleep(1)
            break
        else:

            time.sleep(5)
            x += 5
            print(f'Waiting time: {x}s')

def delete_file(folder):
    file_list = os.listdir(folder)
    #go through each file and remove them
    for file in file_list:
        file_path = os.path.join(folder, file)
        os.remove(file_path)
        print(f'Removed {file_path}')
    print(f'Complete removed all files in {folder}')

def move_and_extract(download_folder, destination_folder):
    #read files in download folder
    file_list_in_download_folder = os.listdir(download_folder)
    #go through each file and move them
    for file_name in file_list_in_download_folder:
      file_path = os.path.join(download_folder, file_name)
      # Check if the file is a zip file
      if file_name.endswith('.zip'):
          # Extract the zip file to destination folder
          print('Extracting files')
          with zipfile.ZipFile(file_path, 'r') as zip_ref:
              zip_ref.extractall(destination_folder)
          # Move the zip file to destination folder
          print('Moving files')
          shutil.move(file_path, destination_folder)
      # Check if the file is an Excel file
      elif file_name.endswith('.xlsx'):
          # Move the Excel file to folder 
          shutil.move(file_path, destination_folder)
    print('Completed moving files')

def read_files(folder):
    #check xlsx files in the folder
    excel_files = [file for file in os.listdir(folder) if file.endswith('.xlsx')]
    #create blank df
    df = []
    #read each file and append the the blank df
    for file in excel_files:
        file_path = os.path.join(folder, file)
        data = pd.read_excel(file_path)
        print(data)
        df.append(data)
    print('Completed reading all files')  
    #return a completely concated df  
    return pd.concat(df, ignore_index=True)
gc = gspread.service_account(r"C:\Users\tam.hoangthanh\Data\api_gsheet.json")
ck_sh = gc.open_by_key('18exqsmwyQuahwE9PYvJiOZmkKhHMM1NNW0P8MY7EdAQ')
ck_ws = ck_sh.worksheet('cookies')
cookies_vns = ck_ws.acell("B2").value
cookies_vnb = ck_ws.acell("B3").value
cookies_vnn = ck_ws.acell("B4").value
cookies_pms = ck_ws.acell("B5").value
token_pms = ck_ws.acell("B6").value
fbs_cookies = ck_ws.acell("B7").value


download_folder = r"C:\Users\tam.hoangthanh\Downloads"
vns_cells_list = r"C:\Users\tam.hoangthanh\data\ABC\vns_cells_list"
vns_sku_list = r"C:\Users\tam.hoangthanh\data\ABC\vns_sku_item_list"
vns_picking_inv_map = r"C:\Users\tam.hoangthanh\data\ABC\vns_picking_inv_map"

sku_item_payload = "{\"export_module\":8,\"task_type\":82,\"extra_data\":\"{\\\"timeRange\\\":1,\\\"module\\\":8,\\\"taskType\\\":82}\"}"

picking_inv_map_payload = "{\"export_module\":3,\"task_type\":86,\"extra_data\":\"{\\\"timeRange\\\":1,\\\"module\\\":3,\\\"taskType\\\":86,\\\"include_batch\\\":\\\"N\\\",\\\"export_angle\\\":1,\\\"pickup_type\\\":1}\"}"

cell_list_payload = "{\"export_module\":8,\"task_type\":29,\"extra_data\":\"{\\\"timeRange\\\":1,\\\"module\\\":8,\\\"taskType\\\":29}\"}"

folder_list = [download_folder,vns_cells_list,vns_sku_list,vns_picking_inv_map]
for fol in folder_list:
    delete_file(fol)

get_wms_report(cookies=cookies_vns, payloads=sku_item_payload, download_folder=download_folder)
move_and_extract(download_folder=download_folder, destination_folder=vns_sku_list)

get_wms_report(cookies=cookies_vns, payloads=picking_inv_map_payload, download_folder=download_folder)
move_and_extract(download_folder=download_folder, destination_folder=vns_picking_inv_map)

get_wms_report(cookies=cookies_vns, payloads=cell_list_payload, download_folder=download_folder)
move_and_extract(download_folder=download_folder, destination_folder=vns_cells_list)

df_vns_sku_list = read_files(vns_sku_list)
df_vns_picking_inv_map = read_files(vns_picking_inv_map)
df_vns_cells_list = read_files(vns_cells_list)


df_inv_use = df_vns_picking_inv_map[['SKU ID','SKU Name','Zone id','Location','Location ABC Classification','On-hand Qty']]
df_sku_use = df_vns_sku_list[['SKU ID','PMS/FBS UPC','UPC barcode3','UPC barcode4','Name','Volume(ml)', 'Net Weight(kg)']]
df_cell_use = df_vns_cells_list[['zone_id','location_id','abc_classification','cell_status','max_sku_qty_per_location','max_capacity(cu.cm.)', 'max_load(kg)']]
df_cell_use = df_cell_use[df_cell_use['cell_status'] == 'Normal']
df_merge1 = pd.merge(df_inv_use, df_cell_use, left_on='Location', right_on='location_id', how='outer')
df_merge2 = pd.merge(df_merge1, df_sku_use, left_on='SKU ID', right_on='SKU ID', how='left')
df_merge2['total_item_volume'] = df_merge2['Volume(ml)'] * df_merge2['On-hand Qty']
df_merge2['total_item_weight'] = df_merge2['Net Weight(kg)'] * df_merge2['On-hand Qty']

df_group_loc = df_merge2.groupby(['zone_id','location_id','abc_classification','cell_status','max_sku_qty_per_location','max_capacity(cu.cm.)', 'max_load(kg)']).agg({
    'SKU ID':'nunique',
    'total_item_volume':'sum',
    'total_item_weight':'sum'
}).reset_index()
df_group_loc = df_group_loc.rename(columns={'SKU ID':'count_sku_id'})

def apply_conditions(row):
    conditions = []
    if row['count_sku_id'] >= row['max_sku_qty_per_location']:
        conditions.append('max sku reached')
    if row['total_item_volume'] > row['max_capacity(cu.cm.)']:
        conditions.append('max volume reached')
    if row['total_item_weight'] > row['max_load(kg)']:
        conditions.append('max weight reached')
    if len(conditions) == 0:
        return 'available'
    else:
        return '-'.join(conditions)
    
df_group_loc['availability_check'] = df_group_loc.apply(apply_conditions, axis=1)


abc_datasheet = gc.open_by_key('17ohUl4pXtBNeXUjILwFeYj5X_w-aUDG6YpKR88uhNhs')
vns_avai_loc = abc_datasheet.worksheet('vns_available_location')
vns_avai_loc.clear()
df_group_loc = df_group_loc.astype(str)
vns_avai_loc.update("A1", [df_group_loc.columns.values.tolist()] + df_group_loc.values.tolist())

vns_sku_list = abc_datasheet.worksheet('vns_sku_list')
vns_sku_list.clear()
df_sku_use = df_sku_use.astype(str)
vns_sku_list.update("A1", [df_sku_use.columns.values.tolist()] + df_sku_use.values.tolist())
