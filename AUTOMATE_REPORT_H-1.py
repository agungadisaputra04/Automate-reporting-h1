"AUTHOR : AGUNG ADI SAPUTRA"
"DAILY REPORT INGW"

import os
import xlwings as xw
import pandas as pd
from datetime import datetime, timedelta
import csv
import re
import sys
import pandas as pd
import warnings
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string
import time
import pyfiglet
import itertools
import psutil
import subprocess
from openpyxl.chart import BarChart, Reference
import shutil
import colorama
from colorama import init, Fore, Back, Style


open_workbooks = []
DIR_HOME = os.getcwd()
DATE_TO_PROCESS_H_MIN_1 = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
DAY = (datetime.now() - timedelta(days=1)).strftime('%d')
DAY_STR = DAY.lstrip('0').zfill(2) 
YEAR = datetime.now().strftime('%Y')
TODAY = datetime.now().strftime('%d')
MONTH_MAP = {
    '01': 'Januari', '02': 'Februari', '03': 'Maret', '04': 'April', '05': 'Mei', '06': 'Juni',
    '07': 'Juli', '08': 'Agustus', '09': 'September', '10': 'Oktober', '11': 'November', '12': 'Desember'
}
MONTH = MONTH_MAP[datetime.now().strftime('%m')]

FILE_NAME_NOW = [
    f'pivot-cpu_mem_http_diameter_{DATE_TO_PROCESS_H_MIN_1}.csv',
    'sheet-limiters-TPS-GY.csv',
    'sheet-top_err-rest.csv',
    'sheet-ccr-diam.csv',
    f'sheet-ccr-diam_ingw-ocs-{DATE_TO_PROCESS_H_MIN_1}.csv',
    f'dblatency-diam-{DATE_TO_PROCESS_H_MIN_1}_pivot.csv',
    f'dblatency-http-{DATE_TO_PROCESS_H_MIN_1}_pivot.csv',
    f'db_latency_{DATE_TO_PROCESS_H_MIN_1}.csv',
    f'cpu_mem_http_diameter_{DATE_TO_PROCESS_H_MIN_1}.csv',
    f'sheet-tpssession-gy_{DATE_TO_PROCESS_H_MIN_1}.csv',
    f'sheet-tpssession-ro_{DATE_TO_PROCESS_H_MIN_1}.csv',
    f'sheet-tpssession-tps_{DATE_TO_PROCESS_H_MIN_1}.csv',
    'sheet-TPS_HTTP.csv',
    'sheet-TPS_OSS.csv',
    'sheet-wifi.csv'
]

FILE_NAME_H_MIN_1 = [
    f'Daily-Report-INGW-{DAY_STR}-{MONTH}-{YEAR}.xlsx',
]

DIR_H_MIN_1 = os.path.join(DIR_HOME, 'H-1')
DIR_NOW = os.path.join(DIR_HOME, 'NOW')

excel_file_h1 = os.path.join(DIR_H_MIN_1, f'Daily-Report-INGW-{DAY}-{MONTH}-{YEAR}.xlsx')
csv_file_now = os.path.join(DIR_NOW, 'sheet-limiters-TPS-GY.csv')
csv_file_ccr_diam = os.path.join(DIR_NOW, 'sheet-ccr-diam.csv')
csv_file_ccr_errrest= os.path.join(DIR_NOW, 'sheet-top_err-rest.csv')
csv_file_ccr_diam_ocs = os.path.join(DIR_NOW, f'sheet-ccr-diam_ingw-ocs-{DATE_TO_PROCESS_H_MIN_1}.csv')
csv_file_db_diam = os.path.join(DIR_NOW, f'dblatency-diam-{DATE_TO_PROCESS_H_MIN_1}_pivot.csv')

def open_workbook(file_path):
    global open_workbooks
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    open_workbooks.append(wb)
    return wb

def check_folder(directory):
    if not os.path.exists(directory):
        print(f'Folder {directory} tidak ditemukan')
        print(f'Membuat folder {directory}')
        os.makedirs(directory)
        return f'Folder {directory} | Created\n'
    else:
        return f'Folder {directory} - \033[92mEXIST!\033[0m\n'
    
def close_wps_if_running():
    wps_running = False
    
    for proc in psutil.process_iter(['pid', 'name']):
        if 'wps.exe' in proc.info['name'].lower():
            wps_running = True
            break
    
    if wps_running:
        input_confirm = input("Proses ini butuh untuk menutup WPS Office. Apakah Anda ingin menutupnya? (Y/N): ").strip().lower()
        if input_confirm == 'y':
            for proc in psutil.process_iter(['pid', 'name']):
                if 'wps.exe' in proc.info['name'].lower():
                    subprocess.run(['taskkill', '/f', '/pid', str(proc.info['pid'])])
                    print("WPS Office Closed.")
                    break
        else:
            print("Operasi ditolak oleh pengguna.")
            sys.exit('\nProses di batalkan\n')
    else:
        print("WPS Office tidak sedang berjalan.Proses dilanjutkan.")

close_wps_if_running()

######################################### Function to check files #####################################
def check_files(directory, files):
    result = ""
    files_not_found = []
    for file in files:
        if not os.path.exists(os.path.join(directory, file)):
            result += f'File {file} | \033[91mNot found\033[0m\n'
            files_not_found.append(file)
        else:
            result += f'File {file} - \033[92mFILE OK!\033[0m\n'

    if files_not_found:
        result += '\nProses di batalkan\n'
        print("File berikut tidak ditemukan, silahkan periksa kembali:\n")
        for file_not_found in files_not_found:
            print(file_not_found, "\033[91mPastikan file tanggal ini!\033[0m")
        sys.exit('\nProses di batalkan\n')
    return result
################################ CPU HEAP  ############################
def pindah_cpuheap(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['CPUHEAPSSR']
        start_row = 26
        end_row = 169
        num_cols = 181

        data = sheet.range((start_row, 1), (end_row, num_cols)).value
        sheet.range((2, 1), (end_row - start_row + 2, num_cols)).value = data
        sheet.range('146:169').api.Delete()

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat memindahkan data di sheet CPUHEAPSSR: {e}")
        raise

############################### CCR DIAM ########################
def hapus_kolom_ccrdiam(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['CCR-DIAM']

        for row_idx in range(2, 26):
            for col_idx in range(1, 6):
                sheet.range((row_idx, col_idx)).value = None

        wb.save()
        wb.close()
        app.quit()
        
    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data di kolom/baris di sheet CCR-DIAM: {e}")

#
def pindahkan_baris_ccrdiam(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['CCR-DIAM']
        max_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

        if max_row < 26:
            print("Tidak ada cukup baris untuk dipindahkan.")
            return

        for row_idx in range(26, max_row + 1):
            for col_idx in range(1, 6):
                cell_value = sheet.range((row_idx, col_idx)).value
                sheet.range((row_idx - 24, col_idx)).value = cell_value
                sheet.range((row_idx, col_idx)).value = None

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat memindahkan data di sheet CCR-DIAM: {e}")

def text_to_columns_and_moveccrdiam(excel_file, csv_file):
    try:
        app = xw.App(visible=False)
        df = pd.read_csv(csv_file, sep=';', header=None)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['CCR-DIAM']

        start_row = 146
        for idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                values = str(value).split(';')
                for sub_col_idx, sub_value in enumerate(values):
                    sheet.range((start_row + idx, col_idx + 1 + sub_col_idx)).value = sub_value

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat memproses file: {e}")

################################################### CCR OCS ###########################

def hapus_kolom_ccrdiamocs(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['CCR-OCS']

        for row_idx in range(2, 26):
            for col_idx in range(1, 6):
                sheet.range((row_idx, col_idx)).value = None

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data di kolom/baris di sheet CCR-DIAM: {e}")

def pindahkan_baris_ccrdiamocs(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['CCR-OCS']
        max_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

        if max_row < 26:
            print("Tidak ada cukup baris untuk dipindahkan.")
            return

        for row_idx in range(26, max_row + 1):
            for col_idx in range(1, 6):
                cell_value = sheet.range((row_idx, col_idx)).value
                sheet.range((row_idx - 24, col_idx)).value = cell_value
                sheet.range((row_idx, col_idx)).value = None

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat memindahkan data di sheet CCR-DIAM: {e}")

def text_to_columns_and_moveccrdiamocs(excel_file, csv_file):
    try:
        app = xw.App(visible=False)
        df = pd.read_csv(csv_file, sep=';')
        wb = xw.Book(excel_file)
        sheet = wb.sheets['CCR-OCS']

        start_row = 146
        for idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                values = str(value).split(';')
                for sub_col_idx, sub_value in enumerate(values):
                    sheet.range((start_row + idx, col_idx + 1 + sub_col_idx)).value = sub_value

        wb.save()
        wb.close()
        app.quit()
    except Exception as e:
        print(f"Terjadi kesalahan saat memproses file: {e}")

################################################### limiter  ########################
def read_data_from_files(excel_file, csv_file):
    try:
        app = xw.App(visible=False)
        df_excel = pd.read_excel(excel_file, sheet_name='Limiter', usecols="M:O", skiprows=4, nrows=32)
        df_csv = pd.read_csv(csv_file, header=None, skiprows=0, nrows=32, usecols=[1, 2, 3])
        return df_excel, df_csv
    except Exception as e:
        print(f"Terjadi kesalahan saat membaca file: {e}")
        return None, None

def replace_data_in_excel(excel_file, excel_sheetname, csv_data):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets[excel_sheetname]
        start_row = 5
        start_col = 13
        for idx, row in csv_data.iterrows():
            for col_idx, value in enumerate(row):
                sheet.range((start_row + idx, start_col + col_idx)).value = value

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat memperbarui data: {e}")

##################################################### Err rest HTTP #######################

def hapus_kolom_errhttp_a(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResHTTP']

        for row_idx in range(3, 13):
            sheet.range(f"B{row_idx}").clear_contents()

        for row_idx in range(17, 38):
            sheet.range(f"B{row_idx}").clear_contents()

        for col_idx in range(4, 38):
            col_letter = xw.utils.col_name(col_idx)
            sheet.range(f"{col_letter}4").clear_contents()

        wb.save()
        wb.close()
        app.quit()

        wb_check = xw.Book(excel_file)
        sheet_check = wb_check.sheets['TopErr&ResHTTP']

        for col_idx in range(4, 38):
            col_letter = xw.utils.col_name(col_idx)
            if sheet_check.range(f"{col_letter}4").value is not None:
                print(f"Cell {col_letter}4 not cleared")

    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data di kolom/baris di sheet TopErr&ResHTTP: {e}")

def moving_row_http(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResHTTP'] 
        start_col = 4 
        end_col = 40  
        start_row = 5
        end_row = 10

        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                col_letter = xw.utils.col_name(col)
                sheet.range(f"{col_letter}{row - 1}").value = sheet.range(f"{col_letter}{row}").value
        
        for col_idx in range(start_col, end_col + 1):
            col_letter = xw.utils.col_name(col_idx)
            sheet.range(f"{col_letter}{end_row}").clear_contents()

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f'Terjadi kesalahan saat memindahkan baris di sheet TopErr&ResHTTP: {e}')

def text_to_columns_and_move_errhttp(excel_file, csv_file):
    try:
        with open(csv_file, 'r', encoding='utf-8-sig') as file:
            reader = csv.reader(file, delimiter=';')
            data = list(reader)
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResHTTP']

        start_row_excel = 3
        start_row_csv = 2
        max_rows = min(len(data) - start_row_csv, 11)
        for i in range(max_rows):
            if len(data[start_row_csv + i]) > 1: 
                value = data[start_row_csv + i][1]  
                sheet.range((start_row_excel + i, 2)).value = value

        values_to_copy = [sheet.range((start_row_excel + i, 2)).value for i in range(max_rows)]
        start_row_target = 10
        start_col_target = 5 
        for i, value in enumerate(values_to_copy):
            sheet.range((start_row_target, start_col_target + i)).value = value

        h_min_1_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        sheet.range('D10').value = h_min_1_date

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

def text_to_columns_and_move_errhttpb(excel_file, csv_file):
    try:

        with open(csv_file, 'r', encoding='utf-8-sig') as file:
            reader = csv.reader(file, delimiter=';')
            data = list(reader)
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResHTTP']

        start_row_excel = 17  
        start_row_csv = 17 
        max_rows = min(len(data) - start_row_csv, 20) 
        for i in range(max_rows):
            if len(data[start_row_csv + i]) > 1:
                value = data[start_row_csv + i][1] 
                sheet.range((start_row_excel + i + 2, 2)).value = value 

        values_to_copy = [sheet.range((start_row_excel + i, 2)).value for i in range(max_rows)]
        start_row_target = 10 
        start_col_target = 18 
        for i, value in enumerate(values_to_copy):
            sheet.range((start_row_target, start_col_target + i)).value = value

        h_min_1_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        sheet.range('Q10').value = h_min_1_date

        wb.save()
        wb.close()
        app.quit()
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")


########################################################## Err Rest OSS #########################
def hapus_kolom_oss(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResOSS']

        for row_idx in range(3, 16):
            sheet.range(f"C{row_idx}").clear_contents()

        for row_idx in range(20, 40):
            sheet.range(f"C{row_idx}").clear_contents()

        for col_idx in range(5, 43):
            col_letter = xw.utils.col_name(col_idx)
            sheet.range(f"{col_letter}4").clear_contents()

        wb.save()
        wb.close()
        app.quit()

        wb_check = xw.Book(excel_file)
        sheet_check = wb_check.sheets['TopErr&ResOSS']

        for col_idx in range(5, 43):
            col_letter = xw.utils.col_name(col_idx)
            if sheet_check.range(f"{col_letter}4").value is not None:
                print(f"Cell {col_letter}4 not cleared")

    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data di kolom/baris di sheet TopErr&ResOSS: {e}")

def moving_row_oss(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResOSS'] 
        start_col = 5 
        end_col = 40
        start_row = 5
        end_row = 10

        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                col_letter = xw.utils.col_name(col)
                sheet.range(f"{col_letter}{row - 1}").value = sheet.range(f"{col_letter}{row}").value
        
        for col_idx in range(start_col, end_col + 1):
            col_letter = xw.utils.col_name(col_idx)
            sheet.range(f"{col_letter}{end_row}").clear_contents()

        wb.save()
        wb.close()
        app.quit()
    except Exception as e:
        print(f'Terjadi kesalahan saat memindahkan baris di sheet TopErr&ResOSS: {e}')

def text_to_columns_and_move_erross(excel_file, csv_file):
    try:
        with open(csv_file, 'r', encoding='utf-8-sig') as file:
            reader = csv.reader(file, delimiter=';')
            data = list(reader)
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResOSS']

        start_row_excel = 2  
        start_row_csv = 48  

        max_rows = min(len(data) - start_row_csv, 14) 
        for i in range(max_rows):
            if len(data[start_row_csv + i]) > 1: 
                value = data[start_row_csv + i][1] 
                sheet.range((start_row_excel + i, 3)).value = value 
        values_to_copy = [sheet.range((row, 3)).value for row in range(3, 17)]
        start_row_target = 10 
        start_col_target = 6 
        for i, value in enumerate(values_to_copy):
            sheet.range((start_row_target, start_col_target + i)).value = value

        h_min_1_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        sheet.range('E10').value = h_min_1_date

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

def text_to_columns_and_move_errossb(excel_file, csv_file):
    try:
        with open(csv_file, 'r', encoding='utf-8-sig') as file:
            reader = csv.reader(file, delimiter=';')
            data = list(reader)
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResOSS']

        start_row_excel = 19
        start_row_csv = 64

        max_rows = min(len(data) - start_row_csv, 21) 
        for i in range(max_rows):
            if len(data[start_row_csv + i]) > 1: 
                value = data[start_row_csv + i][1]  
                sheet.range((start_row_excel + i, 3)).value = value  
        values_to_copy = [sheet.range((row, 3)).value for row in range(20, 41)]
        start_row_target = 10
        start_col_target = 22
        for i, value in enumerate(values_to_copy):
            sheet.range((start_row_target, start_col_target + i)).value = value

        h_min_1_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        sheet.range('U10').value = h_min_1_date

        wb.save()
        wb.close()
        app.quit()
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

############################################ Err Rest Diam ##############################
def hapus_kolom_errdiam(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResDiam']

        for row_idx in range(2, 8):
            sheet.range(f"B{row_idx}").clear_contents()

        for col_idx in range(4, 11):
            col_letter = xw.utils.col_name(col_idx)
            sheet.range(f"{col_letter}3").clear_contents()

        wb.save()
        wb.close()
        app.quit()

        wb_check = xw.Book(excel_file)
        sheet_check = wb_check.sheets['TopErr&ResDiam']

        for col_idx in range(4, 11):
            col_letter = xw.utils.col_name(col_idx)
            if sheet_check.range(f"{col_letter}3").value is not None:
                print(f"Cell {col_letter}3 not cleared")

    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data di kolom/baris di sheet TopErr&ResDIAM: {e}")

def moving_row_diam(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResDiam'] 
        start_col = 4
        end_col = 10
        start_row = 4
        end_row = 9

        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                col_letter = xw.utils.col_name(col)
                sheet.range(f"{col_letter}{row - 1}").value = sheet.range(f"{col_letter}{row}").value
        
        for col_idx in range(start_col, end_col + 1):
            col_letter = xw.utils.col_name(col_idx)
            sheet.range(f"{col_letter}{end_row}").clear_contents()

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f'Terjadi kesalahan saat memindahkan baris di sheet TopErr&ResOSS: {e}')


def text_to_columns_and_move_errdiam(excel_file, csv_file):
    try:
        with open(csv_file, 'r', encoding='utf-8-sig') as file:
            reader = csv.reader(file, delimiter=';')
            data = list(reader)
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TopErr&ResDIAM']
        ##e3002
        csv_row_index = 39 
        if len(data) > csv_row_index and len(data[csv_row_index]) > 1:
            value = data[csv_row_index][1]
        else:
            value = ''
        excel_row_index = 1 
        sheet.range((excel_row_index + 1, 2)).value = value
        ##e5003
        csv_row_index = 44 
        if len(data) > csv_row_index and len(data[csv_row_index]) > 1:
            value = data[csv_row_index][1] 
        else:
            value = ''
        excel_row_index = 2
        sheet.range((excel_row_index + 1, 2)).value = value
        ##e5003
        csv_row_index = 45
        if len(data) > csv_row_index and len(data[csv_row_index]) > 1:
            value = data[csv_row_index][1] 
        else:
            value = ''
        excel_row_index = 4
        sheet.range((excel_row_index + 1, 2)).value = value
        ##e5012
        csv_row_index = 40
        if len(data) > csv_row_index and len(data[csv_row_index]) > 1:
            value = data[csv_row_index][1] 
        else:
            value = ''
        excel_row_index = 5
        sheet.range((excel_row_index + 1, 2)).value = value
        ##e5030
        csv_row_index = 42
        if len(data) > csv_row_index and len(data[csv_row_index]) > 1:
            value = data[csv_row_index][1] 
        else:
            value = ''
        excel_row_index = 6
        sheet.range((excel_row_index + 1, 2)).value = value

        values_to_copy = [sheet.range((row, 2)).value for row in range(2, 8)]
        start_row_target = 9 
        start_col_target = 5 
        for i, value in enumerate(values_to_copy):
            sheet.range((start_row_target, start_col_target + i)).value = value
        h_min_1_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        sheet.range('D9').value = h_min_1_date

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

############################################### DB DIAM ########

def hapus_kolom_dbdiam(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['dbLatency&QueueDiam']
        for row_idx in range(2, 145): 
            for col_idx in range(1, 20):
                sheet.range((row_idx, col_idx)).value = None
        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data di kolom/baris di sheet dbLatency&QueueDiam: {e}")

def pindahkan_baris_dbdiam(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['dbLatency&QueueDiam']
        max_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        if max_row < 145:
            print("Tidak ada cukup baris untuk dipindahkan.")
            return
        data = sheet.range((145, 1), (1005, 20)).value
        sheet.range((2, 1), (2 + (1005 - 145), 20)).value = data
        #buka ini kalo mau hapus sheet.range((145, 1), (1146, 20)).value = None
        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat memindahkan data di sheet dbLatency&QueueDiam: {e}")

def text_to_columns_and_dbdiam(excel_file, csv_file):
    try:
        app = xw.App(visible=False)
        df = pd.read_csv(csv_file, header=None, usecols=[0, 1, 2, 3], skiprows=1, sep=';')
        wb = xw.Book(excel_file)
        sheet = wb.sheets['dbLatency&QueueDiam']
        start_excel_row = 863
        for idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                sheet.range((start_excel_row + idx, col_idx + 1)).value = value
        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat memproses file: {e}")

def text_to_columns_and_dbdiamb(excel_file, csv_file):
    try:
        app = xw.App(visible=False)
        df = pd.read_csv(csv_file, header=None, usecols=[4, 5, 6], skiprows=1, sep=';')
        wb = xw.Book(excel_file)
        sheet = wb.sheets['dbLatency&QueueDiam'] 
        start_excel_row = 863
        for idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                sheet.range((start_excel_row + idx, col_idx + 8)).value = value
        wb.save()
        wb.close()
        app.quit()
    except Exception as e:
        print(f"Terjadi kesalahan saat memproses file: {e}")

################################################### 
#              DB HTTP, TPS IS SOON               #

def hapus_kolom_a_tpshttp(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TPS-HTTP']
        for row_idx in range(2, 4): 
            for col_idx in range(1, 4):
                sheet.range((row_idx, col_idx)).value = None
        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data di kolom/baris di sheet dbLatency&QueueDiam: {e}")

def hapus_kolom_b_tpshttp(excel_file):
    try:
        app = xw.App(visible=False)
        wb = xw.Book(excel_file)
        sheet = wb.sheets['TPS-HTTP']
        for row_idx in range(2, 145): 
            for col_idx in range(18, 21):
                sheet.range((row_idx, col_idx)).value = None
        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data di kolom/baris di sheet dbLatency&QueueDiam: {e}")


# def pindahkan_baris_a_tpshttp(excel_file):
#     try:
#         app = xw.App(visible=False)
#         wb = xw.Book(excel_file)
#         sheet = wb.sheets['TPS-HTTP']
#         max_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
#         if max_row < 145:
#             print("Tidak ada cukup baris untuk dipindahkan.")
#             return
#         data = sheet.range((145, 1), (1005, 20)).value
#         sheet.range((2, 1), (2 + (1005 - 145), 20)).value = data
#         #buka ini kalo mau hapus sheet.range((145, 1), (1146, 20)).value = None
#         wb.save()
#         wb.close()
#         app.quit()

#     except Exception as e:
#         print(f"Terjadi kesalahan saat memindahkan data di sheet dbLatency&QueueDiam: {e}")
# def pindahkan_baris_b_tpshttp(excel_file):
#     try:
#         app = xw.App(visible=False)
#         wb = xw.Book(excel_file)
#         sheet = wb.sheets['TPS-HTTP']
#         max_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
#         if max_row < 145:
#             print("Tidak ada cukup baris untuk dipindahkan.")
#             return
#         data = sheet.range((145, 1), (1005, 20)).value
#         sheet.range((2, 1), (2 + (1005 - 145), 20)).value = data
#         #buka ini kalo mau hapus sheet.range((145, 1), (1146, 20)).value = None
#         wb.save()
#         wb.close()
#         app.quit()

#     except Exception as e:
#         print(f"Terjadi kesalahan saat memindahkan data di sheet dbLatency&QueueDiam: {e}")

# def text_to_columns_and_tpshttp(excel_file, csv_file):
#     try:
#         app = xw.App(visible=False)
#         df = pd.read_csv(csv_file, header=None, usecols=[0, 1, 2, 3], skiprows=1, sep=';')
#         wb = xw.Book(excel_file)
#         sheet = wb.sheets['TPS-HTTP']
#         start_excel_row = 863
#         for idx, row in df.iterrows():
#             for col_idx, value in enumerate(row):
#                 sheet.range((start_excel_row + idx, col_idx + 1)).value = value
#         wb.save()
#         wb.close()
#         app.quit()

#     except Exception as e:
#         print(f"Terjadi kesalahan saat memproses file: {e}")

# def text_to_columns_and_tpshttp(excel_file, csv_file):
#     try:
#         app = xw.App(visible=False)
#         df = pd.read_csv(csv_file, header=None, usecols=[4, 5, 6], skiprows=1, sep=';')
#         wb = xw.Book(excel_file)
#         sheet = wb.sheets['dbLatency&QueueDiam'] 
#         start_excel_row = 863
#         for idx, row in df.iterrows():
#             for col_idx, value in enumerate(row):
#                 sheet.range((start_excel_row + idx, col_idx + 8)).value = value
#         wb.save()
#         wb.close()
#         app.quit()
#     except Exception as e:
#         print(f"Terjadi kesalahan saat memproses file: {e}")




###################################################
init(autoreset=True)

def animated_process(message, duration=20):
    end_time = time.time() + duration
    spinner = itertools.cycle([
        Fore.RED + '|' + Style.RESET_ALL,
        Fore.YELLOW + '/' + Style.RESET_ALL,
        Fore.GREEN + '-' + Style.RESET_ALL,
        Fore.BLUE + '\\' + Style.RESET_ALL,
    ])

    while time.time() < end_time:
        sys.stdout.write(f'\r{message} {next(spinner)}')
        sys.stdout.flush()
        time.sleep(0.1)

    sys.stdout.write(f'\r{message}   ')
    sys.stdout.flush()

def run_bro(process_name, function, *args):
    try:
        animated_process(f'{Back.YELLOW + Fore.WHITE}Processing{Style.RESET_ALL} {process_name}', duration=10)
        function(*args)
        sys.stdout.write(f'\r{process_name} {Back.GREEN + Fore.WHITE} SUCCEEDED {Style.RESET_ALL}\n')
        sys.stdout.flush()
    except Exception as e:
        sys.stdout.write(f'\r{process_name} {Back.RED + Fore.WHITE} FAILED {Style.RESET_ALL}: {e}\n')
        sys.stdout.flush()
        sys.exit(1)

def headertext(text, footer):
    graffiti_text = pyfiglet.figlet_format(text)
    colors = [Fore.MAGENTA]
    graffiti_lines = graffiti_text.split('\n')
    
    for i, line in enumerate(graffiti_lines):
        color = colors[i % len(colors)]
        if i == len(graffiti_lines) - 2: 
            print(color + line + Style.RESET_ALL + ' ' + footer)
        elif i == len(graffiti_lines) - 1:
            print(color + line + Style.RESET_ALL)
        else:
            print(color + line + Style.RESET_ALL)

def save_and_close_workbooks():
    global open_workbooks
    for wb in open_workbooks:
        wb.save()
        wb.close()
    if open_workbooks:
        open_workbooks[0].app.quit()

def rename(lama, baru):
    try:
        if os.path.exists(lama):
            shutil.move(lama, baru)
            print(f" ")
        else:
            print(f'Daily-Report-INGW-{TODAY}-{MONTH}-{YEAR}.xlsx')
            
    except Exception as e:
            print("Terjadi kesalahan")

if __name__ == "__main__":
    try:
        headertext("INGW REPORT", "powered by Icharming")
        time.sleep(2)

        print(Fore.YELLOW + Back.BLUE + 'Checking Directories...' + Style.RESET_ALL)
        time.sleep(1)
        print(check_folder(DIR_NOW))
        print(check_folder(DIR_H_MIN_1))
        
        print(Fore.YELLOW + Back.BLUE + 'Checking Files...' + Style.RESET_ALL)
        time.sleep(1)
        print(check_files(DIR_NOW, FILE_NAME_NOW))
        print(check_files(DIR_H_MIN_1, FILE_NAME_H_MIN_1))
        
        print(Fore.YELLOW + Back.BLUE + 'Processing Your Report' + Style.RESET_ALL)
        print('--------------------------')
        time.sleep(1)

        #heap
        run_bro('Deleting Rows CPU HEAP', pindah_cpuheap, excel_file_h1)
        #Limiter
        df_excel, df_csv = read_data_from_files(excel_file_h1, csv_file_now)
        if df_excel is not None and df_csv is not None:
            run_bro('Updating Data Limiter', replace_data_in_excel, excel_file_h1, 'Limiter', df_csv)

        # CCRDIAM
        run_bro('Deleting Rows CCRDIAM', hapus_kolom_ccrdiam, excel_file_h1)
        run_bro('Moving Data CCRDIAM', pindahkan_baris_ccrdiam, excel_file_h1)
        run_bro('Updating Data CCRDIAM', text_to_columns_and_moveccrdiam, excel_file_h1, csv_file_ccr_diam)

        # CCROCS
        run_bro('Deleting Rows CCROCS', hapus_kolom_ccrdiamocs, excel_file_h1)
        run_bro('Moving Data CCROCS', pindahkan_baris_ccrdiamocs, excel_file_h1)
        run_bro('Updating Data CCROCS', text_to_columns_and_moveccrdiamocs, excel_file_h1, csv_file_ccr_diam_ocs)

        # RESTHTTP
        run_bro('Deleting Rows RESTHTTP', hapus_kolom_errhttp_a, excel_file_h1)
        run_bro('Moving Data RESTHTTP', moving_row_http, excel_file_h1)
        run_bro('Updating Data A RESTHTTP', text_to_columns_and_move_errhttp, excel_file_h1, csv_file_ccr_errrest)
        run_bro('Updating Data B RESTHTTP', text_to_columns_and_move_errhttpb, excel_file_h1, csv_file_ccr_errrest)

        # RESTOSS
        run_bro('Deleting Rows RESTOSS', hapus_kolom_oss, excel_file_h1)
        run_bro('Moving Data RESTOSS', moving_row_oss, excel_file_h1)
        run_bro('Updating Data A RESTOSS', text_to_columns_and_move_erross, excel_file_h1, csv_file_ccr_errrest)
        run_bro('Updating Data B RESTOSS', text_to_columns_and_move_errossb, excel_file_h1, csv_file_ccr_errrest)

        # RESTDIAM
        run_bro('Deleting Rows RESTDIAM', hapus_kolom_errdiam, excel_file_h1)
        run_bro('Moving Data RESTDIAM', moving_row_diam, excel_file_h1)
        run_bro('Updating Data RESTDIAM', text_to_columns_and_move_errdiam, excel_file_h1, csv_file_ccr_errrest)

        # DB_DIAM
        run_bro('Deleting Rows dbLatency_DIAM', hapus_kolom_dbdiam, excel_file_h1)
        run_bro('Moving Data dbLatency_DIAM', pindahkan_baris_dbdiam, excel_file_h1)
        run_bro('Updating Data A dbLatency_DIAM', text_to_columns_and_dbdiam, excel_file_h1, csv_file_db_diam)
        run_bro('Updating Data B dbLatency_DIAM', text_to_columns_and_dbdiamb, excel_file_h1, csv_file_db_diam)
        print('--------------------------')
        print('UPDATEING ALL DATA ' + Back.GREEN + Fore.BLACK + ' SUCCESSEFULLY ' + Style.RESET_ALL)
        

        lama_file = os.path.join(DIR_H_MIN_1, f'Daily-Report-INGW-{DAY_STR}-{MONTH}-{YEAR}.xlsx')
        baru_file = os.path.join(DIR_H_MIN_1, f'Daily-Report-INGW-{TODAY}-{MONTH}-{YEAR}.xlsx')
        rename(lama_file, baru_file)
        print(Back.BLUE + Fore.WHITE + ' NEW FILE ' + Style.RESET_ALL, baru_file)
        print("\n")
        print("Note: if there is appsid activity please fill it manually.")
        print("\n")
        save_and_close_workbooks()
    except Exception as e:
        print(f"\nTerjadi kesalahan saat memproses file: {e}\n")