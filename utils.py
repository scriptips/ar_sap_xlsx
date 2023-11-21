import glob
import logging
import os
import re as re
import shutil
import sys
import time
from datetime import datetime
from functools import wraps
from logging import handlers
from pathlib import Path
from typing import Callable

import numpy as np
import openpyxl as opx
import pandas as pd
import pywintypes
import win32com.client
import pythoncom


class PdExcel:
    '''
    data = {'Column 1': [1, 2, 3], 'Column 2': [4, 5, 6]}
    df = pd.DataFrame(data)
    with PdExcel("data.xlsx") as writer:
        df.to_excel(writer, index=False)
    '''
    def __init__(self, filename):
        self.filename = filename

    def __enter__(self):
        self.writer = pd.ExcelWriter (self.filename, engine='xlsxwriter')
        return self.writer

    def __exit__(self, type, value, traceback):
        self.writer.close()
class Open_Pyxl:
    '''
    Customized class for openpyxl context manager
    '''
    def __init__(self, filename):
        self.filename = filename
        self.wb = None

    def __enter__(self):
        self.wb = opx.load_workbook(self.filename)
        return self.wb

    def __exit__(self, *args):
        self.wb.save(self.filename)
        self.wb.close()
class Frontline:
    '''
    Frontline' class stored in the Utils.
    '''
    def __init__(self, abrev, code, sync_dir, email_main, email_cc, shrp_path):
        self.abrev = abrev
        self.code = code
        self.sync_dir = sync_dir
        self.email_main = email_main
        self.email_cc = email_cc
        self.shrp_path = shrp_path

def time_it(func):
    """
    This decorator function is to control execution time of the whole Script, and
    print an output in Minutes and Seconds upon full completion of the Script.
    Provides indication of how much time was needed for the SAP queries to be prepared, etc.
    """
    def wrapper(*args, **kwargs):
        start = time.time()
        func(*args, **kwargs)
        end = time.time()
        min = int((end - start) // 60)
        sec = int((end - start) % 60)
        def plur(m):
            return '' if m == 1 else "s"
        return print(f"\nThis Script completed in {min} minute{plur(min)} and {sec} second{plur(sec)}.\n")
    return wrapper

def send_email(main_receivers_arg, cc_receivers_arg, entity_updated_arg, link_to_upd_file_arg, if_to_send_arg='no'): 
    """
    This util function is for TR file user email notifications only.
    Few add features not used there, f.e. html or adding of attachments:
        - mail.HTMLBody = '<h3>This is HTML Body</h3>'
        - mail.Attachments.Add('c:\\sample.xlsx')
        - mail.Attachments.Add('c:\\sample2.xlsx').
    """
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = main_receivers_arg
    mail.Subject = f'{entity_updated_arg.upper()} TR file updated!'
    mail.Body = f"Link to the updated TR file here:" \
                f'{link_to_upd_file_arg}.\n\nComments are welcome in the light-green cells of the Column "AG" in the sheet "Customer Line Items".\nAvoid editing the file beyond the Comment cells, make an offline copy instead, if necessary.'
    mail.CC = cc_receivers_arg
    if if_to_send_arg == 'yes':
        print(f'{entity_updated_arg} Notifications sent to: {main_receivers_arg}"; "{main_receivers_arg}!!')
        return mail.Send()
    else: 
        print('Mails not sent!')

def clear_temp(dir):
    """
    The deleter() function checks the temporary file folder, and clears the old temporary files out of there.
    Better to place it before the start of the main script execution, so that temporary folder gets emptied only
    a moment before the new set of temporary files would replace the old ones. Reason for it is that somethimes old
    temporary files might turn out to be helpful in case they would be needed for some ad-hoc reports beyond the TR. 
    Like, f.e., 'SAP Query-De-Luxe' report (qdl file.xlsx, etc).
    """
    fileList = glob.glob(''.join([str(dir),"/*"]))
    for filePath in fileList:
        try:
            os.remove(filePath)
        except:
            print("Error while deleting file : ", filePath)

def move_file(src_file, dst_file):
    """
    move_file function is for removing the last TR version from ctry dir(-> synced dir)\n
    and placing it in tmp dir. Once the new TR is created, it is moved to the ctry dir(-> synced dir)
    """
    shutil.move(src_file, dst_file)

def select_frontlines(input_string:str , fl:str=['koe', 'kla', 'kli']) ->str:
        
        pattern = re.compile(r'\b(k[o0][e3]|k[l1][a4]|k[l1][i1]|b[a4][l1])\b', re.IGNORECASE)
        words = re.split(r'[\s,]+', input_string.lower())
        matches = [word for word in words if pattern.fullmatch(word)]
        while True:
            if input_string == 'bal' or input_string == fl or input_string == '':
                print(*['koe', 'kla', 'kli'])
                return ['koe', 'kla', 'kli']
            elif not matches:
                input_string = input("Error: Invalid input. Please try again..")
            else:
                print(matches)
                return matches
    
def set_now_date(args, nv=datetime.now().date()):
    if args =='':
        print(nv)
        return nv
    else:
        args = datetime.strptime(args, '%d%m%y').date() #parsing
        print(args)
        return args

def set_comparison_date(args:str, cd=datetime(year=((datetime.now()).year - 1), month=12, day=31).date()) -> datetime:
    if args == '':
        print(cd)
        return cd
    else:
        args = datetime.strptime(str(args), '%d%m%y').date().strftime('%Y-%m-%d')  #parsing
        print(args)
        return args  

def set_overdue_days(od:int=120) ->int:
    '''
    The function prompts user to input number of overdue days.\n
    The default function parameter is 120 overdue days.
    '''
    number_of_overdue_days = input("Key-in a number of overdue days for ageing analysis, and press Enter to continue...")
    if number_of_overdue_days == '':
        print(od)
        return od
    else:
        print(number_of_overdue_days)
        return number_of_overdue_days

def format_ordin_dt(df_head_string):
    dd = str(df[df_head_string].str.slice(start=8, stop=-9))
    mm = str(df[df_head_string].str.slice(start=5, stop=-12))
    yyyy = str(df[df_head_string].str.slice(start=0, stop=-15))
    return '.'.join([dd, mm, yyyy])

def format_bill_docs_in_df(df: pd.DataFrame):
    conditions = [
        (df['Reference'].str.len() != 10) | df['Reference'].str.len() == 10,
        (df['Reference'].str.len() == 10) & (df['Reference'].str.isdigit())]
    values = [df['Reference'].str.slice(start=0), df['Reference'].str.slice(start = 1)
              ]
    df['Reference'] = np.select(conditions, values)

def close_sap_excel_file(file_path):
    '''
    Intended for closing sap opened workbooks of the data exports. 
    SAP on it's own triggers excel, thus context manager approach
    was not known how to apply in that case.
    '''    
    times = 0 
    while times <= 30:
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
            excel.Workbooks.Close()
            excel.Quit()
            break
        except pywintypes.com_error as e:
            # print(f'The following error occurred {e}. Giving it a bit more time and retrying in a second ...')
            time.sleep(1)
            times += 1

def copy_file(src, dst):
    '''
    Full path to be specified for src and dst.
    Returns path of a copy.
    '''
    copied_to = shutil.copy(src, dst)
    return copied_to

def return_list_of_frontl_props(set_comp_func_arg, my_entities_arg):
    # prompts for user inputs if func is put in the below list wo assigning it to the variable.
    frontline_props = [Frontline(*x) for x in my_entities_arg if x[0] in set_comp_func_arg]
    return frontline_props

def rename_ar_fullrep_tmp(rename_ar_fullrep_tmp_arg, new_name_ar_fullrep_tmp_arg):
     renamed = os.rename(rename_ar_fullrep_tmp_arg, new_name_ar_fullrep_tmp_arg)
     return renamed

def prompt_continue()-> bool:
    '''
    Function prompts user to unput Space to continue the Script execution\n
    or any other key to stop the Script from further execution and exit.
    '''
    response = input("Press 'Space' to continue or any other to quit: ")
    if response.lower() == ' ':
        return True
    else:
        return False

def setup_logging(logfile_path) -> logging.Logger:
    '''
    TimedRotatingFileHandler class, rotates the log file every month (when='M')
    and keeps a maximum of 12 backup log files (backupCount=12). The interval argument specifies the interval
    between the creation of new log files. For example, if interval=1, a new log file
    will be created every month. Note that the TimedRotatingFileHandler class
    creates a new log file based on the system time, so it is important to make
    sure that the system time is accurate and synchronized. If the system time is
    incorrect, the log files may be created at unexpected times or not at all.
    With this setup, the log file will be cleared every month, and a new log file
    with a timestamp in the filename will be created. You can customize the filename
    format and other parameters of the TimedRotatingFileHandler class to suit your needs.
    
    logger = setup_logging(logfile_path)
    logger.debug('This is a debug message')
    logger.info('This is an info message')
    logger.warning('This is a warning message')
    logger.error('This is an error message')
    logger.critical('This is a critical message')
    
    '''
    # Create a logger and set the log level to INFO
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    # Create a timed rotating file handler and set the log level to INFO
    now = datetime.now()
    log_date_time = now.strftime('%d_%m_%Y')
    log_filename = Path(logfile_path, f'errors_{log_date_time}.log')

    file_handler = logging.handlers.TimedRotatingFileHandler(
    log_filename, when='D', backupCount=30) 

    file_handler = logging.FileHandler(log_filename)
    file_handler.setLevel(logging.INFO)

    # Create a stream handler to print messages to the console
    stream_handler = logging.StreamHandler(sys.stdout)

    # Create a formatter to specify the format of the log messages
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    stream_handler.setFormatter(formatter)

    # Add the file handler and stream handler to the logger
    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    
    return logger
    
def sap_connection_required(func):
    def wrapper(*args, **kwargs):
        times = 0
        while times <= 10:
            try:
                session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
                return func(session, *args, **kwargs)
            except pythoncom.com_error:
                time_str=datetime.now().strftime('%H:%M:%S')
                input(f'WARNING: {time_str} >>> Connect to SAP and Press Enter to retry.')
                times += 1
                if times == 5:
                    print(f'ERROR: {time_str} >>> Exiting after too many failed attempts.')
                    time.sleep(3)
                    sys.exit(1)
    return wrapper