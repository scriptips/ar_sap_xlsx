import os
import re
import warnings
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import psutil
import xlsxwriter
import xlwings as xw
from const import (ar_data_tmp, arch_dir, bill_doc_list_tmp, bridge_data_tmp,
                   bubble_2d_tmp, bubble_3d_tmp, cust_line_items_tmp,
                   cust_master_tmp, date_str, doc_type_descrps,
                   generate_file_name, gl_descrps, header_list,
                   my_column_head_dscptns, my_entities,
                   my_qdl_column_head_dscptns, old_ar_data_tmp,
                   old_cust_line_items_tmp, qdl_sap_raw, qdl_tmp,
                   stack_data_tmp, temp_dir, time_str, wtf_chart_tmp, xlsx)
from openpyxl.utils import column_index_from_string
from plotly import graph_objects as go
from utils import (PdExcel, clear_temp, close_sap_excel_file, copy_file,
                   format_bill_docs_in_df, move_file, rename_ar_fullrep_tmp,
                   return_list_of_frontl_props, send_email, sap_connection_required)

@sap_connection_required
def prep_sap_qdl_file(session, fl_code_arg, temp_path_parent_arg, temp_path_name_arg):
    """
    SAP may not always recognize the Path object from the pathlib library.
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(temp_path_parent_arg)
    prep_qdl(qdl_tmp.parent, qdl_tmp.name)
    """
    # session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
    session.findById("wnd[0]").resizeWorkingPane(354, 42, False)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZV_QUERY_DL"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = ""
    session.findById("wnd[1]/usr/txtENAME-LOW").text = f""
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 5
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = f"{fl_code_arg} ALL PROJ"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 12
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(temp_path_parent_arg)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = temp_path_name_arg
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey(0)
    close_sap_excel_file(Path(temp_path_parent_arg, temp_path_name_arg))

@sap_connection_required
def prep_sap_cust_mast_data_file(session, fl_code_arg, temp_path_parent_arg, temp_path_name_arg):
    # session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
    session.findById("wnd[0]").resizeWorkingPane(354, 42, False)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nsqvi"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050").getAbsoluteRow(0).selected = False
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050").getAbsoluteRow(1).selected = True
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050/txtRS38R-QNAME1[0,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050/txtRS38R-QNAME1[0,1]").caretPosition = 0
    session.findById("wnd[0]/usr/btnP1").press()
    session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = fl_code_arg  # no
    session.findById("wnd[0]/usr/ctxtSP$00002-HIGH").text = fl_code_arg  # līdz...
    session.findById("wnd[0]/usr/ctxtSP$00002-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtSP$00002-HIGH").caretPosition = 3
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    cust_accts = pd.read_clipboard()
    cust_accts = cust_accts.drop(cust_accts.columns[list(range(1,5))], axis=1)
    cust_accts = cust_accts.iloc[7:].rename(columns={cust_accts.columns[0]: "Customer"})
    cust_accts["Customer"] = cust_accts["Customer"].str.replace("\|", "", regex=True)
    cust_accts.reset_index(drop=True, inplace=True)
    cust_accts.set_index("Customer", inplace=True)
    session.findById("wnd[0]").resizeWorkingPane(354, 42, False)
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050").getAbsoluteRow(1).selected = False
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050").getAbsoluteRow(0).selected = True
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050/txtRS38R-QNAME1[0,0]").setFocus()
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050/txtRS38R-QNAME1[0,0]").caretPosition = 0
    session.findById("wnd[0]/usr/btnP1").press()
    session.findById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").press()
    cust_accts.to_clipboard()  # nosūta uz atmiņu pirms peistošanas sapā
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]").resizeWorkingPane(354,42,False)  # nosūtīt Customer+Nme1 tabulu uz clipboard
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(temp_path_parent_arg)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = temp_path_name_arg
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey(0)
    close_sap_excel_file(Path(temp_path_parent_arg, temp_path_name_arg))

@sap_connection_required
def prep_sap_cust_line_items_file(session, fl_code_arg, temp_path_parent_arg, temp_path_name_arg, set_date_arg):
    """
    Prepares Customer Line item report from SAP by using tcode FBL5N.
    It is separated from pd df part because it could be wanted as a standalone SAP FBL5N excel export.
    Excel file later closed with the uttil func 'close_sap_excel_file()'.
    """
    # session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
    session.findById("wnd[0]").resizeWorkingPane(354, 42, False)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00010")
    session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = fl_code_arg
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-HIGH").text = fl_code_arg
    session.findById("wnd[0]/usr/ctxtPA_STIDA").text = set_date_arg.strftime('%d%m%y')
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "KLA AR LIST"
    session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 11
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(temp_path_parent_arg) # SAP sometimes does not accept Path lib objs
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = str(temp_path_name_arg)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey(0)
    close_sap_excel_file(Path(temp_path_parent_arg, temp_path_name_arg))

@sap_connection_required
def prep_sap_bill_so_tab(session, cust_line_items_arg, temp_path_parent_arg, temp_path_name_arg):
    '''
    Copies billing document numbers from the earlier saved Customer Line Item report, formats the df,\n
    and uses as selections in the SAP SQVI table query BAL_VBRP and creates a matching table of blld/so numbers.\n
    It is later saved in the xlsx file: "bill_doc_list_tmp dd-mm-yyyy (hh.mm.ss).xlsx".
    '''
    # session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)

    bill_doc_df = pd.read_excel(cust_line_items_arg)
    bill_doc_df = bill_doc_df[['Reference']]  # Delete all but one column? Use [[..]] - and other columns of the df will be garbage collected.
    bill_doc_df.dropna(inplace=True)  # use inplace or reasign variable
    bill_doc_df.drop_duplicates(inplace=True)

    conditions = [
        (bill_doc_df['Reference'].str.len() != 10) | bill_doc_df['Reference'].str.len() == 10,
        (bill_doc_df['Reference'].str.len() == 10) & (bill_doc_df['Reference'].str.isdigit())]
    
    # diabling warnings for that part of the code..... where slicing is mentioned to get rid of "Slicer List extension is not supported and will be removed warn(msg)" message.
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    values = [bill_doc_df['Reference'].str.slice(start=0), bill_doc_df['Reference'].str.slice(start = 1)
            ]
    bill_doc_df['Reference'] = np.select(conditions, values)

    bill_doc_df['is_digit'] = bill_doc_df['Reference'].str.isdigit()  # add suplementary boolean evaluationary column to the frame
    bill_doc_df['complies_with_length'] = bill_doc_df['Reference'].str.len() == 9
    bill_doc_df.drop(bill_doc_df[bill_doc_df['is_digit'] == False].index, inplace=True)
    bill_doc_df.drop(bill_doc_df[bill_doc_df['complies_with_length'] == False].index, inplace=True)
    bill_doc_df = bill_doc_df[['Reference']]  # repeat to send to garbage the 'My_Boolean' column that has become redundant
    bill_doc_df.to_clipboard(index=False, header=False)
    session.findById("wnd[0]").resizeWorkingPane(354, 42, False)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00176"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00176")
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050").getAbsoluteRow(0).selected = False
    session.findById("wnd[0]/usr/tblSAPMS38RTV3050").getAbsoluteRow(2).selected = True
    session.findById("wnd[0]/usr/btnP1").press()
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtSP$00001-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(temp_path_parent_arg)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = temp_path_name_arg
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey(0)
    close_sap_excel_file(Path(temp_path_parent_arg, temp_path_name_arg))

def prep_qdl_sheet_file(source_file_arg:str, dest_file_arg:str):
    '''
    This function uses sap qdl raw excel file to add four additional columns:
    - "C2C",
    - "Actual Status",
    - "Actual Status Date",
    - "Projects (entire) not yet fully Billed?"
    The sheet 'Qdl' from this temporary file will be added to the TR file.
    '''
    df_qdl = pd.read_excel(source_file_arg)

    with pd.ExcelWriter(dest_file_arg) as writer:
        df_qdl.to_excel(writer, sheet_name='Qdl', startrow=0, index=False, header=True)
        worksheet = writer.sheets['Qdl']
        r = len(df_qdl.index)
        c = len(df_qdl.columns)
        worksheet.add_table(0, 0, r, len(my_qdl_column_head_dscptns)-1, {'name': 'qdl_table',
                                        'header_row': True,
                                        'style': None,
                                        'columns': my_qdl_column_head_dscptns})  # šis glabājas atsevišķā modulī man.
        # appends qdl table with four additoinal columns 
        x = 1
        while x <= r:
            worksheet.write_formula(x, c, f'=AL{x+1}-AQ{x+1}')
            worksheet.write_formula(x, c+1, f'=IF(AB{x+1}="", IF(Z{x+1}="", "Act. 0B", "Act. 3C"), "Act. 5")')
            worksheet.write_formula(x, c+2, f'=IF(IF(AB{x+1}="", IF(Z{x+1}="", X{x+1}, Z{x+1}), AB{x+1})=0, "", IF(AB{x+1}="", IF(Z{x+1}="", X{x+1}, Z{x+1}), AB{x+1}))')
            worksheet.write_formula(x, c+3, f"=(SUMIFS(AL2:AL{r+1}, C2:C{r+1}, C{x+1})-SUMIFS(AC2:AC{r+1}, C2:C{r+1}, C{x+1}))<>0")
            x += 1
    return dest_file_arg

def prep_df_and_wrt_ar_file(qdl_tmp_arg, now_date_arg, cust_line_items_arg, cust_master_data_arg, bill_doc_list_arg, comm_src_arg, ar_data_arg):
    # pd read files:
    # - creates pd df from the Customer line items excel file (note: converters are important to format the fields "Account" and "Posting key").
    line_items_df = pd.read_excel(cust_line_items_arg, converters={'Account': str, 'Posting Key': str}) #target_path
    customers_df = pd.read_excel(cust_master_data_arg)
    bd_so_temp_df = pd.read_excel(bill_doc_list_arg, converters = {'Bill.Doc.':str,'Sales Doc.':str})
    qdl_tmp_df = pd.read_excel(qdl_tmp_arg, converters={'SO Doc N°':str})

    customers_df_dict = dict(zip(customers_df.iloc[:, 0], customers_df.iloc[:, 1]))

    # imported custom util func from the utils module
    format_bill_docs_in_df(line_items_df)

    # reindexing the lines in the df
    line_items_df = line_items_df.reindex(columns=header_list)  

    # fills blank cells with string 'NABU' that are not applicable to business line (or unit, in other words)
    line_items_df['Business Area1'] = line_items_df['Business Area']
    line_items_df['Business Area1'] = line_items_df['Business Area1'].replace(np.nan, 'NABU', regex=True)


    line_items_df['GL Account Description'] = line_items_df['G/L Account']
    line_items_df['GL Account Description'] = line_items_df['GL Account Description'].map(gl_descrps)
    line_items_df['Type'] = line_items_df['Document Type']
    line_items_df['Type'] = line_items_df['Type'].map(doc_type_descrps)
    line_items_df['Customer Name'] = line_items_df['Account']
    line_items_df['Customer Name'] = line_items_df['Customer Name'].map(customers_df_dict)

    # Convert to int.
    line_items_df['Company Code'] = line_items_df['Company Code'].astype(float).round().astype(int) 
    line_items_df['Document Number'] = line_items_df['Document Number'].astype(float).round().astype(int)

    line_items_df['Line item'] = line_items_df['Line item'].map(str) #


    for date_column in ['Document Date', 'Posting Date']:
        line_items_df[date_column] = line_items_df[date_column].dt.strftime('%d.%m.%Y') # Format the dates as strings in the dd.mm.yyyy format


    line_items_df['IndexKey'] = line_items_df['Company Code'].map(str) +\
                                line_items_df['Document Number'].map(str) +\
                                line_items_df['Line item'].map(str) +\
                                line_items_df['Document Date'] +\
                                line_items_df['Posting Date']

    # before concatenating replaces NaN by blanks to align with  the old non py xlsx outputs..
    line_items_df['Reference'] = line_items_df['Reference'].replace(np.nan, '', regex = True)
    line_items_df['Concat'] = line_items_df['Reference'].map(str) +\
                                line_items_df['Company Code'].map(str)


    # converts data strings to dates for the following date columns:
    for date_str in ['Document Date', 'Posting Date', 'Net due date']:
        line_items_df[date_str] = pd.to_datetime(line_items_df[date_str], dayfirst=True).dt.date
        

    line_items_df['Overdue Days'] = now_date_arg
    line_items_df['Overdue Days'] = pd.to_datetime(line_items_df['Overdue Days'], format='%Y-%m-%d')
    line_items_df['Net due date'] = pd.to_datetime(line_items_df['Net due date'], format='%Y-%m-%d')
    line_items_df['Overdue Days'] = line_items_df['Overdue Days'] - line_items_df['Net due date']
    line_items_df['Overdue Days'] = (line_items_df['Overdue Days']).dt.days
    line_items_df['Net due date'] = pd.to_datetime(line_items_df['Net due date'], format='%Y-%m-%d %H:%M:%S').dt.date # converts from hh:mm:ss to simple date format like yyyy-mm-dd.

    # adds new column 'Ageing' works out categories from the given list of ageing tresholds. List comprehension.. Bins...
    ranges = [1, 31, 61, 91, 121, 181, 361, 1000000]
    values = ['       Not Overdue', '      1-30', '     31-60', '    61-90', '   91-120', '  121-180', ' 181-360', '>360']
    line_items_df['Ageing'] = pd.cut(line_items_df['Overdue Days'], bins=[-1000000] +  ranges, # list comprehension
                                    labels=values, right=False)
    
    # Normalizing company name str. Needed if comp name in other languages (french, etc)
    line_items_df['ic?'] = line_items_df['Customer Name']
    norm_series = pd.Series(line_items_df['ic?'])
    normalization = norm_series.str.normalize('NFKD')\
                                .str.encode('ascii', errors='ignore')\
                                .str.decode('utf-8')
    line_items_df['ic?'] = normalization
    conditions = [
        (line_items_df['ic?'].str.find('KONE') != -1),
        (line_items_df['ic?'].str.find('KONE') == -1)]
    values = ['IC', 'Ext']
    line_items_df['ic?'] = np.select(conditions, values)
    
    comment_source_df = pd.read_excel(comm_src_arg, sheet_name='Customer Line Items', skiprows=[0])
    comment_source_df = comment_source_df.dropna(subset=['IndexKey'])

    line_items_df['COMMENTS(to be added in the Sheet "Customer Line Items", column "AG"…)'] = line_items_df['IndexKey'].map(comment_source_df.set_index('IndexKey')['COMMENTS(to be added in the Sheet "Customer Line Items", column "AG"…)'])
    line_items_df['Bad Debt Accruals'] = line_items_df['IndexKey'].map(comment_source_df.set_index('IndexKey')['Bad Debt Accruals'])


    # The below section makes lookup to the project description, Sales person and SO data in the df.
    proj_name_dict = dict(zip(qdl_tmp_df.iloc[:, 2], qdl_tmp_df.iloc[:, 4]))
    proj_name_to_main =  dict(zip(bd_so_temp_df.iloc[:, 0], bd_so_temp_df.iloc[:, 1]))

    line_items_df['Project Name'] = line_items_df['Reference']
    line_items_df['Project Name'] = line_items_df['Project Name'].map(proj_name_to_main)
    line_items_df['Project Name'] = line_items_df['Project Name'].map(proj_name_dict)
    line_items_df['Project_name_prev'] = line_items_df['IndexKey'].map(comment_source_df.set_index('IndexKey')['Project Name'])
    conditions = [
        (line_items_df['Project Name'].astype(str) == line_items_df['Project_name_prev'].astype(str)),
        (line_items_df['Project_name_prev'].astype(str) != line_items_df['Project Name'].astype(str)) & (line_items_df['Project_name_prev'].astype(str) != 'nan'),
        (line_items_df['Project_name_prev'].astype(str) != line_items_df['Project Name'].astype(str)) & (line_items_df['Project_name_prev'].astype(str) == 'nan')]
    values = [line_items_df['Project Name'].astype(str).replace('nan', np.nan, regex=True),
                line_items_df['Project_name_prev'].astype(str).replace('nan', np.nan, regex=True),
                line_items_df['Project Name'].astype(str).replace('nan', np.nan, regex=True)]
    line_items_df['Project Name'] = np.select(conditions, values)
    line_items_df.drop('Project_name_prev', inplace=True, axis=1)

    # Transfers Sales Persons names from the old AR file to the new (tmp).
    spers_name_dict = dict(zip(qdl_tmp_df.iloc[:, 2], qdl_tmp_df.iloc[:, 10]))
    spers_name_to_main = dict(zip(bd_so_temp_df.iloc[:, 0], bd_so_temp_df.iloc[:, 1]))

    line_items_df['Sales Person'] = line_items_df['Reference']
    line_items_df['Sales Person'] = line_items_df['Sales Person'].map(spers_name_to_main)
    line_items_df['Sales Person'] = line_items_df['Sales Person'].map(spers_name_dict)
    line_items_df['Sales_Person_prev'] = line_items_df['IndexKey'].map(comment_source_df.set_index('IndexKey')['Sales Person'])
    conditions = [
        (line_items_df['Sales Person'].astype(str) == line_items_df['Sales_Person_prev'].astype(str)),
        (line_items_df['Sales_Person_prev'].astype(str) != line_items_df['Sales Person'].astype(str)) & (line_items_df['Sales_Person_prev'].astype(str) != 'nan'),
        (line_items_df['Sales_Person_prev'].astype(str) != line_items_df['Sales Person'].astype(str)) & (line_items_df['Sales_Person_prev'].astype(str) == 'nan')]
    values = [line_items_df['Sales Person'].astype(str).replace('nan', np.nan, regex=True),
                line_items_df['Sales_Person_prev'].astype(str).replace('nan', np.nan, regex=True),
                line_items_df['Sales Person'].astype(str).replace('nan', np.nan, regex=True)]
    line_items_df['Sales Person'] = np.select(conditions, values)
    line_items_df.drop('Sales_Person_prev', inplace=True, axis=1)

    so_to_main = dict(zip(bd_so_temp_df.iloc[:, 0], bd_so_temp_df.iloc[:, 1]))
    line_items_df['Sales Order'] = line_items_df['Reference']
    line_items_df['Sales Order'] = line_items_df['Sales Order'].map(so_to_main)


    # calls context manager and writes to excel.
    with pd.ExcelWriter(ar_data_arg) as writer:
        # startrow=1 to free a space for the Total Amount cell on the top of the table.
        line_items_df.to_excel(writer, sheet_name='Customer Line Items', startrow=1, index=False, header=True)

        workbook = writer.book
        worksheet = writer.sheets['Customer Line Items']

        r = len(line_items_df.index)
        c = len(line_items_df.columns)

        worksheet.add_table(1, 0, r+1, c-1, {'name': 'ar_line_items',
                                            'header_row': True,
                                            'style': None,
                                            'columns': my_column_head_dscptns})  
        # Table header formatting.
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'bg_color': '#D9D9D9'})  # Header background color based on HEX

        for col_num, value in enumerate(line_items_df.columns.values):
            worksheet.write(1, col_num, value, header_format)

        celles_formats_boldi_melns = workbook.add_format({'bold': True, 'font_color': 'black'})

        # below, 'value': 99999999' questionable workaround to make conditional formatting 
        komentaru_lauka_adrese = f'AG3:AG{r+2}'
        komentaru_lauka_fons = workbook.add_format({'bg_color': '#EBF1DE'})
        worksheet.conditional_format(komentaru_lauka_adrese, {'type': 'cell', 'criteria': '!=',
                                                            'value': 999999999999,
                                                            'format': komentaru_lauka_fons})
        summu_lauka_adrese = f'N3:N{r+2}'
        summu_lauka_ciparu_formats = workbook.add_format({'num_format': '### ### ##0.00'})
        worksheet.conditional_format(summu_lauka_adrese, {'type': 'cell', 'criteria': '!=',
                                                        'value': 0,
                                                        'format': summu_lauka_ciparu_formats})
        summas_lauka_adrese = 'N1:N1'
        summas_lauka_cipara_formats = workbook.add_format({'num_format': '### ### ##0.00', 'bg_color': '#F1D6B1'})
        worksheet.conditional_format(summas_lauka_adrese, {'type': 'cell', 'criteria': '!=',
                                                        'value': 0,
                                                        'format': summas_lauka_cipara_formats})

        worksheet.write(0, 13, '=SUBTOTAL(9,ar_line_items[Amount in local currency])', celles_formats_boldi_melns)
        worksheet.write(0, 28, '=SUBTOTAL(9,ar_line_items[Bad Debt Accruals])', celles_formats_boldi_melns)
        worksheet.freeze_panes(2, 4)
        worksheet.set_column(9, 10, 9.86)  # formatting columns I and J, setting width to 9,86
        worksheet.set_column(12, 12, 9.86)
        worksheet.set_column(0, 0, 5.0)
        worksheet.set_column(1, 1, 5.0)
        worksheet.set_column(0, 13, 12.0)

        my_hidden_columns = [3, 16, 18, 21, 26, 27, 29, 30, 31]  # Hide util.columns.
        for hc in my_hidden_columns:
            worksheet.set_column(hc, hc, None, None, {'hidden': 1})

        new_format = workbook.add_format()
        new_format.set_align('right')
        worksheet.set_column('T:T', 20, new_format)

def prep_stack_sh_file(ar_data_tmp_arg, bubble_2d_tmp_arg, bubble_3d_tmp_arg, stack_data_tmp_arg):
        stack_df = pd.read_excel(ar_data_tmp_arg, sheet_name='Customer Line Items', skiprows=1)
        stack_df = stack_df[['Document Type', 'Document Number', 'Customer Name', 'Business Area1', 'Ageing', 'Amount in local currency', 'ic?']]
        stack_df = stack_df.loc[stack_df['Business Area1'] != 'V1']
        stack_df = stack_df.loc[stack_df['Business Area1'] != 'VE']
        stack_df = stack_df.loc[stack_df['Business Area1'] != 'VB']

        stack_df["DocCount"] = stack_df['Document Type']
        conditions = [stack_df['DocCount'] == 'RV', stack_df['DocCount'] == 'XE', stack_df['DocCount'] == 'XI']
        values = [1, 1, 1]
        stack_df['DocCount'] = np.select(conditions, values)
        stack_df["DocCount"] = stack_df.groupby('Customer Name')['DocCount'].transform('sum')
        stack_df["AgeCount"] = stack_df['Ageing']

        conditions = [stack_df['Document Type'] == 'RV', stack_df['Document Type'] == 'XE', stack_df['Document Type'] == 'XI']
        values = [stack_df['Ageing'], stack_df['Ageing'], stack_df['Ageing']]
        stack_df['AgeCount'] = np.select(conditions, values, default=np.NaN)  #  np.NaN bija svarīgs, lai nunique neskaitītu tukšumus kā unikālās vērtības

        conditions = [stack_df['Document Type'] == 'RV', stack_df['Document Type'] == 'XE', stack_df['Document Type'] == 'XI']
        values = [stack_df['Ageing'], stack_df['Ageing'], stack_df['Ageing']]
        stack_df['AgeCount'] = np.select(conditions, values, default=np.NaN)

        stack_df = stack_df.groupby(['Customer Name', 'ic?']).agg({'AgeCount':'nunique',"DocCount":'unique', 'Amount in local currency': 'sum'})
        stack_df = stack_df[np.in1d(stack_df.index.get_level_values(1), ['Ext'])]
        stack_df = stack_df.loc[stack_df['AgeCount'] > 1]
        stack_df = stack_df.explode('AgeCount')
        stack_df = stack_df.explode('DocCount')   # "eksplode" pievienots, lai konvertētu no listes uz parastiem vienumiem, pretējā gadījumā katra celle ir atsevišķa liste.....

        stack_df = stack_df[(stack_df['Amount in local currency'] > 0)]  # atlasa tikai tos klientu parādus, kuru konsolidētā summa lielāka par nulli

        max_top = stack_df['AgeCount'].count()  # 19.05.21 mainīgais topa vienumu max skaita noteikšanai pirms topa veidošanas
        # Definē metodi, lai mazāka skaita iespējamo topa vienumu gadījumā izvēlēties tieši to skaitu, vai arī max 10 vienumu limitu.
        if max_top < 10:
            max_top
        else: max_top = 10

        stack_df = stack_df.sort_values(by=['AgeCount'], ascending=False).head(max_top)  # izveido klientu parādu top 'x', atkarībā no tā cik x ir ar lielāku vecuma kategoriju nekā '2'
        # piešķir kopīgus mainīgos 2d un 3d čārtiem reizē
        customer = [x[0] for x in stack_df.index]  #  List comprehension
        y_ageing_cat = stack_df['AgeCount'].tolist()
        x_numb_of_inv = stack_df['DocCount'].tolist()
        z_amount = stack_df['Amount in local currency'].tolist()
        top_10_colors = ['#8d9194', '#0071b9', '#004987', '#58ab27', '#86c2e6', '#ffc627', '#c6d600', '#e51a92', '#0071b9', '#004987']
        # pievienoju no stackoverflow..
        font = {'family': 'Arial',
                'weight': 'normal',   # may set also 'bold' and etc...to_excel
                'size': 14}
        fig = plt.figure(figsize=(14, 8))
        z = 0
        while z < max_top:
            plt.scatter(x=x_numb_of_inv[z], y=y_ageing_cat[z], marker='o', s=z_amount[z], alpha=0.7, edgecolors='none', color=top_10_colors[z])
            z = z + 1
        lgnd = plt.legend(customer, loc='center left', bbox_to_anchor=(1.07, 0.5), fontsize=12)
        for handle in lgnd.legendHandles:
            handle.set_sizes([100])
        plt.grid()
        plt.title(f'TOP-{max_top} of the VA TRs with more than one Ageing category and accounting document')
        plt.ylabel("Number of Ageing Categories")
        plt.xlabel("Number of Accounting Documents")
        plt.savefig(bubble_2d_tmp_arg, bbox_inches='tight', dpi=100)

        font = {'family': 'Arial',
                'weight': 'normal',   # may set also 'bold' and etc...
                'size': 12}
        fig = plt.figure(figsize=(14,8))
        ax = fig.add_subplot(111, projection='3d')
        z = 0
        while z < max_top:
            ax.scatter(x_numb_of_inv[z], y_ageing_cat[z], z_amount[z]/1000, marker='o', s=np.pi*int(z_amount[z])/10, alpha=0.7, color=top_10_colors[z])
            z = z + 1
        lgnd = plt.legend(customer, loc='center left', bbox_to_anchor=(1.07, 0.5), fontsize=12) # bbox_to_anchor - nodrošina, ka čārta leģenda tiek iznesta ārpus rāmja, š.g. pa labi
        for handle in lgnd.legendHandles:
            handle.set_sizes([100])
            fig.tight_layout()
            fig.subplots_adjust(right=0.8)

        ax.set_xlabel('Number of Accounting Documents')
        ax.set_ylabel('Number of Ageing Categories')
        ax.set_zlabel('Debt Amount, T.Eur')
        plt.savefig(bubble_3d_tmp_arg)

        stack_df = stack_df.droplevel(['ic?'])  
        with xlsxwriter.Workbook(stack_data_tmp_arg) as workbook:
    
            worksheet = workbook.add_worksheet('Stacked')
            worksheet.insert_image('A1', bubble_2d_tmp_arg, {'x_scale': 1.0, 'y_scale': 1.0})
            worksheet.insert_image('A39', bubble_3d_tmp_arg, {'x_scale': 1.0, 'y_scale': 1.0})
            chart_xlRange = 'A1:AZ100'
            chart_xlRange_bkgrd = workbook.add_format({'bg_color': '#FFFFFF'})
            worksheet.conditional_format(chart_xlRange, {'type': 'cell', 'criteria': '!=',
                                                            'value': 999999999999,
                                                            'format': chart_xlRange_bkgrd})

            stack_tab_sum_adr = 'AC10:AC20'
            stack_tab_sum_form = workbook.add_format({'num_format': '# ##0;-# ##0'})
            worksheet.conditional_format(stack_tab_sum_adr, {'type': 'cell',
                                                                    'criteria': '!=',
                                                                    'value': 0,
                                                                    'format': stack_tab_sum_form})
            stack_tab_adr = 'Z10:AC20'
            stack_tab_form = workbook.add_format({'font_color': '#A6A6A6'})
            worksheet.conditional_format(stack_tab_adr, {'type': 'cell',
                                                                'criteria': '!=',
                                                                'value': 0,
                                                                'format': stack_tab_form})
            # write column names     
            start_row = 9
            start_col = column_index_from_string('Y')
            header_row = ['Customer', 'AgeCount', 'DocCount', 'Amount']
            for col_num, header_cell in enumerate(header_row):
                worksheet.write(start_row, start_col + col_num, header_cell)

            # write indices
            start_row = 10
            start_col = column_index_from_string('Y')
            for row_num, df_index in enumerate(stack_df.index):
                worksheet.write(start_row + row_num, start_col, df_index)    

            # Write the data
            start_row = 10
            start_col = column_index_from_string('Y')
            for row_num, row_data in enumerate(stack_df.values):
                for col_num, cell_data in enumerate(row_data):
                    worksheet.write(start_row + row_num , start_col + 1 + col_num, cell_data)

            # workaround for autofit in xlsxwriter
            for col_num, col_data in enumerate(stack_df.columns):
                worksheet.set_column(col_num + column_index_from_string('Y')-2, col_num + column_index_from_string('Y')-2, len(col_data) + 2)

def prep_change_sh_file(bridge_data_tmp_arg, wtf_chart_tmp_arg, ar_data_tmp_arg, old_ar_data_tmp_arg, now_date_arg, hist_date_arg, overd_days_arg):
    df1 = pd.read_excel(old_ar_data_tmp_arg, sheet_name='Customer Line Items', skiprows = 1)
    if overd_days_arg == '':
        pass
    else: df1 = df1.loc[df1['Overdue Days'] > int(overd_days_arg)]
    df1['Amount in local currency'] = df1['Amount in local currency']*-1
    df2 = pd.read_excel(ar_data_tmp_arg, sheet_name='Customer Line Items', skiprows = 1)
    if overd_days_arg == '':
        pass
    else: df2 = df2.loc[df2['Overdue Days'] > int(overd_days_arg)]
    # df3 = df1.append(df2) 
    df3 = pd.concat([df1, df2])
    df4 = pd.pivot_table(data=df3, index=['ic?', 'GL Account Description', 'Customer Name'], values='Amount in local currency', columns= ['Business Area1'], margins=True, aggfunc=np.sum, fill_value=0)
    df4 = df4.round(decimals=2).astype(object)
    df4 = df4[df4['All'] != 0]  # pivots ģenerē summāro kolonnu un pēc tās var nomest nulles ar šo koda rindu...
    df4 = df4.reset_index()
    df5 = pd.pivot_table(data=df3, index=['Business Area1'], values='Amount in local currency', margins=True, aggfunc=np.sum, fill_value=0)
    df5 = df5.round(decimals=2).astype(object)
    df5 = df5[df5['Amount in local currency'] != 0]  # pivots ģenerē summāro kolonnu un pēc tās var nomest nulles ar šo koda rindu...
    df5 = df5.reset_index()
    net_ch_chart_label = f"Net changes of the TR over {overd_days_arg} days old, {hist_date_arg.strftime('%d.%m.%Y.')} - {now_date_arg.strftime('%d.%m.%Y.')}"
    # piešķir mainīgo regexam, kas meklē un pārveido listē stringu relativerelative.....
    rel_str = (len(df5['Amount in local currency'])-1)*'relative'
    relative = re.findall('\welative+', rel_str)
    relative.append('total')
    x_vNames = df5['Business Area1'][:-1].tolist() # jāpārtaisa par listi lai var pievienot 'total' papildus aprakstu
    x_vNames.append('Total')
    y_values = df5['Amount in local currency'][:-1].tolist()
    y_values.append(0)
    text_values = df5['Amount in local currency'].astype(int).astype(str).tolist()

    fig = go.Figure(go.Waterfall(
    text=text_values,
    measure=relative,
    x=x_vNames,
    y=y_values,
    connector={"line": {"color": "rgb(63, 63, 63)"}},
    increasing={"marker": {"color": "rgb(229, 26, 146)"}},
    decreasing={"marker": {"color": "rgb(88, 171, 39)"}},
    totals={"marker": {"color": "rgb(141, 145, 148)"}},
        textposition="outside"))

    fig.update_layout(
            title=net_ch_chart_label,
            plot_bgcolor='rgba(0,0,0,0)')  # Šo pieliku pašrocīgi
    
    chart = Path(temp_dir, wtf_chart_tmp_arg, engine='kaleido') # kaleido version had to be upgraded, otherwise it hang the script (wo error handling)
    # fig.show() - this would produce image in a internet browser
    fig.write_image(file =chart)

    with PdExcel(bridge_data_tmp_arg) as writer:        
        df4.to_excel(writer, sheet_name='Net Changes', startrow=30, index=False, header=True)
        workbook = writer.book
        worksheet = writer.sheets['Net Changes']
        worksheet.insert_image('A1', chart)
        chart_xlRange = 'A1:AZ100'
        chart_xlRange_bkgrd = workbook.add_format({'bg_color': '#FFFFFF'})
        worksheet.conditional_format(chart_xlRange, {'type': 'cell', 'criteria': '!=',
                                                        'value': 999999999999,
                                                        'format': chart_xlRange_bkgrd})
        # Auto-adjust columns' width
        for column in df4:
            column_width = max(df4[column].astype(str).map(len).max(), len(column)+3) # +3 added to fit nicer
            col_idx = df4.columns.get_loc(column)
            writer.sheets['Net Changes'].set_column(col_idx, col_idx, column_width)

def compile_ar_fullrep(ar_fullrep_tmp_path_arg, ar_data_tmp_arg, old_ar_data_tmp_arg, stack_data_tmp_arg, bridge_data_tmp_arg, tr_qdl_sheet_arg, new_fullrep_tmp_arg):
    # Create a single instance of the xw.App object
    app = xw.App(visible=False)
    app.display_alerts = False

    # deleting old sheets, not touching Pivot.
    sheet_names = ['Customer Line Items', 'Customer Line Items (2)', 'Stacked', 'Net Changes', 'Qdl']
    tgt_wb = xw.Book(ar_fullrep_tmp_path_arg)

    # Open the target workbook and delete the old sheets
    tgt_wb = xw.Book(ar_fullrep_tmp_path_arg)
    for sheet_to_del in sheet_names:
        tgt_wb.sheets[sheet_to_del].delete()

    # Copy the sheets from the source workbooks to the target workbook
    src_wb = xw.Book(ar_data_tmp_arg)
    src_sheet = src_wb.sheets['Customer Line Items']
    src_sheet.api.Copy(Before=tgt_wb.sheets['Top10'].api)
    src_wb.save()
    src_wb.close()

    src_wb = xw.Book(old_ar_data_tmp_arg)
    src_sheet = src_wb.sheets['Customer Line Items']
    src_sheet.api.Copy(After=tgt_wb.sheets['C2C'].api)
    src_wb.save()
    src_wb.close()

    src_wb = xw.Book(stack_data_tmp_arg)
    src_sheet = src_wb.sheets['Stacked']
    src_sheet.api.Copy(After=tgt_wb.sheets['NEB VB'].api)
    src_wb.save()
    src_wb.close()

    src_wb = xw.Book(bridge_data_tmp_arg)
    src_sheet = src_wb.sheets['Net Changes']
    src_sheet.api.Copy(After=tgt_wb.sheets['Stacked'].api)
    src_wb.save()
    src_wb.close()

    src_wb = xw.Book(tr_qdl_sheet_arg)
    src_sheet = src_wb.sheets['Qdl']
    src_sheet.api.Copy(Before=tgt_wb.sheets['C2C'].api)
    src_wb.save()
    src_wb.close()

    # autofits the sheets like Comments and Cust Names, etc..
    for af in ['Customer Line Items', 'Customer Line Items (2)']:
        tgt_wb.sheets[af].select()
        autofit = ('N:N', 'W:W', 'X:X', 'Y:Y', 'Z:Z', 'AG:AG', 'T:T')
        for x in autofit:
            xw.Range(x).columns.autofit()

    tgt_wb.sheets['Customer Line Items (2)'].api.Visible = False
    tgt_wb.sheets['Qdl'].api.Visible = False

    # Refresh all pivot tables in the workbook
    for sheet in tgt_wb.sheets:
        for pt in sheet.api.PivotTables():
            pt.PivotCache().Refresh()

    tgt_wb.sheets['Customer Line Items'].select()  # navigate to the front sheet once everything is done
    xw.Range('N1:N1').select()   # navigate to the front sheet once everything is done

    # Save the changes to the target workbook and close the workbooks
    tgt_wb.save()
    tgt_wb.close()


    for process in psutil.process_iter():
       if process.name() == "EXCEL.EXE":
       # Terminate the process
        process.kill()
    # Quit the Excel application

    rename_ar_fullrep_tmp(ar_fullrep_tmp_path_arg, new_fullrep_tmp_arg)

def process_the_files(frontline_input_args, now_date_arg, hist_date_arg, overd_days_arg, send_email_arg):
    
    clear_temp(temp_dir)
    
    frontline_properties = return_list_of_frontl_props(frontline_input_args, my_entities)
    for frontline in frontline_properties:
        # Path variables initiated
        ar_fullrep_name = os.listdir(frontline.sync_dir)[0] # considering only single file on the dir
        ar_fullrep_sync_path = Path(frontline.sync_dir, ar_fullrep_name)
        ar_fullrep_tmp_path = Path(temp_dir, ar_fullrep_name)
        new_ar_fullrep_tmp = generate_file_name(temp_dir, f'{(frontline.abrev).upper()}({frontline.code}) AR Data', date_str, time_str, xlsx)
        fl_dir = ''.join([frontline.abrev,'_ar_temp'])
        ar_fullrep_arch_path = Path(arch_dir / fl_dir, ar_fullrep_name)

        # Move the AR file from sync to tmp dir
        move_file(ar_fullrep_sync_path, temp_dir) 
        copy_file(Path(temp_dir, ar_fullrep_name), ar_fullrep_arch_path)

        # add frontline abbreviations to tmp files during runtime:
 
        def fl(join_arg):
            return ' '.join([frontline.abrev, str(join_arg)]) #wrapping in str type just in case
        # temp files that are later used for the AR report compilation.
        
        prep_sap_qdl_file(frontline.code, qdl_sap_raw.parent, fl(qdl_sap_raw.name))
        prep_sap_cust_mast_data_file(frontline.code, cust_master_tmp.parent, fl(cust_master_tmp.name))
        prep_sap_cust_line_items_file(frontline.code, temp_dir, fl(cust_line_items_tmp.name), now_date_arg)
        prep_sap_cust_line_items_file(frontline.code, temp_dir, fl(old_cust_line_items_tmp.name), now_date_arg)
        prep_sap_bill_so_tab(Path(cust_line_items_tmp.parent, fl(cust_line_items_tmp.name)),\
                              bill_doc_list_tmp.parent, fl(bill_doc_list_tmp.name))    

        prep_qdl_sheet_file(Path(qdl_sap_raw.parent, fl(qdl_sap_raw.name)), Path(qdl_tmp.parent, fl(qdl_tmp.name)))

        prep_df_and_wrt_ar_file(Path(qdl_tmp.parent, fl(qdl_tmp.name)), now_date_arg, Path(cust_line_items_tmp.parent, fl(cust_line_items_tmp.name)),
                                 Path(cust_master_tmp.parent, fl(cust_master_tmp.name)), Path(bill_doc_list_tmp.parent, fl(bill_doc_list_tmp.name)),
                                  ar_fullrep_tmp_path, Path(ar_data_tmp.parent, fl(ar_data_tmp.name)))
        
        prep_df_and_wrt_ar_file(Path(qdl_tmp.parent, fl(qdl_tmp.name)), hist_date_arg, Path(old_cust_line_items_tmp.parent, fl(old_cust_line_items_tmp.name)),
                                 Path(cust_master_tmp.parent, fl(cust_master_tmp.name)), Path(bill_doc_list_tmp.parent, fl(bill_doc_list_tmp.name)),
                                  ar_fullrep_tmp_path, Path(old_ar_data_tmp.parent, fl(old_ar_data_tmp.name)))
        
        prep_stack_sh_file(Path(ar_data_tmp.parent,  fl(ar_data_tmp.name)), Path(bubble_2d_tmp.parent, fl(bubble_2d_tmp.name)),
                           Path(bubble_3d_tmp.parent, fl(bubble_3d_tmp.name)), Path(stack_data_tmp.parent, fl(stack_data_tmp.name)))
        
        prep_change_sh_file(Path(bridge_data_tmp.parent, fl(bridge_data_tmp.name)), Path(wtf_chart_tmp.parent, fl(wtf_chart_tmp.name)),
                            Path(ar_data_tmp.parent, fl(ar_data_tmp.name)), Path(old_ar_data_tmp.parent, fl(old_ar_data_tmp.name)),
                            now_date_arg, hist_date_arg, overd_days_arg)


        # AR file compilation
        compile_ar_fullrep(ar_fullrep_tmp_path, Path(ar_data_tmp.parent, fl(ar_data_tmp.name)),
                           Path(old_ar_data_tmp.parent, fl(old_ar_data_tmp.name)), Path(stack_data_tmp.parent, fl(stack_data_tmp.name)), Path(bridge_data_tmp.parent, fl(bridge_data_tmp.name)),
                             Path(qdl_tmp.parent, fl(qdl_tmp.name)), new_ar_fullrep_tmp)

        # move the AR file back from tmp to the sync dir where it could be accessed by the users (..once synced).
        move_file(new_ar_fullrep_tmp, frontline.sync_dir)
        
        # Notification in the terminal
        print(f'\n{frontline.abrev} TR file created on: {frontline.sync_dir}.. !!')

        # Send Email or ommit. Depending on user's input. method called at the end, just in case if script stops unexpectedly.
        send_email(frontline.email_main, frontline.email_cc, frontline.abrev, frontline.shrp_path, send_email_arg)    