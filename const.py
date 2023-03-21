import json
import os
from datetime import datetime
from pathlib import Path

CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.json')
with open(CONFIG_FILE, 'r') as f:
    config = json.load(f)

# masked
arch_dir = Path(config['arch_dir'])
temp_dir = config['temp_dir']
err_log_file_path = config['err_log_file_path']
ent_abrevs = config['ent_abrevs'] # entity literal abbreviations;
ent_codes = config['ent_codes'] # entity numerical codes;
sync_folders = config['sync_folders'] # list of str , paths
test_sync_folders = config['test_sync_folders'] # list of str , paths
main_receivers = config['main_receivers'] # list of str, email addresses
bal_mt = config['bal_mt'] # list of str, email addresses
shrp_folders = config['shrp_folders'] # list of str, paths
test_rec = config['test_rec'] # str, email address
ctry_code_names = config['ctry_code_names'] # dict, country codes and names
bal_fls = config['bal_fls'] # list of str
ba_descrps = config['ba_descrps'] # dict, Business Area descriptions

my_entities = [[ent_abrevs[0], ent_codes[0], sync_folders[0], main_receivers[0], bal_mt[0], shrp_folders[0]],
                [ent_abrevs[1], ent_codes[1], sync_folders[1], main_receivers[1], bal_mt[0], shrp_folders[1]],
                [ent_abrevs[2], ent_codes[2], sync_folders[2], main_receivers[2], bal_mt[0], shrp_folders[2]]]

# for testing purposes
my_entities_ = [[ent_abrevs[0], ent_codes[0], test_sync_folders[0], test_rec, test_rec, shrp_folders[0]],
                [ent_abrevs[1], ent_codes[1], test_sync_folders[1], test_rec, test_rec, shrp_folders[1]],
                [ent_abrevs[2], ent_codes[2], test_sync_folders[2], test_rec, test_rec, shrp_folders[2]]]

start_user_alert = "\nFor the script to run properly, BEFORE CONTINUING, ensure the following: ...\n\n\
                    -  SAP is set to the starting screen, and only a single SAP session is-on\n\
                    -  Close all MS Excel workbooks;\n\
                    -  Check that country target folders contain ONLY A SINGLE version of the AR file.\n\n\n"

cstr_colmns = ['Customer Name', 'Account']

gl_colmns = ['GL Account Description', 'G/L Account']

gl_descrps = {
    130100: 'AR - Domestic', 
    130000: 'AR - Foreign', 
    226000: 'PIR - External',  
    130300: 'AR - Intra_Corporate',  
    }

doc_type = ['Type', 'Document Type']

doc_type_descrps = {
    'AA': 'Asset posting',
    'AB': 'All reversal docs',
    'AC': 'Reversal  RN(net iv)',
    'AF': 'Dep. postings',
    'AN': 'Net asset posting',
    'BG': 'Bank guarantee -note',
    'BV': 'Intracomp.inv., SD',
    'BX': 'Intracomp.cr.note SD',
    'C1': 'Other Cust Documents',
    'C2': 'Man.Cust.Inv',
    'C3': 'Man.Cust Cr Notes',
    'C4': 'Auxil. Inv. from FI',
    'CH': 'Contract Settlement',
    'CL': 'Credit losses',
    'D1': 'AR pmyt (ref.mand.)',
    'D2': 'Doors invoices',
    'D4': 'Auxil. Inv. from SD',
    'DA': 'AR General/Clearing',
    'DB': 'Incoming bank tfrs',
    'DC': 'Dir.deb.,Bill of exc',
    'DD': 'Incoming bank tfs 2',
    'DE': 'RR postings',
    'DF': 'Bill of exch.(elect)',
    'DG': 'Customer credit memo',
    'DH': 'Bill of exch.(paper)',
    'DI': 'Customer interest',
    'DJ': 'Incoming cheques',
    'DK': 'Cust./vend. netting',
    'DL': 'Payt. to customer',
    'DM': 'Payt.on acc.,no inv',
    'DN': 'Check without credit',
    'DO': 'Incoming cash',
    'DP': 'Unpaid bill of exchg',
    'DQ': 'B/E (paper) correct.',
    'DR': 'Pmt.on ac-inv not id',
    'DS': 'Inc. pmt. reversal',
    'DT': 'AR payments',
    'DU': 'Pmt. on acc-reversal',
    'DV': 'Unpaid reversal',
    'DW': 'Pmt/ac-rev.In.not id',
    'DX': 'Unpaid B/E 2',
    'DY': 'Unpaid B/E 3',
    'DZ': 'Down pmnt request/AR',
    'E1': 'Other Vend. Docs EU',
    'E2': 'Vend Inv. EU',
    'E3': 'Vend. Cr. Notes EU',
    'EA': 'AP Rev.EU inv.ot.ctr',
    'EB': 'AP EU inv.fr ot.ctry',
    'EC': 'AP EU cre.fr ot.ctry',
    'ER': 'Vend. Intracomp Inv',
    'EU': 'EURO conversion post',
    'EX': 'External Number',
    'F1': 'Other Vend. Docs ITA',
    'F2': 'Vend. Inv. ITA',
    'F3': 'Vend. Cr. Notes ITA',
    'F4': 'AP Self-invoice',
    'F5': 'AP Rev. self-invoice',
    'F6': 'Vend. Inv. ITA',
    'F7': 'Vend. Cr. Notes ITA',
    'F8': 'AP Self-inv.Material',
    'F9': 'AP Rev. self-inv.Mat',
    'FA': 'AP Rev.inv.fr o ctry',
    'FB': 'AP inv. fr oth.cntry',
    'FC': 'AP cre. fr oth.cntry',
    'FL': 'Mth. End Post',
    'H2': 'Hydro invoices',
    'IW': 'Incoming wire trsf.',
    'KA': 'Vendor clearing/gen.',
    'KB': 'Consignment Clearing',
    'KC': '',
    'KE': 'AP invoices',
    'KG': 'Vendor Credit Memo',
    'KN': 'Net Vendors',
    'KP': 'Account Maintenance',
    'KR': 'Vendor pmyt., B/E',
    'KT': 'Vend. paym., transf.',
    'KZ': 'Vend. paym., cheque',
    'ML': 'ML Settlement',
    'PC': '',
    'PP': 'Preauthorized Pymts',
    'PR': 'Price Change',
    'RA': 'Sub.Cred.Memo Stlmt',
    'RB': 'Reserve for Bad Debt',
    'RE': 'Invoice receipt/cred',
    'RF': 'R/cred TAXI KIE&TLI',
    'RN': 'Invoice Receipt(net)',
    'RS': 'BilldocSDplant abrod',
    'RT': 'CustCr Note Plt Abrd',
    'RV': 'Bill.doc.transfer/SD',
    'RX': 'Cust.cred.note/SD,SM',
    'SA': 'Closing Posting',
    'SB': 'G/L Bank Posting',
    'SK': 'G/L petty cash post.',
    'SP': 'Split Payment Italy',
    'SU': 'Subseqnt debit doc.',
    'SZ': 'Closing postings',
    'TB': 'TR Reversal docs.',
    'TI': 'Treasury interface',
    'TM': 'TR interface-balance',
    'TR': 'TR document',
    'UE': 'Data Transfer',
    'VA': 'Credit Card Payment',
    'VB': 'Misc. Cash Receipt',
    'VC': 'Gelgo',
    'WA': 'Goods issue',
    'WE': 'Goods receipt',
    'WI': 'Inventory Document',
    'WL': 'Goods Issue/Delivery',
    'WN': 'Net Goods Receipt',
    'XA': 'G/L items-conversion',
    'XB': 'Vendor inv. from Con',
    'XC': 'Vendor Cr.note-Conv.',
    'XD': 'Pays.to vendors-conv',
    'XE': 'Cust.inv.from conver',
    'XF': 'Cust.CR.note-convers',
    'XG': 'Pays.from cust.-conv',
    'XH': 'Adjustment postings',
    'XI': 'Cust.Prog.inv. only!',
    'XJ': 'Cust.Prog.Paym.only!',
    'XL': 'Main.billing arrears',
    'XO': 'MAC spares cr. notes',
    'XP': 'Closing bal./mer',
    'XR': 'Opening bal./merg',
    'XS': 'B/E, conversion',
    'XT': 'MAC spares billing',
    'XX': 'Deferred tax transfe',
    'Y1': 'Vendor direct debit',
    'Y2': 'Customer dir. debit',
    'Y3': 'G/L general postings',
    'Y4': 'Intercorp.invoice/AR',
    'Y5': 'Intercorp.pr.inv./AR',
    'Y6': 'Intercorp. credit/AR',
    'Y7': 'Intercorp.invoice/AP',
    'Y8': 'Intercorp. credit/AP',
    'Y9': 'Cross company postin',
    'YC': 'G/L clearings',
    'YI': 'Intracomp.post./IDoc',
    'YP': 'Revenue periodising',
    'YS': 'Settlement postings',
    'YW': 'WIP postings/CO',
    'Z1': 'Cl./dt.,Prg.inv.rec.',
    'Z2': 'AR V5 Invoice',
    'Z3': 'Incoming cash receip',
    'Z4': 'Outgoing cash paymen',
    'Z5': 'Incoming cash st.pet',
    'Z6': 'Outgoing cash st.pet',
    'Z7': 'Travel exp.Interface',
    'Z8': 'Pmt.on ac-inv not id',
    'Z9': 'Payroll postings int',
    'ZA': 'Cust.inv.-interface?',
    'ZB': 'Cus.cr.nt.-interface',
    'ZC': 'Ven.inv-Rev.Charge',
    'ZD': 'VenCR.Nte-R.Char.',
    'ZE': 'Ven.(ITA)in-int.face',
    'ZF': 'Ven(ITA)CR.nt-i/face',
    'ZG': 'G/L posting-interfac',
    'ZH': 'Travel.exps-interfac',
    'ZI': 'Bank Guarantees',
    'ZJ': 'AP-eInvoice rec./cr.',
    'ZK': 'Intracomp.Inv.AP',
    'ZL': 'Tender Time',
    'ZP': 'Payment posting',
    'ZR': 'Bank Reconciliation',
    'ZS': 'Payment by Check',
    'ZT': 'TR payments',
    'ZV': 'General Payments',
    'ZW': 'TR clearing/pmnts',
    'ZZ': 'Down pmnt request/AP'
    }

ctry_colmns = ['Country', 'Company Name']

ba_colmns = ['Business Area1', 'BA - Description']

all_descr_maps = [(doc_type, doc_type_descrps), (gl_colmns, gl_descrps), (ctry_colmns, ctry_code_names), (ba_colmns, ba_descrps)]

now_is = datetime.now()

date_str = "-".join([str(now_is.strftime("%d")), str(now_is.strftime("%m")), str(now_is.strftime("%Y"))])

time_str = datetime.now().strftime("%H.%M.%S")

def generate_file_name(temp_dir, prefix, date_str, time_str, extension):
    return Path(temp_dir, ' '.join([prefix, date_str, ''.join(["(", time_str,")", extension])]))
xlsx = '.xlsx'
png = '.png'
cust_master_str = 'cust_master_tmp'
bill_doc_list_str = 'bill_doc_list_tmp'
cust_line_items_str = 'cust_line_items_tmp'
old_cust_line_items_str = 'old_cust_line_items_tmp'
ar_data_str = 'ar_data_tmp'
old_ar_data_str = 'old_ar_data_tmp'
bubble_2d_str = 'bubble_2d_tmp'
bubble_3d_str = 'bubble_3d_tmp'
bubble_3d_str = 'bubble_3d_tmp'
stacked_str = 'stacked_tmp'
wtf_chart_str = 'wtf_chart_tmp'
bridge_data_str = 'bridge_data_tmp'
qdl_sap_raw_str = 'qdl_sap_raw'
qdl_tmp_str = 'qdl_tmp'

cust_master_tmp = generate_file_name(temp_dir, cust_master_str, date_str, time_str, xlsx)
bill_doc_list_tmp = generate_file_name(temp_dir, bill_doc_list_str, date_str, time_str, xlsx)

cust_line_items_tmp = generate_file_name(temp_dir, cust_line_items_str, date_str, time_str, xlsx)
old_cust_line_items_tmp = generate_file_name(temp_dir, old_cust_line_items_str , date_str, time_str, xlsx)

ar_data_tmp = generate_file_name(temp_dir, ar_data_str, date_str, time_str, xlsx)
old_ar_data_tmp = generate_file_name(temp_dir, old_ar_data_str, date_str, time_str, xlsx)

bubble_2d_tmp = generate_file_name(temp_dir, bubble_2d_str, date_str, time_str, png)
bubble_3d_tmp = generate_file_name(temp_dir, bubble_3d_str, date_str, time_str, png)
stack_data_tmp = generate_file_name(temp_dir, stacked_str, date_str, time_str, xlsx)

wtf_chart_tmp = generate_file_name(temp_dir, wtf_chart_str, date_str, time_str, png)
bridge_data_tmp = generate_file_name(temp_dir, bridge_data_str, date_str, time_str, xlsx)

qdl_sap_raw = generate_file_name(temp_dir, qdl_sap_raw_str, date_str, time_str, xlsx)
qdl_tmp = generate_file_name(temp_dir, qdl_tmp_str, date_str, time_str, xlsx)

file_ext_tplist = [(qdl_sap_raw_str, xlsx), (qdl_tmp_str, xlsx), (cust_master_str, xlsx), (cust_line_items_str, xlsx), (old_cust_line_items_str, xlsx), (bill_doc_list_str, xlsx),
(ar_data_str, xlsx), (old_ar_data_str,  xlsx), (stacked_str, xlsx), (bubble_2d_str, png), (bubble_3d_str, png)]

header_list = [
    'Document Type', 'Company Code', 'Posting Key', 'Business Area',
    'G/L Account', 'Account', 'Document Number', 'Line item', 'Reference',
    'Document Date', 'Posting Date', 'Terms of Payment', 'Net due date',
    'Amount in local currency', 'Document Header Text', 'Text', 'User Name',
    'Business Area1', 'Business Area2', 'Ageing', 'Overdue Days', 'Type', 'Customer Name',
    'Project Name', 'Sales Person', 'Sales Order', 'GL Account Description',
    'BA - Description', 'Bad Debt Accruals', 'Country', 'IndexKey', 'Concat',
    'COMMENTS(to be added in the Sheet "Customer Line Items", column "AG"…)', 'ic?']

my_column_head_dscptns = [
    {'header': x} for x in [
    'Document Type',
    'Company Code',
    'Posting Key',
    'Business Area',
    'G/L Account',
    'Account',
    'Document Number',
    'Line item',
    'Reference',
    'Document Date',
    'Posting Date',
    'Terms of Payment',
    'Net due date',
    'Amount in local currency',
    'Document Header Text',
    'Text',
    'User Name',
    'Business Area1',
    'Business Area2',
    'Ageing',
    'Overdue Days',
    'Type',
    'Customer Name',
    'Project Name',
    'Sales Person',
    'Sales Order',
    'GL Account Description',
    'BA - Description',
    'Bad Debt Accruals',
    'Country',
    'IndexKey',
    'Concat',
    'COMMENTS(to be added in the Sheet "Customer Line Items", column "AG"…)',
    'ic?']
    ]

my_qdl_column_head_dscptns = [
    {'header': x} for x in [
    "WBS Element",
    "Network Number",
    "SO Doc N°",
    "Material Number",
    "Project name",
    "SO item text",
    "Division",
    "Sales Off",
    "Sales Group",
    "Profit Center",
    "Sales Employee name",
    "Customer",
    "KONE Project id",
    "Business type",
    "Category required for Rush",
    "Technical platform",
    "Equipment in service",
    "Supervisor",
    "Supervisor name",
    "Elevator type code",
    "Creation date",
    "Order booked date",
    "Sched. 0B",
    "Act. 0B",
    "Sched. 3C",
    "Act. 3C",
    "Sched. 5",
    "Act. 5",
    "Pl. Revenue (0)",
    "Pl. Corp Mat Costs (0)",
    "Pl Ext Mat Costs (0)",
    "Pl Oth costs (0)",
    "Pl Lab Costs (0)",
    "Pl Tot Costs (0)",
    "CM2 (0)",
    "CM2 % (0)",
    "Plnd Hrs for INST",
    "Actual Revenue",
    "Actual Corp Mat Costs",
    "Actual Ext Mat Costs",
    "Actual Oth costs",
    "Actual Lab Costs",
    "Actual Tot Costs (Act.)",
    "Actual CM2",
    "Actual CM2 %",
    "Used Hrs",
    "Object status",
    "Work center",
    "Equip N°",
    "Revenue (-2)",
    "Corp Mat Costs (-2)",
    "Extern Mat Costs (-2)",
    "Oth costs (-2)",
    "Lab Costs (-2)",
    "Tot Costs (-2)",
    "CM2 (-2)",
    "CM2 % (-2)",
    "Your Reference",
    "WIP",
    "Late Cost Provision",
    "Loss Contract Provision",
    "Technical platform2",
    "Subcontracted",
    "Building state",
    "Number of units",
    "Labour rate",
    "Technically complete date",
    "C2C",
    "Actual Status",
    "Actual Status Date",
    "Projects (entire) not yet fully Billed?"]
    ]