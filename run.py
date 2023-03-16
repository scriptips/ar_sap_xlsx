import os
import re as re
import sys
import time
 
from const import err_log_file_path, my_entities, start_user_alert 
from sap import process_the_files
from utils import (prompt_continue, return_list_of_frontl_props,
                   select_frontlines, send_email, set_comparison_date,
                   set_now_date, set_overdue_days, setup_logging)

os.chdir(os.path.dirname(os.path.abspath(__file__)))

def main():
    # Terminal alert  before the report runs.
    print(start_user_alert)

    func_defaults = [x.__defaults__ for x in [select_frontlines, set_now_date, set_comparison_date, set_overdue_days, send_email]]
    (select_frontlines_fdefault,), (set_now_date_fdefault,), (set_comparison_date_fdefault,), (set_overdue_days_fdefault,), (send_email_fdefault,), = func_defaults

    apply_defaults =\
        input(f'Press Enter to apply defaults, "Space + Enter" to customize or "e + Enter" just to send the notification emails..\n\
                  Entity: {select_frontlines_fdefault},\n\
                    Date: {set_now_date_fdefault},\n\
               Hist.Date: {set_comparison_date_fdefault},\n\
              Overd.days: {set_overdue_days_fdefault},\n\
                  Emails: {send_email_fdefault}\n')

    if apply_defaults == '':
        # user input variables initiated.
        frontline_inputs = select_frontlines_fdefault
        now_date = set_now_date_fdefault 
        hist_date = set_comparison_date_fdefault 
        overd_days = set_overdue_days_fdefault
        mail_notif_input = send_email_fdefault
        
        process_the_files(select_frontlines_fdefault, set_now_date_fdefault, set_comparison_date_fdefault, set_overdue_days_fdefault, send_email_fdefault)

        print('\nThe Script completed with Default report parameters !!')

    elif apply_defaults == 'e':
        frontlines = return_list_of_frontl_props(select_frontlines(input('Type-in frontline abrevs: ... ')), my_entities)
        for fl in frontlines:
            send_email(fl.email_main, fl.email_cc, fl.abrev, fl.shrp_path, 'yes')
        
    else:
        # user input variables initiated.Notice comma in frontline_inputs, =..., that's for unpacking.
        frontline_inputs =\
              select_frontlines(input(\
             f'\nEnter company codes (e.g. "bal", "koe", "kla" or "kli", or default "bal", press Enter to continue.. '))
        print()
        now_date =\
              set_now_date(input(\
             'Key-in current date "ddmmyy", or press Enter....'))
        print()
        hist_date =\
              set_comparison_date(\
             input('Key-in historic date "ddmmyy", or press Enter....'))
        print()
        overd_days = set_overdue_days()
        mail_notif_input = input('\nSend email notificatoins [yes/no] ?... ') 

        if mail_notif_input == 'yes':
            print('yes\n')
        else: mail_notif_input ='no', print('no\n')

        process_the_files(frontline_inputs, now_date, hist_date, overd_days, mail_notif_input)

    input('\n\nPress "Enter" to exit this screen..') # stops terminal to let user hit any key before it gets auto-closed


if __name__ == '__main__':
        if prompt_continue():
            print("Continuing with the script...")
            logger = setup_logging(err_log_file_path)
            try:
                main()
            except Exception as e:
                logger.exception(e)
                input('\n\nPress "Enter" to exit this screen..')
                
        else:
            print("Quitting the script...")
            time.sleep(3)
            sys.exit()