import email_pupculture as e_pc 
import get_reg as gr 
import os 
this_dir = os.path.dirname(os.path.abspath(__file__))

# CLEAR LOGS
reg_log = os.path.join(this_dir, 'inputlog.txt') 
mail_log = os.path.join(this_dir, 'maillog.txt') 

open(reg_log, 'w').close() 
open(mail_log, 'w').close() 

print("\nLogs cleared!")
