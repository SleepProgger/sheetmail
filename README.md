# sheetmail
Merge mailer reading from spreadsheet files

Install
-------------
Python 2.7 is required. Download f.e. [here](https://www.python.org/download/releases/2.7/)  

Python module openpyxl.  
The easiest way is to use pip like:

>pip install openpyxl

For linux user that should be it.
Windows user might need to follow this steps (Only tested on Win7, but it should be similair on Win>7):  
1. Click the start menu button and search for cmd  
2. Right click it and run as Administrator  
3. Paste the above mentioned command. If this works, awesome, otherwise continue with..  
4. Find the place where you installed python (in most cases *c:\python27* or something like *C:\program files\python*)  
5. Paste the following into the cmd window (replace "*c:\python27*" with your python path)  

> c:\python27\Scripts\pip.exe install openpyxl


ATTENTION 
-------------
This script updates the original spreadsheet file, so i strongly suggest doing a backup before running.
Also it will replace function like *=concat("A", "B")* with the result. I.e the cell would contain "AB".  

Due to a bug in openpyxl which could potentially corrupt your excel files, i strongly advise in always suplying the *--cleancomments* parameter. This will remove all comments from the spreadsheet file.

To test if everything works as expected there is the *--test* parameter which tests the mail server connectivity and also verifies the spreadsheet file and tests for the mentioned comment bug.

Also the *--nosend* parameter could be useful to test if the correct data would be send.

Usage
-------------
This script send mails with data suplied by spreadsheet files (Only tested with xlsx files so far !).  
The spreadsheet file need to suply:  
1. The recipient  
2. The message  
3. The subject (not required when the *--staticsubject* parameter is used)  
4. A column containing the information if the data of this row should be send. (0 to send, eveything else to ignore).  This column will get updated by the script when the mail was sucesfuly send (value 1) or the recipient was invalid (value 2)  
The connection details for the mail server are stored in a json file.

Configuration file
-------------
Example for gmail account:

    {  
     "mail_user": [  
      {  
       "host": "smtp.gmail.com",   
       "port": 465, 
       "username": "your_account@gmail.com"  
       "password": "super_secret",  
       "sender_addr": "your_account@gmail.com",  
       "timeframe": 86400,  
       "allowed_requests": 2000,  
       "use_fixed_delay": true,  
       "update_config": false,  
       "user_ssl": true,  
       "timeframe_end": 1445647266.221,  
       "remaining_requests": 0,  
      }  
     ]  
    }  

*TODO: Write something about the quota*


Example
-------------
Send mails from the file test.xlsx where the recipient is in the first column, the message in the second and the status field in the third column. The subject will be the same "Hello from sheetmail" for every message.
The used sheet is the first one (set with *--sheetindex *). *--cleancomments* is used to prevent the mentioned bug in the ATTENTION section.

    python sheetmail.py --config config.json -m 0 -b 1 -o 2 --cleancomments --staticsubject "Hello from sheetmail" xour_xlsxfile.xlsx

Parameter
-------------

    D:\mailer>"c:\Program Files (x86)\python\python.exe" sheetmail.py --help
    usage: exel_test.py [-h] --config CONFIG [--loglvl {DEBUG,INFO,WARN,ERROR}]
                        [--logfile LOGFILE] [--colmail COLMAIL]
                        [--colsubject COLSUBJECT] [--colbody COLBODY]
                        [--colsend COLSEND] [--staticsubject STATICSUBJECT]
                        [--sheetindex SHEETINDEX] [--cleancomments] [--test]
                        [--nosend]
                        excel_file
    
    Sends emails with data supplied by excel files.
    
    positional arguments:
      excel_file            The excel file to get data from.
    
    optional arguments:
      -h, --help            show this help message and exit  
      --config CONFIG, -c CONFIG  
                            Choose the configuration file.  
      --loglvl {DEBUG,INFO,WARN,ERROR}, -l {DEBUG,INFO,WARN,ERROR}  
                            Set the log level.  
     --logfile LOGFILE, -f LOGFILE  
                            Also write log to file.  
      --colmail COLMAIL, -m COLMAIL  
                            The column containing the email address.  
      --colsubject COLSUBJECT, -s COLSUBJECT  
                            The column containing the email subjects.  
      --colbody COLBODY, -b COLBODY  
                            The column containing the email message.  
      --colsend COLSEND, -o COLSEND  
                            The column used to mark if the mail was send  
      --staticsubject STATICSUBJECT, -x STATICSUBJECT  
                            Can be used to use a static subject.  
      --sheetindex SHEETINDEX, -i SHEETINDEX  
                            The sheet to use.  
      --cleancomments       Remove comments from file. Openpyxl has a bug leading  
                            to corrupt files otherwise.  
      --test                Only test all mail accounts and the spreadsheet file
                            and exit.  
      --nosend              Do not send the mails. Used for testing.  

