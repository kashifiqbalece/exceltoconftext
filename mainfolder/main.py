import pandas as pd
from jinja2 import Template
import get_config_sheet_function,run_config_create_perconfig

###Prompt for User Input for File Name######
excel_sheet=input("Enter Excel File with complete path or with name if it exists in same directory of this tool:\n")

######## Condition check for File Extension
if excel_sheet.split('.')[-1]=='xlsx' or  excel_sheet.split('.')[-1]=='xls':
    #output file name for true case
    output_filename=input("Enter Output File name (either Compelete Path or JustName with Extention):\n")
    sheet_names=get_config_sheet_function.get_sheet_names(excel_sheet)
    print(f'Following Sheets Will be Used To generate Configuration in {output_filename}\n{sheet_names}')
    #permission check to continue
    permission_continue=(input("\n ******Would You Like To Continue to Run This Tool To Generate The Config ******** (Y/N): "))
    if permission_continue.lower()=="y":
       #run function for sheets
        print(f"Running Tool For {len(sheet_names)} Sheets one by one ")
        for sheet in sheet_names:
            header_value_global = 0
            #getting Config Types
            configtypelist=get_config_sheet_function.get_config_type_insheet(excel_sheet,sheet)
            #counter for while loop
            configtype=0
            # running function for each sheet and each config type
            # Condition for Optional Paramter display and contnue accordinlgy
            if (input("Do you want Optional Parameter (y/n)").lower())=='y':
                print("Optional parameter part here")
            else:
                while configtype<=len(configtypelist)-1:
                    print(f"Doing Stuff for sheet {sheet} and Config Type {configtypelist[configtype]}")
                    x = "NN_CONFIG_" + configtypelist[configtype]
                    #print(x)
                    header_value_global=getattr(run_config_create_perconfig,x)(excel_sheet,sheet,configtypelist[configtype],output_filename,header_value_global)
                    header_value_global=header_value_global+1
                    configtype = configtype + 1
                print(f"Print Config Ready For {sheet} and config Type - {configtypelist}")
            print(f"***************Your Output for Config File is saved on file : {output_filename}**************")
            print("***Program Ends Here***")
    else:
        print("Exit Bye")
else:
    print("sorry Wrong File type Only XLSX and XLS are supported")