import os
import pandas as pd
import openpyxl as op
import csv
import re
import csv
from tkinter import *
from tkinter import filedialog

tds_file_path=""


def clear_file_paths():
    global tds_file_path
    
    books_path_label = Label(root,text = ' '*len(tds_file_path)*3)
    books_path_label.config(font=("Arial", 12))
    books_path_label.place(x=240,y=50)



def get_books_file_path():
    global tds_file_path
    tds_file_path = filedialog.askopenfilename(initialdir = "/", title = "Select a File", filetypes = (("XLSX File", "*.xlsx*"),("CSV File", "*.csv*"),("Excel", "*.xls*"),("All files", "*.*"))) 

    books_path_label = Label(root,text = tds_file_path)
    books_path_label.config(font=("Arial", 12))
    books_path_label.place(x=240,y=50)

    if os.path.isfile(tds_file_path):
        message_desc=Label(root,text="Books File exists in the given path")
        message_desc.config(font=("Arial", 12),foreground="green")
        message_desc.place(x=20,y=250)
        message_desc.after(1000,message_desc.destroy)
    else:
        message_desc=Label(root,text="File does not exist in the given path")
        message_desc.config(font=("Arial", 12),foreground="red")
        message_desc.place(x=20,y=250)
        message_desc.after(1000,message_desc.destroy)

def get_tds_consolidation():
    global tds_file_path 
    cwd=os.getcwd()
    #tds_file_path=os.path.join(cwd,"data","tds data.xlsx")
    tds_wb=op.load_workbook(tds_file_path)
    tds_df=pd.DataFrame(tds_wb.active.values)
    tds_df.columns=tds_df.iloc[4]
    tds_df=tds_df[5:-3]
    
    output_file_path=os.path.join(cwd,"temp","tds_consolidation_output.csv")
    
    if os.path.exists(output_file_path):
        os.remove(output_file_path)

    pan_master=dict()
    for index,row in tds_df.iterrows():
        pan_master[row["PAN"].strip()]=row["Party Account"]

    
    tds_category_df = {k: v for k, v in tds_df.groupby("TDS Transaction Nature")}
    tds_salary_category=""
    tds_salary_df=pd.DataFrame()
    for key,value in tds_category_df.items():
        if "salar" in key.lower():
            tds_salary_df=pd.DataFrame(value)
            tds_salary_category=key
            
    del tds_category_df[tds_salary_category]
    
    category_final_master=dict()
    category_final_total=dict()
    category_final_tds=dict()

    for key,value in tds_category_df.items():
        category_final_master[key]=None
        category_final_total[key]=None
        category_final_tds[key]=None



    for key,value in tds_category_df.items():
        tds_final_df = pd.DataFrame(columns=["Party Account","TDS Transaction Nature","PAN","Assessable Value","TDS Rate","TDS Amount"])
        tds_sum_df={k: v for k, v in value.groupby("PAN")}
        temp=key.split("-")
        tds_category=temp[0]
        tds_percentage=re.findall("\d+\.*\d*\%", temp[len(temp)-1])
        
        if len(tds_percentage)>0:
            tds_percentage=float(tds_percentage[0][:-1])
        else:
            tds_percentage=0.00
    
        for key1,value1 in tds_sum_df.items():
            assessable_value=sum(value1["Assessable Value"])
            tds_total=(assessable_value*tds_percentage)/100   
            row_dict={"Party Account":pan_master[key1.strip()],"TDS Transaction Nature":key,"PAN":key1,"Assessable Value":float(assessable_value),
            "TDS Rate":float(tds_percentage),"TDS Amount":float(tds_total)}
            tds_final_df=tds_final_df.append(row_dict,ignore_index=True)
            
        category_final_master[key]=tds_final_df
        category_final_total[key]=float(sum(tds_final_df["Assessable Value"]))
        category_final_tds[key]=float(sum(tds_final_df["TDS Amount"]))

    for key,value in category_final_master.items():
        temp=key.split("-")
        tds_category=temp[0]+" - "+temp[1]
        
        row_dict={"Party Account":"Total for "+tds_category,"TDS Transaction Nature":"","PAN":"","Assessable Value":category_final_total[key],
            "TDS Rate":"","TDS Amount":category_final_tds[key]}
        category_final_master[key]=category_final_master[key].append(row_dict,ignore_index=True)
        row_dict={"Party Account":" ","TDS Transaction Nature":"","PAN":"","Assessable Value":"","TDS Rate":"","TDS Amount":""}
        category_final_master[key]=category_final_master[key].append(row_dict,ignore_index=True)

    tds_final_df = pd.DataFrame(columns=["Party Account","TDS Transaction Nature","PAN","Assessable Value","TDS Rate","TDS Amount"])
    tds_final_df.to_csv(output_file_path, mode='a', index = False)
    
    for key,value in category_final_master.items():
        value.to_csv(output_file_path, mode='a', index = False, header=None)

    tds_salary_df_final=pd.DataFrame(columns=["Party Account","TDS Transaction Nature","PAN","Assessable Value","TDS Rate","TDS Amount"])
    for index, row in tds_salary_df.iterrows(): 
        row_dict={"Party Account":row["Party Account"],"TDS Transaction Nature":row["TDS Transaction Nature"],"PAN":row["PAN"],
                "Assessable Value":float(row["Assessable Value"]),"TDS Rate":float(0),"TDS Amount":float(row["TDS Amount"])}
        
        tds_salary_df_final=tds_salary_df_final.append(row_dict,ignore_index=True)
    
    tds_sal_total=sum(list(tds_salary_df_final["Assessable Value"]))
    tds_tds_total=sum(list(tds_salary_df_final["TDS Amount"]))
    
    row_dict={"Party Account":"Total for SALARIES - 92B","TDS Transaction Nature":"","PAN":"","Assessable Value":tds_sal_total,
            "TDS Rate":"","TDS Amount":tds_tds_total}
    
    tds_salary_df_final=tds_salary_df_final.append(row_dict,ignore_index=True)

    tds_salary_df_final.to_csv(output_file_path, mode='a', index = False, header=None)

    tds_final_one=0
    for key,value in category_final_total.items():
        tds_final_one=tds_final_one+value

    tds_final_one=tds_final_one+tds_sal_total

    tds_final_two=0
    for key,value in category_final_tds.items():
        tds_final_two=value+tds_final_two
    
    tds_final_two=tds_final_two+tds_tds_total

    temp_list=["Total"," "," ",tds_final_one," ",tds_final_two]
    with open(output_file_path, "a") as fp:
        wr = csv.writer(fp, dialect='excel')
        wr.writerow(temp_list)

    os.startfile(output_file_path)
    

if __name__=="__main__":
    root = Tk()
    root.title("Twills Clothing Pvt. Ltd.")
    root.geometry("550x300")

    header_label_one=Label(root,text="TDS Consolidation Tool",anchor="w")
    header_label_one.config(font=("Arial", 16))
    header_label_one.place(x=10,y=10)

    instruction_button = Button(root, text="Instructions")
    instruction_button.config(font=("Arial", 12))
    instruction_button.place(x=400,y=10)

    books_label=Label(root,text="TDS Data : ",font=("bold",10))
    books_label.config(font=("Arial", 12))
    books_label.place(x=10,y=50)

    books_data_file = Button(root,text = "Choose File",command=get_books_file_path)
    books_data_file.config(font=("Arial", 12))
    books_data_file.place(x=140,y=50)

    button=Button(root,text="Consolidate TDS", command=get_tds_consolidation)
    button.config(font=("Arial", 12))
    button.place(x=10,y=150)

    button=Button(root,text="Clear",command=clear_file_paths)
    button.config(font=("Arial", 12))
    button.place(x=180,y=150)

    message_label = Label(root,text = "Message :")
    message_label.config(font=("Arial", 12))
    message_label.place(x=10,y=200)

    message_label=Label(root,text="Welcome !!")
    message_label.config(font=("Arial", 12),fg="blue")
    message_label.place(x=20,y=250)
    message_label.after(1000,message_label.destroy)
    
    root.mainloop()
