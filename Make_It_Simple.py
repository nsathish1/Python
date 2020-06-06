import tkinter as tk
import numpy as np
import pandas as pd
import glob
import pathlib
from datetime import datetime
import re


## Creating a greeting note

def greetings():
    
    now = datetime.now()
    hour = now.hour

    if hour < 12:
        greeting = "Good morning!"
    elif hour < 18:
        greeting = "Good afternoon!"
    elif hour < 22:
        greeting = "Good evening!"   
    else:
        greeting = "Good night!"

    return greeting   

all_data = pd.DataFrame()

## Return Current time for folder name creation

def Current_Time():
    
    current_time = datetime.now().strftime("%d%m%H%M%S")
    return current_time

## Creates folder

def Create_Folder(Folder_Path,Folder_Name):
    
    path = pathlib.Path(Folder_Path+"\\"+Folder_Name)
    path.mkdir()
    Final_path = Folder_Path+"\\"+Folder_Name
    return Final_path

## Main logic flow function
def ops_main(Task_name,S_path,D_path,Split_count,Col_name,Value):

## Validating the input data for empty string

    if (S_path.isspace() or len(S_path) == 0 or D_path.isspace() or len(D_path) == 0 
        or Split_count.isspace() or len(Split_count) == 0 or Col_name.isspace() or len(Col_name) == 0
        or Value.isspace() or len(Value) == 0):
        
        res = "error: Required fields are empty"
        myText.set(res)
        return
    
    S_path=S_path.replace("\"",'')
    D_path=D_path.replace("\"",'')
        
## Validating the presence of end charcter("\") in input string

    if(S_path[-1] == '\\'):
        S_path = S_path[:-1]
        
    if(D_path[-1] == '\\'):
        D_path = D_path[:-1]
        
## Validating the input file format(.xlsx)
    
    if (Task_name == 'Split_excel' or Task_name == 'Filter_excel'):
        
        Match = S_path.find(".xlsx")
        
        if (Match <= 0):
            res = "error: Source file not in .xlsx format "
            myText.set(res)
            return   
        
## Creating criteria for button navigation
    
    if pathlib.Path(S_path).exists():
        
        if pathlib.Path(D_path).exists():
            
            if (Task_name == 'Combine_excel' ):
                ops_combine_Excel(S_path,D_path)
            if (Task_name == 'Split_excel' ):
                ops_split_excel(S_path,D_path,Split_count)
            if (Task_name == 'Filter_excel' ):
                ops_filter_excel(S_path,D_path,Col_name,Value)
            
        else:
             res = "error: Destination folder does not exist"
             myText.set(res)
    else:
    
        res = "error: Source folder does not exist"
        myText.set(res)

def ops_combine_Excel(S_path,D_path):
               
    all_data = pd.DataFrame()

##Looping through source folder for .xlsx file(reading and updating the data in df-dataframe)     
    
    for f in glob.glob(S_path +"\*.xlsx"):
        
        df = pd.read_excel(f)
        all_data = all_data.append(df,ignore_index=False)

    Folder_Name = Current_Time()
    Final_path = Create_Folder(D_path,Folder_Name)

    all_data.reset_index(drop = True, inplace = True)
    writer = pd.ExcelWriter(Final_path+"\Output.xlsx")
    all_data.to_excel(writer,'sheet1',index = False)
    writer.save()
    
    if pathlib.Path(Final_path+"\Output.xlsx").is_file():
      
        res = "Success"
    else:
        res = "Failure"
        
    myText.set(res)  
   
def ops_split_excel(S_path,D_path,Split_count):
   
    Folder_Name = Current_Time()
    Final_path = Create_Folder(D_path,Folder_Name)
   
  ##Finding the file name from the given source excel path
    File_name= re.search(r'[^\\/:*?"<>|\r\n]+$',S_path).group().replace('.xlsx',"")
    data_xlsx = pd.read_excel(S_path)
    
    ##converting the file to csv for chunk
    csv_path = Final_path+"\\"+File_name+".csv"

    data_xlsx.to_csv(csv_path, encoding='utf-8',index=False)

    i=0
    split = int(Split_count)
    for df in pd.read_csv(csv_path,chunksize=split):
        
         df.to_excel(Final_path+"\\"+File_name+"_"+'{:02d}.xlsx'.format(i), index=False)
         i += 1
        
    res = "Success"
    myText.set(res) 

def ops_filter_excel(S_path,D_path,Col_name,Value):
    
    Folder_Name = Current_Time()
    Final_path = Create_Folder(D_path,Folder_Name)
    
    File_name= re.search(r'[^\\/:*?"<>|\r\n]+$',S_path).group().replace('.xlsx',"")
  
    df = pd.read_excel(S_path)
    d_type = df[Col_name].dtype

    if (d_type == 'object'):
        grouped = df.groupby(df[Col_name])
        Out = grouped.get_group(str(Value))

    if (d_type == 'int64'):
        grouped = df.groupby(df[Col_name])
        Out = grouped.get_group(int(Value))

    elif (d_type == 'float64'):
        grouped = df.groupby(df[Col_name])
        Out = grouped.get_group(float(Value))
    
    writer = pd.ExcelWriter(D_path+"\\"+File_name+"_"+Value+"_Filter.xlsx")
    Out.to_excel(writer,'sheet1',index = False)
    writer.save()
    
    res = "Success"
    myText.set(res)
   
def combine_excel():
    
    tk.Label(master,text="Please enter the below details to proceed.", bg="grey90",fg="black",
             font=(None, 10)).grid(row=10, sticky=tk.W,pady=10)
    
    tk.Label(master,text="SOURCE FOLDER PATH        : ", bg="grey90",fg="black",font=(None, 10)).grid(row=15, sticky=tk.W,pady=5)
    tk.Label(master,text="DESTINATION FOLDER PATH   : ",bg="grey90",fg="black", font=(None, 10)).grid(row=17, sticky=tk.W,pady=5)
    tk.Label(master,text="RESULT                    :",bg="grey90",fg="black", font=(None, 10)).grid(row=25, sticky=tk.W,pady=5)
    tk.Label(master,text="                                                          ",bg="grey90",fg="black", font=(None, 10)).grid(row=19, sticky=tk.W,pady=5)
    tk.Label(master,text="                                                            ",bg="grey90",fg="black", font=(None, 10)).grid(row=19,column = 1, sticky=tk.W,pady=5)
    
    tk.Label(master,text="                                                          ",bg="grey90",fg="black", font=(None, 10)).grid(row=21, sticky=tk.W,pady=5)
    tk.Label(master,text="                                                            ",bg="grey90",fg="black", font=(None, 10)).grid(row=21,column = 1, sticky=tk.W,pady=5)

    
    e1 = tk.Entry(master,width=35)
    e2 = tk.Entry(master,width=35)
    e1.grid(row=15, column=1)
    e2.grid(row=17, column=1)
    
    tk.Button(master,text='RUN',command= lambda: ops_main("Combine_excel",e1.get(),e2.get(),"Null","Null","Null")
              ,bg="grey80",fg="black",font=(None, 10),height = 1, 
          width = 7).grid(row=22,column=1,sticky=tk.W,pady=12)
    result=tk.Label(master, text="", textvariable=myText,bg="grey90",fg="black", font=(None,10)).grid(row=25,column=1, sticky=tk.W,pady=10)
    
def split_excel():
       
    tk.Label(master,text="SOURCE EXCEL FILE PATH    : ", bg="grey90",fg="black",font=(None, 10)).grid(row=15, sticky=tk.W,pady=5)
    tk.Label(master,text="DESTINATION FOLDER PATH   : ",bg="grey90",fg="black", font=(None, 10)).grid(row=17, sticky=tk.W,pady=5)
    tk.Label(master,text="                                            ",bg="grey90",fg="black",font=(None, 10)).grid(row=19,column = 0, sticky=tk.W,pady=5)
    tk.Label(master,text="                                            ",bg="grey90",fg="black",font=(None, 10)).grid(row=19,column = 1, sticky=tk.W,pady=5)
    tk.Label(master,text="                                            ",bg="grey90",fg="black",font=(None, 10)).grid(row=21, sticky=tk.W,pady=5)
    tk.Label(master,text="                                                       ",bg="grey90",fg="black",font=(None, 10)).grid(row=21,column =1, sticky=tk.W,pady=5)
    tk.Label(master,text="SPLIT COUNT               : ",bg="grey90",fg="black", font=(None, 10)).grid(row=19, sticky=tk.W,pady=5)
    tk.Label(master,text="RESULT                    :",bg="grey90",fg="black", font=(None, 10)).grid(row=25, sticky=tk.W,pady=5)


    e3 = tk.Entry(master,width=35)
    e4 = tk.Entry(master,width=35)
    e5 = tk.Entry(master,width=35)
    e3.grid(row=15, column=1)
    e4.grid(row=17, column=1)
    e5.grid(row=19, column=1)
    
    tk.Button(master,text='RUN',command=lambda: ops_main("Split_excel",e3.get(),e4.get(),e5.get(),"Null","Null")
              ,bg="grey80",fg="black",font=(None, 10),height = 1, 
          width = 7).grid(row=22,column=1,sticky=tk.W,pady=12)
    result=tk.Label(master, text="", textvariable=myText,bg="grey90",fg="black", font=(None, 10)).grid(row=25,column=1, sticky=tk.W,pady=10)
       
def filter_excel():
       
    tk.Label(master,text="SOURCE EXCEL FILE PATH    : ", bg="grey90",fg="black",font=(None, 10)).grid(row=15, sticky=tk.W,pady=5)
    tk.Label(master,text="DESTINATION FOLDER PATH   : ",bg="grey90",fg="black", font=(None, 10)).grid(row=17, sticky=tk.W,pady=5)
    #tk.Label(master,text="                                                          ",bg="grey90",fg="black", font=(None, 10)).grid(row=19, sticky=tk.W,pady=5)
    #tk.Label(master,text="                                                            ",bg="grey90",fg="black", font=(None, 10)).grid(row=19,column = 1, sticky=tk.W,pady=5)
    tk.Label(master,text="COLUMN NAME               : ",bg="grey90",fg="black", font=(None, 10)).grid(row=19, sticky=tk.W,pady=5)
    tk.Label(master,text="VALUE TO BE FILTERED         : ",bg="grey90",fg="black", font=(None, 10)).grid(row=21, sticky=tk.W,pady=5)
    tk.Label(master,text="RESULT                    :",bg="grey90",fg="black", font=(None, 10)).grid(row=25, sticky=tk.W,pady=5)
        
    e3 = tk.Entry(master,width=35)
    e4 = tk.Entry(master,width=35)
    e5 = tk.Entry(master,width=35)
    e6 = tk.Entry(master,width=35)
    e3.grid(row=15, column=1)
    e4.grid(row=17, column=1)   
    e5.grid(row=19, column=1)  
    e6.grid(row=21, column=1)  
    
    tk.Button(master,text='RUN',command=lambda: ops_main("Filter_excel",e3.get(),e4.get(),"Null",e5.get(),e6.get())
              ,bg="grey80",fg="black",font=(None, 10),height = 1, 
          width = 7).grid(row=22,column=1,sticky=tk.W,pady=12)       
    result=tk.Label(master, text="", textvariable=myText,bg="grey90",fg="black", font=(None, 10)).grid(row=25,column=1, sticky=tk.W,pady=10)

def close_window():
    
    master.destroy()      
    
##############TK box-UI DESIGN###################################################

master = tk.Tk()
myText=tk.StringVar();
master.geometry("600x490+400+100")
master.title("Make It Simple!")
master.configure(background='grey90')

tk.Label(master, text= "Hi "+greetings(),font=(None, 15),bg="grey90",fg="Black").grid(row=0,column=0)
tk.Label(master, text= "Please choose your option from the below menu buttons",font=(None, 10),bg="grey90",fg="Black").grid(row=2,column=0)



tk.Button(master,text='COMBINE EXCEL',command=combine_excel,bg="grey80",fg="black",
          font=(None, 11)).grid(row=4,column=0,sticky=tk.W,pady=10)

tk.Button(master,text='SPLIT EXCEL', command=split_excel,bg="grey80",fg="black",
          font=(None, 11)).grid(row=6,column=0,sticky=tk.W,pady=10)

tk.Button(master,text='FILTER EXCEL DATA', command=filter_excel,bg="grey80",fg="black",
          font=(None, 11)).grid(row=8,column=0,sticky=tk.W,pady=10)

tk.Button(master,text='EXIT', command=close_window,bg="grey80",fg="black",font=(None, 11),height = 1, 
          width = 7).grid(row=8,column=1,sticky=tk.W,pady=10)

tk.Label(master,text="*Designed and developed by Sathish Nageshwaran",bg="grey90",fg="black", font=(None, 7)).grid(row=400,column=1,sticky=tk.W,pady=5)
#myFont = font.Font(size=30)
##


tk.mainloop()
