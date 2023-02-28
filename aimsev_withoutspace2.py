import pandas as pd
import csv
from tkinter import *
from tkinter import filedialog as fd
import os
from pathlib import Path


ext = ('.csv')
win = Tk()
win.title('AIMS EVENT READER .CSV')
win.geometry("400x400+10+20")
lbl=Label(win, text="CSV DATA CONVERTER by Prateek", fg='Blue', font=("Helvetica", 10)).place(x=150,y=330)

def btn1():
    dirnamee = fd.askdirectory()
    for filename in os.listdir(dirnamee):
       if filename.endswith(ext):
            name = dirnamee+'/'+filename
            filenamefirst = name
            
            #print(filenamefirst)
            dirname = os.path.dirname(filenamefirst).strip("\"")
            #print(dirname)
            fileName1 = filenamefirst.strip().strip("'")
            #print(fileName1)
            basename = os.path.basename(filenamefirst)
            #print(basename)
            filename=os.path.splitext(basename)[0]
            finalfilename=dirname+'\\'+filename+'-READABLE.xlsx'

            df = pd.read_csv(filenamefirst)
            matrix2 = df[df.columns[0]].to_numpy()
            list1 = matrix2.tolist()
            list2 = matrix2.tolist()
            lst = []
            lstc = []
            lstst = []
            lstid = []
            lstrtc = []
            for i in range(2,102):
                try:
        
                    list=list1[i].split(";")
     
                    lst.append(list[3])
                except:
                    lstrtc.append("-")

    
            for j in range(102,202):
                try:
                    
                    list=list1[j].split(";")
                
                    lstc.append(list[3])
                except:
                    lstrtc.append("-")
    
  
            for j in range(202,302):
                try:
        
                    list=list1[j].split(";")
                    lstst.append(list[3])
                except:
                    lstrtc.append("-")
 
    
            for j in range(302,402):
                try:
        
                    list=list1[j].split(";")
    
                    lstid.append(list[3])
                except:
                    lstrtc.append("-")
        
    

            for j in range(402,502):
                try:
                    list=list1[j].split(";")
                    lstrtc.append(list[3])
                except:
                    lstrtc.append("-")

    
 
    

            lst.insert(0, "Faults")   
 
            lstc.insert(0, "Command")  

            lstst.insert(0, "Status")

            lstid.insert(0, "ID")

            lstrtc.insert(0, "RTC")








            def ctb(n):
                return "{0:b}".format(int(n))

            def split(word):
                return [char for char in word]

   
            lstb=[]  

            for i in range(1,101):
                newww=split(ctb(int(lstst[i])))
                c=len(newww)-1
                listnew=['0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0']
                for j in range( len(newww)) :
                    listnew[15-j]=newww[c]
                    c=c-1 
                lstb.append(listnew)
 
            

            lst15 = []
            lst14 = []
            lst13 = []
            lst12 = []
            lst11 = []
            lst10 = []
            lst9 = []
            lst8 = []
            lst7 = []
            lst6 = []
            lst5 = []
            lst4 = []
            lst3 = []
            lst2 = []
            lst1 = []
            lst0 = []



            for i in range(0,100):
                    try:
                        if(lstb[i][15]=='1'):
                            lst15.append('Th')
                        else:
                            lst15.append('-')
                    except IndexError:
                        lst15.append('NO DATA')
                     
                    try:
                        if(lstb[i][14]=='1'):
                            lst14.append('Phase Loss')
                        else:
                            lst14.append('-')
            
                    except IndexError:
                        lst14.append('NO DATA')
            
                    try:
                        if(lstb[i][13]=='1'):
                          lst13.append('Remote Pos')
                        else:
                            lst13.append('-')
                    except IndexError:
                        lst13.append('NO DATA')
            
                    try:
                        if(lstb[i][12]=='1'):
                          lst12.append('Local Pos')
                        else:
                            lst12.append('-')
                    except IndexError:
                        lst12.append('NO DATA')
            
                    try:
                        if(lstb[i][11]=='1'):
                          lst11.append('LSO')
                        else:
                            lst11.append('-')
                    except IndexError:
                        lst11.append('NO DATA')
             
                    try:
                        if(lstb[i][10]=='1'):
                          lst10.append('LSC')
                        else:
                            lst10.append('-')
                    except IndexError:
                        lst10.append('NO DATA')
            
                    try:
                        if(lstb[i][9]=='1'):
                          lst9.append('TSO')
                        else:
                            lst9.append('-')
                    except IndexError:
                        lst9.append('NO DATA')
            
                    try:
                        if(lstb[i][8]=='1'):
                          lst8.append('TSC')
                        else:
                            lst8.append('-')
                    except IndexError:
                        lst8.append('NO DATA')
            
                    try:
                        if(lstb[i][7]=='1'):
                          lst7.append('Opened')
                        else:
                            lst7.append('-')
                    except IndexError:
                        lst7.append('NO DATA')
            
                    try:
                        if(lstb[i][6]=='1'):
                          lst6.append('Closed')
                        else:
                            lst6.append('-')
                    except IndexError:
                        lst6.append('NO DATA')
            
                    try:
                        if(lstb[i][5]=='1'):
                          lst5.append('Setpoint Reached')
                        else:
                            lst5.append('-')
                    except IndexError:
                        lst5.append('NO DATA')
            
                    try:
                        if(lstb[i][4]=='1'):
                          lst4.append('ESD')
                        else:
                            lst4.append('-')
                    except IndexError:
                        lst4.append('NO DATA')
            
                    try:
                        if(lstb[i][3]=='1'):
                          lst3.append('Running open ')
                        else:
                            lst3.append('-')
                    except IndexError:
                        lst3.append('NO DATA')
            
                    try:
                        if(lstb[i][2]=='1'):
                          lst2.append('Running close')
                        else:
                            lst2.append('-')
                    except IndexError:
                        lst2.append('NO DATA')
            
                    try:
                        if(lstb[i][1]=='1'):
                          lst1.append('Warning')
                        else:
                            lst1.append('-')
                    except IndexError:
                        lst1.append('NO DATA')
            
                    try:
                        if(lstb[i][0]=='1'):
                          lst0.append('Fault')
                        else:
                            lst0.append('-')
                    except IndexError:
                        lst0.append('NO DATA')
            lstcomm=[]          
            for i in range(1,101):
                list=int(lstc[i]) 
                if(list==0):
                    lstcomm.append('NO-CMD')
                elif(list==1):
                    lstcomm.append('C-DCS-OPEN')
                elif(list==2):
                    lstcomm.append('C-DCS-CLOSE')
                elif(list==3):
                    lstcomm.append('C-DCS-STOP')
                elif(list==4):
                    lstcomm.append('C-DCS-SET-POINT')
                elif(list==5):
                    lstcomm.append('C-HMI-OPEN')
                elif(list==6):
                    lstcomm.append('C-HMI-CLOSE')
                elif(list==7):
                    lstcomm.append('C-HMI-STOP')
                elif(list==8):
                    lstcomm.append('C-HMI-SET-POINT')
                elif(list==9):
                    lstcomm.append('C-HMI-ESD')
            lstb.insert(0,"Binary")
            lstcomm.insert(0, "Commands")

            lst0.insert(0, "Fault ")
            lst1.insert(0, "Warning")
            lst2.insert(0, "Running close ")
            lst3.insert(0, "Running open ")
            lst4.insert(0, "ESD ")
            lst5.insert(0, "Setpoint Reached ")
            lst6.insert(0, "Closed ")
            lst7.insert(0, "Opened ")
            lst8.insert(0, "TSC ")
            lst9.insert(0, "TSO ")
            lst10.insert(0, "LSC ")
            lst11.insert(0, "LSO ")
            lst12.insert(0, "Local Pos ")
            lst13.insert(0, "Remote Pos ")
            lst14.insert(0, "Phase Loss ")
            lst15.insert(0, "Th ")

            lstfa=[]          
            for i in range(1,101):
                list=int(lst[i]) 
                if(list==0):
                         lstfa.append('0')
                elif(list==1):
                         lstfa.append('FAULT')
                else:
                     lstfa.append(list)

    

            lstfa.insert(0, "General Fault")
            new_list = [
            lstid,
            lstrtc,
            lstcomm,
            lstfa,
            lstst,
            lst0, 
            lst1,
            lst2, 
            lst3, 
            lst4, 
            lst5, 
            lst6, 
            lst7, 
            lst8, 
            lst9, 
            lst10, 
            lst11,
            lst12,
            lst13,
            lst14,
            lst15, ]

            import xlsxwriter
   
            with xlsxwriter.Workbook(finalfilename) as workbook:  #generate file test.xlsx
                worksheet = workbook.add_worksheet()

                for col_num, data in enumerate(new_list):
                    worksheet.write_column(0 , col_num, data)
            
def btn2():


            filenamefirst = fd.askopenfilename()
            #print(filenamefirst)
            dirname = os.path.dirname(filenamefirst).strip("\"")
            #print(dirname)
            fileName1 = filenamefirst.strip().strip("'")
            #print(fileName1)
            basename = os.path.basename(filenamefirst)
            #print(basename)
            filename=os.path.splitext(basename)[0]
            finalfilename=dirname+'\\'+filename+'-READABLE.xlsx'

            df = pd.read_csv(filenamefirst)
            matrix2 = df[df.columns[0]].to_numpy()
            list1 = matrix2.tolist()
            lst = []
            lstc = []
            lstst = []
            lstid = []
            lstrtc = []
            for i in range(2,102):
                try:
        
                    list=list1[i].split(";")
     
                    lst.append(list[3])
                except:
                    lstrtc.append("-")

    
            for j in range(102,202):
                try:
                    
                    list=list1[j].split(";")
                
                    lstc.append(list[3])
                except:
                    lstrtc.append("-")
    
  
            for j in range(202,302):
                try:
        
                    list=list1[j].split(";")
                    lstst.append(list[3])
                except:
                    lstrtc.append("-")
 
    
            for j in range(302,402):
                try:
        
                    list=list1[j].split(";")
    
                    lstid.append(list[3])
                except:
                    lstrtc.append("-")
        
    

            for j in range(402,502):
                try:
                    list=list1[j].split(";")
                    lstrtc.append(list[3])
                except:
                    lstrtc.append("-")

    
 
    

            lst.insert(0, "Faults")   
 
            lstc.insert(0, "Command")  

            lstst.insert(0, "Status")

            lstid.insert(0, "ID")

            lstrtc.insert(0, "RTC")








            def ctb(n):
                return "{0:b}".format(int(n))

            def split(word):
                return [char for char in word]

   
            lstb=[]  

            for i in range(1,101):
                newww=split(ctb(int(lstst[i])))
                c=len(newww)-1
                listnew=['0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0']
                for j in range( len(newww)) :
                    listnew[15-j]=newww[c]
                    c=c-1 
                lstb.append(listnew)
 
            

            lst15 = []
            lst14 = []
            lst13 = []
            lst12 = []
            lst11 = []
            lst10 = []
            lst9 = []
            lst8 = []
            lst7 = []
            lst6 = []
            lst5 = []
            lst4 = []
            lst3 = []
            lst2 = []
            lst1 = []
            lst0 = []



            for i in range(0,100):
                    try:
                        if(lstb[i][15]=='1'):
                            lst15.append('Th')
                        else:
                            lst15.append('-')
                    except IndexError:
                        lst15.append('NO DATA')
                     
                    try:
                        if(lstb[i][14]=='1'):
                            lst14.append('Phase Loss')
                        else:
                            lst14.append('-')
            
                    except IndexError:
                        lst14.append('NO DATA')
            
                    try:
                        if(lstb[i][13]=='1'):
                          lst13.append('Remote Pos')
                        else:
                            lst13.append('-')
                    except IndexError:
                        lst13.append('NO DATA')
            
                    try:
                        if(lstb[i][12]=='1'):
                          lst12.append('Local Pos')
                        else:
                            lst12.append('-')
                    except IndexError:
                        lst12.append('NO DATA')
            
                    try:
                        if(lstb[i][11]=='1'):
                          lst11.append('LSO')
                        else:
                            lst11.append('-')
                    except IndexError:
                        lst11.append('NO DATA')
             
                    try:
                        if(lstb[i][10]=='1'):
                          lst10.append('LSC')
                        else:
                            lst10.append('-')
                    except IndexError:
                        lst10.append('NO DATA')
            
                    try:
                        if(lstb[i][9]=='1'):
                          lst9.append('TSO')
                        else:
                            lst9.append('-')
                    except IndexError:
                        lst9.append('NO DATA')
            
                    try:
                        if(lstb[i][8]=='1'):
                          lst8.append('TSC')
                        else:
                            lst8.append('-')
                    except IndexError:
                        lst8.append('NO DATA')
            
                    try:
                        if(lstb[i][7]=='1'):
                          lst7.append('Opened')
                        else:
                            lst7.append('-')
                    except IndexError:
                        lst7.append('NO DATA')
            
                    try:
                        if(lstb[i][6]=='1'):
                          lst6.append('Closed')
                        else:
                            lst6.append('-')
                    except IndexError:
                        lst6.append('NO DATA')
            
                    try:
                        if(lstb[i][5]=='1'):
                          lst5.append('Setpoint Reached')
                        else:
                            lst5.append('-')
                    except IndexError:
                        lst5.append('NO DATA')
            
                    try:
                        if(lstb[i][4]=='1'):
                          lst4.append('ESD')
                        else:
                            lst4.append('-')
                    except IndexError:
                        lst4.append('NO DATA')
            
                    try:
                        if(lstb[i][3]=='1'):
                          lst3.append('Running open ')
                        else:
                            lst3.append('-')
                    except IndexError:
                        lst3.append('NO DATA')
            
                    try:
                        if(lstb[i][2]=='1'):
                          lst2.append('Running close')
                        else:
                            lst2.append('-')
                    except IndexError:
                        lst2.append('NO DATA')
            
                    try:
                        if(lstb[i][1]=='1'):
                          lst1.append('Warning')
                        else:
                            lst1.append('-')
                    except IndexError:
                        lst1.append('NO DATA')
            
                    try:
                        if(lstb[i][0]=='1'):
                          lst0.append('Fault')
                        else:
                            lst0.append('-')
                    except IndexError:
                        lst0.append('NO DATA')
            lstcomm=[]          
            for i in range(1,101):
                list=int(lstc[i]) 
                if(list==0):
                    lstcomm.append('NO-CMD')
                elif(list==1):
                    lstcomm.append('C-DCS-OPEN')
                elif(list==2):
                    lstcomm.append('C-DCS-CLOSE')
                elif(list==3):
                    lstcomm.append('C-DCS-STOP')
                elif(list==4):
                    lstcomm.append('C-DCS-SET-POINT')
                elif(list==5):
                    lstcomm.append('C-HMI-OPEN')
                elif(list==6):
                    lstcomm.append('C-HMI-CLOSE')
                elif(list==7):
                    lstcomm.append('C-HMI-STOP')
                elif(list==8):
                    lstcomm.append('C-HMI-SET-POINT')
                elif(list==9):
                    lstcomm.append('C-HMI-ESD')
            lstb.insert(0,"Binary")
            lstcomm.insert(0, "Commands")

            lst0.insert(0, "Fault ")
            lst1.insert(0, "Warning")
            lst2.insert(0, "Running close ")
            lst3.insert(0, "Running open ")
            lst4.insert(0, "ESD ")
            lst5.insert(0, "Setpoint Reached ")
            lst6.insert(0, "Closed ")
            lst7.insert(0, "Opened ")
            lst8.insert(0, "TSC ")
            lst9.insert(0, "TSO ")
            lst10.insert(0, "LSC ")
            lst11.insert(0, "LSO ")
            lst12.insert(0, "Local Pos ")
            lst13.insert(0, "Remote Pos ")
            lst14.insert(0, "Phase Loss ")
            lst15.insert(0, "Th ")

            lstfa=[]    
                  
            for i in range(1,101):
                list=int(lst[i]) 
                if(list==0):
                         lstfa.append('0')
                elif(list==1):
                         lstfa.append('FAULT')
                else:
                     lstfa.append(list)

            lstfa.insert(0, "General Fault")
            new_list=[lstid,lstrtc,lstcomm,lstfa,lstst,
            lst0, 
            lst1,
            lst2, 
            lst3, 
            lst4, 
            lst5, 
            lst6, 
            lst7, 
            lst8, 
            lst9, 
            lst10, 
            lst11,
            lst12,
            lst13,
            lst14,
            lst15, ]

            import xlsxwriter
   
            with xlsxwriter.Workbook(finalfilename) as workbook:  #generate file test.xlsx
                worksheet = workbook.add_worksheet()

                for col_num, data in enumerate(new_list):
                    worksheet.write_column(0 , col_num, data) 

                   
button1 =  Button(win, text="ALL CSV FILES IN DIRECTORY", command=btn1,bd=4, ).place(x=110, y=60)
button2 =  Button(win, text="SINGLE CSV FILE", command=btn2,bd=4).place(x=140, y=150)



win.mainloop()            

    
  
    
    
   
    
  
    
    
