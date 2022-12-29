import numpy as np
import pandas as pd
import datetime
from datetime import date

reg = pd.read_excel("Document_register.xlsx")
reg.columns = reg.iloc[1]

Clough_Status = reg["Document Status"]
for i in Clough_Status.index:
    if (Clough_Status[i]== "canceled"):
        reg = reg.drop(i)

today = np.datetime64("today" , "D")

def outlist(reg):

    TM_History = reg["TM History TM(Rev)"]
    TM_History = TM_History.fillna("")

    for x in TM_History.index:
        if (TM_History[x] != ""):
            reg = reg.drop(x)

    regC = reg.drop(0)
    regC = regC.dropna(subset = ["Clough Doc No"])
    regC["Forecast Submission Date2"] = pd.to_datetime(regC["Forecast Submission Date2"] , errors="coerce")
    regC = regC.sort_values(by = "Forecast Submission Date2")

    return regC
    print(regC["Forecast Submission Date2"])

def overdue(reg):

    reg = outlist(reg)

    due_date = (reg["Forecast Submission Date2"])
    due_date = due_date.dropna()
    overdue=(due_date - today)
    overdue = overdue.astype("timedelta64[D]").astype(int)

    overduefinal = []

    for j in overdue.index:
        if int(overdue[j])< 0:
            overduefinal.append(j)

    print("")
    print("")
    print("There are ", len(overduefinal), " overdue documents as of today(" , today, ") aas listed below: ")
    print("")
    print("")

    k=1
    for j in overduefinal:

        print(k, "-" , reg.Title[j], "( ", reg["Clough Doc No"][j]," )")
        print(overdue[j], " days overdue (" , due_date[j].date(), ")")
        print("")
        k +=1

def outstanding(reg):
    reg = outlist(reg)

    print("")
    print("")
    print(" and there are ", reg.shape[0], "outstanding documents as listed below: ")
    print("")
    print("")
    
    k=1
    for x in reg.index:
        due_date = (reg["Forecast Submission Date2"][x])
        od = (due_date - today).days
        
        if od >0 :
            
            
            print(k, "-" , reg["Title"][x], "(" , reg["Clough Doc No"][x], ") Due in " , od , "days") 
            print("") 
            
        elif od<0:
            print(k, "-" , reg["Title"][x], "(" , reg["Clough Doc No"][x], ") overdue " , -(od) , "days") 
            print("")
            
        
       
            
    k+=1

def report(reg):

    regout = outlist(reg)

    due_date = (reg["Forecast Submission Date2"])
    due_date = due_date.dropna()

    submitted = due_date.shape[0] - regout.shape[0]

    print(" We have Submitted " , submitted, "out of " , due_date.shape[0], "documents in contract and: ")
    
    overdue(reg)
    outstanding(reg)

report(reg)
