import pandas as pd
import numpy as np
import time
from gurobipy import*
import datetime
from datetime import datetime 

operation=pd.read_excel("processtime.xls")
order=pd.read_excel("order.xls")
attendance=pd.read_excel("attendance.xls")
resource=pd.read_excel("resource.xls")
#----------
order=order.drop(["orderCode1"],axis=1)
order['notBefore'] = order['notBefore'].astype("datetime64")
order['notAfter'] = order['notAfter'].astype("datetime64")
print(order.dtypes)
#----------
operation=operation.drop(["sequence1"],axis=1)
#----------
resource=resource.drop(["weekday"],axis=1)
#----------
#------改變時間類型-------
attendance["start"] = pd.to_timedelta(attendance["start"]+ ':00')
attendance["end"] = pd.to_timedelta(attendance["end"] + ':00')
print(attendance.dtypes)
#------------------------
df = pd.DataFrame()
df=pd.merge(order,operation, on="productCode", how="outer")
df=df.drop(["resourceCode"],axis=1)

df = df.reindex(columns=['orderCode',"productCode","sequence","operationCode","resourceCode1","operationTime","prepareTime","notBefore", "notAfter"])
df=pd.merge(df,resource, on="resourceCode1", how="inner")


df = df.reindex(columns=['orderCode',"productCode","sequence","operationCode","resourceCode1","attendanceCode","prepareTime","operationTime","notBefore", "notAfter"])
df=df.sort_values(by=["orderCode","sequence"])
df["Total time"]=df["prepareTime"]+df["operationTime"]
df = df.reindex(columns=['orderCode',"productCode","sequence","operationCode","resourceCode1","attendanceCode","Total time","notBefore", "notAfter"])

df=pd.merge(df,attendance, on="attendanceCode", how="outer")
df=df.sort_values(by=["orderCode","sequence"])
df = df.reindex(columns=['orderCode',"productCode","sequence","operationCode","resourceCode1","attendanceCode","start","Total time","end","notBefore", "notAfter"])

df1 = pd.DataFrame(df)
df1.to_excel("output0.xlsx") 
