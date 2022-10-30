import pandas as pd
import numpy as np
import os  
import datetime
import calendar
import csv 
from datetime import date
from datetime import datetime,timedelta
import itertools 
from csv import writer
from csv import DictWriter

df1=pd.read_csv("input_attendance.csv")
df2=pd.read_csv("input_registered_students.csv")

df1=df1.dropna()
df1['Timestamp'].str.split(' ',expand=True)
df1[['Dates','Time']]=df1['Timestamp'].str.split(' ',expand=True)
#df1[['Roll_No','Name']]=df1['Attendance'].str.split(' ',1,expand=True)
df1['Roll'] = df1.Attendance.str.split(' ', expand = True)[0]
df1[['hour','minute','second']]=df1['Time'].str.split(':',expand=True)
df1['Dates'] = pd.to_datetime(df1['Dates'],format = "%d/%m/%Y")
df1['dayOfWeek'] = df1['Dates'].dt.day_name()



temp1=df1['Dates'].iloc[0]
temp2=df1['Dates'].iloc[-1]
final1 = temp1.strftime('%d-%m-%Y')
final2=temp2.strftime('%d-%m-%Y')
unique_date=[]
for i in df1['Dates']:
  unique_date.append(i)
res=[*set(unique_date)]
total_no_of_lectures=0
for i in res:
  d = pd.Timestamp(i)
  if d.day_name()=='Thursday' or d.day_name()=='Monday' :
    total_no_of_lectures+=1
#print(total_no_of_lectures)  
 
time1='14:00:00'
time2='15:00:00'



TIME1 = datetime.strptime(time1, '%H:%M:%S')
TIME2=datetime.strptime(time2, '%H:%M:%S')



x=[]
for i in df1['Dates']:
  x.append(i)
y=[]
for i in df1['Roll']:
  y.append(i)
p=[]
for i in df2['Roll No']:
  p.append(i)
q=[]
for i in df2['Name']:
  q.append(i)
r=[]
for i in df1['dayOfWeek']:
  r.append(i)
s=[]
for i in df1['Time']:
  s.append(i)

col=['Roll','Name','total_no_of_lectures','attendance_count_actual','attendance_count_fake','Percentage (attendance_count_actual/total_lecture_taken) 2 digit decimal ']
with open('output/attendance_report_consolidated.csv','w')as file1: ##-->creating a file1.csv 
        csvWriter=csv.writer(file1,delimiter=',')
        csvWriter.writerow(col)

        for (c,d) in zip(p,q):
          attendance_count_actual=0
          attendance_count_fake=0
          arr=[]
          for uD in unique_date:
            for (a,b,e,f) in zip(x,r,y,s):
              datetime_object = datetime.strptime(f, '%H:%M:%S')
              if ((b=='Monday'or b=='Thursday') and e==c and(datetime_object>=TIME1 and datetime_object<=TIME2)and a==uD):
                attendance_count_actual+=1
              elif((b!='Monday'or b!='Thursday') and e==c and a==uD):
                attendance_count_fake+=1 
            if attendance_count_actual>1:
              attendance_count_fake+=(attendance_count_actual-1)
              attendance_count_actual=1
          attendance_count_absent=total_no_of_lectures-attendance_count_actual
          percentage=(attendance_count_actual/total_no_of_lectures)*100
          arr=[c,d,total_no_of_lectures,attendance_count_actual,attendance_count_fake,percentage]
          csvWriter.writerow(arr)
          


df=pd.read_csv('output/attendance_report_consolidated.csv') 
df.to_csv('output/attendance_report_consolidated.csv',mode='w',header=True,index=False)

for (c,d,i,j,k,l) in zip(p,q,df['total_no_of_lectures'],df['attendance_count_actual'],df['attendance_count_fake'],df['Percentage (attendance_count_actual/total_lecture_taken) 2 digit decimal ']):
  with open('file2.csv','w')as file2: ##-->creating a file1.csv 
          csvWriter=csv.writer(file2,delimiter=',')
          csvWriter.writerow(col)
          
          
          arr=[c,d,i,j,k,l]
          csvWriter.writerow(arr)
  
  df3=pd.read_csv('file2.csv') 
  df3.to_csv('output/'+c+'.csv',mode='w',header=True,index=False)    
          
  os.remove('file2.csv')       
          
                  


df1.to_csv('output.csv',index=False)