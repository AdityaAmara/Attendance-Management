import pandas as pd
import xlsxwriter
from datetime import date

df = pd.read_excel("data.xlsx")
n=len(df.index)
dlis = df['Roll.no']
_dict = {}
for stu_no in range(0,n):
    _dict[str(dlis[stu_no])] = stu_no

#print(_dict)
#_dict = { '1':0,'3':1,'5':2,'7':3,'8':4,'9':5,'11':6,'14':7,'15':8 }

def format():
    df = pd.read_excel("data.xlsx")
    samlist= list(df.columns)
    for a in range(2,len(samlist)):
        df[samlist[a]]=df[samlist[a]].fillna(0)
    df.to_excel("data.xlsx",index=False)
    print(df.to_string(index = False))

def attendance():
    today = date.today()
    dat = today.strftime("%d/%m/%Y")
    print(dat)
    l=list()
    for i in range(0,n):
        l.append(0)
    print("Enter the Roll.no's of Present Students:\nEnter X to terminate")
    while True:
        try:
            take = input()
            if(take=='X'):
                break
            else:
                ind = _dict.get(take)
                l[ind]=1
                print("Next")
        except:
            print("Enter a valid Roll.no")
        
    temp_att = l
    df[dat]= temp_att
    df.to_excel("data.xlsx",index=False)
    print(df)
    x=temp_att.count(1)
    print("The no of Presenties people is " + str(x))

def extraclass():
    today = date.today()
    dat = today.strftime("%d/%m/%Y")
    print("Extra class on: " + str(dat))
    dat = dat + 'Extra Class'
    l=list()
    for i in range(0,n):
        l.append(0)
    print("Enter the Roll.no's of Present Students:\nEnter X to terminate")
    while True:
        try:
            take = input()
            if(take=='X'):
                break
            else:
                ind = _dict.get(take)
                l[ind]=1
                print("Next")
        except:
            print("Enter a valid Roll.no")
        
    temp_att = l
    df[dat]= temp_att
    df.to_excel("data.xlsx",index=False)
    print(df)
    x=temp_att.count(1)
    print("The no of Presenties people is " + str(x))

def check(take):
    try: 
        take = str(take)
        ind = _dict.get(take)
        df2= pd.read_excel("data.xlsx",index_col="Roll.no")
        row= df2.iloc[ind]
        sum = 0
        for k in range(1,len(row)):
            sum = sum + row[k]
        tot=len(row)-1
        per = (sum/tot)*100
        return per
    except:
        print("Error !\nPlease enter checkper(Valid Roll.no)")

def checkper(take):
    try:
        take = str(take)
        ind = _dict.get(take)
        df2= pd.read_excel("data.xlsx",index_col="Roll.no")
        row= df2.iloc[ind]
        print(row[0])
        sum = 0
        for k in range(1,len(row)):
            sum = sum + row[k]
        print("Total classes Attended: "+ str(sum))
        print("Total classes Conducted: " + str(len(row)-1))
        tot=len(row)-1
        per = sum/tot*100
        per = str(per)
        per = per[0:5]
        print("Percentage: " + str(per))
    except:
        print("Error !\nPlease enter checkper(Valid Roll.no)")

def checkfull():
    templis = list()
    key_list = list(_dict.keys())
    print("Percentages:")
    for i in _dict:
        templis.append(check(i))
    j=0
    for i in templis:
        temp = str(templis[j])
        temp = temp[0:5]
        print(str(key_list[j]) + " " + str(temp))
        j=j+1

def debar():
    templis = list()
    key_list = list(_dict.keys())
    print("The Debar list (<75%)")
    for i in _dict:
        templis.append(check(i))
    #print(templis)
    j=0
    for i in templis:
        if(i<75):
            temp = str(templis[j])
            temp = temp[0:5]
            print(str(key_list[j]) + " " + str(temp))
        j=j+1


