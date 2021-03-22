import pandas as pd
import openpyxl
from openpyxl import load_workbook
z = pd.read_excel('Data.xlsx', sheet_name=['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5'])
tmp = []
df1=pd.DataFrame()
df4=pd.DataFrame()
df5=pd.DataFrame()
df6=pd.DataFrame()
x1 = pd.DataFrame(list(z.items()))
df4=pd.DataFrame(x1[1][0].columns)
df6=df6.append(df4)
for i in range (1,5):
    df5=pd.DataFrame(x1[1][i].columns[3:10])
    df6=df6.append(df5)
df=pd.DataFrame(df6)
df3=pd.DataFrame()
n = int(input('Enter no of inputs:-'))
count=0
for _ in range(n):
    tmp1 = []
    h = int(input('Enter your ps no:-'))
    name = str(input('Enter the Name:-'))
    email = str(input('Enter the email:-'))
    tmp1.append(h)
    tmp1.append(name)
    tmp1.append(email)
    tmp.append(tmp1)

#df1 = pd.DataFrame(columns=[z.columns])


for i in tmp:
    h, name, email = i
    y = z['Sheet1']
    y = y[(y['PS number'] == h) & (y['Display Name'] == name) & (y['Official Email Address'] == email)]
    if len(y) == 0:
        print('No match')
    else:
        df = pd.DataFrame(y, columns = ['SL#','PS number','Display Name','Official Email Address'])
        for i in z.keys():
            x = z[i]
            t = x[(x['PS number'] == h) & (x['Display Name'] == name) & (x['Official Email Address'] == email)]
            col = x.columns
            for j in col:
                df[j] = t[j]
                count=count+1
                df3.at[i,'No of columns']=count
    df1=df1.append(df)
df2 = df1.describe()
df3.at[1,'A']=(len(df1.columns)*n)
book = load_workbook('Data.xlsx')
writer = pd.ExcelWriter('Data.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
df1.to_excel(writer, sheet_name='master', index=False)
df3.to_excel(writer, sheet_name='summary',index=False)
writer.save()