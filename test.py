import pandas as pd

data=pd.read_excel('WithEmail.xlsx')

if 'Email' in data.columns:
    emails=list(data['Email'])
# print(emails)
    c=[]
    for i in emails:
      if pd.isnull(i)==False:
        # print(i)
          c.append(i)

    emails=c
    print(emails)     
else:
    print('Not Exist')    

       