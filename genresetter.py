import pandas as pd
import numpy as np
import openpyxl as op
df = pd.read_excel("final.xlsx")
wb = op.Workbook()
st = wb.active
k=0
genre = ["Action","Horror","Thriller","Drama","Biopic","Fantasy","Sci/Fi","Comedy","Animated","Romantic"]
for i in range(2,df['title'].size):
    if(df.loc[i,'Genre'] == 'Romantic'):
        l = df.loc[i,:].values.tolist()
        k+=1
        for j in range(1,len(l)):
            st.cell(row = k,column= j).value = l[j]
wb.save("Romantic.xlsx")