import pandas as pd
import openpyxl as op
import numpy as np
df = pd.read_excel("final.xlsx")
dp = pd.read_excel("FrontEnd\Book.xlsx")
wb = op.Workbook()
st = wb.active
k=0
for i in range(df['title'].size):
    if(dp.loc[0,0] == df.loc[i,'Mood']):
        if(dp.loc[0,1] == df.loc[i,'Occasion']):
            if(dp.loc[0,2] == df.loc[i,'Genre']):
                if(dp.loc[0,3] == df.loc[i,'IMDB']):
                    if(dp.loc[0,4] == df.loc[i,'Year']):
                        l = df.loc[i,:].values.tolist()
                        k+=1
                        for j in range(1,len(l)):
                            st.cell(roq = k, column=j).value = l[j]
wb.save("Answer.xlsx")