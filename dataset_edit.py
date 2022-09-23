import pandas as pd
import numpy as np
from sklearn.preprocessing import LabelEncoder
import random
mood = ["Happy","Sad","Neutral"]
genre = ["Action","Horror","Thriller","Drama","Fantasy","Scif","Comedy","Romance"]
age = ["gRated","pgRated","pg13Rated","rRated"]
ocn = ["Family","Friends","Date","Alone"]
year = ["before80","8000","0020","recent"]
idb = [5,5.1,5.2,5.3,5.4,5.5,5.6,5.7,5.8,5.9,6,6.1,6.2,6.3,6.4,6.5,6.6,6.7,6.8,6.9,7,7.1,7.2,7.3,7.4,7.5,7.6,7.7,7.8,7.9,8,8.1,8.2,8.3,8.4,8.5,8.6,8.7,8.8,8.9,9,9.1,9.2,9.3,9.4,9.5,9.6,9.7,9.8,9.9]
data = pd.read_excel("movies.xlsx")
updt = pd.ExcelWriter("final.xlsx")
data.drop('genres',axis=1,inplace=True)
data.to_excel(updt,'Sheet1')
# updt.save()
moodf = []
for i in range(data['title'].size):
    moodf.append(random.randint(0,2))
moods=[]
for i in moodf:
    moods.append(mood[i])
df = data.assign(Mood = moods)
df.to_excel(updt,'Sheet1')
genf = []
for i in range(data['title'].size):
    genf.append(random.randint(0,7))
gens=[]
for i in genf:
    gens.append(genre[i])
df2 = df.assign(Genre = gens)
df2.to_excel(updt,'Sheet1')
agef = []
for i in range(data['title'].size):
    agef.append(random.randint(0,3))
ages=[]
for i in agef:
    ages.append(age[i])
df3 = df2.assign(Age = ages)
df3.to_excel(updt,'Sheet1')

ocnf = []
for i in range(data['title'].size):
    ocnf.append(random.randint(0,3))
ocns=[]
for i in ocnf:
    ocns.append(ocn[i])
df4 = df3.assign(Occasion = ocns)
df4.to_excel(updt,'Sheet1')

yrf = []
for i in range(data['title'].size):
    yrf.append(random.randint(0,3))
yrs=[]
for i in yrf:
    yrs.append(year[i])
df5 = df4.assign(Year = yrs)
df5.to_excel(updt,'Sheet1')

idbf =[]
for i in range(data['title'].size):
    idbf.append(random.randint(0,49))
idbs=[]
for i in idbf:
    idbs.append(idb[i])
df6 = df5.assign(IMDB = idbs)
df6.to_excel(updt,'Sheet1')
updt.save()