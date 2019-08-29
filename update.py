import pandas as pd
import csv


with open('data.csv') as f:
    data = csv.reader(f)
    lines = list(data)
    for line in lines:
        line.pop(0)
    with open('data.csv','w') as g:
        writer = csv.writer(g,lineterminator='\n')
        writer.writerows(lines)
        
df = pd.read_csv('data.csv')
df.to_excel('data.xlsx',index = False)