import pandas as pd
import math

df = pd.read_excel("sagatave_eksamenam.xlsx", sheet_name='Lapa_0')

data = df.iloc[1:]

data.iloc[:, 11] = pd.to_numeric(data.iloc[:, 11], errors='coerce') 
data.iloc[:, 13] = pd.to_numeric(data.iloc[:, 13], errors='coerce')  

filtered = data[
    (data.iloc[:, 5] == 'Korporatīvais') &  
    (data.iloc[:, 11] >= 40) &             
    (data.iloc[:, 11] <= 50)              
]

total_sum = filtered.iloc[:, 13].sum()  
result = math.floor(total_sum)

print(f"Answer: {result}")
