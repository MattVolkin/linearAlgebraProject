import pandas as pd 
import numpy as np
all_sheets = pd.read_excel('carcosone_ranking_matrix .xlsx', sheet_name=None)
for sheet in all_sheets :
    all_sheets[sheet] = all_sheets[sheet].drop(all_sheets[sheet].columns[0], axis=1)

df_key = all_sheets["Key"]
df_connections = all_sheets["Unique connections"]
df_road = all_sheets["Road conections"]
df_city = all_sheets["City conections"]
df_grass = all_sheets["Grass conections"]
df_sheid = all_sheets["Sheild Matrix"]
cityM = df_city.to_numpy()
connectionsM = df_connections.to_numpy()
roadM = df_road.to_numpy()
grassM = df_grass.to_numpy()
sheildM = df_sheid.to_numpy()
print(roadM)
calcM = (connectionsM @ roadM)#((connectionsM @ cityM)*2 )@ (sheildM*2) 
print (roadM)

