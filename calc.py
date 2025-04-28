import pandas as pd  
name = input("do you want the discription of the tile in the results Y/N ")
if name == "Y":
    index =1
else:
    index =0
all_sheets = pd.read_excel('carcosone_ranking_matrix  (8).xlsx', sheet_name=None)
for sheet in all_sheets :
    if sheet != "Key":
        all_sheets[sheet] = all_sheets[sheet].drop(all_sheets[sheet].columns[0], axis=1)
df_key = all_sheets["Key"]
df_connections = all_sheets["Unique connections"]
df_road = all_sheets["Road conections"]
df_city = all_sheets["City conections"]
df_roadp = all_sheets["Road Posible"]
df_cityp = all_sheets["City Posible"]
df_church = all_sheets["monistariy connections"]
df_sheid = all_sheets["Sheild Matrix"]
df_grass = all_sheets["Grass posible"]
cityM = df_city.to_numpy()
connectionsM = df_connections.to_numpy()
roadM = df_road.to_numpy()
roadPM = df_roadp.to_numpy()
cityPM = df_cityp.to_numpy()
curchM = df_church.to_numpy()
sheildM = df_sheid.to_numpy()
grassM = df_grass.to_numpy()
R = (roadPM) @ roadM * 2
C = (cityPM@ cityM) * 2 + (sheildM) * 2
M =  curchM.T * 2
G = (grassM *2) + (grassM  * 2) + (sheildM) * 2


calcM = ((R + M + C+ G) / 6)* connectionsM
answer = list()
df =pd.DataFrame((roadPM))
df_city.index = df_key.iloc[:, 1]   # for row names
df_city.columns = df_key.iloc[:, 1]
final = []

for row in range(len(calcM)):
    for col in range(row , len(calcM)):  # only look at col > row
        val1 = calcM[row][col]
        val2 = calcM[col][row]
        max_val = max(val1, val2)

        string = df_key.iloc[row, index] + " and " + df_key.iloc[col, index] + " score is: " + str(max_val)
        final.append(string)

# Sort in descending order by the number at the end
final_sorted = sorted(final, key=lambda x: float(x.split(":")[-1].strip()), reverse=True)
file = open("results.txt","w")
# Print the sorted list
for item in final_sorted:
    file.write(item+"\n")

file.close()
