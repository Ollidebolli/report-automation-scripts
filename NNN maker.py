import pandas as pd
import numpy as np
cloud = pd.read_csv('Cloud Details.csv', usecols=['Planning Entity ID','Opp ID'])
onprem = pd.read_csv('On-Prem Details.csv', usecols=['Planning Entity ID','Opp ID'])
if str(cloud['Opp ID'][cloud.index[0]])[0] != '0':
    cloud['Opp ID'] = cloud['Opp ID'].apply(lambda x: str(x).zfill(len(str(x))+1))
    cloud['Opp ID'] = cloud['Opp ID'].apply(lambda x: x[:-2])
if str(onprem['Opp ID'][onprem.index[0]])[0] != '0':
    onprem['Opp ID'] = onprem['Opp ID'].apply(lambda x: str(x).zfill(len(str(x))+1))
    onprem['Opp ID'] = onprem['Opp ID'].apply(lambda x: x[:-2])   
data = pd.concat([cloud,onprem], ignore_index=True)
PEID = pd.read_excel('PID.xlsx', usecols=['PE ID'])['PE ID']
data['NNN'] = np.where(np.isin(data['Planning Entity ID'], PEID),1,'NNN')
data.drop_duplicates('Opp ID', inplace=True)
data[data['NNN'] == 'NNN'][['Opp ID','NNN']].set_index('Opp ID').to_csv('NNN list.csv')