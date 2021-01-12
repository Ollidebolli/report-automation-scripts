import numpy as np
import pandas as pd
import glob

#open current report (should set header to row to as first row is usually empty)

On_prem = pd.read_csv('Details View On Prem.csv', header=1, thousands=',')
Cloud = pd.read_csv('Details View Cloud.csv', header=1, thousands=',')

OP_new = On_prem[(On_prem['IAC'] == 'GB - Lower') & (On_prem['Distribution Channel'] != 'OEM Outbound direct')].copy()
C_new = Cloud[(Cloud['IAC'] == 'GB - Lower') & (Cloud['Distribution Channel'] != 'OEM Outbound direct')].copy()

#if there is no zero before op Id add it and set index as opp id
if str(OP_new['Opp ID (Hyperlink)'][OP_new.index[0]])[0] != '0':
    OP_new['Opp ID (Hyperlink)'] = OP_new['Opp ID (Hyperlink)'].apply(lambda x: str(x).zfill(len(str(x))+1))
    OP_new['Opp ID (Hyperlink)'] = OP_new['Opp ID (Hyperlink)'].apply(lambda x: x[:-2])
if str(C_new['Opp ID (Hyperlink)'][C_new.index[0]])[0] != '0':
    C_new['Opp ID (Hyperlink)'] = C_new['Opp ID (Hyperlink)'].apply(lambda x: str(x).zfill(len(str(x))+1))
    C_new['Opp ID (Hyperlink)'] = C_new['Opp ID (Hyperlink)'].apply(lambda x: x[:-2])

#set index for easy joining and remove deuplicate axis
OP_new.set_index('Opp ID (Hyperlink)', inplace=True)
C_new.set_index('Opp ID (Hyperlink)', inplace=True)
OP_new = OP_new[~OP_new.index.duplicated()]
C_new = C_new[~C_new.index.duplicated()]

#Get the presales investment details
df = pd.read_csv(glob.glob('Employee* *.csv')[0],header= 1)
cols = df.columns[-1:]
df.rename(columns={cols[0]:"Time Days",}, inplace = True)

Resource = df

Managers = np.array(['Veronica Bastianon', 'Alaa ELGANAGY', 'Alexandra Jovovic'])
Nature = np.where(np.isin(Resource['Manager Name'], Managers) == True, 'GB Presales','Field Presales')
Resource['NATURE'] = Nature

#Create a table thats sorts working hours based on field or GB
new_IR = Resource.groupby(['Opportunity Id','NATURE'])['Time Days'].sum().unstack()

# check if first is 0 else change
if str(new_IR.index[0])[0] != '0':
    new_IR.index = pd.Series(new_IR.index).apply(lambda x: str(x).zfill(len(str(x))+1))
    
new_IR = new_IR[~new_IR.index.duplicated()]

POP = pd.read_csv('Presales Individual lvl (OP).csv')
PCL = pd.read_csv('Presales Individual lvl (CL).csv')

POP_new = POP[['Name','Opp ID']].copy()
PCL_new = PCL[['Name','Opp ID']].copy()

#drop NA rows (no opp id values) to add 0 if first is not zero, set index to opp id, remove duplicates
PCL_new.dropna(subset=['Opp ID'], inplace=True)
POP_new.dropna(subset=['Opp ID'], inplace=True)
if str(POP_new['Opp ID'][0])[0] != '0':
    POP_new['Opp ID'] = POP_new['Opp ID'].apply(lambda x: str(x).zfill(len(str(x))+1))   
if str(PCL_new['Opp ID'][0])[0] != '0':
    PCL_new['Opp ID'] = PCL_new['Opp ID'].apply(lambda x: str(int(x)).zfill(10))    
POP_new.set_index('Opp ID', inplace=True)
PCL_new.set_index('Opp ID', inplace=True)
POP_new = POP_new[~POP_new.index.duplicated()]
PCL_new = PCL_new[~PCL_new.index.duplicated()]

#Get Net new names
NNN = pd.read_csv('NNN list.csv').set_index('Opp ID')

#Open old deal execution report
print('reading old report')
Oldxl = pd.ExcelFile('old_report.xlsx')
Old_prem = pd.read_excel(Oldxl, 'ON PREM')
Old_cloud = pd.read_excel(Oldxl, 'CLOUD')

#check if zero set opp id as index
if str(Old_prem['Opp ID (Hyperlink)'][0])[0] != '0':
    Old_prem['Opp ID (Hyperlink)'] = Old_prem['Opp ID (Hyperlink)'].apply(lambda x: str(x).zfill(len(str(x))+1))
if str(Old_cloud['Opp ID (Hyperlink)'][0])[0] != '0':
    Old_cloud['Opp ID (Hyperlink)'] = Old_cloud['Opp ID (Hyperlink)'].apply(lambda x: str(x).zfill(len(str(x))+1))

Old_prem.set_index('Opp ID (Hyperlink)', inplace=True)
Old_cloud.set_index('Opp ID (Hyperlink)', inplace=True)

Old_prem = Old_prem[~Old_prem.index.duplicated()]
Old_cloud = Old_cloud[~Old_cloud.index.duplicated()]

#create the final dataframes that will be exported and set index so that things will match up!
FINAL_ON_PREM = pd.DataFrame(index=OP_new.index)
FINAL_CLOUD = pd.DataFrame(index=C_new.index)

#match from Deal execution report
FINAL_ON_PREM['Region L3']          = OP_new['Region L3']
FINAL_ON_PREM['Company Name']       = OP_new['Company Name']
FINAL_ON_PREM['Deal Size (EUR)']    = OP_new['Deal Size (EUR)']
FINAL_ON_PREM['Amount']             = OP_new['Amount']
FINAL_ON_PREM['Quarter']            = OP_new['Quarter']
FINAL_ON_PREM['Opp Close Date']     = OP_new['Opp Close Date']
FINAL_ON_PREM['Opp Description']    = OP_new['Opp Description']
FINAL_ON_PREM['Opp ID (Hyperlink)'] = OP_new.index
FINAL_ON_PREM['FC Category']        = OP_new['FC Category']
#continue from deal exectuion
FINAL_ON_PREM['Distribution Channel'] = OP_new['Distribution Channel']
FINAL_ON_PREM['Channel Partner']      = OP_new['Channel Partner']
FINAL_ON_PREM['Opportunity Owner']    = OP_new['Opportunity Owner']
#Net New Name
FINAL_ON_PREM['Net New Name']         = NNN
#Match managers to Opportunity owner
FINAL_ON_PREM['Manager']              = ' '
#continue match from Deal execution report
FINAL_ON_PREM['Presales Lead Name']   = OP_new['Presales Lead Name']
FINAL_ON_PREM['Lead Sales Bag']       = OP_new['Lead Sales Bag']
#match from Investment resource report
FINAL_ON_PREM['Number of MD from Field Presales'] = new_IR['Field Presales']
FINAL_ON_PREM['Number of MD from GB Presales'] = new_IR['GB Presales']
#match from pipeline report
FINAL_ON_PREM['Presales Attached'] = POP_new['Name']
#match from old deal exectuion report
FINAL_ON_PREM['Tier'] = Old_prem['Tier']
FINAL_ON_PREM['To be checked by'] = Old_prem['To be checked by']
FINAL_ON_PREM['Comment'] = Old_prem['Comment']


#match from Deal execution report
FINAL_CLOUD['Region L3'] = C_new['Region L3']
FINAL_CLOUD['Company Name'] = C_new['Company Name']
FINAL_CLOUD['Deal Size (EUR)'] = C_new['Deal Size (EUR)']
FINAL_CLOUD['Amount'] = C_new['Amount']
FINAL_CLOUD['Quarter'] = C_new['Quarter']
FINAL_CLOUD['Opp Close Date'] = C_new['Opp Close Date']
FINAL_CLOUD['Opp Description'] = C_new['Opp Description']
FINAL_CLOUD['Opp ID (Hyperlink)'] = C_new.index
FINAL_CLOUD['FC Category'] = C_new['FC Category']
#Continue from deal exectuion
FINAL_CLOUD['Distribution Channel'] = C_new['Distribution Channel']
FINAL_CLOUD['Channel Partner'] = C_new['Channel Partner']
FINAL_CLOUD['Opportunity Owner'] = C_new['Opportunity Owner']
#Net New Name
FINAL_CLOUD['Net New Name'] = NNN
#Match managers to Opportunity owner
FINAL_CLOUD['Manager'] = ' '
#continue match from Deal execution report
FINAL_CLOUD['Presales Lead Name'] = C_new['Presales Lead Name']
FINAL_CLOUD['Lead Sales Bag'] = C_new['Lead Sales Bag']
#match from Investment resource report
FINAL_CLOUD['Number of MD from Field Presales'] = new_IR['Field Presales']
FINAL_CLOUD['Number of MD from GB Presales'] = new_IR['GB Presales']
#match from pipeline report
FINAL_CLOUD['Presales Attached'] = PCL_new['Name']
#match from old deal exectuion report
FINAL_CLOUD['Tier'] = Old_cloud['Tier']
FINAL_CLOUD['To be checked by'] = Old_cloud['To be checked by']
FINAL_CLOUD['Comment'] = Old_cloud['Comment']

#remove N/A and fill with space
FINAL_ON_PREM.fillna(value=' ', inplace=True)
FINAL_CLOUD.fillna(value=' ', inplace=True)

# Write each dataframe to a different worksheet.
print('Saving new report')
writer = pd.ExcelWriter('AUTO_REPORT.xlsx', engine='xlsxwriter')
FINAL_ON_PREM.to_excel(writer, sheet_name='ON PREM', index=True)
FINAL_CLOUD.to_excel(writer, sheet_name='CLOUD', index=True)
writer.save()