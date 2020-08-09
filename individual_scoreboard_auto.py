import pandas as pd
import numpy as np
import glob

quarter = input('what quarter ? (YYYY-QQ) : ' )

#Open Pipeline supprt reports
OP = pd.read_csv(glob.glob('*(OP)' + '*.csv')[0])
CL = pd.read_csv(glob.glob('*(CL)' + '*.csv')[0])
#Sort Pipeline support report, drop last row, combine them together etc etc
OP.drop(OP.tail(1).index,inplace=True)
CL.drop(CL.tail(1).index,inplace=True)

OP['nature'] = 'on prem'
CL['nature'] = 'cloud'
#set index so that we can concatenate
CL.set_index(np.arange(len(OP.index),len(OP.index)+len(CL.index)),inplace=True)
#column names are different (capital C)
OP.rename(columns={'DRM category':'DRM Category'}, inplace=True)
OPCL = pd.concat([OP,CL], sort=False)

def make_all_float(column):
    L = []
    for a,i in column.fillna(0).iteritems():
        if type(i) == str:
            i = i.replace(',', '')
        elif type(i) == int:
            i = int(i)
        L.append(float(i))
    return L

OPCL['1x'] = make_all_float(OPCL['kEUR']) 
OPCL['2x'] = make_all_float(OPCL['ACV kEUR'])
OPCL.drop(['kEUR','ACV kEUR'], axis=1, inplace = True)

won = OPCL[OPCL['DRM Category'] == 'Booked/Won'].groupby(['Name','Closing Quarter'])['Opp ID'].nunique()
lost = OPCL[OPCL['DRM Category'] == 'Disc/Lost'].groupby(['Name','Closing Quarter'])['Opp ID'].nunique()

deals = pd.DataFrame(OPCL.groupby(['Name','Closing Quarter'])['1x'].sum()).join(won).rename(columns={'Opp ID':'won'}).join(lost).rename(columns={'Opp ID':'lost'}).drop('1x',axis=1)

won_revenue = OPCL[OPCL['DRM Category'] == 'Booked/Won'].groupby(['Name','Closing Quarter'])['1x','2x'].sum()

#weird solution due to indexeing issues
final = deals.join(won_revenue)
final.reset_index(inplace=True)
hello = final.copy()
final['1x'] = hello['1x'] + hello['2x']
final['2x'] = hello['1x'] + hello['2x'] * 2.5

Q = final[np.isin(final['Closing Quarter'], quarter)] 
Q = Q.groupby(['Name'])[['won','lost','1x','2x']].sum()

template = pd.read_csv('individual_scoreboard_template.csv')
template.set_index('Presales', inplace=True)

template['Revenue Supported LGB Revenue \n(OP & 1X ACV) '] = Q['1x']
template['Productivity  Supported LGB Revenue \n(OP & 2,5X ACV) '] = Q['2x']
template['# deals won'] = Q['won']
template['# deals lost or discontinued'] = Q['lost']
template['Win rate'] = template['# deals won'] / (template['# deals won'] + template['# deals lost or discontinued']) 

data = pd.read_csv('Presales Investment - Resource Details.csv')
data['Investment Quarter'] = data['Investment Quarter'].fillna(0).apply(lambda x: str(int(x))[:4] + '-Q' + str(int(x))[-1] if (x != 0) else ' ')

#get rid of rows that arent related to our managers
Managers = np.array(['Veronica Bastianon', 'Alaa ELGANAGY','Alexandra Jovovic'])
data = data[np.isin(data['Resource Manager'], Managers)]

data = data[np.isin(data['Investment Quarter'], quarter)]

customers = data.groupby('Resource Display Name')['Opportunity Id','Global Ultimate Name'].nunique()

#create a filter and get all the customer facing deals sorted by name
customer_facing_filter = np.array(['Business Development - CF','Opportunity Support - CF','Consumption & Renewal - CF'])
customer_facing = data[np.isin(data['Task Type Desc'], customer_facing_filter)].groupby('Resource Display Name')['Task Type Desc'].count()

#create a filter and get days invested
days_invested_filter = np.array(['Opportunity Support - CF','Opportunity Support - Prep','POC','RFx'])
days_invested = data[np.isin(data['Task Type Desc'], days_invested_filter)].groupby('Resource Display Name')['Activity Days'].sum()

#put in the rest of the stuff, last two are series becuase they needed differnet filtering.
template['# Deals supported']              = customers['Opportunity Id']
template['# touched customers']            = customers['Global Ultimate Name']
template['# Customer facing interactions'] = customer_facing
template['Days invested in deal support']  = days_invested

template['Average days per deal '] =  template['Days invested in deal support'] / template['# Deals supported']
template['Average days per customer'] =  template['Days invested in deal support'] / template['# touched customers']

performance = pd.read_csv('Presales Activity Analysis.csv')
performance = performance[performance['Quarter'] == quarter].set_index('Employee Full Name')

template['Utilization in %'] = performance['Utilization %']
template['Deal support in %'] = performance['Total Deal Execution Utilization']
template['Bus. Dev in %'] = performance['Business Development Utilization']

MH = pd.read_csv('Presales Missing Hours Details.csv')
MH = MH[MH['Quarter'] == quarter].set_index('Employee Name')

template['Missing hours at end of quarter'] = MH['Missing Hours']

template.to_excel(f'Auto_Scoreboard {quarter}.xlsx')