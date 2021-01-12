#Before you run this file you should have the BUD.xlsx file in the same folder and a folder named "Individual"
import pandas as pd 
import numpy as np
import glob

Alexandra = ['Tim Hatcher','Conor Gribbin','Adriaan Welgraven','Flavio La Terza','Ouiam Haddou','Hamza BOULLOUS','Girish Kini']
Veronica =  ['Ayse Feyzanur Ermurat','Raffaele Talarico','Daniele Sarno','Francesc Callejas','Francesco Bastianello','Hur Bakan','Jorge Sanchez','Josefina Furrer','Josep Martin','Luca Fatigati','Marco De Rossi','Selene Hernandez','Raquel Lopez','Jonathan Soret']
Alaa =      ['Abdollah MOSTAFA','Ahmed ABDELAAL','Mariam YEHIA','Mohamed SADAT','Monica YOUSSEF','Obada SAYED','Samar ALEZABY','Clotilde Demeyer','Ilias Bensaid','Vianney Dufour','Kim Do Tri']

quarters = ['2020-Q1','2020-Q2','2020-Q3','2020-Q4','2021-Q1','2021-Q2']

names = Alaa + Alexandra + Veronica

def shift(df, name):
    """function that shifts columns, needed to make things look good when putting in excel"""
    df[name] = df.index
    return df[list(df.columns)[-1:] + list(df.columns)[:-1]]

def date_format(series):
    """Changes the format to 'YYYY-Q' from YYYYQ"""
    return series.fillna(0).apply(lambda x: str(int(x))[:4] + '-Q' + str(int(x))[-1] if (x != 0) else ' ')

#Open BUD file
bud = pd.read_excel(glob.glob('*BUD' + '*.xlsx')[0])

#Open Pipeline supprt reports
OP = pd.read_csv(glob.glob('*(OP)' + '*.csv')[0])
CL = pd.read_csv(glob.glob('*(CL)' + '*.csv')[0])

#Open investment Resource report 
IR = pd.read_csv(glob.glob('*Employee' + '*.csv')[0],header= 1, low_memory = False)
cols = IR.columns[-1:]
IR.rename(columns={cols[0]:"Time Days",}, inplace = True)
IR['Investment Quarter'] = date_format(IR['Investment Quarter'])

#Open Missing hours report
MH = pd.read_csv(glob.glob('*Missing' + '*.csv')[0], header=1)
cols = MH.columns[-1:]
MH.rename(columns={cols[0]:"Missing Hours"}, inplace = True)

#Open Activity analysis/ Time record analysis report,read data and sort table columns etc
df = pd.read_csv(glob.glob('*Time Recorded' + '*.csv')[0],header=1)
cols = df.columns[-3:]
df.rename(columns={cols[0]:"Actual Capacity Days",
                   cols[1]:"Time Recorded Days",
                   cols[2]:"Utilized Days",}, inplace = True)

#Create a table with multiindex that has every name for every quarter
AA = pd.DataFrame(index=pd.MultiIndex.from_product([np.sort(df["Year-Quarter YYYY-'Q'Q"].unique()),df['Employee Name'].unique()], names=['Quarter', 'Employee Name']))
#Get Utilization rate
UTI = df.groupby(["Year-Quarter YYYY-'Q'Q",'Employee Name'])[['Actual Capacity Days','Utilized Days']].sum()
AA['Utilization'] = UTI['Utilized Days'] / UTI['Actual Capacity Days']
#Get the deal support rate
task_types = ['Opportunity Support - Prep','Opportunity Support - CF','RFx']
AA['Deal support Rate'] = df[np.isin(df['Task Type'],task_types)].groupby(["Year-Quarter YYYY-'Q'Q",'Employee Name'])['Time Recorded Days'].sum() / UTI['Actual Capacity Days']
#Get the demand generation/bussiness development rate
task_types = ['Business Development - Prep','Business Development - CF']
AA['Demand Generation']  = df[np.isin(df['Task Type'],task_types)].groupby(["Year-Quarter YYYY-'Q'Q",'Employee Name'])['Time Recorded Days'].sum() / UTI['Actual Capacity Days']
AA.reset_index(inplace=True)

#Sort Pipeline support report, drop last row, combine them together etc etc
OP.drop(OP.tail(1).index,inplace=True)
CL.drop(CL.tail(1).index,inplace=True)
OP['nature'] = 'on prem'
CL['nature'] = 'cloud'
#set index so that we can concatenate
CL.set_index(np.arange(len(OP.index),len(OP.index)+len(CL.index)),inplace=True)
#column names are different (capital C)
OP.rename(columns={'DRM category':'DRM Category'}, inplace=True)
OPCL = pd.concat([OP,CL], ignore_index=True, sort=False)

#there is probably better way of doing this but it works :)
def make_all_float(column):
    L = []
    for a,i in column.fillna(0).iteritems():
        if type(i) == str:
            i = i.replace(',', '')
        elif type(i) == int:
            i = int(i)
        L.append(float(i))
    return L

OPCL['kEUR'] = make_all_float(OPCL['kEUR']) 
OPCL['ACV kEUR'] = make_all_float(OPCL['ACV kEUR'])
OPCL['value'] = OPCL['kEUR'] + OPCL['ACV kEUR']

#Build the "universal" bud est in and expected revenue that will be seen in all dashboards
onprem = pd.DataFrame(columns=quarters,index=['Budget','Estimated in','Budget Achievement','Presales coverage'])
cloud = pd.DataFrame(columns=quarters,index=['Budget','Estimated in','Budget Achievement','Presales coverage'])

bud = bud.set_index('Unnamed: 0').transpose()

cloud.iloc[0] = bud.iloc[0]
cloud.iloc[1] = bud.iloc[1]
cloud.iloc[3] = bud.iloc[2]
onprem.iloc[0] = bud.iloc[3]
onprem.iloc[1] = bud.iloc[4]
onprem.iloc[3] = bud.iloc[5]

for column in onprem.columns:
    try:
        onprem[column].iloc[2] = onprem[column].iloc[1] / onprem[column].iloc[0]
        onprem[column][2:] = onprem[column][2:].astype(float).apply(lambda x:np.round(x,decimals=2))
        onprem[column][2:] = (onprem[column][2:]*100).apply(lambda x: str(x)[:3] + '%')
    except:pass

for column in cloud.columns:
    try:
        cloud[column].iloc[2] = cloud[column].iloc[1] / cloud[column].iloc[0]
        cloud[column][2:] = cloud[column][2:].astype(float).apply(lambda x:np.round(x,decimals=2))
        cloud[column][2:] = (cloud[column][2:]*100).apply(lambda x: str(x)[:3] + '%')
    except:pass

onprem = shift(onprem, 'All EMEA - All LoB - Forecast vs Actuals - ON PREM')
cloud = shift(cloud, 'All EMEA - All LoB - Forecast vs Actuals - CLOUD')

writer = pd.ExcelWriter(f'Individual\Grouped.xlsx',engine='xlsxwriter')   
workbook = writer.book

for name in names:
    try:

        #function that matches all columns with matching column names at index in 2 DataFrames
        def match_up(df1,df2):
            for q in df1.columns:
                for nq in df2.columns:
                    if q == nq:
                        df1[q] = df2[nq]

        #get the expected pipeline revenue
        DRM = np.array(['Probable', 'Upside', 'Committed'])
        pipeline_calc = OPCL[(OPCL['Name'] == name) & (np.isin(OPCL['DRM Category'],DRM))].groupby(['Closing Quarter','nature'])['value'].sum().unstack(0)
        pipeline = pd.DataFrame(columns=quarters,index=['on prem','cloud','total'])
        match_up(pipeline,pipeline_calc)
        pipeline.fillna(0, inplace=True)
        pipeline.iloc[2] = pipeline.iloc[0] + (pipeline.iloc[1] *2.5)

        #get the committed contribution revenue
        DRM = np.array(['Booked/Won'])
        contribution_calc = OPCL[(OPCL['Name'] == name) & (np.isin(OPCL['DRM Category'],DRM))].groupby(['Closing Quarter','nature'])['value'].sum().unstack(0)
        contribution = pd.DataFrame(columns=quarters,index=['on prem','cloud','total'])
        match_up(contribution,contribution_calc)
        contribution.fillna(0, inplace=True)
        contribution.iloc[2] = contribution.iloc[0] + (contribution.iloc[1] *2.5)

        #make an investment Resource sorted by name
        IRN = IR[IR['Employee Name'] == name]

        try:
            #Find top countries where time is spent and sort them by current quarter and % out of total time spent.
            countries = pd.DataFrame(IRN.groupby(['Investment Quarter', 'MU'])['Time Days'].sum()).unstack(0)
            countries.columns = countries.columns.droplevel()
            countries.sort_values(by=countries.columns[0], ascending=False, inplace=True)
            countries = countries / countries.sum(axis=0) * 100
            top_countries = pd.DataFrame(columns=quarters,index=countries.index)
            match_up(top_countries, countries)

            #display as percantages (string format)
            for quarter in top_countries.columns:
                top_countries[quarter] = top_countries[quarter].fillna(0).apply(lambda x: str(x)[:3] + '%' if (x != 0) else ' ')

        except: 
            top_countries = pd.DataFrame(columns=quarters) 
            print(f'{name} not found in investment resource report')

        #make a OPCL based on name
        OPCLN = OPCL[OPCL['Name'] == name].copy()

        #Get top 3 lost Opportunity owners
        top_lost = pd.DataFrame(columns=quarters,index=['top1','top2','top3'])
        lost = OPCLN[OPCLN['DRM Category'] == 'Disc/Lost']
        lost = lost.groupby(['Closing Quarter','Opportunity Owner'])['Opp ID'].nunique().unstack(0)
        for quarter in quarters:
            try:
                x = 3 - len(np.array(lost.sort_values(by=quarter, ascending=False).index[:3]))
                top_lost[quarter] = np.pad(np.array(lost.sort_values(by=quarter, ascending=False).index[:3]),(0,x),'constant', constant_values=('-', ' '))
            except: pass
        
        #Get top 3 won opportunity owners
        top_won = pd.DataFrame(columns=quarters,index=['top 1', 'top 2', 'top 3'])
        won = OPCLN[OPCLN['DRM Category'] == 'Booked/Won']
        won = won.groupby(['Closing Quarter','Opportunity Owner'])['Opp ID'].nunique().unstack(0)
        for quarter in quarters:
            try:
                x = 3 - len(np.array(won.sort_values(by=quarter, ascending=False).index[:3]))
                top_won[quarter] = np.pad(np.array(won.sort_values(by=quarter, ascending=False).index[:3]),(0,x),'constant', constant_values=('-', ' '))
            except: pass
        
        #get deal impact
        supported = pd.DataFrame(IRN.groupby(['Investment Quarter'])['Opportunity Id'].nunique()).transpose()
        touched = pd.DataFrame(IRN.groupby(['Investment Quarter'])['Opp Global Ultimate Name'].nunique()).transpose()
        customer_facing_filter = np.array(['Business Development - CF','Opportunity Support - CF','Consumption & Renewal - CF'])
        customer_facing = IRN[np.isin(IRN['Task Type Desc'], customer_facing_filter)].groupby('Investment Quarter')['Task Type Desc'].count()

        DRM = ['Booked/Won']
        nr_won_deals = pd.DataFrame(OPCL[(OPCL['Name'] == name) & (np.isin(OPCL['DRM Category'],DRM))].groupby(['Closing Quarter'])['Opp ID'].nunique()).transpose()
        DRM = ['Disc/Lost']
        nr_lost_deals = pd.DataFrame(OPCL[(OPCL['Name'] == name) & (np.isin(OPCL['DRM Category'],DRM))].groupby(['Closing Quarter'])['Opp ID'].nunique()).transpose()

        #Build Deal impact frame
        deal_impact = pd.DataFrame(columns=quarters,index=['nr of supported deals','nr of touched customers','nr of customer interactions','nr of deals won','nr of deals lost','Win rate'])

        deal_impact.iloc[0] = supported.iloc[0]
        deal_impact.iloc[1] = touched.iloc[0]
        deal_impact.iloc[2] = customer_facing
        deal_impact.iloc[3] = nr_won_deals.iloc[0]
        deal_impact.iloc[4] = nr_lost_deals.iloc[0]
        deal_impact.iloc[5] = deal_impact.iloc[3] / (deal_impact.iloc[4] + deal_impact.iloc[3])
        deal_impact.iloc[5] = (deal_impact.iloc[5] * 100).fillna(0).apply(lambda x: str(x)[:3] + '%' if (x != 0) else ' ')

        productivity = pd.DataFrame(columns=quarters, index=['Utilization Rate - >75%','Deal Support Rate - >60%','Business Dev Rate - >15%','Nber of Missing Hours','Minimum MD','Maximum MD','Average MD','# Dry runs','# Reusable assets','# on-site customer meetings','# Enablement sessions','# Demand Generation events'])

        #Makes sure AA is in right order in case of weird extract (just in case)
        AA_name = AA[AA['Employee Name'] == name].set_index('Quarter').transpose()

        #put in bus dev and uti rates and turn them into percentages
        productivity.iloc[0] = (AA_name.iloc[1]*100).astype(str)
        productivity.iloc[1] = (AA_name.iloc[2]*100).astype(str)
        productivity.iloc[2] = (AA_name.iloc[3]*100).astype(str)

        #display as percantages (string format)
        for quarter in productivity.columns:
            productivity[quarter] = productivity[quarter].fillna(0).apply(lambda x: str(x)[:3] + '%' if (x != 0) else ' ')

        #Missing hours
        MH.rename(columns={"Year-Quarter YYYY-'Q'Q":"Quarter"},inplace=True)
        MH_name = MH[MH['Employee Name'] == name].set_index('Quarter').transpose()
        productivity.iloc[3] = MH_name.iloc[1]

        #min max and average MD
        name_MD = IRN.groupby(['Investment Quarter','Opp Global Ultimate Name'])['Time Days'].sum().unstack(0)
        productivity.iloc[4] = name_MD.min()
        productivity.iloc[5] = name_MD.max()
        productivity.iloc[6] = name_MD.mean()


        #Make index into columns and put it first so that index has a label
        pipeline = shift(pipeline, 'Your Current Pipeline')
        contribution = shift(contribution, 'Your Current Contribution To Closed Revenue')
        top_countries = shift(top_countries, 'Distribution of your support by MU')
        top_won = shift(top_won, 'Opp Owner where you had the most won deals')
        top_lost = shift(top_lost, 'Opp Owner where you had the most lost deals')
        deal_impact = shift(deal_impact, 'Your deal impact')
        productivity = shift(productivity, 'Your productivity')


        worksheet = workbook.add_worksheet(name)
        writer.sheets[name] = worksheet

        row = 2
        onprem.to_excel(writer,sheet_name=name,startrow=row , startcol=1, float_format="%.2f", index=False)

        row += len(onprem)+2
        cloud.to_excel(writer,sheet_name=name,startrow=row, startcol=1, float_format="%.2f", index=False)

        row += len(cloud)+2
        pipeline.to_excel(writer,sheet_name=name,startrow=row, startcol=1, float_format="%.2f", index=False)

        row += len(pipeline)+2
        contribution.to_excel(writer,sheet_name=name,startrow=row, startcol=1, float_format="%.2f", index=False)

        row += len(contribution)+2
        top_countries.to_excel(writer,sheet_name=name,startrow=row, startcol=1, float_format="%.2f", index=False)

        row += len(top_countries)+2
        top_won.to_excel(writer,sheet_name=name,startrow=row, startcol=1, index=False)
       
        row += len(top_won)+2
        top_lost.to_excel(writer,sheet_name=name,startrow=row, startcol=1, index=False)

        row += len(top_lost)+2
        deal_impact.to_excel(writer,sheet_name=name,startrow=row, startcol=1, float_format="%.2f", index=False)

        row += len(deal_impact)+2
        productivity.to_excel(writer,sheet_name=name,startrow=row, startcol=1, float_format="%.2f", index=False)
        
    except: print(f'error with {name}')

writer.save()