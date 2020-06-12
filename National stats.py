import requests
import time
import json 
import pandas as pd 

# for getting time stamp as parameter
def gettime():
    return int(round(time.time()*1000))

# to get the json data for the data table
def GetJsonData(DataYears,DataCode,Period):
    url = r'http://data.stats.gov.cn/easyquery.htm'
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36'}
    
    # Set keyvalue
    keyvalue={}
    # Different keyvalue for yearly, monthly or quarterly data
    keyvalue['m'] = 'QueryData'
    if Period == 'Yearly':
        keyvalue['dbcode'] = 'hgnd'
    elif Period == 'Monthly':
        keyvalue['dbcode'] = 'hgyd'
    elif Period == 'Quarterly':
        keyvalue['dbcode'] = 'hgjd'
    else:
        print('Wrong Period, please input Yearly, Monthly or Quarterly')
    keyvalue['rowcode'] = 'zb'
    keyvalue['colcode'] = 'sj'
    keyvalue['wds'] = '[]'
    # DataCode defines what table to extract
    keyvalue['dfwds'] = '[{"wdcode":"zb","valuecode":"' + DataCode +'"}]'
    keyvalue['k1'] = str(gettime())
    keyvalue['h']=1
    # Start a session
    s = requests.session()
    if DataYears == 10 or DataYears==13:

        response = s.get(url,headers=headers, params = keyvalue)
        

    elif DataYears == 20 or DataYears == 36:# Get the 20 years or 36 months data
        response = s.get(url,headers=headers, params = keyvalue)
        if DataYears ==20:
            keyvalue['dfwds']='[{"wdcode":"sj","valuecode":"LAST20"}]'
        else:
            keyvalue['dfwds']='[{"wdcode":"sj","valuecode":"LAST36"}]'
        keyvalue['k1']=str(gettime())
        # parameter of 'h' is not needed here
        keyvalue.pop('h')
        # s.cookies.update()
        response = s.get(url,headers=headers, params = keyvalue)
    else:
        print('Error of number of years')
    JsonData = dict(json.loads(response.text))['returndata']
    return JsonData

# convert the json data into pandas dataframe
def ExtratTable(JsonData,DataYears):
    datanodes = JsonData['datanodes']
    wdnodes = JsonData['wdnodes']
    items = wdnodes[0]['nodes']
    itemlist = []
    itemUnit=[]
    for i in range(0,len(items)):
        itemlist.append(items[i]['cname'])
        itemUnit.append(items[i]['unit'])
    
    data=[]
    year=[]
    code = [] 
    for i in range(0,len(datanodes)):
        data.append(datanodes[i]['data']['strdata'])
        year.append(datanodes[i]['wds'][1]['valuecode'])
        code.append(datanodes[i]['wds'][0]['valuecode'])  

    # Convert the string data into float data
    for i in range(0,len(data)):
        try:
            data[i]=float(data[i])
        except:
            continue
    
    # Rebuild the table from json file
    table = pd.DataFrame({'Items':itemlist, 'Unit':itemUnit})
    for i in range(0,DataYears):
        table[str(year[i])]=0    
    for i in range(0,len(itemlist)):
        table.iloc[i,2:]=data[i*DataYears:(i+1)*DataYears]
    
    # Reverse the order of data
    Table_Reorder=table[['Items','Unit']]
    for i in range(2,len(table.columns)):
        Table_Reorder = Table_Reorder.join(table.iloc[:,[len(table.columns)-i+1]])

    Table_Reorder.set_index('Items', inplace=True)

    return Table_Reorder 

def main():
    # define the table for yearly data
    DataCodeYearly ={'Energy Prod':'A070B',
                'Oil Balance':'A070Q',
                'Gas Balance':'A0710',
                'Crude Balance':'A070U',
                'Energy Consumption':'A070E',
                'Energy Import':'A0707',
                'Energy Investment':'A070A'
                }
    # define the table for monthly data
    DataCodeMonthly ={
                'Crude Montly Prod':'A030102',
                'Gas Monthly Prod':'A030103',
                'CBM Monthly Prod':'A030104',
                'LNG prod':'A030105',
                'PMI Index':'A0B03'
                }
    
    # set the length of data: 10 or 20 years, 13 or 36 months
    DataYears = 20
    DataMonths = 36
    # Write the tables into excel
    with pd.ExcelWriter('National Stats Data.xlsx') as writer:
        for i in DataCodeYearly:
            Table = ExtratTable(GetJsonData(DataYears,DataCodeYearly[i],'Yearly'),DataYears)
            Table.to_excel(writer,sheet_name=i)
        for i in DataCodeMonthly:
            Table = ExtratTable(GetJsonData(DataMonths,DataCodeMonthly[i],'Monthly'),DataMonths)
            Table.to_excel(writer,sheet_name=i)
        

if __name__ == '__main__':
    main()    
    

