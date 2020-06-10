import requests
import time
import json 
import pandas as pd 

def gettime():
    return int(round(time.time()*1000))


def GetJsonData(DataYears,DataCode,Period):
    url = r'http://data.stats.gov.cn/easyquery.htm'
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36'}
    
    # Set keyvalue
    keyvalue={}
    keyvalue['m'] = 'QueryData'
    if Period == 'Yearly':
        keyvalue['dbcode'] = 'hgnd'
    elif Period == 'Monthly':
        keyvalue['dbcode'] = 'hgyd'
    elif Period == 'Quarterly':
        keyvalue['dbcode'] = 'hgjd'
    else:
        Print('Wrong Period, please input Yearly, Montly or Quarterly')
    keyvalue['rowcode'] = 'zb'
    keyvalue['colcode'] = 'sj'
    keyvalue['wds'] = '[]'
    keyvalue['dfwds'] = '[{"wdcode":"zb","valuecode":"' + DataCode +'"}]'
    keyvalue['k1'] = str(gettime())
    keyvalue['h']=1
    # Start a session
    s = requests.session()
    if DataYears == 10 or DataYears==13:

        response = s.get(url,headers=headers, params = keyvalue)
        # Change the keyvalue for 20 years of data

    elif DataYears == 20:# Get the 20 years data
        response = s.get(url,headers=headers, params = keyvalue)
        keyvalue['dfwds']='[{"wdcode":"sj","valuecode":"LAST20"}]'
        keyvalue['k1']=str(gettime())
        response = s.get(url,headers=headers, params = keyvalue)
    else:
        print('Error of number of years')
    JsonData = dict(json.loads(response.text))['returndata']
    return JsonData

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
    # data = [float(i) for i in data]
    table = pd.DataFrame({'Items':itemlist, 'Unit':itemUnit})
    for i in range(0,DataYears):
        table[str(year[i])]=0    
    for i in range(0,len(itemlist)):
        table.iloc[i,2:]=data[i*DataYears:(i+1)*DataYears]
    
    # Reverse the order of data
    Table_Reorder=table[['Items','Unit']]
    for i in range(2,len(table.columns)):
        Table_Reorder = Table_Reorder.join(table.iloc[:,[len(table.columns)-i+1]])

    return Table_Reorder 

def main():
    DataCodeYearly ={'ResourceProd':'A070B',
                'OilBalance':'A070Q',
                'GasBalance':'A0710',
                'CrudeBalance':'A070U',
                'EnergyConsumption':'A070E',
                'EnergyImport':'A0707'
                }
    DataCodeMonthly ={
                'Crude Montly Prod':'A030102',
                'Gas Monthly Prod':'A030103',
                'CBM Monthly Prod':'A030104',
                'LNG prod':'A030105'
                }
    
    DataYears = 10
    DataMonths = 13
    with pd.ExcelWriter('National Stats Data.xlsx') as writer:
        for i in DataCodeYearly:
            Table = ExtratTable(GetJsonData(DataYears,DataCodeYearly[i],'Yearly'),DataYears)
            Table.to_excel(writer,sheet_name=i)
        for i in DataCodeMonthly:
            Table = ExtratTable(GetJsonData(DataMonths,DataCodeMonthly[i],'Monthly'),DataMonths)
            Table.to_excel(writer,sheet_name=i)
        
    # Table.to_excel('National Stats.xlsx')
    # print(Table)

if __name__ == '__main__':
    main()    
    

