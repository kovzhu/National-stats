import requests
import time
import json 
import pandas as pd 

def gettime():
    return int(round(time.time()*1000))


def GetJsonData():
    url = r'http://data.stats.gov.cn/easyquery.htm'
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36'}
    
    # Set keyvalue
    keyvalue={}
    keyvalue['m'] = 'QueryData'
    keyvalue['dbcode'] = 'hgnd'
    keyvalue['rowcode'] = 'zb'
    keyvalue['colcode'] = 'sj'
    keyvalue['wds'] = '[]'
    keyvalue['dfwds'] = '[{"wdcode":"zb","valuecode":"A070B"}]'
    keyvalue['k1'] = str(gettime())
    keyvalue['h']=1
    # Start a session
    s = requests.session()
    response = s.get(url,headers=headers, params = keyvalue)
    # Change the keyvalue for 20 years of data
    keyvalue['dfwds']='[{"wdcode":"sj","valuecode":"LAST20"}]'
    # Get the 20 years data
    response = s.get(url,headers=headers, params = keyvalue)
    JsonData = dict(json.loads(response.text))['returndata']
    return JsonData

def ExtratTable(JsonData):
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
    table = pd.DataFrame({'Items':itemlist, 'Unit':itemUnit})
    for i in range(0,20):
        table[str(year[i])]=0    
    for i in range(0,len(itemlist)):
        table.iloc[i,2:]=data[i*20:(i+1)*20]

    return table 

def main():
    JsonData = GetJsonData()
    Table = ExtratTable(JsonData)
    Table.to_excel('National Stats.xlsx')

if __name__ == '__main__':
    main()    
    

