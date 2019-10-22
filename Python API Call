import requests
import ast
import openpyxl

#API Query Call
url = "http://interseller.io/api/campaigns/5d0a93c689430c7c8d4ee681/stats"

querystring = {" ":"","%20":""}

headers = {
    'x-api-key': "******************************",
    'User-Agent': "PostmanRuntime/7.15.0",
    'Accept': "*/*",
    'Cache-Control': "no-cache",
    'Postman-Token': "93967a3a-ec58-458a-a80b-4ce84cf22288,af433049-fca8-4746-a885-29d6a335fed6",
    'cookie': "__cfduid=df8197ec7026ef093e829dabb61cd18b71560796175",
    'accept-encoding': "gzip, deflate",
    'referer': "http://interseller.io/api/campaigns/5d0a93c689430c7c8d4ee681/stats?%20=",
    'Connection': "keep-alive",
    'cache-control': "no-cache"
    }

response = requests.request("GET", url, headers=headers, params=querystring)


#Testing the data call
#print(response.text)

#Transforming data into dictionary form
data=ast.literal_eval(response.text)

#Opening the workbook
wb = openpyxl.load_workbook('Campaign Progress.xlsx')

#Opening the worksheet
ws = wb['Sheet1']

#Printing to the excel sheet
ws['K3'] = data["state"]['total']
ws['K4'] = data["state"]['ongoing']
ws['K5'] = data["state"]['bounced']
ws['K6'] = data["state"]['messaged']
ws['K7'] = data["state"]['viewed']
ws['K8'] = data["state"]['visited']
ws['K9'] = data["state"]['replied']
ws['K10'] =data["state"]['booked']
ws['K11'] =data["state"]['error']

#Saving changes to the workbook 
wb.save("Campaign Progress.xlsx")

