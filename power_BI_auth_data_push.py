import pandas as pd
import requests
import json
import msal


data = pd.read_excel(r'C:\Users\AD35221\OneDrive - Lumen\Documents\PON_port_forcast2024new.xlsx', sheet_name='XGS-PON Forecast')
df = pd.DataFrame(data=data)
df['OLT_LAT'] = df['OLT_LAT'].fillna(0)
df['OLT_LONG'] = df['OLT_LONG'].fillna(0)

#GET TOKEN
tenant_id ='72b17115-9915-42c0-9f1b-4f98e5a4bcd2'
authority_url = f'https://login.microsoftonline.com/{tenant_id}'
# authentication details
client_id= 'client_id' 
client_secret = 'client_sercret' # needed if authenticating with acquire_token_for_client function
username = 'username@domain' # needed if authenticating with acquire_token_by_username_password function
password = 'password' # needed if authenticating with acquire_token_by_username_password function
scope =['https://analysis.windows.net/powerbi/api/.default']

def get_token_for_client(scope):
	app = msal.ConfidentialClientApplication(client_id,authority=authority_url,client_credential=client_secret)
	result = app.acquire_token_for_client(scopes=scope)
	if 'access_token' in result:
		return(result['access_token'])
	else:
		print('Error in get_token_username_password:',result.get("error"), result.get("error_description"))

token = get_token_for_client(scope=scope)

headers = {
	'authorization': f"Bearer {token}",
    'content-type': "application/json"}

datasetId = 'b2a8be16-700d-4534-bc20-f5c4328d1060'
tableName = 'RealTimeData'
groupId = 'a2d6b97c-9a4b-4541-9eab-21b15c461d27'
#XGS
datasetId = '15084088-5e57-47e1-8a5d-041e60802577'
url = f"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/tables/{tableName}/rows"
#url = 'https://api.powerbi.com/beta/72b17115-9915-42c0-9f1b-4f98e5a4bcd2/datasets/15084088-5e57-47e1-8a5d-041e60802577/rows?experience=power-bi&key=bmW55IsCsQOjqrCTO%2BrBJbTiStYHM8ItoeHi5rP9Fyoj3%2B%2BuJDKwhjZfCYtt4a%2FWD6oSFXgZcPSzlPRH%2Fs0scg%3D%3D'
payload = df.to_json(orient='records')
parsed = json.loads(payload)
data = json.dumps(parsed, indent = 4)
#make request, clean and display response
response = requests.request("POST", url, headers=headers, data=data)
#response = requests.request("DELETE", url, headers=headers)
parsed = json.loads(response.text)
print(json.dumps(parsed, indent = 4))
