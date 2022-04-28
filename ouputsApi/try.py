import requests
import json

url = 'https://api.bademeister-jan.pro/outputs/store'




myobj = {'projectid': "123", "tabid": "songXrevXhalf",
         "data": "a123"}  # 1 sec
response = requests.post(url, data=myobj)
print(response)
