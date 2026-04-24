import requests

TOKEN = "Wmw5xFBAyj82NOIYU2yPObLR8FtRXm6c6QgV4"

headers = {"Authorization": f"Bearer {TOKEN}"}

url = "https://api.smartsheet.com/2.0/sheets"

response = requests.get(url, headers=headers)

print("STATUS:", response.status_code)
print(response.text)
