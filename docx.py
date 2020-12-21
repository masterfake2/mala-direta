import requests
import io
endpoint = "https://api2.docconversionapi.com/jobs/create"
headers = {
    'X-ApplicationID': 'f96659f5-bf99-4e1e-b363-49631226de82',
    'X-SecretKey': '8a530a1b-c742-4f6e-9c35-ec779d2598c3'
}
file = open("D:\Desenvolvimento\mala-direta\mala-direta-arquivos\ADONIAS DA SILVA SANTOS\CONTRATOS\ADONIAS DA SILVA SANTOS.docx", "rb")
data = {
    'outputFormat': 'pdf',
    'async': 'false',
    'conversionParameters': '{}'
}
files = {
    'inputFile': (file.read())
}
r = requests.post(url=endpoint, data=data, headers=headers, files = files)
response = r.text
print(r.text)