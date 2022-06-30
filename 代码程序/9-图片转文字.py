from aip import AipOcr
import re
APP_ID = "待补充，百度智能云账号"
API_KEY = "待补充，百度智能云账号"
SECRET_KEY = "待补充，百度智能云账号"

client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
i = open(r'C:\Users\DELL\Desktop\识别1.png','rb')
img = i.read()
# message = client.basicGeneral(img);   #通用文字识别（标准版）接口
message = client.basicAccurate(img);   #通用文字识别（高精度版）接口
message.get('words_result')
for i in message.get('words_result'):
    print(i.get('words'))
