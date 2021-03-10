import openpyxl
import json
import requests
import datetime
import dateparser
# filejson = 'c:\\Users\\jazzt\\desktop\\NAM-IP\\Ih8ZoAQ4.json'
# filexlsx = 'c:\\Users\\jazzt\\desktop\\NAM-IP\\tasks.xlsx'
filename = "expo-micro-"+str(datetime.date.today())+".xlsx"
url ="https://trello.com/b/Ih8ZoAQ4.json"
font = openpyxl.styles.Font(size=12,
                            bold=True)

wrap = openpyxl.styles.Alignment(wrap_text=True)