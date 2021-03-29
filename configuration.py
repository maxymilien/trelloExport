import openpyxl
import json
import requests
import datetime
import dateparser
filename = "stage-"+str(datetime.date.today())+".xlsx"
url ="https://trello.com/b/Ih8ZoAQ4.json"
urlStage ="https://trello.com/b/mX2aY5jC.json"
font = openpyxl.styles.Font(size=12,
                            bold=True)

wrap = openpyxl.styles.Alignment(wrap_text=True)