import gspread
from gspread_formatting  import *
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import datetime

try:
    from googlesearch import search
except ImportError:
    print("No module named 'google' found")


# to search



"""keywords = ["retail pos vs cloud","how to market online covid-19 confidence","how to improve checkout process","merchandising move","more internet presence store",
           "insights from google reviews large retailer","crisis and incident management in retail","retail business during the holidays tips","how to attract customers on the street",
            "online retail presence","merchandising sayings eye level is buy level","retail search engine management","loss prevention tips","cloud based pos vs traditional pos",
            "retail loss prevention tips","what are skus","is cloud based pos software a good idea","retail merchandising vs inventory","loss prevention ideas for retail",
            "tips incrose during holiday season","marketingtactics.io","back to school shopping seo tips","thanksgiving pos","digital marketer retail store","loss prevention blog",
            "local seo for retail","father's day slow ecommerce sales","retail inventory loss prevention","loss prevention in retail stores","what is merchandising in retail",
            "how to attract customers to a store","visual merchandising vs inventory","best practices for google short name","best prasctices for google short name",
            "cloud based retail pos","halloween product display ideas","easy checkout point of sale","how to attract customers retail","google pos software",
            "retail seo optimization","how do i create a short name on google business","cloud retail pos","managing multiple retail stores","privacy management for small businesses",
            "how to attract customers in retail store","how to attract customers in shop","seo for retail","what's merchandising","what is merchandising","retail merchandise",
            "point of sale or retail store","why is pos essential in retail business","how to attract customers to shop"]
"""
keywords = ["what is merchandising","retail merchandise","point of sale or retail store","more internet presence store",
            "insights from google reviews large retailer"]



scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Python").sheet1

def dostuff(s):

    query = s
    n = 0
    x = 0

    for j in search(query, tld="co.in", num=30, stop=30, pause=4):
        n+=1
        if "takulabs" in j:
            x+=1
            print(n, s ,j)
            insertRow = [j, s, n]
            sheet.insert_row(insertRow, x+1)


    ranks = sheet.col_values(3)
    ranks.remove('Rank')
    print(ranks)



headers = ["URL", "Keyword", "Rank"]
blank = ["","",""]
sheet.insert_row(blank)
sheet.insert_row(headers)
sheet.format('A1:C1', {'textFormat': {'bold': True}})
sheet.format('A1:C1', {"backgroundColor": {"red": 0.0,"green": 1.0, "blue": 0.0}})

for i in keywords:
    dostuff(i)
    sheet.format('A2:C2', {'textFormat': {'bold': False}})
    sheet.format('A5:C5', {'textFormat': {'bold': False}})
    sheet.format('A2:C2', {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}})
    sheet.format('A3:C3', {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}})

#sheet.format('A6:C6', {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0}})





