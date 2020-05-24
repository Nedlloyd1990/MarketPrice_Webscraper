import bs4 
import requests
from urllib.request import urlopen
import xlwings as xs
import datetime
import xlsxwriter
import os



count=1
dir_path = os.path.dirname(os.path.realpath(__file__))

wb=xs.Book()
sht= wb.sheets('Sheet1')
while True:
    webData=urlopen("https://finance.yahoo.com/quote/AAPL/")
    soup=bs4.BeautifulSoup(webData.read(), 'lxml')
    sht.range("A" + str(count)).value=str(datetime.datetime.now().time())
    sht.range("B" + str(count)).value= [i.text for i in soup.find_all('span',{'class':'Trsdu(0.3s) Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(b)'})]
    count=count+1