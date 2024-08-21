from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd
import time
from datetime import datetime
import re
#----------------------------------------------------------------------------------------------------------------
years=input("輸入爬蟲年份 ex:2021 2022....:")
month=input("輸入爬蟲月份 ex:01 02 03...12:")
start=int(input("要從第幾筆開始爬?"))
amount=int(input("爬幾筆資料?"))
print(datetime.now())

op=Options()
op.add_argument('--blink-settings=imagesEnabled=false')

op.chrome_executable_path="./chromedriver.exe"
driver=webdriver.Chrome(options=op)

url = 'https://isbn.ncl.edu.tw/NEW_ISBNNet/H52_BrowsingByPubMonth.php?Pact=SelectPubMonth' 
driver.get(url) 

yearSelect= Select(driver.find_element(By.NAME,'FO_選擇年份'))
yearSelect.select_by_value(years)
monthSelect= Select(driver.find_element(By.NAME,'FO_選擇月份'))
monthSelect.select_by_value(month)
driver.find_element(By.TAG_NAME,"input").click()
time.sleep(0.5)



title=list()
author=list()
publisher=list()
version=list()
callNumber=list()
topic=list()
target=list()
keyword=list()
isbn=list()
outline=list()
information=list()



page=int(start//10)
startLine=start%10
number=driver.find_element(By.XPATH,"/html/body/section/div/div/div[2]/div[2]/form[1]/div/div[1]/input")
number.clear()
if start%10==0:
    number.send_keys(f'{page}') 
else:
    number.send_keys(f'{page+1}')
driver.find_element(By.XPATH,"/html/body/section/div/div/div[2]/div[2]/form[1]/div/div[1]/a").click()

if start%10==0:
    driver.find_element(By.XPATH,f"/html/body/section/div/div/div[2]/div[2]/form[1]/div/table/tbody/tr[11]/td[3]/a").click()
else:
    driver.find_element(By.XPATH,f"/html/body/section/div/div/div[2]/div[2]/form[1]/div/table/tbody/tr[{startLine+1}]/td[3]/a").click()
time.sleep(0.5)


for i in range(amount):

    tmpList=list()

    for x in driver.find_elements(By.TAG_NAME,"tr"):
        for y in x.find_elements(By.TAG_NAME,"td"):
            tmpList.append(y.text)

    title.append(tmpList[1])
    author.append(tmpList[3])
    publisher.append(tmpList[5])
    version.append(tmpList[7])
    callNumber.append(tmpList[9])
    topic.append(tmpList[11])
    target.append(tmpList[13])
    keyword.append(tmpList[15])
    isbn.append(tmpList[16:-1])
    outline.append(tmpList[-1])
    information.append(tmpList)
    
    startData=str(driver.find_element(By.XPATH,"/html/body/section/div/div/div[2]/div[2]/form/p[1]").text)
    dataCount=re.findall(r"\d+",startData)
    dataCount=str(dataCount[0])
    print("完成第",dataCount,"筆")
    print(datetime.now())

    driver.find_element(By.XPATH,"/html/body/section/div/div/div[2]/div[2]/form/div[1]/div/a[3]/span").click()
    time.sleep(0.5)


driver.close()
#----------------------------------------------------------------------------------------------------------------

df = pd.DataFrame(
{
"書名":title,
"作者":author,
"出版機構":publisher,
"出版版次":version,
"圖書類號":callNumber,
"主題標題":topic,
"適讀對象":target,
"關鍵字詞":keyword,
"ISBN與其他資訊":isbn,
"內容簡介":outline,
"全部資訊":information
}
)

df.to_excel(f"{years+month}.xlsx",index=False)

print(datetime.now())