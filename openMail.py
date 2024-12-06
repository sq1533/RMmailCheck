import os
import sys
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import time as t
from datetime import datetime
import requests
import json
import pandas as pd
from bs4 import BeautifulSoup
with open('C:\\Users\\USER\\ve_1\\DB\\3loginInfo.json', 'r', encoding='utf-8') as f:
    login_info = json.load(f)
with open('C:\\Users\\USER\\ve_1\\DB\\restDay.json',"r") as f:
    restday = json.load(f)
tele_bot = pd.Series(login_info['RMbot'])
works_login = pd.Series(login_info['worksMail'])
#숫자 콤마넣기
def comma(x) -> None:
    return '{:,}'.format(round(x))
def read_mail(soup):
    #RM한도 증액 제외 가맹점
    ignoreName = ["이지피쥐","핀테크링크​","엘피엔지​","코리아결제시스템"]
    ignoreOrder = ["오프라인"]
    RM_market = soup.find_all('td')
    text_count = len(RM_market)
    i : int = 2
    j : int = 3
    k : int = 5
    l : int = 13
    marketID : list = []
    marketName : list = []
    marketPrice : list = []
    order : list = []
    #판다스 데이터프레임 정렬
    while l < text_count:
        marketID.append(str(RM_market[i]).replace('<td>','').replace('</td>',''))
        marketName.append(str(RM_market[j]).replace('<td>','').replace('</td>',''))
        marketPrice.append(str(RM_market[k]).replace('<td>','').replace('</td>','').replace(',',''))
        order.append(str(RM_market[l]).replace('<td>','').replace('</td>',''))
        i = i + 14
        j = j + 14
        k = k + 14
        l = l + 14
        if l >= text_count:
            break
    newdata = pd.DataFrame(data={"상점ID":marketID,"상점명":marketName,"월한도":marketPrice,"비고":order})
    #불필요 및 중복 데이터 분류
    RM_month = pd.read_json('C:\\Users\\USER\\ve_1\\DB\\rmMail.json',orient='records',dtype={'상점ID':str,'상점명':str,'월한도':str,'비고':str})
    lastID = RM_month['상점ID'].tolist()
    for n in newdata.index.tolist():
        if any(nm in str(newdata.loc[n]["상점명"]) for nm in ignoreName):
            newdata.drop(n, inplace=True)
        elif any(ord in str(newdata.loc[n]["비고"]) for ord in ignoreOrder):
            newdata.drop(n, inplace=True)
        elif any(id == str(newdata.loc[n]["상점ID"]) for id in lastID):
            newdata.drop(n, inplace=True)
        else:
            pass
    return newdata
#매월 1일 데이터 초기화
def reset() -> None:
    resets = {
        "상점ID":"T_ID",
        "상점명":"T_Name",
        "월한도":"1000000",
        "비고":"",
    }
    pd.DataFrame(resets,index=[0]).to_json('C:\\Users\\USER\\ve_1\\DB\\rmMail.json',orient='records',force_ascii=False,indent=4)
    requests.get(f"https://api.telegram.org/bot{tele_bot['token']}/sendMessage?chat_id={tele_bot['chatId']}&text=초기화_완료")
    t.sleep(61)
#페이지 로드
def getHome(page) -> None:
    t.sleep(1)
    id_box = page.find_element(By.XPATH,'//input[@id="user_id"]')
    login_button_1 = page.find_element(By.XPATH,'//button[@id="loginStart"]')
    id = works_login['id']
    ActionChains(page).send_keys_to_element(id_box, '{}'.format(id)).click(login_button_1).perform()
    t.sleep(1)
    password_box = page.find_element(By.XPATH,'//input[@id="user_pwd"]')
    login_button_2 = page.find_element(By.XPATH,'//button[@id="loginBtn"]')
    password = works_login['pw']
    ActionChains(page).send_keys_to_element(password_box, '{}'.format(password)).click(login_button_2).perform()
    t.sleep(1)
def newMail(page) -> None:
    page.get("https://mail.worksmobile.com/#/my/102")
    t.sleep(2)
    mailHome_soup = BeautifulSoup(page.page_source,'html.parser')
    if mailHome_soup.find('li', attrs={'class':'notRead'}) != None:
        newMail = page.find_element(By.XPATH,"//li[contains(@class,'notRead')]//div[@class='mTitle']//strong[@class='mail_title']")
        ActionChains(page).click(newMail).perform()
        t.sleep(1)
        mail_soup = BeautifulSoup(page.page_source,'html.parser')
        if read_mail(mail_soup).empty:
            tell = "{일}일 {시간}시 증액 필요 가맹점 없음".format(일=datetime.now().day,시간=datetime.now().hour)
            requests.get(f"https://api.telegram.org/bot{tele_bot['token']}/sendMessage?chat_id={tele_bot['chatId']}&text={tell}")
        else:
            for update in read_mail(mail_soup).index.tolist():
                tell = '{일}일 {시간}시 {상점명}[{상점ID}] 한도 증액필요\n월한도 {한도}원 / 증액 {증액}원'.format(
                    일=datetime.now().day,
                    시간=datetime.now().hour,
                    상점명=read_mail(mail_soup).loc[update]["상점명"],
                    상점ID=read_mail(mail_soup).loc[update]["상점ID"],
                    한도=comma(int(read_mail(mail_soup).loc[update]["월한도"])),
                    증액=comma(int(read_mail(mail_soup).loc[update]["월한도"])*120/100))
                requests.get(f"https://api.telegram.org/bot{tele_bot['token']}/sendMessage?chat_id={tele_bot['chatId']}&text={tell}")
                RM_month = pd.read_json('C:\\Users\\USER\\ve_1\\DB\\rmMail.json',orient='records',dtype={'상점ID':str,'상점명':str,'월한도':str,'비고':str})
                if update == read_mail(mail_soup).index.tolist()[-1]:
                    resurts = pd.concat([RM_month,read_mail(mail_soup)],ignore_index=True)
                    resurts.to_json('C:\\Users\\USER\\ve_1\\DB\\rmMail.json',orient='records',force_ascii=False,indent=4)
                else:pass
    else:pass
def emailClick(page):
    page.get("https://mail.worksmobile.com/#/my/102")
    t.sleep(2)
    mailHome_soup = BeautifulSoup(page.page_source,'html.parser')
    if mailHome_soup.find('li', attrs={'class':'notRead'}) != None:
        newMail = page.find_element(By.XPATH,"//li[contains(@class,'notRead')]//div[@class='mTitle']//strong[@class='mail_title']")
        ActionChains(page).click(newMail).perform()
    else:pass
workTime = ["08:00","10:00","12:00","14:00","16:00"]
restTime = ["00:00","02:00","04:00","06:00","18:00","20:00","22:00"]
def main():
    while True:
        try:
            if datetime.now().strftime('%d %H:%M') == "01 01:00":
                reset()
            else:pass
            if datetime.now().strftime('%d') in restday[datetime.now().strftime('%m')]:
                if datetime.now().strftime('%H:%M') in list(set(workTime)|set(restTime)):
                    options = webdriver.ChromeOptions()
                    options.add_argument("--headless")
                    options.add_argument('--disable-gpu')
                    options.add_argument('--disable-extensions')
                    options.add_argument('--blink-settings=imagesEnabled=false')
                    driver = webdriver.Chrome(options=options)
                    driver.get("https://mail.worksmobile.com/")
                    getHome(driver)
                    for i in range(10):
                        newMail(driver)
                        t.sleep(3)
                    driver.quit()
                    t.sleep(3000)
                else:pass
            else:
                if datetime.now().strftime('%H:%M') in workTime:
                    options = webdriver.ChromeOptions()
                    options.add_argument("--headless")
                    options.add_argument('--disable-gpu')
                    options.add_argument('--disable-extensions')
                    options.add_argument('--blink-settings=imagesEnabled=false')
                    driver = webdriver.Chrome(options=options)
                    driver.get("https://mail.worksmobile.com/")
                    getHome(driver)
                    for i in range(10):
                        emailClick(driver)
                        t.sleep(3)
                    driver.quit()
                    t.sleep(3000)
                elif datetime.now().strftime('%H:%M') in restTime:
                    options = webdriver.ChromeOptions()
                    options.add_argument("--headless")
                    options.add_argument('--disable-gpu')
                    options.add_argument('--disable-extensions')
                    options.add_argument('--blink-settings=imagesEnabled=false')
                    driver = webdriver.Chrome(options=options)
                    driver.get("https://mail.worksmobile.com/")
                    getHome(driver)
                    for i in range(10):
                        newMail(driver)
                        t.sleep(3)
                    driver.quit()
                    t.sleep(3000)
                else:pass
            t.sleep(0.5)
        except Exception:
            t.sleep(2)
            os.execl(sys.executable, sys.executable, *sys.argv)
if __name__ == "__main__":main()