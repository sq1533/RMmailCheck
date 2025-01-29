import os
import json
import pandas as pd
import time as t
from datetime import datetime
import requests
import sys
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup

loginPath = os.path.join(os.path.dirname(__file__),"..","loginInfo.json")
restDayPath = os.path.join(os.path.dirname(__file__),"..","restDay.json")
rmMailPath = os.path.join(os.path.dirname(__file__),"DB","rmMail.json")
with open(loginPath, 'r', encoding='utf-8') as f:
    login_info = json.load(f)
with open(restDayPath,"r") as f:
    restday = json.load(f)
works_login = pd.Series(login_info['worksMail'])
tele_bot = pd.Series(login_info['RMbot'])

#숫자 콤마넣기
def comma(x):
    return '{:,}'.format(round(x))

#매월 1일 데이터 초기화
def reset() -> None:
    resets = {
        "상점ID":"T_ID",
        "상점명":"T_Name",
        "월한도":"1000000",
        "비고":""
    }
    pd.DataFrame(resets,index=[0]).to_json(rmMailPath,orient='records',force_ascii=False,indent=4)
    requests.get(f"https://api.telegram.org/bot{tele_bot['token']}/sendMessage?chat_id={tele_bot['chatId']}&text=초기화_완료")
    t.sleep(61)

#RM한도 증액 가맹점 분류
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
    RM_month = pd.read_json(rmMailPath,orient='records',dtype={'상점ID':str,'상점명':str,'월한도':str,'비고':str})
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

#페이지 로드
def getHome(page) -> None:
    page.get("https://mail.worksmobile.com/")
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
    page.get("https://mail.worksmobile.com/#/my/102")

#RM메일 확인
def newMail(page) -> None:
    page.refresh()
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
            page.get("https://mail.worksmobile.com/#/my/102")
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
                RM_month = pd.read_json(rmMailPath,orient='records',dtype={'상점ID':str,'상점명':str,'월한도':str,'비고':str})
                if update == read_mail(mail_soup).index.tolist()[-1]:
                    resurts = pd.concat([RM_month,read_mail(mail_soup)],ignore_index=True)
                    resurts.to_json(rmMailPath,orient='records',force_ascii=False,indent=4)
                else:pass
            page.get("https://mail.worksmobile.com/#/my/102")
    else:
        pass

#영업일 메일 체크
def emailClick(page) -> None:
    page.refresh()
    t.sleep(2)
    mailHome_soup = BeautifulSoup(page.page_source,'html.parser')
    if mailHome_soup.find('li', attrs={'class':'notRead'}) != None:
        newMail = page.find_element(By.XPATH,"//li[contains(@class,'notRead')]//div[@class='mTitle']//strong[@class='mail_title']")
        ActionChains(page).click(newMail).perform()
        page.get("https://mail.worksmobile.com/#/my/102")
    else:
        pass

workTime = ["08:00","10:00","12:00","14:00","16:00"]
restTime = ["00:00","02:00","04:00","06:00","18:00","20:00","22:00"]
def main():
    try:
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-extensions')
        driver = webdriver.Firefox(options=options)
        getHome(driver)
        while True:
            today = datetime.now()
            #데이터 리셋
            if today.strftime('%d %H:%M') == "01 01:00":
                reset()
            else:
                pass
            #영업일 구분
            if (today.weekday() == 5) or (today.weekday() == 6) or (today.strftime('%d') in restday[today.strftime('%m')]):
                if today.strftime('%H:%M') in list(set(workTime)|set(restTime)):
                    for i in range(10):
                        newMail(driver)
                        t.sleep(3)
                    t.sleep(3000)
                else:
                    pass
            else:
                if today.strftime('%H:%M') in workTime:
                    for i in range(10):
                        emailClick(driver)
                        t.sleep(3)
                    t.sleep(3000)
                elif today.strftime('%H:%M') in restTime:
                    for i in range(10):
                        newMail(driver)
                        t.sleep(3)
                    t.sleep(3000)
                else:
                    pass
            t.sleep(0.5)
    except Exception:
        driver.quit()
        t.sleep(2)
        os.execl(sys.executable, sys.executable, *sys.argv)

if __name__ == "__main__":
    main()