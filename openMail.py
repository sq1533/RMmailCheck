import time as t
import requests
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
from datetime import datetime, time
options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument('--disable-gpu')
options.add_argument("--disable-javascript")
options.add_argument('--disable-extensions')
options.add_argument('--blink-settings=imagesEnabled=false')
driver = webdriver.Chrome(options=options)
with open('C:\\Users\\USER\\ve_1\\DB\\3loginInfo.json', 'r', encoding='utf-8') as f:
    login_info = json.load(f)
with open('C:\\Users\\USER\\ve_1\\DB\\6faxInfo.json', 'r',encoding='utf-8') as f:
    fax_info = json.load(f)
with open('C:\\Users\\USER\\ve_1\\DB\\8restDay.json',"r") as f:
    restday = json.load(f)
tele_bot = pd.Series(login_info['bot'])
works_login = pd.Series(login_info['worksMail'])
fax = pd.DataFrame(fax_info)
#숫자 콤마넣기
def comma(x):
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
    RM_month = pd.read_json('C:\\Users\\USER\\ve_1\\DB\\7rmMail.json',orient='records',dtype={'상점ID':str,'상점명':str,'월한도':str,'비고':str})
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
class RM:
    #매월 1일 데이터 초기화
    def reset(self):
        resets = {
            "상점ID":"T_ID",
            "상점명":"T_Name",
            "월한도":"1000000",
            "비고":"",
        }
        pd.DataFrame(resets,index=[0]).to_json('C:\\Users\\USER\\ve_1\\DB\\7rmMail.json',orient='records',force_ascii=False,indent=4)
        #텔레그램 API 전송
        requests.get(f"https://api.telegram.org/bot{tele_bot['token']}/sendMessage?chat_id={tele_bot['chatId']}&text=초기화_완료")
        t.sleep(61)
    def getHome(self):
        driver.get("https://mail.worksmobile.com/")
        t.sleep(1)
        id_box = driver.find_element(By.XPATH,'//input[@id="user_id"]')
        login_button_1 = driver.find_element(By.XPATH,'//button[@id="loginStart"]')
        id = works_login['id']
        ActionChains(driver).send_keys_to_element(id_box, '{}'.format(id)).click(login_button_1).perform()
        t.sleep(1)
        password_box = driver.find_element(By.XPATH,'//input[@id="user_pwd"]')
        login_button_2 = driver.find_element(By.XPATH,'//button[@id="loginBtn"]')
        password = works_login['pw']
        ActionChains(driver).send_keys_to_element(password_box, '{}'.format(password)).click(login_button_2).perform()
        t.sleep(1)
    def newMail(self):
        try:
            driver.get("https://mail.worksmobile.com/#/my/102")
            t.sleep(2)
            mailHome_soup = BeautifulSoup(driver.page_source,'html.parser')
            if mailHome_soup.find('li', attrs={'class':'notRead'}) != None:
                newMail = driver.find_element(By.XPATH,"//li[contains(@class,'notRead')]//div[@class='mTitle']//strong[@class='mail_title']")
                ActionChains(driver).click(newMail).perform()
                t.sleep(1)
                mail_soup = BeautifulSoup(driver.page_source,'html.parser')
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
                        RM_month = pd.read_json('C:\\Users\\USER\\ve_1\\DB\\7rmMail.json',orient='records',dtype={'상점ID':str,'상점명':str,'월한도':str,'비고':str})
                        if update == read_mail(mail_soup).index.tolist()[-1]:
                            resurts = pd.concat([RM_month,read_mail(mail_soup)],ignore_index=True)
                            resurts.to_json('C:\\Users\\USER\\ve_1\\DB\\7rmMail.json',orient='records',force_ascii=False,indent=4)
                        else:pass
                t.sleep(60*58)
            else:pass
        except TimeoutException:
            driver.quit()
            t.sleep(5)
            RM.getHome()
    def emailClick(self):
        try:
            driver.get("https://mail.worksmobile.com/#/my/102")
            t.sleep(2)
            mailHome_soup = BeautifulSoup(driver.page_source,'html.parser')
            if mailHome_soup.find('li', attrs={'class':'notRead'}) != None:
                newMail = driver.find_element(By.XPATH,"//li[contains(@class,'notRead')]//div[@class='mTitle']//strong[@class='mail_title']")
                ActionChains(driver).click(newMail).perform()
                driver.get("https://mail.worksmobile.com/#/my/102")
                t.sleep(60*58)
            else:pass
        except TimeoutException:
            driver.quit()
            t.sleep(5)
            RM.getHome()
    #종료
    def logout(self):
        try:
            logout_profile = driver.find_element(By.XPATH,'//div[@class="profile_area"]')
            logout_btn = driver.find_element(By.XPATH,'//a[@class="btn logout"]')
            ActionChains(driver).click(logout_profile).click(logout_btn).perform()
            t.sleep(1)
        except TimeoutException:
            driver.quit()
            t.sleep(5)
            RM.getHome()
    def login(self):
        try:
            password_box = driver.find_element(By.XPATH,'//input[@id="user_pwd"]')
            login_button_2 = driver.find_element(By.XPATH,'//button[@id="loginBtn"]')
            password = works_login['pw']
            ActionChains(driver).send_keys_to_element(password_box, '{}'.format(password)).click(login_button_2).perform()
            t.sleep(1)
        except TimeoutException:
            driver.quit()
            t.sleep(5)
            RM.getHome()
RMmail = RM()
if __name__ == "__main__":
    RMmail.getHome()
    while True:
        if datetime.now().strftime('%d %H:%M') == "01 01:00":
            RMmail.reset()
        else:pass
        for i in range(100):
            if datetime.now().strftime('%d') in restday[datetime.now().strftime('%m')]:
                RMmail.newMail()
                t.sleep(5)
            else:
                if time(8,0)<datetime.now().time()<=time(18,0):
                    RMmail.emailClick()
                    t.sleep(5)
                else:
                    RMmail.newMail()
                    t.sleep(5)
        RMmail.logout
        RMmail.login
        t.sleep(0.5)