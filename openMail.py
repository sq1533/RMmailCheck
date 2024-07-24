import time
import requests
import json
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

#로그인 정보 호출
works_login = pd.read_json('C:\\Users\\USER\\ve_1\\RMmailCheck\\login.json',orient='records')

#숫자 콤마넣기
def comma(x):
    return '{:,}'.format(round(x))
#매월 1일 데이터 초기화
def reset():
    if datetime.now().day == 1:
        resets = {
            "상점ID":"T_ID",
            "상점명":"T_Name",
            "월한도":"1000000",
            "비고":"",
        }
        pd.DataFrame(resets,index=[0]).to_json('C:\\Users\\USER\\ve_1\\RMmailCheck\\RMdata.json',orient='records',force_ascii=False,indent=4)
        #텔레그램 API 전송
        requests.get(f"https://api.telegram.org/bot{works_login['bot']['token']}/sendMessage?chat_id={works_login['bot']['chatId']}&text=초기화_완료")
        time.sleep(2)
    else:
        pass
#임시한도 증액 메일 텍스트 데이터 읽어오기
def mailCheck():
    #RM한도 증액 제외 가맹점
    ignoreName = ["이지피쥐","핀테크링크​","엘피엔지​","코리아결제시스템"]
    ignoreOrder = ["오프라인"]
    #크롬 드라이버 옵션 설정
    driver = webdriver.Chrome(options=webdriver.ChromeOptions().add_argument('--blink-settings=imagesEnabled=false'))
    #크롬 드라이버 실행
    url = "https://mail.worksmobile.com/#/my/104"
    driver.get(url)
    driver.implicitly_wait(1)
    #로그인 정보입력(아이디)
    id_box = driver.find_element(By.XPATH,'//input[@id="user_id"]')
    login_button_1 = driver.find_element(By.XPATH,'//button[@id="loginStart"]')
    ActionChains(driver)
    id = works_login['works']['id']
    ActionChains(driver).send_keys_to_element(id_box, '{}'.format(id)).click(login_button_1).perform()
    time.sleep(1)
    #로그인 정보입력(비밀번호)
    password_box = driver.find_element(By.XPATH,'//input[@id="user_pwd"]')
    login_button_2 = driver.find_element(By.XPATH,'//button[@id="loginBtn"]')
    password = works_login['works']['pw']
    ActionChains(driver).send_keys_to_element(password_box, '{}'.format(password)).click(login_button_2).perform()
    time.sleep(5)
    driver.refresh()
    time.sleep(3)
    #타임 스케쥴 진행
    html = driver.page_source
    mailHome_soup = BeautifulSoup(html,'html.parser')
    #신규 메일 확인
    if mailHome_soup.find('li', attrs={'class':'notRead'}) != None:
        newMail = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,"//li[contains(@class, 'notRead')]//div[@class='mTitle']//strong[@class='mail_title']")))
        ActionChains(driver).click(newMail).perform()
        time.sleep(1)
        #필요 데이터 가져오기
        html = driver.page_source
        mail_soup = BeautifulSoup(html,'html.parser')
        RM_market = mail_soup.find_all('td')
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
        RM_month = pd.read_json('C:\\Users\\USER\\ve_1\\RMmailCheck\\RMdata.json',
                                orient='records',
                                dtype={'상점ID':str,'상점명':str,'월한도':str,'비고':str})
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
        #증액 필요 가맹점 파일 업데이트 및 텔레그램 전송
        if newdata.empty:
            #텔레그램 API 전송
            tell = "{일}일 {시간}시 증액 필요 가맹점 없음".format(일=datetime.now().day,시간=datetime.now().hour)
            requests.get(f"https://api.telegram.org/bot{works_login['bot']['token']}/sendMessage?chat_id={works_login['bot']['chatId']}&text={tell}")
            driver.quit()
        else:
            for update in newdata.index.tolist():
                tell = '{일}일 {시간}시 {상점명}[{상점ID}] 한도 증액필요\n월한도 {한도}원 / 증액 {증액}원'.format(
                                                                                        일=datetime.now().day,
                                                                                        시간=datetime.now().hour,
                                                                                        상점명=newdata.loc[update]["상점명"],
                                                                                        상점ID=newdata.loc[update]["상점ID"],
                                                                                        한도=comma(int(newdata.loc[update]["월한도"])),
                                                                                        증액=comma(int(newdata.loc[update]["월한도"])*120/100))
                #텔레그램 API 전송
                requests.get(f"https://api.telegram.org/bot{works_login['bot']['token']}/sendMessage?chat_id={works_login['bot']['chatId']}&text={tell}")
                time.sleep(1)
                #Json파일 업로드
                if update == newdata.index.tolist()[-1]:
                    resurts = pd.concat([RM_month,newdata],ignore_index=True)
                    resurts.to_json('C:\\Users\\USER\\ve_1\\RMmailCheck\\RMdata.json',orient='records',force_ascii=False,indent=4)
                    driver.quit()
                else:
                    pass
    else:
        requests.get(f"https://api.telegram.org/bot{works_login['bot']['token']}/sendMessage?chat_id={works_login['bot']['chatId']}&text=이메일 없음")
        driver.quit()
        pass
with open('C:\\Users\\USER\\ve_1\\RMmailCheck\\restDay.json',"r") as f:
    restday = json.load(f)
Timeline1 = ["00:00","02:00","04:00","06:00","08:00","10:00","12:00","14:00","16:00","18:00","20:00","22:00"]#주말 및 공휴일 대응
Timeline2 = ["00:00","02:00","04:00","06:00","18:00","20:00","22:00"]#영업시간 미대응
Timeline3 = ["08:00","10:00","12:00","14:00","16:00"]#영업시간 외 대응
if __name__ == "__main__":    
    while True:
        if datetime.now().strftime('%d') in restday[datetime.now().strftime('%m')]:
            if datetime.now().strftime('%H:%M') in Timeline1:
                mailCheck()
                time.sleep(60)
            else:
                pass
        else:
            if datetime.now().strftime('%H:%M') in Timeline2:
                mailCheck()
                time.sleep(60)
            elif datetime.now().strftime('%H:%M') in Timeline3:
                requests.get(f"https://api.telegram.org/bot{works_login['bot']['token']}/sendMessage?chat_id={works_login['bot']['chatId']}&text=영업시간 미대응")
                time.sleep(60)
            else:
                pass
        if datetime.now().strftime('%d %H:%M') == "01 01:00":
            reset()
            time.sleep(60)
        else:
            pass
        time.sleep(1)