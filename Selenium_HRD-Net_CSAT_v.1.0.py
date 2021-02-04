import csv
from openpyxl import load_workbook

from bs4 import BeautifulSoup
from bs4 import re

from selenium import webdriver
from selenium.webdriver import ActionChains

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium. webdriver.support.ui import WebDriverWait

# 주소 입력
URL = 'https://www.hrd.go.kr/hrdp/ti/ptiao/PTIAO0100L.do'

# 불러오기(Driver & Web Load)
driver = webdriver.Chrome(executable_path='chromedriver')
driver.get(url=URL)
driver.maximize_window() # 전체화면
driver.window_handles # Tab List

# 훈련 기관/과정 검색 클릭
chkBox = driver.find_element_by_xpath('//*[@id="keywordType2"]')
chkBox.click()

# 훈련기관명 입력
srchBox1 = driver.find_element_by_xpath('//*[@id="keyword2"]')
srchBox1.send_keys('그린컴퓨터')

# 엑셀 연동
wb = load_workbook(filename='test.xlsx') # 파일명 입력
sht = wb['sheet1']

rowSt = 3 # 행 시작 번호 입력
while rowSt <= 500:
    rowSt = rowSt + 1
    i = str(rowSt)
    rowSrch = sht['C' + i].value # C열 i번째 셀값

    # 훈련과정명 입력
    srchBox2 = driver.find_element_by_xpath('//*[@id="keyword1"]')
    srchBox2.clear()
    srchBox2.send_keys(rowSrch)

    # 개강일자 입력
    srchBox3 = driver.find_element_by_xpath('//*[@id="startDate"]')
    srchBox3.clear()
    srchBox3.send_keys('20200101')

    # 선택항목 검색 클릭
    srch = driver.find_element_by_xpath('//*[@id="searchForm"]/div/div[2]/div[1]/fieldset/div[2]/div/button[2]')
    srch.click()

    # 과정 클릭
    post = driver.find_element_by_xpath('//*[@id="searchForm"]/div/div[2]/div[8]/ul/li/div[2]/p/a')
    post.click()

    # 최근 열린 탭으로 전환
    crntTab = driver.window_handles[-1]
    driver.switch_to.window(window_name=crntTab)

    # 명시적 대기 (5초 동안 0.5초씩 element를 찾을 수 있는지 시도)
    try:
        elem = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="infoTab4"]/button' ))
        )
    finally:

        # 만족도/수강후기 스크롤
        scrllDwn = driver.find_element_by_xpath('//*[@id="infoTab4"]/button')
        ActionChains(driver).move_to_element(scrllDwn).perform()

        # 만족도/수강후기 클릭
        rvw = driver.find_element_by_xpath('//*[@id="infoTab4"]/button')
        rvw.click()

        # 최근 열린 탭 URL 저장
        crntUrl = driver.current_url
        driver.get(crntUrl)

        # 만족도 Data 추출
        dpth_1_dl = driver.find_element_by_xpath('//*[@id="infoDiv"]/div[1]/dl')
        dpth_2_dd = dpth_1_dl.find_elements_by_tag_name("dd")
        print(dpth_2_dd[0].text)

        # 최근 열린 탭 종료 후 기존 탭 활성화
        driver.close()
        firstTab = driver.window_handles[0]
        driver.switch_to.window(window_name=firstTab)

        if rowSt == 4:
            print(dpth_2_dd[0].text)
            break


        # html = driver.page_source
        # bsObj = BeautifulSoup(html, 'html.parser')
        # srch1 = driver.find_elements_by_xpath('//*[@id="infoDiv"]/div[1]/dl/dd/span[1]')
        # srch2 = driver.find_elements_by_class_name('ment')
        # for sample1 in srch1:
        #     print(sample1.text)
        #
        # for sample2 in srch2:
        #     print(sample2.get_attribute('text'))

#         # HTML 추출 후 만족도 정보와 후기 정보 찾기
#         html = driver.page_source
#         bsObj = BeautifulSoup(html, 'html.parser')
#         srch1 = bsObj.find_all("span", {"class": "num"}, limit=1)
#         srch2 = bsObj.find_all("p", {"class": "ment"})
#
#         # 호출한 엑셀 파일에 만족도 정보와 후기 정보를 각각 A열, B열 1행에 입력
#         wrksht1 = workbook.active # 현재 참조중인 엑셀 파일
#         for j in srch1:
#             wrksht1['A'+i] = j.text # A열 i번째 셀에 srch1 정보의 text만 추출하여 입력
#         for j in srch2:
#             wrksht1['B'+i] = j.text # B열 i번째 셀에 srch2 정보의 text만 추출하여 입력
#         rowSt = rowSt + 1
# wrksht1.save('C:/Users/git_200101/Desktop/sample.xlsx') # 엑셀 파일 저장