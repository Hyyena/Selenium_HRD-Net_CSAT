import csv
import pandas as pd
import numpy as np
from openpyxl import load_workbook

from bs4 import BeautifulSoup
from bs4 import re

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

from datetime import datetime

# 주소 입력
URL = 'https://www.hrd.go.kr/hrdp/ti/ptiao/PTIAO0100L.do'

# 브라우저 옵션
driverOpt = webdriver.ChromeOptions()
driverOpt.add_argument('headless') # 창 숨기기
driverOpt.add_argument('disable-gpu') # GPU 사용 안 함

# 불러오기(Driver & Web Load)
driver = webdriver.Chrome(executable_path='chromedriver', options=driverOpt)
driver.get(url=URL)
driver.maximize_window() # 전체화면
driver.window_handles # Tab List

def crawl():

    # 엑셀 연동 (openpyxl)
    wb = load_workbook(filename='test.xlsx') # 파일명 입력
    sht = wb['sheet1']

    # 엑셀 연동 (pandas)
    df1 = pd.read_excel('test.xlsx') # 파일명 입력
    df2 = df1.dropna(subset=['연번']) # NaN 데이터 처리
    df2_col = df2[['연번', '훈련과정명', '주훈련대상', '심사년도회차', '훈련시작일', '훈련종료일']]
    print(df2_col)

    # 엑셀 제목 행 입력
    colList = ['연번', '훈련과정명', '주훈련대상', '회차', '훈련시작일', '훈련종료일', '수강평', '만족도']

    # 오늘 날짜
    date = '15'
    strDate = str(date) # 수정 필요

    # CSV 파일 생성
    with open('HRD-Net_CSAT_'+date+'.csv', 'w', newline='') as f: # newline=''을 통해 개행
        w = csv.writer(f)
        w.writerow(colList)

        # 비어있는 리스트 생성
        rowList1 = []
        rowList2 = []

        rowSt = 3 # 행 시작 번호 입력
        while rowSt <= 500:
            i = str(rowSt)
            rowSrch = sht['C' + i].value # C열 i번째 셀값
            if rowSrch is None:
                break
            rowSt += 1

            # 훈련 기관/과정 검색 클릭
            keyBlck = driver.find_element_by_xpath('//*[@id="searchForm"]/div/div[2]/div[1]/fieldset/div[1]/dl[1]/dd/div[1]').is_displayed()
            chkBox = driver.find_element_by_xpath('//*[@id="keywordType2"]')
            if not keyBlck:
                chkBox.click()

            # 훈련기관명 입력
            srchBox1 = driver.find_element_by_xpath('//*[@id="keyword2"]')
            srchBox1.clear()
            srchBox1.send_keys('그린컴퓨터')

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

            # 명시적 대기 (3초 동안 0.5초씩 element를 찾을 수 있는지 시도)
            try:
                elem = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="infoTab4"]/button' ))
                )
            finally:

                # 만족도/수강후기 스크롤
                scrllDwn = driver.find_element_by_xpath('//*[@id="infoTab4"]/button')
                ActionChains(driver).move_to_element(scrllDwn).perform()

                # 만족도/수강후기 클릭
                rvw = driver.find_element_by_xpath('//*[@id="infoTab4"]/button')
                rvw.click()

                # 만족도 Data 추출
                srch1 = driver.find_elements_by_xpath('//*[@id="infoDiv"]/div[1]/dl/dd/span[1]')
                for i in srch1:
                    rowList1.append(i.text)

                # 수강후기 Data 추출
                dpth_1_dl = driver.find_element_by_xpath('//*[@id="tbodyEpilogue"]')
                dpth_2_dd = dpth_1_dl.find_elements_by_tag_name('dd')

                # 수강후기 Data Indexing
                for i in dpth_2_dd:
                    rowList2.append(i.text)
                    x = []
                    x.append(i.text)
                    if i is dpth_2_dd[-1]: # for문 마지막 배열만 출력
                        print(rowList2)
                        print(len(rowList2))
                        # srch1_df = pd.DataFrame(rowList2)

                # 만족도 Data Indexing
                for i in srch1:
                    rowList1.append(i.text)
                    for j in range(len(rowList2)):
                        rowList1.append(i.text)

            # 최근 열린 탭 종료 후 기존 탭 활성화
            driver.close()
            firstTab = driver.window_handles[0]
            driver.switch_to.window(window_name=firstTab)

            srch1_df = pd.DataFrame(rowList1)
            srch2_df = pd.DataFrame(rowList2)

        print(srch1_df)
        print(srch2_df)

def main():
    crawl()

if __name__ == '__main__':
    main()

#         # 호출한 엑셀 파일에 만족도 정보와 후기 정보를 각각 A열, B열 1행에 입력
#         wrksht1 = workbook.active # 현재 참조중인 엑셀 파일
#         for j in srch1:
#             wrksht1['A'+i] = j.text # A열 i번째 셀에 srch1 정보의 text만 추출하여 입력
#         for j in srch2:
#             wrksht1['B'+i] = j.text # B열 i번째 셀에 srch2 정보의 text만 추출하여 입력
#         rowSt = rowSt + 1
# wrksht1.save('C:/Users/git_200101/Desktop/sample.xlsx') # 엑셀 파일 저장