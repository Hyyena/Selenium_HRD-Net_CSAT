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
driver.maximize_window()

# 훈련 기관/과정 검색 클릭
checkBox = driver.find_element_by_xpath('//*[@id="keywordType2"]')
checkBox.click()

# 훈련기관명 입력
searchBox1 = driver.find_element_by_xpath('//*[@id="keyword2"]')
searchBox1.send_keys('그린컴퓨터')

# 엑셀 연동
workbook = load_workbook(filename='test.xlsx') # 파일명 입력
sheet = workbook['sheet1']

rowStart = 3 # 행 시작 번호 입력
while rowStart <= 500:
    i = str(rowStart)
    rowSearch = sheet['C'+i].value # C열 i번째 셀값

    # 훈련과정명 입력
    searchBox2 = driver.find_element_by_xpath('//*[@id="keyword1"]')
    assert isinstance(rowSearch, object)
    searchBox2.send_keys(rowSearch)

    # 개강일자 입력
    searchBox3 = driver.find_element_by_xpath('//*[@id="startDate"]')
    searchBox3.clear()
    searchBox3.send_keys('20200101')

    # 선택항목 검색 클릭
    search = driver.find_element_by_xpath('//*[@id="searchForm"]/div/div[2]/div[1]/fieldset/div[2]/div/button[2]')
    search.click()

    # 과정 클릭
    posting = driver.find_element_by_xpath('//*[@id="searchForm"]/div/div[2]/div[8]/ul/li/div[2]/p/a')
    posting.click()

    # 최근 열린 탭으로 전환
    driver.switch_to.window(driver.window_handles[-1])

    # 명시적 대기 (5초 동안 0.5초씩 element를 찾을 수 있는지 시도)
    try:
        element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="infoTab4"]/button' ))
        )
    finally:

        # currentUrl = driver.current_url
        # driver.get(currentUrl)
        # html = driver.page_source
        # bsObj = BeautifulSoup(html, 'html.parser')
        # srch1 = driver.find_elements_by_xpath('//*[@id="infoDiv"]/div[1]/dl/dd/span[1]')
        # srch2 = driver.find_elements_by_class_name('ment')
        # for sample1 in srch1:
        #     print(sample1.text)
        #
        # for sample2 in srch2:
        #     print(sample2.get_attribute('text'))

        # 만족도/수강후기 스크롤
        scrollDown = driver.find_element_by_xpath('//*[@id="infoTab4"]/button')
        ActionChains(driver).move_to_element(scrollDown).perform()

        # 만족도/수강후기 클릭
        review = driver.find_element_by_xpath('//*[@id="infoTab4"]/button')
        review.click()

        # HTML 추출 후 만족도 정보와 후기 정보 찾기
        html = driver.page_source
        bsObj = BeautifulSoup(html, 'html.parser')
        srch1 = bsObj.find_all("span", {"class": "num"}, limit=1)
        srch2 = bsObj.find_all("p", {"class": "ment"})

        # 호출한 엑셀 파일에 만족도 정보와 후기 정보를 각각 A열, B열 1행에 입력
        wrksht1 = workbook.active # 현재 참조중인 엑셀 파일
        for j in srch1:
            wrksht1['A'+i] = j.text # A열 i번째 셀에 srch1 정보의 text만 추출하여 입력
        for j in srch2:
            wrksht1['B'+i] = j.text # B열 i번째 셀에 srch2 정보의 text만 추출하여 입력
        rowStart = rowStart + 1
wrksht1.save('C:/Users/git_200101/Desktop/sample.xlsx') # 엑셀 파일 저장