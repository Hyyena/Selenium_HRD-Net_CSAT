from bs4 import BeautifulSoup
from openpyxl import load_workbook

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

        # 만족도/수강후기 스크롤
        scrollDown = driver.find_element_by_xpath('//*[@id="infoTab4"]/button')
        ActionChains(driver).move_to_element(scrollDown).perform()

        # 만족도/수강후기 클릭
        review = driver.find_element_by_xpath('//*[@id="infoTab4"]/button')
        review.click()

        # 만족도 복사
