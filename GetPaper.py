from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait

driver = webdriver.Chrome('C:/Users/kpark/AppData/Local/Programs/Python/chromedriver.exe')
driver.get('http://www.riss.kr/index.do')

def search_riss():
    handler = driver.window_handles

    for i in range(0,4):
        try:
            driver.switch_to_window(handler[i])
            query_input = driver.find_element_by_xpath("//input[@id='basicQuery']")
            
        except Exception as e:
            print('Failure')
            driver.close()

    driver.switch_to_window(handler[0])

    # Input of Keywords & Search
    sleep(2)
    query = "우버"
    query_input.send_keys(query)
    search_btn = driver.find_element_by_xpath("//input[@src='/main/images/sc_btn.gif']")
    search_btn.click()
 
    # 학위논문
    graduate_sec = driver.find_element_by_xpath("//a[contains(text(), '학위논문')]")
    graduate_sec.click()

    #100개씩 보여주기
    show_hundred = driver.find_element_by_xpath("//a[contains(text(), '100개씩 출력')]")
    show_hundred.click()

def scrap_data():
    title_candidates = driver.find_elements_by_xpath("//p[@class='txt']") #전체 돌아가는 숫자
    val = 0 #저자, 소속을 위한 임시 변수

    for element in range(0, len(title_candidates)):
        paper_title = title_candidates[element].text

        info_temp = driver.find_elements_by_xpath("//span[@class='etc']/a")

        print(val)
        paper_author = info_temp[val].text
        val = val + 1
        print(val)
        paper_association = info_temp[val].text
        val = val + 1 
        print(val)

        try:
            year_temp = driver.find_elements_by_xpath("//span[@class='etc']")[element].text
            one_tmp = year_temp.split(' ')
            print(one_tmp)
            two_tmp = one_tmp[2].split(',')
            print(two_tmp)
            paper_year = two_tmp[1][1:-1]
        
        except Exception as e:
            year_temp = driver.find_elements_by_xpath("//span[@class='etc']")[element].text
            one_tmp = year_temp.split(' ')
            print(one_tmp)
            two_tmp = one_tmp[1].split(',')
            print(two_tmp)
            paper_year = two_tmp[1][1:-1]            

        print(paper_title)
        print(paper_author)
        print(paper_association)
        print(paper_year)

def search_list_by_clicking():
    #전체 페이퍼 항목들을 보여줌
    graduate_paper_list = driver.find_elements_by_xpath("//p[@class='txt']")
    
    for element in range(0, len(graduate_paper_list)):
        print(element)
        graduate_paper_list = driver.find_elements_by_xpath("//p[@class='txt']")
        graduate_paper_list[element].click()
        sleep(1)

        #제목
        paper_title = driver.find_element_by_xpath("//p[@id='vTop02_tit']").text

        #저자
        author_temp = driver.find_elements_by_xpath("//a[@class='text']")
        '''
        # Test Code
        for i in range(0,len(author_temp)):
            print(i)
            print(author_temp[i].text)
        '''
        paper_author = author_temp[0].text
        sleep(1)

        #발행년도
        year_tmp = driver.find_elements_by_xpath("//ul[@class='report']/li/p")
        paper_year = year_tmp[7].text
        sleep(1)

        #URL
        paper_url = driver.find_element_by_xpath("//p[@class='copy']/a").text

        print(paper_title)
        print(paper_author)
        print(paper_year)
        print(paper_url)

        driver.execute_script("window.history.go(-1)") # Backspacing is not working
        driver.refresh()
        sleep(1)    







search_riss()
scrap_data()