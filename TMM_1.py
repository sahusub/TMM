
# *********************Trial Monitoring*********************
# **********************************************************

import xlrd
import datetime
import collections
from selenium import webdriver, common
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.select import Select
import xlrd
import unittest
from selenium.webdriver.common.action_chains import ActionChains


chrome_options = Options()
chrome_options.add_argument('--enable-automation')

global driver
driver = webdriver.Chrome(
    r'C:\Users\sahusub\Downloads\chromedriver_win32 - latest_78\chromedriver.exe')
driver.set_window_size(1024, 600)
driver.maximize_window()


def loginForDEVQA():

    driver.get('https://novartis3-sb.pvcloud.com/testing/')

    DSN = Select(driver.find_element_by_name('DSN'))
    DSN.select_by_visible_text('NVTSB1QA')
    username = driver.find_element_by_id('Username')
    password = driver.find_element_by_id('UserPass')
    username.send_keys('sahusub')
    password.send_keys('horizon')
    password.submit()
    print('Login is done !!')


def login():
    driver.get('https://novartis.pvcloud.com/planview/')
    print('Login tp PROD PV is done !!')


def openAndcopyTemplate():
    openWP('4007544')
    TMMAL = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[@title=">: Trial Monitoring Timesheet Activities List"]')))
    TMMAL.click()
    time.sleep(1)
    actions = ''
    actions = ActionChains(driver)
    actions.key_down(Keys.CONTROL)
    actions.send_keys('c')
    actions.key_up(Keys.CONTROL)
    time.sleep(1)
    actions.perform()


def searchWP(wpcode):
    print('Opening WP: ', wpcode)
    searchbox = driver.find_element_by_id('bannerSearchBox')
    searchbox.clear()
    time.sleep(1)
    for i in wpcode:
        searchbox.send_keys(i)
        time.sleep(0.5)
    time.sleep(5)
    searchbox.send_keys(Keys.ENTER)
    WPCode_in_SearchResult = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[@id="searchResults"]/tbody/tr[2]/td[4]')))


def openWP(wpcode):
    #print('Opening WP: ', wpcode)
    searchbox = driver.find_element_by_id('bannerSearchBox')
    time.sleep(4)
    searchbox.clear()
    time.sleep(1)
    searchbox.send_keys(wpcode)
    time.sleep(5)
    searchbox.send_keys(Keys.ENTER)
    WPCode_in_SearchResult = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[@id="searchResults"]/tbody/tr[2]/td[4]')))
    if WPCode_in_SearchResult.text == wpcode:

        ele1 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, r'//*[@class="form-field"]/span/span')))
        ele1.click()
        ele2 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, r'//*[@id="pvPopup1pvPopup"]/ul/li[2]')))
        ele2.click()
        time.sleep(15)

    else:
        print("Error: WP Code Searched: "+wpcode +
              " But WP Opened: "+WPCode_in_SearchResult.text)
        searchbox = driver.find_element_by_id('bannerSearchBox')
        searchbox.clear()
        time.sleep(1)
        for i in wpcode:
            searchbox.send_keys(i)
            time.sleep(0.5)
        time.sleep(5)
        searchbox.send_keys(Keys.ENTER)
        WPCode_in_SearchResult = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
            (By.XPATH, r'//*[@id="searchResults"]/tbody/tr[2]/td[4]')))
        if WPCode_in_SearchResult.text == wpcode:

            ele1 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                (By.XPATH, r'//*[@class="form-field"]/span/span')))
            ele1.click()
            ele2 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                (By.XPATH, r'//*[@id="pvPopup1pvPopup"]/ul/li[2]')))
            ele2.click()
            time.sleep(15)
        else:
            searchWP(wpcode)
            WPCode_in_SearchResult = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
                (By.XPATH, r'//*[@id="searchResults"]/tbody/tr[2]/td[4]')))

            if WPCode_in_SearchResult.text == wpcode:

                ele1 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                    (By.XPATH, r'//*[@class="form-field"]/span/span')))
                ele1.click()
                ele2 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                    (By.XPATH, r'//*[@id="pvPopup1pvPopup"]/ul/li[2]')))
                ele2.click()
                time.sleep(15)
            else:
                searchWP(wpcode)
                WPCode_in_SearchResult = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
                    (By.XPATH, r'//*[@id="searchResults"]/tbody/tr[2]/td[4]')))
                if WPCode_in_SearchResult.text == wpcode:

                    ele1 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                        (By.XPATH, r'//*[@class="form-field"]/span/span')))
                    ele1.click()
                    ele2 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                        (By.XPATH, r'//*[@id="pvPopup1pvPopup"]/ul/li[2]')))
                    ele2.click()
                    time.sleep(15)


def pasteTemplate():
    TimeReporting = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[@title=">: Time Reporting Summary" or @title=">: Time reporting (SUMMARY)" or @title=">: Time Reporting (SUMMARY)" or @title=">: TIme Reporting Summary"]')))
    TimeReporting.click()
    actions = ''
    actions = ActionChains(driver)
    actions.key_down(Keys.CONTROL)
    actions.send_keys('v')
    actions.key_up(Keys.CONTROL)
    time.sleep(1)
    actions.perform()
    time.sleep(2)
    Save = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[contains(@id,"dialog")]/div[3]/div[2]/button[2]')))
    Save.click()
    time.sleep(4)


def setPredecessor():
    time.sleep(5)
    Search = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[contains(@id, "text-filter-")]')))
    time.sleep(2)
    Search.click()
    Search.send_keys('Fe')
    time.sleep(2)
    Predecessor = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="grid1"]/div[5]/div[3]/div/div[3]/div')))
    Predecessor.click()

    actions = ''
    actions = ActionChains(driver)
    time.sleep(1)
    actions.send_keys(Keys.ENTER)
    time.sleep(1)
    actions.perform()
    time.sleep(1)

    E1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="multival-main-input"]')))
    time.sleep(1)
    # E1.click()
    time.sleep(1)
    E1.send_keys('ctt')

    time.sleep(2)

    actions = ''
    actions = ActionChains(driver)
    time.sleep(1)
    actions.key_down(Keys.SHIFT)
    actions.key_down(Keys.DOWN)
    time.sleep(1)
    actions.key_down(Keys.ENTER)
    actions.perform()
    actions.key_up(Keys.SHIFT)
    actions.key_up(Keys.DOWN)
    actions.key_up(Keys.ENTER)
    actions.perform()
    time.sleep(1)
    actions.send_keys(Keys.ENTER)
    time.sleep(1)
    actions.perform()
    time.sleep(1)
    date_stamp = str(datetime.datetime.now()).split('.')[0]
    date_stamp = date_stamp.replace(
        " ", "_").replace(":", "_").replace("-", "_")
    file_name = date_stamp + ".png"
    driver.save_screenshot(file_name)
    time.sleep(1)
    Search.clear()


def openAuthDP():
    Search = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[contains(@id, "text-filter-")]')))
    time.sleep(2)
    Search.click()
    Search.send_keys('Time Report')
    time.sleep(1)
    TimeReport = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[@title=">>: Time Report" or @title=">: Time Report" or @title=">>: Time Reporting" or @title=">>: Time reporting"]')))
    TimeReport.click()

    ActionChains(driver) \
        .key_down(Keys.SHIFT) \
        .key_down(Keys.F9) \
        .key_up(Keys.SHIFT) \
        .key_up(Keys.F9)\
        .perform()

    Auth = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.ID, 'pv-tabs-4')))
    Auth.click()

    NewAuth = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="taskauthorize-0-insertButtonContainer"]/span[2]')))
    NewAuth.click()

    time.sleep(2)
    MainWindow = driver.window_handles[0]
    DataPicker = driver.window_handles[1]
    driver.switch_to_window(DataPicker)
    Search = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[@id="pickerTabBar"]/li[2]/a')))
    Search.click()
    time.sleep(2)
    driver.switch_to_frame('iframeSearchView')
    driver.switch_to_frame('frameAttributes')


def addGDO(Authorization):

    SearchDescription = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.ID, 'attribute_description')))
    SearchDescription.clear()
    SearchDescription.send_keys(Authorization)
    SearchButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.ID, '_search')))
    SearchButton.click()

    driver.switch_to_default_content()
    driver.switch_to_frame('iframeSearchView')
    driver.switch_to_frame('frameSearchList')
    time.sleep(2)
    '''RC = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[@type="checkbox" and @name="sel_list"]')))
    #print('len is: ', len(driver.find_elements_by_xpath(r'//*[@type="checkbox" and @name="sel_list"]')))
    '''
    RC1 = len(driver.find_elements_by_xpath(r'//*[contains(@id, "sel_link_")]'))
    print('len is: ',RC1)
    '''RC1 = WebDriverWait(driver, 15).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[contains(@id, "sel_link_")]')))'''
    
    if RC1>0:

        '''ResourceCheckBox = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, r'//*[@type="checkbox" and @name="sel_list"]')))
        ResourceCheckBox.click()'''
        RC2= WebDriverWait(driver, 15).until(EC.element_to_be_clickable(
        (By.XPATH, r'//*[contains(@id, "sel_link_")]')))
        time.sleep(2)
        RC2.click()

        time.sleep(1)
        driver.switch_to_default_content()
        driver.switch_to_frame('iframeSearchView')
        driver.switch_to_frame('frameAttributes')
        SearchDescription.clear()

    else:

        print("Atrribute already added : ", Authorization)
        time.sleep(1)
        driver.switch_to_default_content()
        time.sleep(2)
        driver.switch_to_frame('iframeSearchView')
        driver.switch_to_frame('frameAttributes')
        # driver.find_element_by_id('cancel').click()
        # time.sleep(2)
        #MainWindow = driver.window_handles[0]
        # driver.switch_to_window(MainWindow)


def closeAuthDP():
    driver.switch_to_default_content()

    OKButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.ID, 'OK')))
    OKButton.click()
    time.sleep(2)
    MainWindow = driver.window_handles[0]
    driver.switch_to_window(MainWindow)


def delete(GDO):
    Schedule = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@class = "pvSelectContainer form-field tray-button pivot-select"]')))
    time.sleep(2)
    Schedule.click()
    actions = ''
    actions = ActionChains(driver)
    actions.key_down(Keys.DOWN)
    actions.key_down(Keys.ENTER)
    actions.key_up(Keys.DOWN)
    time.sleep(1)
    actions.perform()
    time.sleep(1)
    Search = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[contains(@id, "text-filter-")]')))
    Search.click()
    Search.clear()
    Search.send_keys(GDO)

    time.sleep(2)
    GDOElement = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[text()="Global Development Operations"]')))
    GDOElement.click()

    actions = ''
    actions = ActionChains(driver)
    actions.key_down(Keys.DELETE)
    actions.key_up(Keys.DELETE)
    time.sleep(1)
    actions.perform()
    time.sleep(1)

    Save = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[text()="Yes"]')))
    Save.click()


def mainTMM():

    openAndcopyTemplate()
    workbook = xlrd.open_workbook(
        r'C:\TMM\TMMTestData.xlsx')
    sheet = workbook.sheet_by_name("Sheet1")
    rowcount = sheet.nrows
    for curr_row in range(1, rowcount, 1):
        for curr_col in range(0, 1, 1):
            WPC = str(sheet.cell_value(curr_row, curr_col))
            print('WP Code is: ', WPC)

            openWP(WPC)
            pasteTemplate()
            time.sleep(15)
            setPredecessor()
            openAuthDP()

            addGDO('I10Y')
            addGDO('I100')
            addGDO('I11Z')
            addGDO('I120')
            addGDO('I130')
            addGDO('I140')
            addGDO('I150')
            addGDO('I151')
            closeAuthDP()
            delete('Global Development Operations')
            print('Completed WP: ', WPC)


# loginForDEVQA()
login()
mainTMM()
