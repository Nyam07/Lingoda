# Imports
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.chrome.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time

# stop warnings
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(service=Service(
    ChromeDriverManager().install()), options=options)


df = pd.read_excel("Lingoda_input.xls", dtype=str)

#function to click the cookies button


def click_cookies():
    #accept terms an conditions
    try:
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="onetrust-accept-btn-handler"]'))).click()
    except TimeoutException:
        pass


for i in range(len(df)):
    fname = df['First Name'].iloc[i]
    lname = df['Last Name'].iloc[i]
    email = df["Email "].iloc[i]
    user_password = str(df['Password'].iloc[i])
    card_no = df["Credit card number"].iloc[i]
    expiration_date = df["Expiration date of the credit card"].iloc[i]
    cvc = str(df['CVC'].iloc[i])
    date_of_course = df["Date of the course"].iloc[i]
    time_of_course = df["Time of the course"].iloc[i]
    instructor_notes = df["Notes for the instructor"].iloc[i]
    user_language = df["Language"].iloc[i]
    user_level = df["Level"].iloc[i]

    # loop untill the elements for registration are present
    while True:
        driver = webdriver.Chrome(service=Service(
            ChromeDriverManager().install()), options=options)
        url = "https://www.lingoda.com/en/"
        driver.get(url)

        click_cookies()

        # click the dropdown
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="tw-block-hero-home-0"]/div[2]/div[3]/div[1]/div[1]'))).click()

        get_started = driver.find_element(
            By.XPATH, '//*[@id="tw-block-hero-home-0"]/div[2]/div[3]/div[2]')

        if user_language == "English":
            select_language = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="tw-block-hero-home-0"]/div[2]/div[3]/div[1]/div[2]/ul/li[1]')))
            select_language.click()
            time.sleep(2)
            try:
                get_started.click()
            except:
                print('Please check your internet connection.')

            try:
                #page2
                #click_cookies()

                if user_level == "A1":
                    level = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div[2]/button[1]')))
                elif user_level == "A2":
                    level = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div[2]/button[2]')))
                elif user_level == "B1":
                    level = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div[2]/button[3]')))
                elif user_level == "B2":
                    level = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div[2]/button[4]')))
                else:
                    level = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div[2]/button[5]')))
                level.click()

                #click_cookies()
                class_type = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/button[1]')))
                class_type.click()

                later = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div/div/div[4]/button')))
                later.click()
            except TimeoutException:
                driver.close()
                continue
            break

        elif user_language == "Business English":
            try:
                select_language = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="tw-block-hero-home-0"]/div[2]/div[3]/div[1]/div[2]/ul/li[2]')))
                select_language.click()
                time.sleep(2)

                try:
                    get_started.click()
                except:
                    print('Please check your internet connection.')

                if user_level == "A1":
                    level = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div[2]/button[1]')))
                elif user_level == "A2":
                    level = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div[2]/button[2]')))
                else:
                    level = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div[2]/button[3]')))

                level.click()

                class_type = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/button[1]')))
                class_type.click()
            except TimeoutException:
                driver.close()
                continue
            break
    # free trial error button
    try:
        free_trial = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div[2]/div[1]/div/div/div[5]/button')))
        free_trial.click()
    except TimeoutException:
        free_trial = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="no-cashback"]/div[2]/div/div/div[1]/div[2]/div[4]/a[2]')))
        free_trial.click()

    signup_btn = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/div/div[2]/div/div[3]/button')))
    signup_btn.click()

    #enter details
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.NAME, 'firstName'))).send_keys(fname)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.NAME, 'lastName'))).send_keys(lname)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.NAME, 'email'))).send_keys(email)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.NAME, 'password'))).send_keys(user_password)

    #SIGNUP_BTN
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/div/div[2]/div/div[3]/div/form/div/div[5]/button'))).click()

    try:
        # switch to the iframe containing credit card details
        WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '//*[@id="lingoda-card"]/div/iframe')))
    except TimeoutException:
        print('Email already used!!!')
        continue

    #fill credit card details
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/form/div/div[2]/span[1]/span[2]/div/div[2]/span/input'))).send_keys(card_no)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/form/div/div[2]/span[2]/span/span/input'))).send_keys(expiration_date)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/form/div/div[2]/span[3]/span/span/input'))).send_keys(cvc)

    # switch to main frame and accept terms and conditions
    driver.switch_to.default_content()
    driver.find_element(By.NAME, 'termsAndConditions').click()

    #start free trial
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, '#root > div.MuiBox-root.css-1obf64m > div.MuiBox-root.css-ab8yd1 > div.MuiBox-root.css-b95f0i > div.MuiBox-root.css-14d5zh9 > div > div > div > div.MuiGrid-root.MuiGrid-item.MuiGrid-grid-xs-12.MuiGrid-grid-md-7.sc-bdvvtL.dNcRQs.css-1u92e3e > div > div > form > div > div.MuiBox-root.css-14x72jt > div > div > button > div'))).click()

    except TimeoutException:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, '//*[@id="root"]/div/div[1]/div[4]/div[2]/div/div/div/div[1]/div/div/form/div/div[3]/div/div/button/div'))).click()
    #CONTINUE BUTTON
    try:
        #click_cookies()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/button'))).click()
    except TimeoutException:
        print('Credit card error!! .Please check that you entered the correct details and that the card has some funds.')
        break

    #other
    #click_cookies()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/div[15]/button'))).click()

    # done button
    #click_cookies()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[3]/button'))).click()

    # book classes
    #click_cookies()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="root"]/div/div[1]/div[1]/header/div/div[1]/div/div/a[2]'))).click()

    # book customized private class
    time.sleep(2)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sidebar"]/div/div/div[2]/div[2]/div/a'))).click()

    # instructor notes
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[8]/div[3]/div/div[2]/div[1]/div/div/div/textarea[1]'))).send_keys(instructor_notes)

    #date
    driver.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/div[2]/div[3]/div/div[1]/div/div/div/input').send_keys(date_of_course)

    #time
    time_el= driver.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/div[2]/div[3]/div/div[2]/div/div/div/input')
    time_el.send_keys(Keys.CONTROL,"a", Keys.DELETE)
    time_el.send_keys(time_of_course)

    #set time
    driver.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/div[2]/div[3]/div/div[2]/div/div/div/div/button').click()

    # book class
    driver.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/div[3]/div/div[1]/button').click()

    #confirm button
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[3]/div/div[3]/div/div[1]/button'))).click()
        print('Success')
    except TimeoutException:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'body > div:nth-child(37) > div.MuiDialog-container.MuiDialog-scrollPaper.css-ekeie0 > div > div.MuiBox-root.css-1lnk26a > div > div.MuiGrid-root.MuiGrid-item.MuiGrid-grid-xs-12.MuiGrid-grid-sm-auto.sc-bdvvtL.llWXcJ.css-110lzp > button'))).click()

    driver.close()