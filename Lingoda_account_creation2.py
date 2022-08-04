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
from selenium.common.exceptions import TimeoutException
import pandas as pd
import datetime
import calendar


# stop warnings
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])


# read the csv file
df = pd.read_excel("Lingoda_input.xls", dtype=str)

#function to click the cookies button
def click_cookies():
    #accept terms an conditions
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="onetrust-accept-btn-handler"]'))).click()
    except TimeoutException:
        pass

#loop though the excel file
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

    date_of_course= "2022/"+date_of_course
    course_date = datetime.datetime.strptime(date_of_course, '%Y/%m/%d')
    course_month = calendar.month_name[course_date.month]
    course_day = str(course_date.day)
    course_month


    #START BROWSER
    driver = webdriver.Chrome(service=Service(
        ChromeDriverManager().install()), options=options)

    # loop untill the elements for registration are present
    while True:

        if user_language == "English":
            url = "https://learn.lingoda.com/en/get-started/level?sectionSlug=english"

            driver.get(url)
            try:
                #page2
                click_cookies()

                #select the language level
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

                #select the class type
                class_type = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/button[1]')))
                class_type.click()

                later = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/div/div/div[4]/button')))
                later.click()
            except TimeoutException:
                continue
            break

        elif user_language == "Business English":
            try:
                url = "https://learn.lingoda.com/en/get-started/level?sectionSlug=business-english"
                driver.get(url)

                click_cookies()

                #SELECT THE USER LEVEL
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

                #SELECT THE CLASS TYPE
                class_type = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/button[1]')))
                class_type.click()
            except TimeoutException:
                continue
            break
        else:
            print("Please select either English or Business English")
            continue

    # CLICK THE FREE TRIAL BUTTON
    try:
        free_trial = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, '#root > div > div.MuiBox-root.css-dmmwmi > main > div.MuiBox-root.css-dvxtzn > div.MuiGrid-root.sc-bdvvtL.css-dwbo9d > div > div:nth-child(1) > div > div > div:nth-child(5) > button')))
        free_trial.click()
    except TimeoutException:
        free_trial = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="no-cashback"]/div[2]/div/div/div[1]/div[2]/div[4]/a[2]')))
        free_trial.click()

    # CLICK THE SIGNUP BUTTON
    signup_btn = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/div/div[2]/div/div[3]/button')))
    signup_btn.click()

    # ENTER PERSONAL DETAILS
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.NAME, 'firstName'))).send_keys(fname)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.NAME, 'lastName'))).send_keys(lname)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.NAME, 'email'))).send_keys(email)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.NAME, 'password'))).send_keys(user_password)

    # CLICK SIGNUP_BTN
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/div/div[2]/div/div[3]/div/form/div/div[5]/button'))).click()

    try:
        # switch to the iframe containing credit card details
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '//*[@id="lingoda-card"]/div/iframe')))
    except TimeoutException:
        print('Email already used!!!')
        continue

    # FILL CREDIT CARD DETAILS
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/form/div/div[2]/span[1]/span[2]/div/div[2]/span/input'))).send_keys(card_no)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/form/div/div[2]/span[2]/span/span/input'))).send_keys(expiration_date)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/form/div/div[2]/span[3]/span/span/input'))).send_keys(cvc)

    # SWITCH TO MAIN FRAME AND ACCEPT TERMS AND CONDITIONS
    driver.switch_to.default_content()
    driver.find_element(By.NAME, 'termsAndConditions').click()

    # CLICK THE START FREE TRIAL BUTTON
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="root"]/div/div[1]/div[4]/div[2]/div/div/div/div[1]/div/div/form/div/div[3]/div/div/button'))).click()

    #CONTINUE BUTTON
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div/button'))).click()
    except TimeoutException:
        print('Credit card error!! .Please check that you entered the correct details and that the card has some funds.')
        break

    #other
    WebDriverWait(driver, 20).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[2]/div[15]/button'))).click()

    # done button
    WebDriverWait(driver, 20).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="root"]/div/div[1]/main/div[2]/div/div[3]/button'))).click()

    driver.close()