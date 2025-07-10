import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from urllib.parse import quote

df = pd.read_excel("Book1.xlsx")

driver = webdriver.Chrome()

driver.get("https://web.whatsapp.com")
print("Scan QR with phone")
time.sleep(20)

for i,row in df.iterrows():
    number = row['Phone']
    message = row['Message']
    encoded_msg = quote(" ")

    url = f"https://web.whatsapp.com/send?phone={number}&text={encoded_msg}"
    driver.get(url)
    time.sleep(10)
    try:
        time.sleep(10)
        attach_btn = WebDriverWait(driver,15).until(EC.element_to_be_clickable((By.XPATH,"//button[@title='Attach' and @type='button']//span[@data-icon='plus-rounded']")))
        attach_btn.click()
        time.sleep(1)

        poll_btn = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "//li[.//span[text()='Poll']]")))
        poll_btn.click()
        time.sleep(1)

        textboxes = WebDriverWait(driver,10).until(
            EC.presence_of_all_elements_located((By.XPATH,"//div[@role='textbox' and @contenteditable='true']"))
        )
        if len(textboxes) >= 3:
            textboxes[0].click()
            textboxes[0].send_keys(message)
            time.sleep(0.5)
            textboxes[1].click()
            textboxes[1].send_keys("Yes")
            time.sleep(0.5)
            textboxes[2].click()
            textboxes[2].send_keys("No")
            time.sleep(0.5)

            toggle = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"polls-single-option-switch")))
            if toggle.get_attribute("aria-checked") == 'true':
                toggle.click()
            send_btn = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//div[@role='button' and @aria-label='Send']")))
            send_btn.click()
        else:
            print(f"Poll not loaded for {number}")
        send_poll = driver.find_element(By.XPATH,"//span[@data-icon='send']")
        send_poll.click()
        print(f"Poll sent to {number}")
        df.at[i,'Status'] = 'sent'
    except Exception:
        print(f"Failed for {number}")
        df.at[i,'Status'] = "Failed"
    time.sleep(5)