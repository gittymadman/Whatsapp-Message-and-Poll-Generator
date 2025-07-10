import selenium
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from urllib.parse import quote

df = pd.read_excel("Book1.xlsx")

driver = webdriver.Chrome()

driver.get("https://web.whatsapp.com")
print("Scan QR with phone")
time.sleep(20)

for i,row in df.iterrows():
    number = row['Phone']
    message = row['Message']
    encoded_msg = quote(message)

    url = f"https://web.whatsapp.com/send?phone={number}&text={encoded_msg}"
    driver.get(url)
    time.sleep(10)

    try:
        continue_btn = driver.find_elements(By.LINK_TEXT,"Continue to Chat")
        if continue_btn:
            continue_btn[0].click
            time.sleep(5)
        
        use_web = driver.find_elements(By.LINK_TEXT, "use WhatsApp Web")
        if use_web:
            use_web[0].click
            time.sleep(10)

        send_button = driver.find_element(By.XPATH,"//span[@data-icon='wds-ic-send-filled']")
        send_button.click()
        print(f"Message sent to {number}")

        time.sleep(5)
    except Exception:
        print(f"Failed to send to {number}")
