import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from urllib.parse import quote
import re

df = pd.read_excel("Book1.xlsx")

driver = webdriver.Chrome()

driver.get("https://web.whatsapp.com")
print("Scan QR with phone")
time.sleep(20)

for i,row in df.iterrows():
    number = row['Phone']
    encoded_msg = quote(" ")
    url = f"https://web.whatsapp.com/send?phone={number}&text={encoded_msg}"
    driver.get(url)
    time.sleep(15)
    try:
        poll_messages = driver.find_elements(By.XPATH, "//div[contains(@aria-label, 'Poll from you')]")
        if not poll_messages:
            print("No poll messages found")
            vote = "Not Found"
        else:
            # Get the latest poll (assumes last one is most recent)
            latest_poll = poll_messages[-1]
            aria_label = latest_poll.get_attribute("aria-label")

            match = re.search(r"Yes:\s*(\d+),\s*No:\s*(\d+)",aria_label)
            if match:
                yes_count = int(match.group(1))
                No_count = int(match.group(2))

                if yes_count > 0:
                    vote="Yes"
                elif No_count > 0:
                    vote="No"
                else:
                    vote="No Response"
            else:
                print("Unable to check")
                vote = "Unable to check"

        df.at[i, 'Response'] = vote
        print(f"Vote for {number}:{vote}")

    except Exception:
        print("Error extracting vote:")
        df.at[i, 'Response'] = "Error"


    