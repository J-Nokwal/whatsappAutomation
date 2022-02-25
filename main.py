from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
from pathlib import Path
import win32clipboard
from PIL import Image
from io import BytesIO
from selenium.webdriver.common.action_chains import ActionChains

print("app started")
def send_to_clipboard(clip_type, filepath):
    try:
        image = Image.open(filepath)
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()
    except:
        print("-Error while coping image")
        print("-Please paste correct image with name myImage.png or data_image.jpg")
        raise

sleepTime = input("-Input Time Gap Between Message (Minimum-5)\n-Press Enter to go For Default (Default - 10 sec)\n-: ")
if sleepTime == '':
    sleepTime = 10
else:
    sleepTime = int(sleepTime)
if sleepTime < 5:
    sleepTime = 5
profile = "user-data-dir=" + str(Path.home()) + '''\\AppData\\Local\\Google\\Chrome\\User Data\\Default'''

print(profile)
try:
    options = webdriver.ChromeOptions()
    # options.add_argument('''user-data-dir=C:\\Users\\J-Nokwal\\AppData\\Local\\Google\\Chrome\\User Data\\Default''')
    # print(options.arguments)
    options.add_argument(profile)
    driver = webdriver.Chrome('chromedriver', options=options)
except:
    print("Chrome Driver Option Error")

driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 60)

data = pd.read_csv("data_file.csv")
data_dict = data.to_dict('list')
pnums = data_dict['Number']
messages = data_dict['Message']
tlen = len(messages)
status = [False for i in range(tlen)]

inp_xpath = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]'
send_xpath = '//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[2]/div[2]/div/div'
imgPath = "data_image.jpg"
# driver.get("http://web.whatsapp.com/send?phone=+919855400573&text=hii")
send_to_clipboard(win32clipboard.CF_DIB, imgPath)
actions = ActionChains(driver)
haveError = False
print("-Automation Started")
for i in range(tlen):
    try:
        time.sleep(sleepTime - 3)
        driver.get("http://web.whatsapp.com/send?phone=+{}&text={}".format(pnums[i], messages[i]))
        input_box = wait.until(EC.presence_of_element_located((By.XPATH, inp_xpath)))
        input_box.send_keys(Keys.CONTROL + 'v')
        time.sleep(1)
        send_box = wait.until((EC.presence_of_element_located((By.XPATH, send_xpath))))
        send_box.send_keys(Keys.ENTER)
        status[i] = True
        time.sleep(2)
        status[i] = True
    except:
        haveError = True
data['Status'] = status
data.to_csv('data_file.csv', index=False)
if haveError:
    print("-An exception occurred")
    print("-Please make sure you have active internet connectio")
    print("-phone nubmer should contain country code without '+'")
    print("-Do Not Copy Any text or Image or any thing while process is going on")
    print("-Close previously opened window opened by this sofware")
    print("-data_file.csv is updated for Status")
driver.quit()
print("-Automation Ends")