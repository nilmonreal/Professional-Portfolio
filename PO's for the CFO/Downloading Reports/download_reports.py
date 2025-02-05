from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from driver_wait import wait_for_downloads_to_complete
import time
import os


USERID = "USER"
ACCOUNT_PASSWORD = "PASS"
PHONE = "MOBILE"

# Define the download directory
download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

# Optional - Keep the browser open if the script crashes.
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)

# Set the download directory and disable download prompts
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)


################# EZAir Approval Site ###########################
driver = webdriver.Chrome(options=chrome_options)
driver.get("http://ezanoc1cho:91/PO/PurchaseOrders.aspx")


rq_button = driver.find_element(by=By.NAME, value="ctl00$ContentPlaceHolder1$chk_pendingRequirements")
rq_button.click()

# Open the dropdown by clicking it
dropdown_button = driver.find_element(By.XPATH, "//button[@title='Level 1-Tactical Buyer , Level 2-Purchasing Coordinator, Level 3-Senior Manager, Level 4-CFO, Level 5-CEO, Finanzas']")
dropdown_button.click()

time.sleep(1)

# Click the "Select All" checkbox
select_all_checkbox = driver.find_element(By.XPATH, "//li[@class='multiselect-item multiselect-all active']/a/label/input[@value='multiselect-all']")
select_all_checkbox.click()

# Allow some time for the action to register
time.sleep(1)

# Click the "Level 4-CFO" checkbox
level_4_checkbox = driver.find_element(By.XPATH, "//li[@class='']/a/label[@title='Level 4-CFO']/input[@value='4']")
level_4_checkbox.click()

# Optional: Wait to observe the result or perform additional actions
time.sleep(1)

id_button = driver.find_element(by=By.NAME, value="ctl00$ContentPlaceHolder1$txt_F_POID")
id_button.click()
id_button.send_keys(Keys.ENTER)

time.sleep(3)

export_button = driver.find_element(By.ID, "ContentPlaceHolder1_btn_Excel")
export_button.click()

time.sleep(3)


################# RY4WEB #######################

driver = webdriver.Chrome(options=chrome_options)
driver.get("http://cdroyal4.cdzodiac.com/")

r4_web_link = driver.find_element(By.LINK_TEXT, "R4 Web EZ Interiors")
r4_web_link.click()

time.sleep(2)

try:
    user_input = driver.find_element(By.ID, "cuser")
    user_input.send_keys(USERID)
except NoSuchElementException:
    print("Element not found, refreshing page")
    driver.refresh()
    time.sleep(2)
    user_input = driver.find_element(By.ID, "cuser")
    user_input.send_keys(USERID)

pass_input = driver.find_element(By.ID, "cpass")
pass_input.send_keys(ACCOUNT_PASSWORD)
pass_input.send_keys(Keys.ENTER)

time.sleep(2)

r4_web_po = driver.find_element(By.LINK_TEXT, "Purchase Orders")
r4_web_po.click()


select_element = driver.find_element(By.ID, "inquiry")
select_inquiry = Select(select_element)
select_inquiry.select_by_value("http://ezroyal4.ezairinterior.com/cgi-bin/wspd_cgi.sh/WService=wsliveez/wpo/poInq.r")

go_button = driver.find_element(By.XPATH, "//a[@href=\"javascript:GoTo('inquiry')\"]")
go_button.click()

select_element = driver.find_element(By.ID, "poType")
select = Select(select_element)
select.select_by_visible_text("Both")

select_element = driver.find_element(By.ID, "pastDue")
select = Select(select_element)
select.select_by_visible_text("Both")

select_element = driver.find_element(By.ID, "revDue")
select = Select(select_element)
select.select_by_visible_text("Both")

date_input = driver.find_element(By.NAME, "frDt")
date_input.clear()
date_input.send_keys("01/01/15")

date_input = driver.find_element(By.NAME, "toDt")
date_input.clear()
date_input.send_keys("01/01/30")

date_input = driver.find_element(By.NAME, "toPoDt")
date_input.clear()
date_input.send_keys("01/01/30")

submit_button = driver.find_element(By.XPATH, "//td/input[@type='image' and @src='/images/button2/submit.gif']")
submit_button.click()

time.sleep(10)

export_link = driver.find_element(By.XPATH, "//a[@href='javascript:exp();']")
export_link.click()


print("Waiting for the file to download...")

# Wait for all downloads to complete
try:
    wait_for_downloads_to_complete(download_dir)
    print("Download complete!")
except TimeoutError as e:
    print(str(e))

time.sleep(30)







