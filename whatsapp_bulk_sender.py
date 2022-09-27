# Program to send bulk customized message through WhatsApp web application
# Author @inforkgodara

#*************************************************

#time.sleep is main constraint


#*************************************************



from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import pandas
import time

file_path = r"C:\Users\Aasawari\Downloads\Scenery.jpg"#file path
# Load the chrome driver
driver = webdriver.Chrome(executable_path=r"C:\Users\Aasawari\Downloads\chromedriver_win32\chromedriver.exe")
count = 0

# Open WhatsApp URL in chrome browser
driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 20)

# Read data from excel
excel_data = pandas.read_excel('bulk invitation.xlsx', sheet_name='guests')
# message = "hellooooooo!!!!!"
message = excel_data['Message'][0]
print(message)
print(excel_data['Name'].tolist())
# Iterate excel rows till to finish
# for column in excel_data['Name'].tolist():
#     # Locate search box through x_path
#     search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'
#     person_title = wait.until(lambda driver:driver.find_element_by_xpath(search_box))
#
#     # Clear search box if any contact number is written in it
#     person_title.clear()
#
#     # Send contact number in search box
#     person_title.send_keys(str(excel_data['Contact'][count]))
#     count = count + 1
#
#     # Wait for 3 seconds to search contact number
#     time.sleep(10)
#
#     try:
#         # Load error message in case unavailability of contact number
#         element = driver.find_element_by_xpath('//*[@id="pane-side"]/div[1]/div/span')
#     except NoSuchElementException:
#         # Format the message from excel sheet
#         # message = "hellllllloooooo!!!!!!!!"
#         message = message.replace('{name}', column)
#         person_title.send_keys(Keys.ENTER)
#         actions = ActionChains(driver)
#        attachment_section = driver.find_element_by_xpath('//div[@title = "Attach"]')
#        attachment_section.click()
#        image_box = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
#        image_box.send_keys(file_path)
#        time.sleep(3)
#         # send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
#         # send_button=driver.find_element_by_class_name("_1w1m1")
#         # send_button.click()
#
#        send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
#
#         actions.send_keys(message)
#         actions.send_keys(Keys.ENTER)
#         actions.perform()
#         # send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
#         # actions.send_keys(Keys.ENTER)
#         # send_button.click()
#         # send_button.send_keys(Keys.ENTER)
#         # actions.perform()
#         person_title.clear()

# Iterate excel rows till to finish
for column in excel_data['Name'].tolist():
    # Locate search box through x_path
    search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'
    person_title = wait.until(lambda driver:driver.find_element_by_xpath(search_box))

    # Clear search box if any contact number is written in it
    person_title.clear()

    # Send contact number in search box
    time.sleep(5)
    person_title.send_keys(str(excel_data['Contact'][count]))
    count = count + 1

    # Wait for 3 seconds to search contact number
    time.sleep(5)

    try:
        # Load error message in case unavailability of contact number
        element = driver.find_element_by_xpath('//*[@id="pane-side"]/div[1]/div/span')
    except NoSuchElementException:
        # Format the message from excel sheet
        new_message = message.replace('{name}', column)
        print(new_message)
        time.sleep(5)
        person_title.send_keys(Keys.ENTER)
        actions = ActionChains(driver)

        attachment_section = driver.find_element_by_xpath('//div[@title = "Attach"]')
        attachment_section.click()
        image_box = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_box.send_keys(file_path)
        time.sleep(5)
        send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
        actions.send_keys(new_message)
        actions.send_keys(Keys.ENTER)
        actions.perform()
        time.sleep(5)
        # data - icon = "x"
        # close_button = driver.find_element_by_xpath('//span[@data-icon="x"]')
        # actions.send_keys(Keys.ENTER)
        # person_title.clear()


# Close Chrome browser
driver.quit()
