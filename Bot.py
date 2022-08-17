import win32com.client
from selenium import webdriver
import pyautogui
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import os
import winshell
import time

#Creates a folder for the images downloaded from the bot
def createFolder():
    newpath = 'C:\Documents\SavedImages'
    try:
        os.makedirs(newpath)
    except FileExistsError:
        pass

#Creates a shortcut for the folder
def saveFolder():
    desktop = winshell.desktop()
    path = os.path.join(desktop, 'SavedImages.lnk')
    target = "C:\Documents\SavedImages"
    icon = "C:\Documents\SavedImages"
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.IconLocation = icon
    shortcut.save()

#Function for downloading the pictures
def downloadPictures():
    amount = 0
    downloadCount = 0
    input = True
    screenWidth = driver.execute_script('return (window.innerWidth)')
    screenHeight = driver.execute_script('return (window.innerHeight)')

    while (amount == 0):
        driver.execute_script("var a = prompt('Please input the amount of pictures you would like to save!', '');document.body.setAttribute('data-id', a)")
        time.sleep(5)
        amount = int(driver.find_element(By.TAG_NAME, "body").get_attribute('data-id'))

        if(amount == 0):
            driver.execute_script("var a = alert('The Imager Saver Bot will be shutting off!')")
            input = False
            break
        time.sleep(1)

    while(input):
        pyautogui.dragTo(screenWidth / 1.5, screenHeight / 1.3, 0)
        pyautogui.rightClick()
        for x in range(0, 2):
            pyautogui.press('down')
        pyautogui.press('enter')

        if (downloadCount < amount):
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.5)
            downloadCount += 1
            if (downloadCount == amount):
                input = False
                break

            pyautogui.press('right')

#Main Code:
createFolder()
options = webdriver.ChromeOptions()
s = Service('C:\\WebDrivers\\chromedriver.exe')
driver = webdriver.Chrome(service = s, options=options)

driver.command_executor._commands['send_command'] = (
    'POST', '/session/$sessionId/chromium/send_command')
download_path = 'C:\Documents\SavedImages'
params = {
    'cmd': 'Page.setDownloadBehavior',
    'params': { 'behavior': 'allow', 'downloadPath': download_path }
}
driver.execute("send_command", params)

#Currently the code only works for saving pictures on Messenger
driver.get("https://www.messenger.com/login")
driver.maximize_window()

#The 60 seconds sleep time is for a user to login in to their social media account and get to the chat where they want to start saving images
time.sleep(60)

driver.execute_script(
        "var a = confirm('Would you like to start the Image Saver Bot?', '');document.body.setAttribute('data-id', a)")
time.sleep(6)
start = driver.find_element(By.TAG_NAME, "body").get_attribute('data-id')

if (start == "true"):
    start = True
else:
    start = False

if (start):
    saveFolder()
    downloadPictures()
    driver.execute_script("var a = alert('The Imager Saver Bot has finished saving the pictures!')")
else:
    driver.execute_script("var a = alert('The Imager Saver Bot is not running.')")
