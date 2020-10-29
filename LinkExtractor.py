from selenium import webdriver
from selenium.webdriver import firefox
from selenium.webdriver import chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.webdriver.support.expected_conditions import element_to_be_clickable
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
import time
from os import path

#THIS FUNCTION IS USED TO SETUP INITIAL DRIVER
def setup_driver():
    print("Which Browser Do you Have in Your Computer:")
    print("1. Mozilla Firefox")
    print("2. Google Chrome")

    while(1):
        choice = input("\n\tEnter Your Choice 1/2 : ")

        if choice == "1":
            break
        elif choice == "2":
            break
        
        print("\tWrong Number Entered, Try Again.")


    #firefox
    if choice == "1":
        firefox_options = firefox.options.Options()
        firefox_options.add_argument("--headless")
        firefox_path = path.abspath('geckodriver.exe')
        
        profile = webdriver.FirefoxProfile()
        profile.set_preference("general.useragent.override", "Mozilla/5.0 (X11; Linux x86_64; rv:80.0) Gecko/20100101 Firefox/80.0")

        driver = webdriver.Firefox(firefox_profile=profile,executable_path=firefox_path,options=firefox_options)

    elif choice == "2":
        chrome_options = chrome.options.Options()
        chrome_options.add_argument("--headless")
        chrome_path = path.abspath('chromedriver.exe')
        
        driver = webdriver.Chrome(executable_path=chrome_path,options=chrome_options)

    return driver

#THIS FUNCTION IS USED TO LOAD THE PAGE INITIALLY
def load_initial_page(driver,wait):
    #search_region =  driver.find_element_by_xpath("//input[@class='select2-search__field']")
    #search_region.send_keys("Zürich")
    #wait.until(presence_of_element_located((By.XPATH,"(//li[@class='select2-results__option'])[2]")))
    #select_zurich = driver.find_element_by_xpath("(//li[@class='select2-results__option'])[2]")
    #select_zurich.click()
    select_alter = driver.find_element_by_xpath("//button[@data-id='7DB3FD4D-F413-7498-DEFB416468D5914A']")
    select_alter.click()
    search_button = driver.find_element_by_xpath("//span[@class='btn-text']")
    search_button.click()

    #wait.until(presence_of_element_located((By.XPATH, "//img[@class=' lazyloaded']")))
    time.sleep(2)

    button_list = driver.find_element_by_xpath("//button[@class='view-list']")
    driver.execute_script("arguments[0].click();", button_list)
    #button_list.click()

    time.sleep(2)

    driver.set_window_size(1920, 1080)
    #driver.get_screenshot_as_file("screenshot.png")

#THIS FUNCTION IS USED TO GET ALL THE URLS
def get_urls(driver,index,output_file):
    all_urls = driver.find_elements_by_xpath("//article/a")
    num_all_urls = len(all_urls)
    temp = 0

    for i in range(index, num_all_urls):
        temp = temp+1

        string = all_urls[i].get_attribute("href")
        output_file.write(string)
        output_file.write("\n")
        print(f"{i+1} URL WRITTEN TO FILE")

    index += temp
    return index

#THIS FUNCTION IS USED TO CLICK ON SHOW MORE BUTTON
def show_more(driver,wait):    
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        show_more_button = driver.find_element_by_xpath("//button[@class='next btn']")
        show_more_button.click()
        time.sleep(2)
        return False
    except NoSuchElementException:
        return True
    except ElementNotInteractableException:
        return True

def main():

    #SETUP INITIAL DRIVER
    driver = setup_driver()
    
    #SEND THE INITIAL REQUEST
    wait = WebDriverWait(driver,10)
    driver.get("https://www.heiminfo.ch/institutionen")

    #LOADING THE INITIAL PAGE
    load_initial_page(driver,wait)

    #INITIAL INDEX OF URL GETS IS 0
    index = 0

    #CREATING AND OPENING A FILE FOR WRITING
    output_file = open(r"all_urls.txt","w")

    #GETTING THE URLS
    while(1):
        
        #GETTING THE URLS
        index = get_urls(driver,index,output_file)

        #CLICKING ON SHOW MORE
        end = show_more(driver,wait)

        #IF NO MORE TO SHOW
        if end == True:
            break

    #CLOSING THINGS
    driver.close()
    output_file.close()

if __name__ == "__main__":
    main()
