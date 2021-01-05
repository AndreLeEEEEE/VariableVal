import openpyxl
import datetime
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains

def locate_by_name(web_driver, name):
    """Clicks on something by name."""
    WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.NAME, name))).click()
    # Returns nothing

def get_by_name(web_driver, name):
    """Returns something by name."""
    return WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.NAME, name)))

def locate_by_id(web_driver, id):
    """Clicks on something by id."""
    WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.ID, id))).click()
    # Returns nothing

def get_by_id(web_driver, id):
    """Returns something by id."""
    return WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.ID, id)))

def locate_by_class(web_driver, class_name):
    """Clicks on something by class name."""
    WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, class_name))).click()
    # Returns nothing

def get_by_class(web_driver, class_name):
    """Returns something by class name."""
    return WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, class_name)))

def select_drop(web_driver, id, value):
    """Chooses an option by value in a dropdown box by id."""
    select = Select(get_by_id(web_driver, id))
    select.select_by_value(value)
    # Returns nothing

def standard(date):
    """Make sure that the inputted dates are valid."""
    temp_date = date.split('/')  # Failure to perform the split will result in an exception
    for go in range(3):  # Convert all numbers to ints
        temp_date[go] = int(temp_date[go])

    return datetime.date(temp_date[2], temp_date[0], temp_date[1])

def extract(driver, time, span):
    select_drop(driver, "fltGroupBy", "BuildingInvType")
    # Click to unselect these
    driver.find_element_by_xpath("//input[@value='CConsign']").click()
    driver.find_element_by_xpath("//input[@value='Supplies']").click()
    driver.find_element_by_xpath("//input[@value='NonStd']").click()
    for index in range(span):
        # For the inventory date
        inv_box = get_by_id(driver, "fltDate_DTE_RQD")
        inv_box = get_by_id(driver, "fltTime_TME")
        inv_box.send_keys("5:00 PM")
        # For the cost date
        cost_box = get_by_id(driver, "fltCostDate_DTE_RQD")
        cost_box = get_by_id(driver, "fltCostTime_TME")
        cost_box.send_keys("5:00 PM")

        locate_by_id(driver, "SubmitLink")

        time += datetime.timedelta(days=1)  # Increment the day

    input("Program Pause")

def Plex(time1, span):
    try:
        # Getting into Plex
        driver = webdriver.Chrome("chromedriver.exe")
        driver.get("https://www.plexonline.com/modules/systemadministration/login/index.aspx?")
        driver.find_element_by_name("txtUserID").send_keys("w.Andre.Le")
        driver.find_element_by_name("txtPassword").send_keys("OokyOoki2")
        driver.find_element_by_name("txtCompanyCode").send_keys("wanco")
        locate_by_id(driver, "btnLogin")
        # Since logging in through selenium opens up a new window instead of changing the current login screen,
        # go to the new screen
        driver.switch_to.window(driver.window_handles[1])
        action = ActionChains(driver)
        action.key_down(Keys.CONTROL).send_keys('M').key_up(Keys.CONTROL).perform()
        action = 404
        time.sleep(1)
        action = ActionChains(driver)
        action.send_keys("Inventory Valuation at Standard Cost").send_keys(Keys.RETURN).perform()
        action = 404
        time.sleep(1)
        action = ActionChains(driver)
        action.send_keys(Keys.DOWN).send_keys(Keys.RETURN).perform()
        time.sleep(1)
        info = extract(driver, time1, span)
    except Exception as e:
        print("An error was encountered:")
        print(e)
    finally:  # Whether the operation was erroneous or successful, quit the driver
        driver.quit()
    
    try:
        return info
    except:
        print("Data missing due to program bug")
        exit()

def main():
    valid = False
    while not valid:
        try:
            date1 = input("Enter in the first date(MM/DD/YYYY): ")
            date1 = standard(date1)
            date2 = input("Enter in the second date(MM/DD/YYYY): ")
            date2 = standard(date2)
            valid = True
        except:
            print("At least one of the dates entered didn't have a valid format or is impossible")

    if date1 > date2:  # If date1 is more recent
        days = (date1 - date2).days + 1
        date = date2
    elif date1 < date2:  # If date2 is more recent
        days = (date2 - date1).days + 1
        date = date1
    else:  # Both dates are the same
        days = 1
        date = date1
    # The "extra" 1 in days is because the first day is included

    data = Plex(date, days)

main()