from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from data_generator import get_data, time_constrained_function
from state_locater import area_codes, state_to_abbreviation
from data_archive import check_case_completion, add_data_to_excel

error_count = 0
completed_count = 0
completed_old = 0
invoice_number = ''


def offset_rows():
    rows = int(input("How many cases have been completed?: "))
    global completed_old
    completed_old += rows
    for num in range(rows):
        time.sleep(0.05)
        print("Down " + str(num))
        actions.send_keys(Keys.DOWN)
        actions.perform()


# If WMS SSL Certificate is Expired
def wms_certificate_expired():
    print("\nWMS Certificate is Expired.")
    driver.get('chrome://net-internals/#hsts')
    print('Approving WMS...')
    driver.find_element(By.XPATH, '//*[@id="domain-security-policy-view-delete-input"]').send_keys('currentbyar.wms.oncospark.com')
    driver.find_element(By.XPATH, '//*[@id="domain-security-policy-view-delete-submit"]').click()
    # driver.get('https://currentbyar.wms.oncospark.com/Account/Login')
    print('Verifying WMS...')
    # driver.find_element(By.XPATH, '//*[@id="details-button"]').click()
    # driver.find_element(By.XPATH, '//*[@id="proceed-link"]').click()


# Extracts what state the provider is located in by utilizing the area code of the phone number of the provider.
def get_state(field_data):
    try:
        digits = ''.join([char for char in field_data[2] if char.isdigit()][:3])
        area_code = digits[:3]
        region = area_codes.get(area_code)
        if region is not None:
            state_abbr = state_to_abbreviation.get(region)
            return state_abbr
        else:
            return "Select state"
    except Exception as e:
        return "Select state"


# Fills data into their respective fields on the website.
def fill_in_fields(field_data, location):
    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_DOB"]').clear()
    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_DOB"]').send_keys(field_data[0])

    states = Select(driver.find_element(By.XPATH, '//*[@id="Modal_Edit_State"]'))
    states.select_by_visible_text(location)

    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_PhysicianName"]').clear()
    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_PhysicianName"]').send_keys(field_data[1])

    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_phyPhoneNumber"]').clear()
    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_phyPhoneNumber"]').send_keys(field_data[2])

    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_PhyFaxNumber"]').clear()
    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_PhyFaxNumber"]').send_keys(field_data[3])

    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_Request"]').clear()
    driver.find_element(By.XPATH, '//*[@id="Modal_Edit_Request"]').send_keys(field_data[4])


# Navigates to the WMS Website, Asks for user login credentials. Attempts to log in and navigate to the claims.
def access_claims():
    # Navigates to and logs in to WMS
    print("\nAccessing WMS...")
    driver.get('https://currentbyar.wms.oncospark.com/Account/Login')
    # username = input("WMS Username: ")
    # password = input("WMS Password: ")
    driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/form/div[1]/input').send_keys("Nikki")
    driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/form/div[2]/input').send_keys("Nikki@1")
    driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/form/button').click()
    time.sleep(3)

    try:
        # Navigate to where the claims are located
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div[1]/ul[1]/li[2]/a').click()
        time.sleep(3)
    except Exception as e:
        print("Incorrect Login. Please Try Again.")
        access_claims()


try:
    # Initializes browser and actions
    print("Installing Chrome Services...")
    # service = Service(service=ChromeDriverManager().install())
    service = Service("driver/chromedriver")
    print("Initializing Chrome...")
    options = webdriver.ChromeOptions()
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument('--headless=new')
    driver = webdriver.Chrome(service=service, options=options)
    actions = ActionChains(driver)

    wms_certificate_expired()

    # Navigates to WMS and Accesses the Claims
    access_claims()

    print("Successful login.\n")

    # Sorts cases by invoice number
    driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div[2]/div[1]/div[1]/div[1]/div/div[3]/div').click()
    time.sleep(3)


    # Selects the first case on WMS
    try:
        driver.find_element(By.XPATH, '//*[@id="myGrid"]/div/div[2]/div[2]/div[3]/div[2]/div/div/div[1]').click()
    except Exception as e:
        print("There are no cases assigned to this user.")
        quit()


    # Offsets completed rows.
    offset_rows()

    # Loop through all the cells
    while error_count < 5:

        # Opens the case in WMS
        actions.send_keys(Keys.SPACE)
        actions.perform()
        time.sleep(0.6)

        # Checks if this case is the same one as the previous case. Allows us to know when the program has gotten to the bottom of all the cases.
        new_invoice_number = driver.find_element(By.XPATH, '//*[@id="Modal_ClaimInfo_TabContent_Claiminfo"]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td/span').text
        if new_invoice_number == invoice_number:
            error_count += 1
        else:
            error_count = 0
        # Get the patient name and patient number
        patient_name = driver.find_element(By.XPATH, '//*[@id="Modal_Edit_P_PatientName"]').get_attribute("value")
        patient_number = driver.find_element(By.XPATH, '//*[@id="Modal_ClaimInfo_TabContent_Claiminfo"]/div[1]/div/div/div[1]/div/table/tbody/tr[2]/td').text

        # Check if this case has already been completed.
        invoice_number = new_invoice_number
        print("\nWorking on " + patient_name + ". Invoice number: " + invoice_number)
        validity = check_case_completion(invoice_number)
        if not validity:
            # Selects the next case
            completed_old += 1
            print("Number of old cases: " + str(completed_old))
            driver.find_element(By.XPATH, '//*[@id="Modal_Claimdetail_Tabs"]/li[6]/button/span[1]').click()
            time.sleep(0.8)
            actions.send_keys(Keys.DOWN)
            actions.perform()
            continue

        # Extract the comments from the case
        print("extracting comments...")
        comments = (driver.find_element(By.XPATH, '//*[@id="Modal_Comments_List_Div"]')).find_elements(By.XPATH, "./div")
        for x in range(len(comments)):
            # noinspection PyTypeChecker
            comments[x] = comments[x].find_element(By.XPATH, './div/div/p').text

        print("reformatting comments with openAI...")
        # Gets the data in the correct format and fills it into WMS and the Excel
        data = time_constrained_function(get_data(), comments, 10)
        state = get_state(data)

        # print("adding data to database...")
        # add_data_to_excel(invoice_number, patient_number, patient_name, state, data)

        print("filling in fields on WMS...")
        fill_in_fields(data, state)

        # Saves the data into the WMS Case
        print("saving in WMS...")
        driver.find_element(By.XPATH, '//*[@id="Modal_Edit_SaveBtn"]').click()
        time.sleep(1)
        completed_count += 1

        # Closes the current WMS Case
        print("Finished Case. Number of new cases completed: " + str(completed_count) + ". Number of total cases: " + str(completed_old + completed_count) + ".")
        driver.find_element(By.XPATH, '//*[@id="Modal_Claimdetail_Tabs"]/li[6]/button/span[1]').click()
        time.sleep(1)

        # Selects the next case
        actions.send_keys(Keys.DOWN)
        actions.perform()
        time.sleep(2)

    print("Finished All Cases. Number of cases completed: " + str(completed_count + completed_old))

    driver.quit()

except Exception as e:
    driver.quit()
    print(e)
    print("\nError Occured... quit Chrome Driver.")
