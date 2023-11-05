"""Author: Uygar Tolga Kara - Date: Wed Sep 27 10:40:24 2023."""
# -*- coding: utf-8 -*-

from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from time import sleep
from json import dump

# %% Data Source - FaultCodes.co

geckodriver_path = "geckodriver.exe"
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

options = Options()
service = Service(geckodriver_path)
options.binary_location = firefox_path

wd = webdriver.Firefox(service=service, options=options)
webpage = r"https://faultcodes.co/code/"

wd.get(webpage)

input("Press enter to continue...")

web_dict = {}

# Find all buttons with the specific class
buttons_ref = wd.find_elements(By.CSS_SELECTOR, '.btn.btn-success.btn-sm.px-3.py-2')

# Iterate through the list of buttons and click each one
for index in range(len(buttons_ref)):
    buttons = wd.find_elements(By.CSS_SELECTOR, '.btn.btn-success.btn-sm.px-3.py-2')
    if "READ MORE ABOUT" in buttons[index].text:

        code = buttons[index].text.rsplit(maxsplit=1)[-1]
        print(code)

        web_dict[code] = {}

        buttons[index].click()
        sleep(2)

        # Find the paragraph element that contains the seriousness rating
        severity_paragraph = wd.find_element(By.XPATH, "//p[contains(@class, 'lead mb-2') and contains(text(), 'seriousness of')]")
        # Extract the text from the <strong> tag inside the paragraph
        severity_text = severity_paragraph.find_element(By.TAG_NAME, "strong").text
        # Extract the rating number (assumes the format is always 'X/10')
        severity_rating = severity_text.split('/')[0]
        # Add severity rating to nested dictionary
        web_dict[code]["Criticality"] = severity_rating

        # Find the paragraph element with the class "lead mb-0"
        description_paragraph = wd.find_element(By.CLASS_NAME, 'lead.mb-0')
        # Extract the text content of the paragraph
        description_text = description_paragraph.text
        # Add description to nested dictionary
        web_dict[code]["Description"] = description_text

        # Define the ids of the h2 elements
        sections = [
            "long-description",
            "other-signs",
            "the-problem",
            "fixes",
            "seriousness"]

        # Function to get the paragraphs following an h2 until the next h2
        def get_paragraphs_after_header(header_id):
            """pass."""
            header = wd.find_element(By.ID, header_id)
            paragraphs = wd.execute_script("""
                var header = arguments[0];
                var paragraphs = [];
                var sibling = header.nextElementSibling;
                while(sibling && sibling.tagName.toLowerCase() !== 'h2') {
                    if(sibling.tagName.toLowerCase() === 'p') {
                        paragraphs.push(sibling.textContent);
                    }
                    sibling = sibling.nextElementSibling;
                }
                return paragraphs;
            """, header)
            return paragraphs

        # Loop through each section id and get the content
        for section in sections:
            web_dict[code][section] = get_paragraphs_after_header(section)

        wd.back()
        sleep(2)

# Close the WebDriver
wd.quit()

# Write the dictionary to a JSON file
with open("faultcodes_co_powertrain.json", "w") as file:
    dump(web_dict, file)

# %% Data Source - odxdata.com

geckodriver_path = "geckodriver.exe"
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

options = Options()
service = Service(geckodriver_path)
options.binary_location = firefox_path

wd = webdriver.Firefox(service=service, options=options)
webpage = r"https://odxdata.com/diagnostic-trouble-codes-chassis-or-c-codes/"

wd.get(webpage)

# Wait for the page to load (adjust the sleep time as necessary)
sleep(5)

# Initialize an empty dictionary to store the fault codes and descriptions
fault_codes = {}

# Find the table
table = wd.find_element(By.CLASS_NAME, 'wp-block-table')

# Iterate through the rows of the table
for row in table.find_elements(By.TAG_NAME, 'tr'):
    # Find the columns (td elements) in this row
    columns = row.find_elements(By.TAG_NAME, 'td')
    # Check if there are enough columns
    if len(columns) >= 2:
        # The first column is the fault code, the second is the description
        code = columns[0].text
        description = columns[1].text
        # Add them to the dictionary
        fault_codes[code] = description

# Print the fault codes and descriptions
for code, description in fault_codes.items():
    print(f'{code}: {description}')

# Write the dictionary to a JSON file
with open("odxdata_chassis.json", "w") as file:
    dump(fault_codes, file)

# -------------------------------------------------------------

webpage = r"https://odxdata.com/diagnostic-trouble-codes-dtcs/"
wd.get(webpage)

# Wait for the page to load (adjust the sleep time as necessary)
sleep(5)

# Initialize an empty dictionary to store the fault codes and descriptions
fault_codes = {}

# Find the table
table = wd.find_element(By.CLASS_NAME, 'wp-block-table')

# Iterate through the rows of the table
for row in table.find_elements(By.TAG_NAME, 'tr'):
    # Find the columns (td elements) in this row
    columns = row.find_elements(By.TAG_NAME, 'td')
    # Check if there are enough columns
    if len(columns) >= 2:
        # The first column is the fault code, the second is the description
        code = columns[0].text
        description = columns[1].text
        # Add them to the dictionary
        fault_codes[code] = description

# Print the fault codes and descriptions
for code, description in fault_codes.items():
    print(f'{code}: {description}')

# Write the dictionary to a JSON file
with open("odxdata_powertrain.json", "w") as file:
    dump(fault_codes, file)

# --------------------------------------------------------------

webpage = r"https://odxdata.com/diagnotic-trouble-codes-body-or-b-codes/"
wd.get(webpage)

# Wait for the page to load (adjust the sleep time as necessary)
sleep(5)

# Initialize an empty dictionary to store the fault codes and descriptions
fault_codes = {}

# Find the table
table = wd.find_element(By.CLASS_NAME, 'wp-block-table')

# Iterate through the rows of the table
for row in table.find_elements(By.TAG_NAME, 'tr'):
    # Find the columns (td elements) in this row
    columns = row.find_elements(By.TAG_NAME, 'td')
    # Check if there are enough columns
    if len(columns) >= 2:
        # The first column is the fault code, the second is the description
        code = columns[0].text
        description = columns[1].text
        # Add them to the dictionary
        fault_codes[code] = description

# Print the fault codes and descriptions
for code, description in fault_codes.items():
    print(f'{code}: {description}')

# Write the dictionary to a JSON file
with open("odxdata_body.json", "w") as file:
    dump(fault_codes, file)

wd.quit()

# %% Data Source - bads.lt/en

geckodriver_path = "geckodriver.exe"
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

options = Options()
service = Service(geckodriver_path)
options.binary_location = firefox_path

wd = webdriver.Firefox(service=service, options=options)
webpage = r"https://bads.lt/en/blog/technical-information/obd2-codes-and-meanings"

wd.get(webpage)

# Wait for the page to load
sleep(5)

# Initialize an empty dictionary to store the codes and descriptions
codes_dict = {}

# Find the <h4> element
h4 = wd.find_element(By.XPATH, "//h4[contains(text(), 'Want to know more about codes and their meanings?')]")

# Iterate through the sibling elements
for sibling in h4.find_elements(By.XPATH, "following-sibling::*"):
    if sibling.tag_name.lower() == 'p':
        if len(sibling.text.split(' ', 1)) == 2:
            # Split the text content from the first whitespace
            code, description = sibling.text.split(' ', 1)
            # Add them to the dictionary
            codes_dict[code] = description
    else:
        # Stop if we've reached a different type of element
        break

# Print the codes and descriptions
for code, description in codes_dict.items():
    print(f'{code}: {description}')

# Write the dictionary to a JSON file
with open("bads_lt_allcodes.json", "w") as file:
    dump(codes_dict, file)

# Close the WebDriver
wd.quit()

# %% Data Source - obdii.pro

geckodriver_path = "geckodriver.exe"
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

options = Options()
service = Service(geckodriver_path)
options.binary_location = firefox_path

wd = webdriver.Firefox(service=service, options=options)
webpage = r"https://obdii.pro/en/model/all"

wd.get(webpage)

# Wait for the page to load
input("Press enter to continue...")
# sleep(3)

# Find the definition div
definition_div = wd.find_element(By.CLASS_NAME, 'definition')

# Initialize an empty dictionary to store the data
codes_dict = {}

# Find all div elements with class 'code'
code_divs_ref = wd.find_elements(By.CLASS_NAME, 'code')

for index in range(len(code_divs_ref)):

    code_divs = wd.find_elements(By.CLASS_NAME, 'code')
    link = code_divs[index].find_element(By.TAG_NAME, 'a')

    # Get the text
    code = link.text

    # Click the link (uncomment the next line if you need to actually click the links)
    link.click()

    # Add a wait time or handle the new page load as necessary here
    # sleep(2)

    codes_dict[code] = {}

    # ---------------------------------------------------------------------------------

    # Find the definition div
    definition_div = wd.find_element(By.CLASS_NAME, 'definition')

    model_divs = definition_div.find_elements(By.CLASS_NAME, 'model')

    for model_div in model_divs:
        model = model_div.find_element(By.TAG_NAME, 'h2').text
        def_div = model_div.find_element(By.XPATH, 'following-sibling::div[@class="def"]')
        definition = def_div.text if def_div else "Definition not found"
        codes_dict[code][model] = definition

    # ---------------------------------------------------------------------------------

    # Navigate back to the original page if you clicked a link
    wd.back()

    # Wait for the page to reload if necessary
    # sleep(2)

# Print the codes and links
for code, link in codes_dict.items():
    print(f'{code}: {link}')

# Write the dictionary to a JSON file
with open("obdii_allcodes.json", "w") as file:
    dump(codes_dict, file)

# Close the WebDriver
wd.quit()

# %%

geckodriver_path = "geckodriver.exe"
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

options = Options()
service = Service(geckodriver_path)
options.binary_location = firefox_path

wd = webdriver.Firefox(service=service, options=options)
webpage = r"https://repairpal.com/obd-ii-code-chart/bxxxx/1"

wd.get(webpage)
input("continue")

# Wait for the divs to be loaded (modify the CSS selector as necessary)
wait = WebDriverWait(wd, 10)
elements_ref = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div > a.ng-flex')))

# Iterate through each div
for index in range(len(elements_ref)):
    elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div > a.ng-flex')))
    # Get the code and description
    code = elements[index].find_element(By.CSS_SELECTOR, '.code').text
    description = elements[index].find_element(By.CSS_SELECTOR, '.description').text

    print("Code:", code)
    print("Description:", description)

    # Click the link
    elements[index].click()

    # Add a delay or wait for a specific condition here if necessary before continuing to the next page
    # ...

    # Navigate back to the original page if necessary
    wd.back()

# Close the browser window
wd.quit()

# %% FAULTY

geckodriver_path = "geckodriver.exe"
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

options = Options()
service = Service(geckodriver_path)
options.binary_location = firefox_path

wd = webdriver.Firefox(service=service, options=options)
webpage = r"https://www.transmissionrepaircostguide.com/transmission-diagnostic-trouble-codes-dtc/"

wd.get(webpage)

# Wait for the table to be loaded
wait = WebDriverWait(wd, 10)
table = wait.until(EC.presence_of_element_located((By.ID, 'tablepress-109')))

# Find all the rows in the table body
rows = table.find_elements(By.TAG_NAME, 'tr')

code_dict = {}

# Iterate through each row
for idx in range(len(rows)):

    # Wait for the table to be reloaded before continuing to the next row
    table = wait.until(EC.presence_of_element_located((By.ID, 'tablepress-109')))

    # Find all the rows in the table body
    rows = table.find_elements(By.TAG_NAME, 'tr')

    # Get the cells in the row
    cells = rows[idx].find_elements(By.TAG_NAME, 'td')

    # Ensure there are two cells in the row before proceeding
    if len(cells) == 2:
        # Get the text from the cells
        code = cells[0].text
        description = cells[1].text

        code_dict[code] = {}
        code_dict[code]["Description"] = description

        # Print the code and description
        print("Code:", code)
        print("Description:", description)

        # Click the link in the first cell
        link = cells[0].find_element(By.TAG_NAME, 'a')
        link.click()

        # -----------------------------------------------------------------------------

        # Iterate through all the h2, p, and ul elements
        # Initialize the dictionary with predefined keys

        try:
            code_dict[code]["How Serious is the Code?"] = wd.find_element(By.XPATH, "//h2[text()='How Serious is the Code?']/following-sibling::p[1]").text
        except NoSuchElementException:
            pass
        except Exception as e:
            print("ERROR")
            break

        try:
            code_dict[code]["Symptoms"] = [li.text for li in wd.find_element(By.XPATH, "//h2[text()='Symptoms']/following-sibling::ul[1]").find_elements(By.TAG_NAME, "li")]
        except NoSuchElementException:
            pass
        except Exception as e:
            print("ERROR")
            break

        try:
            code_dict[code]["Causes"] = [li.text for li in wd.find_element(By.XPATH, "//h2[text()='Causes']/following-sibling::ul[1]").find_elements(By.TAG_NAME, "li")]
        except NoSuchElementException:
            pass
        except Exception as e:
            print("ERROR")
            break

        try:
            code_dict[code]["How to Diagnose the Code?"] = wd.find_element(By.XPATH, "//h2[text()='How to Diagnose the Code?']/following-sibling::p[1]").text
        except NoSuchElementException:
            pass
        except Exception as e:
            print("ERROR")
            break

        try:
            code_dict[code]["Common Mistakes When Diagnosing"] = wd.find_element(By.XPATH, "//h2[text()='Common Mistakes When Diagnosing']/following-sibling::p[1]").text
        except NoSuchElementException:
            pass
        except Exception as e:
            print("ERROR")
            break

        try:
            code_dict[code]["What Repairs Will Fix"] = [li.text for li in wd.find_element(By.XPATH, "//h2[starts-with(text(), 'What Repairs Will Fix')]/following-sibling::ul[1]").find_elements(By.TAG_NAME, "li")]
        except NoSuchElementException:
            pass
        except Exception as e:
            print("ERROR")
            break

        # -----------------------------------------------------------------------------

        # Add a delay or wait for a specific condition here if necessary before continuing to the next page
        # For example, wait for 2 seconds
        sleep(1)

        # Navigate back to the original page
        wd.back()

        sleep(1)

# Write the dictionary to a JSON file
with open("transmissionrepaircostguide_allcodes.json", "w") as file:
    dump(code_dict, file)


# Close the browser window
wd.quit()

# %%

url = r"https://partsavatar.ca/check-engine-light-scan-read-all-obd-2-trouble-codes-p0-powertrain-p1-p2-diagnose-repair-fix-engine-error"
geckodriver_path = "geckodriver.exe"
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

options = Options()
service = Service(geckodriver_path)
options.binary_location = firefox_path

wd = webdriver.Firefox(service=service, options=options)
wd.get(url)

# Give the page some time to load (adjust as necessary)
sleep(5)

# Locate all <a> elements
links = wd.find_elements(By.XPATH, "//a")

# Filter links based on text length
filtered_links = [link.get_attribute("href") for link in links if len(link.text) == 5]

# --------------------------------------------------------------------------------------

geckodriver_path = "geckodriver.exe"
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

options = Options()
service = Service(geckodriver_path)
options.binary_location = firefox_path

wd = webdriver.Firefox(service=service, options=options)

# Iterate through the filtered links
for link in filtered_links:
    wd.get(link)
    sleep(5)

    # -------------------------------------------------------------------------------

    # Locate the title element based on its containing text
    title_xpath = "//span[contains(., 'What') and contains(., 'causes') and contains(., 'this problem')]"
    title_element = wd.find_element(By.XPATH, title_xpath)

    # Print the title text
    print("Title:", title_element.text)

    # Locate the parent <p> of the title element, then find the following <ul> and its <li> children
    list_items = wd.find_elements(By.XPATH, f"{title_xpath}/ancestor::p/following-sibling::ul[1]/li/span[@lang='EN-IN']")

    # Extract text from each list item and save it in a list
    causes = [item.text for item in list_items]
    causes = [cause for cause in causes if cause != ""]

    # Print the list of causes
    print("Causes:")
    for cause in causes:
        print("-", cause)



    # -------------------------------------------------------------------------------

# Close the browser window
wd.quit()

#%% Joining JSON Files



