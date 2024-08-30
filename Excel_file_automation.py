import datetime
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time

chrome_options = Options()
chrome_options.binary_location = (r"C:\Users\Dream\PycharmProjects\pythonProject\Demo\chrome.exe")  # Specify the path to the Chrome binary
chrome_options.add_argument("--headless")
chrome_options.add_argument("--lang=en-US")
chrome_options.add_argument("--disable-gpu")

service = Service(r'C:\Users\Dream\PycharmProjects\pythonProject\Demo\chromedriver.exe')  # Specify the path to your chromedriver
driver = webdriver.Chrome(service=service, options=chrome_options)

today=datetime.datetime.now()
day_name=today.strftime("%A")

# Open the Excel workbook
wb = load_workbook('excel_file.xlsx')  # Update with your Excel file path

# Initialize variables to store the longest and shortest options
longest_option = ""
shortest_option = None

# Function to perform a Google search and get the suggestions
def google_search(keyword):
    global longest_option, shortest_option  # Declare these variables as global
    driver.get('https://www.google.com')
    search_box = driver.find_element(By.NAME, 'q')
    search_box.send_keys(keyword)
    time.sleep(3)  # Wait for the suggestions to load

    # Retrieve the search suggestions
    suggestions = driver.find_elements(By.CSS_SELECTOR, 'li span')

    # Loop through the suggestions to find the longest and shortest options
    for suggestion in suggestions:
        option_text = suggestion.text.strip()
        if not option_text:
            continue

        if len(option_text) > len(longest_option):
            longest_option = option_text

        if shortest_option is None or len(option_text) < len(shortest_option):
            shortest_option = option_text


for sheet in wb.sheetnames:
    if sheet==day_name:
        sheet = wb.active
        print(sheet)
        for row in range(2, sheet.max_row + 1):  # Assuming data starts from the second row
            keyword = sheet.cell(row=row, column=1).value
            if keyword:
                google_search(keyword)
            # Output the results
            print(keyword)
            sheet.cell(row=row, column=2).value = longest_option
            sheet.cell(row=row, column=3).value = shortest_option
# Save the updated Excel workbook
wb.save('excel_file.xlsx')

# Close the browser
driver.quit()