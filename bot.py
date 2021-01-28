import xlrd
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.select import Select

print('                       BOT STARTED...')
city = []
lic_type = []
d_path = ''
with open('download_path.txt') as file:
    d_path = file.read()
workbook = xlrd.open_workbook('input_data/input_data.xlsx')
sheet = workbook.sheet_by_index(0)

for row in range(sheet.nrows):
    city.append(sheet.cell_value(row, 0))
    lic_type.append(sheet.cell_value(row, 1))

wrong_city = []
for num in range(len(city)):
    print(city[num], ' - ', lic_type[num])
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    options.add_argument('start-maximized')
    preferances = {"download.default_directory": d_path}
    options.add_experimental_option("prefs", preferances)
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.get("https://cslb.ca.gov/OnlineServices/CheckLicenseII/ZipCodeSearch.aspx")
    sleep(1)

    try:
        driver.find_element_by_name('ctl00$MainContent$txtCity').send_keys(city[num])
        selection = Select(driver.find_element_by_name('ctl00$MainContent$ddlLicenseType'))
        selection.select_by_visible_text(lic_type[num])
        driver.find_element_by_name('ctl00$MainContent$btnZipCodeSearch').click()
        sleep(1)
        driver.find_element_by_name('ctl00$MainContent$ibExportToExcell').click()
        sleep(5)
        driver.close()
    except:
        driver.close()
        print(city[num], ' with ', lic_type[num], 'is wrong.')
        wrong_city.append(city[num])

with open('wrong_input/wrong_inputs.txt', 'w') as w_in:
    for i in range(len(wrong_city)):
        w_in.write(f'{wrong_city[i]}, ')

print('                       Files generated successfully...')
