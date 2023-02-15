from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get("https://en.wikipedia.org/wiki/RPA")
driver.maximize_window()

group = driver.find_elements("xpath", "//h2/span[@id]")
main_table = []
for i in range(len(group)):

    #Xpath for subgroup and Link
    currentGroupName = '//*[@id="mw-content-text"]/div[1]/ul[' + str(i + 1) + ']/li'
    currentGroupLink = '//*[@id="mw-content-text"]/div[1]/ul[' + str(i + 1) + ']/li//a'

    subgroupNames = driver.find_elements("xpath", currentGroupName)
    subgroupLinks = driver.find_elements("xpath", currentGroupLink)

    for j in range(len(subgroupNames)):
        #add each subgroup and its link to dict
        temp_row = {'Group': group[i].text,
                    'Subsection': subgroupNames[j].text,
                    'Link': subgroupLinks[j].get_attribute("href")}
        main_table.append(temp_row)

driver.quit()

dt_report = pd.DataFrame(main_table)
# Make pivot table
final_table=pd.pivot_table(dt_report, index=['Group','Subsection','Link'])
final_table = final_table.style.set_properties(**{'text-align': 'left'})
# Write report to file
final_table.to_excel('RPA_Task.xlsx',sheet_name='RPA_Information', index=True)

