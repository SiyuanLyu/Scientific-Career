import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

input_path = "Name-list.xlsx"
chromedriver_path = r"Folder\chromedriver.exe"


chrome_options = Options()
chrome_options.add_argument("--start-maximized")


df = pd.read_excel(input_path)
results = []

def get_award_detail(driver, xpath):
    try:
        return driver.find_element(By.XPATH, xpath).text.strip()
    except:
        return "N/A"

driver = webdriver.Chrome(service=Service(chromedriver_path), options=chrome_options)

for idx, row in df.iterrows():
    first_name = str(row["FirstName"]).strip()
    last_name = str(row["FamilyName"]).strip()
    full_name = str(row["Name"]).strip()
    print(f"Search for：{full_name} ...")

    try:
        driver.get("https://www.nsf.gov/awardsearch/advancedSearch.jsp")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "PIFirstName")))

        driver.find_element(By.ID, "PIFirstName").clear()
        driver.find_element(By.ID, "PIFirstName").send_keys(first_name)
        driver.find_element(By.ID, "PILastName").clear()
        driver.find_element(By.ID, "PILastName").send_keys(last_name)
        checkbox = driver.find_element(By.ID, "Field2")
        if not checkbox.is_selected():
            checkbox.click()
        # Checkbox for ExpiredAwards if needed
        expired_checkbox = driver.find_element(By.ID, "ExpiredAwards")
        if not expired_checkbox.is_selected():
            expired_checkbox.click()

        driver.find_element(By.XPATH, "//input[@type='submit' and @value='Search']").click()
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)


        try:
            driver.find_element(By.ID, "noResultsFoundDiv")
            results.append({
                "Name": full_name,
                "FirstName": first_name,
                "FamilyName": last_name,
                "Status": "Not Found"
            })
            print("No search results")
            continue
        except:
            pass

        # Scrape links for results
        award_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'showAward?AWD_ID')]")
        print(f"{len(award_links)} fundings are found")

        for i, link in enumerate(award_links):
            try:
                award_title = link.get_attribute("title").split(";")[0]
                link.click()
                time.sleep(2)
                driver.switch_to.window(driver.window_handles[-1])

                result = {
                    "Name": full_name,
                    "FirstName": first_name,
                    "FamilyName": last_name,
                    "Award Title": award_title,
                    "Recipient": get_award_detail(driver, "//td[strong[contains(text(), 'Recipient')]]/following-sibling::td[1]"),
                    "Start Date": get_award_detail(driver, "//td[strong[contains(text(), 'Start Date')]]/following-sibling::td[1]"),
                    "End Date": get_award_detail(driver, "//td[strong[contains(text(), 'End Date')]]/following-sibling::td[1]"),
                    "Total Intended Award Amount": get_award_detail(driver, "//strong[contains(normalize-space(.), 'Total Intended Award Amount')]/parent::td/following-sibling::td"),
                    "NSF Program(s)": get_award_detail(driver, "//td[strong[contains(text(), 'NSF Program(s)')]]/following-sibling::td[1]"),
                    "UEI": get_award_detail(driver, "//td[strong[contains(text(), 'Unique Entity Identifier')]]/following-sibling::td[1]"),
                    "History of Investigator": get_award_detail(driver, "//td[strong[contains(text(), 'History of Investigator')]]/following-sibling::td[1]"),
                    "Status": f"Found #{i+1}"
                }

                results.append(result)
                print(f"Scrape funding {i+1}")
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

                # Refresh the page
                award_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'showAward?AWD_ID')]")

            except Exception as e:
                print(f"Fail for funding {i+1}）：{e}")
                try:
                    driver.switch_to.window(driver.window_handles[0])
                except:
                    pass
                continue

    except Exception as e:
        print(f"Error for {full_name}: {e}")
        continue


output_path = "Result.xlsx"
pd.DataFrame(results).to_excel(output_path, index=False)
print(f"Save to：{output_path}")

driver.quit()