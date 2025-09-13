import pandas as pd
import time
import random
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

EXCEL_PATH = "List of Names.xlsx"
OUTPUT_PATH = "ProQuest_Search_Results.xlsx"

CHROMEDRIVER_PATH = r"Folder\chromedriver.exe"

df = pd.read_excel(EXCEL_PATH)
names = df["Name"].dropna().tolist()

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=chrome_options)
driver.get("https://www.proquest.com/dissertations")
time.sleep(5)

# Cookie button
try:
    accept_button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Accept')]")
    ))
    accept_button.click()
    time.sleep(2)
except:
    pass


papers = []


for name in names:
    search_query = f"AU({name})"
    print(f"Searching: {search_query}")
    
    try:
        # Back to search page and wait for 5 sec
        driver.get("https://www.proquest.com/dissertations")
        time.sleep(5)

        
        search_box = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.NAME, "searchTerm"))
        )
        
        # Clear search box
        for _ in range(3):  # Try 3 times
            search_box.clear()
            search_box.send_keys(Keys.CONTROL + "a")
            search_box.send_keys(Keys.DELETE)
            time.sleep(1)
            if search_box.get_attribute("value") == "":
                break
        
        search_box.send_keys(search_query)
        search_box.send_keys(Keys.RETURN)
        time.sleep(5)

        # No results
        try:
            no_results = driver.find_element(By.XPATH, "//*[@id='mainContentLeft']/div[1]/div/p[1]/span[@class='error_message']")
            if "found 0 results" in no_results.text:
                print(f"No search result for {name}")
                papers.append({"Name": name, "Status": f"No search result for {name}"})
                continue
        except:
            print(f"Find {name}")
        
        # Scrape the information needed
        paper_links = []
        while True:
            results = driver.find_elements(By.XPATH, ".//a[contains(@title, 'Abstract/Details')]")
            for result in results:
                paper_links.append(result.get_attribute("href"))
            try:
                next_button = driver.find_element(By.XPATH, "//a[@title='Next Page']")
                driver.execute_script("arguments[0].click();", next_button)
                time.sleep(5)
            except:
                break

        print(f"{len(paper_links)} papers are found")

        # Visit details page
        for paper_url in paper_links:
            driver.get(paper_url)
            time.sleep(5)
            
            def get_text(xpath1, xpath2=None):
                try:
                    return WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, xpath1))
                    ).text.strip()
                except:
                    if xpath2:
                        try:
                            return WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, xpath2))
                            ).text.strip()
                        except:
                            return "N/A"
                    return "N/A"
            
            title = get_text("//div[contains(text(), 'Title')]/following-sibling::div")
            author = get_text("//div[contains(text(),'Author')]/following-sibling::div")
            degree_date = get_text("//div[contains(text(),'Degree date')]/following-sibling::div")
            university = get_text("//div[contains(text(),'University/institution')]/following-sibling::div")
            degree = get_text("//div[normalize-space()='Degree']/following-sibling::div/span")
            subjects = get_text("//div[contains(text(),'Subject')]/following-sibling::div")
            advisor = get_text("//div[contains(text(),'Advisor')]/following-sibling::div")
            dissertation_number = get_text("//div[contains(text(),'Dissertation/thesis number')]/following-sibling::div")
            document_id = get_text("//div[contains(text(),'ProQuest document ID')]/following-sibling::div")
            
            print(f"📜 Title: {title}")
            
            papers.append({
                "Name": name,
                "Title": title,
                "Author": author,
                "Degree Date": degree_date,
                "University": university,
                "Degree": degree,
                "Subjects": subjects,
                "Advisor": advisor,
                "Dissertation Number": dissertation_number,
                "ProQuest Document ID": document_id,
                "Document URL": paper_url,
                "Status": "Found"
            })
    
    except Exception as e:
        print(f"Error for {name}: {e}")
        papers.append({"Name": name, "Status": f"Error: {e}"})

driver.quit()

df_results = pd.DataFrame(papers)
df_results.to_excel(OUTPUT_PATH, index=False)
print(f"Save to {OUTPUT_PATH}")