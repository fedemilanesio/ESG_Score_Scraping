from selenium.webdriver import Chrome
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup as bs
import time
import pandas as pd

#######################
df = pd.read_excel(r'C:\...\RULEX\offical_ESG_data\DJGSGRUP as of Feb 25 20241.xlsx',
                   sheet_name='MEMB_DJGSGRUP')

company_names = df.iloc[:, 1]
company_names = company_names.tolist()
company_names_tr = [names[:-2].strip() for names in company_names]

c_n_one=company_names_tr[211:]
esg_output_file_1 = r"C:\Users\fedem\Desktop\211-.xlsx"

def bottino(list,file):
    dfs = []
    count = 0
    path = ChromeDriverManager().install()
    driver = Chrome(service=Service(path))
    driver.maximize_window()
    driver.get("https://www.sustainalytics.com/esg-ratings")
    for stock in list:
        if count == 0:
            driver.execute_script('window.scrollBy(0,500)', "")
            time.sleep(1)
        if count >= 1 :
            driver.execute_script('window.scrollBy(0,200)', "")
            time.sleep(1)
        input_bar = driver.find_element(By.ID, "searchInput")
        try:
            time.sleep(2)
            input_bar.send_keys(stock)
            if len(stock) >= 10:
                input_bar.send_keys(Keys.BACKSPACE*2)
            elif len(stock) < 6:
                #time.sleep(2)
                input_bar.send_keys(Keys.SPACE)
                input_bar.send_keys(Keys.BACKSPACE)
            else:
                #time.sleep(2)
                input_bar.send_keys(Keys.BACKSPACE)
            time.sleep(3)
            company = driver.find_element(By.CLASS_NAME, "companyName")
            company_name = ((bs(company.get_attribute('innerHTML'), 'html.parser')).get_text()).strip()
            company.click()
            element = driver.find_element(By.XPATH,
                                          "//div[@class='col-6 risk-rating-score']//span")
            tag = element.get_attribute('outerHTML')
            score = (bs(tag, 'html.parser')).get_text()
            industry = driver.find_element(By.CLASS_NAME, 'industry-group')
            industry = ((bs(industry.get_attribute('innerHTML'), 'html.parser')).get_text()).strip()
            results = {'Company': stock, 'Score': score, 'Industry': industry}
            print(company_name, score, industry)
            count += 1
            df_result = pd.DataFrame(results, index=[0])
            dfs.append(df_result)
            time.sleep(2)
        except:
            print(f'{stock} Not Found')
            results = {'Company': stock, 'Score': 'not found', 'Industry': 'not found'}
            time.sleep(2)
            input_bar.clear()
            driver.execute_script('window.scrollBy(0,-200)', "")
            df_result = pd.DataFrame(results, index=[0])
            dfs.append(df_result)
            count += 1
            time.sleep(2)

    driver.quit()
    df_output = pd.concat(dfs, ignore_index=True)
    df_output.to_excel(file, index=False)
    return None

one = bottino(c_n_one,esg_output_file_1)


