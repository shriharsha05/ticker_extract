from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import re

driver = webdriver.Firefox()					

def read_nse_companies():
    df = pd.read_excel('NSE_MAR21.xlsx',index_col=0)
    # print(df)
    # print(df['Symbol'])
    for company in df['Symbol']:
        print("\n", company)
        get_webpage(company)

def get_webpage(company):
    try:
        url = "https://ticker.finology.in/company/"+company+"/"
        driver.get(url)
        result = driver.page_source
        extract_ratings(result)
    except:
        pass

def extract_ratings(result):
    soup = BeautifulSoup(result,'html.parser')
    # print(soup.prettify())
    try:
        overall_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('span',id="mainContent_ltrlOverAllRating"))).group(1)
        management_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_ManagementRating"))).group(1)
        valuation_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_ValuationRating"))).group(1)
        efficiency_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_EfficiencyRating"))).group(1)
        financials_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_FinancialsRating"))).group(1)
        print(overall_rating)
        print(management_rating)
        print(valuation_rating)
        print(efficiency_rating)
        print(financials_rating)
    except:
        pass

if __name__ == "__main__":
    driver = webdriver.Firefox()	
    read_nse_companies()
    driver.close()