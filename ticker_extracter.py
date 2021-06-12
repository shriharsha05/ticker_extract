from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import csv
import re		

def extract_ratings(symbol, company_name, result):
    soup = BeautifulSoup(result,'html.parser')
    # print(soup.prettify())
    try:
        overall_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('span',id="mainContent_ltrlOverAllRating"))).group(1)
        management_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_ManagementRating"))).group(1)
        valuation_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_ValuationRating"))).group(1)
        efficiency_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_EfficiencyRating"))).group(1)
        financials_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_FinancialsRating"))).group(1)
        print(symbol, company_name, overall_rating, management_rating, valuation_rating, efficiency_rating, financials_rating)
        csv_writer.writerow([symbol, company_name, overall_rating, management_rating, valuation_rating, efficiency_rating, financials_rating])
    except:
        pass

def get_webpage(symbol, company_name):
    url = "https://ticker.finology.in/company/"+symbol+"/"
    try:
        driver.get(url)
    except:
        print("Unable to load: ", url)
        pass
    result = driver.page_source
    extract_ratings(symbol, company_name, result)

def read_nse_companies():
    df = pd.read_excel('NSE_MAR21.xlsx',index_col=0)
    # print(df)
    # print(df['Symbol'])
    for symbol, company_name in zip(df['Symbol'],df['Company Name']):
        get_webpage(symbol, company_name)

if __name__ == "__main__":
    driver = webdriver.Firefox()
    csv_file = open('nse_companies_valuations.csv', 'w')
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(['Symbol', 'Company Name', 'Overall Rating','Ownership','Valuation','Efficiency','Financials'])
    read_nse_companies()
    driver.close()
    csv_file.close()