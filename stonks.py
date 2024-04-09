from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import csv
import re		
import warnings
from selenium.webdriver.common.by import By

def extract_ratings(symbol, company_name, result, book_value, stock_cmp, mcap, cash, debt, roe, eps):
    soup = BeautifulSoup(result,'html.parser')
    try:
        if (book_value > stock_cmp) and (eps>0) and (roe>0) and ((mcap+cash)-debt>0):
            discount_percent = round(((book_value-stock_cmp)/book_value)*100,2)
            potential_upside = round(((stock_cmp-book_value)/stock_cmp)*100*(-1),2)
            overall_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('span',id="mainContent_ltrlOverAllRating"))).group(1)
            management_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_ManagementRating"))).group(1)
            valuation_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_ValuationRating"))).group(1)
            efficiency_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_EfficiencyRating"))).group(1)
            financials_rating = re.search(r'Valuation Rating is (.*?) out of 5', str(soup.find('div',id="mainContent_FinancialsRating"))).group(1)
            print(symbol, company_name, overall_rating, management_rating, valuation_rating, efficiency_rating, financials_rating, stock_cmp, book_value, discount_percent, potential_upside, mcap, debt, roe, eps)
            csv_writer.writerow([symbol, company_name, overall_rating, management_rating, valuation_rating, efficiency_rating, financials_rating, stock_cmp, book_value, discount_percent, potential_upside, mcap, debt, roe, eps])
    except Exception as inst:
        print(inst.args)
        pass

def get_webpage(symbol, company_name):
    try:
        url = "https://ticker.finology.in/company/"+symbol+"/"
        driver.get(url)
        result = driver.page_source
        book_value = float(driver.find_element(By.XPATH, '//*[@id="mainContent_updAddRatios"]/div[8]/p/span').get_attribute('innerHTML').replace(',',''))
        stock_cmp = float(driver.find_element(By.XPATH, '//*[@id="mainContent_clsprice"]/span/span').get_attribute('innerHTML').replace(',',''))
        mcap = float(driver.find_element(By.XPATH, '//*[@id="mainContent_updAddRatios"]/div[1]/p/span').get_attribute('innerHTML').replace(',',''))
        cash = float(driver.find_element(By.XPATH, '//*[@id="mainContent_ltrlCash"]/span').get_attribute('innerHTML').replace(',',''))
        debt = float(driver.find_element(By.XPATH, '//*[@id="mainContent_ltrlDebt"]/span').get_attribute('innerHTML').replace(',',''))
        roe = float(driver.find_element(By.XPATH, '//*[@id="mainContent_updAddRatios"]/div[14]/p/span').get_attribute('innerHTML').replace(',',''))
        eps = float(driver.find_element(By.XPATH, '//*[@id="mainContent_updAddRatios"]/div[12]/p/span').get_attribute('innerHTML').replace(',',''))
        extract_ratings(symbol, company_name, result, book_value, stock_cmp, mcap, cash, debt, roe, eps)
    except Exception as inst:
        print(inst.args)
        print("unable to load", url)
        
    
def read_nse_companies():
    df = pd.read_excel('NSE_MAR24.xlsx',index_col=0)
    for symbol, company_name in zip(df['Symbol'],df['Company Name']):
        get_webpage(symbol, company_name)

if __name__ == "__main__":
    warnings.filterwarnings("ignore", category=DeprecationWarning) 
    driver = webdriver.Firefox()
    csv_file = open('attractive_stonks_10_apr_2024.csv', 'a')
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(['Symbol', 'Company Name', 'Overall Rating','Ownership','Valuation','Efficiency','Financials', 'CMP', 'Book Value', 'Discount %', 'Potential Upside %', 'MCAP', 'Debt', 'ROE', 'EPS'])
    read_nse_companies()
    driver.close()
    csv_file.close()