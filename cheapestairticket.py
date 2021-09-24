from selenium import webdriver
import pandas as pd
import numpy as np
import datetime
import time
from datetime import date
import sys
from selenium.webdriver.chrome.options import Options

# chrome_options = Options()
# chrome_options.add_argument("--disable-extensions")
# chrome_options.add_argument("--disable-gpu")
# # chrome_options.add_argument("--no-sandbox") # linux only
# chrome_options.add_argument("--headless")
# chrome_options.headless = True # also works
# driver = webdriver.Chrome('./chromedriver', options=chrome_options)

driver = webdriver.Chrome('./chromedriver')

month = [10,10,10,10,10,10,10]
month = [str(x).zfill(2) for x in month]
day = [4,5,6,7,8,9,10]
day = [str(x).zfill(2) for x in day]
year = [2021,2021,2021,2021,2021,2021,2021]
year = [str(x).zfill(4) for x in year]

# Creating an empty dataframe called 'Flight_Prices'
Flight_Prices = pd.DataFrame()

# Creating a for loop which will iterate through all the required pages of the website
for a,b,c in zip(day,month,year):
    print('execting', a, b, c)
    driver.get("https://www.makemytrip.com/flight/search?tripType=O&itinerary=DEL-BOM-{}/{}/{}&paxType=A-1_C-0_I-0&cabinClass=E&sTime=1597828876664&forwardFlowRequired=true".format(a,b,c))
    time.sleep(5)
    
    lenOfPage = driver.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
    match=False
    while(match==False):
        lastCount = lenOfPage
        time.sleep(1)
        lenOfPage = driver.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
        if lastCount==lenOfPage:
            match=True
    
    time.sleep(10)
    # Extracting all the Airline names using xpath
    fe = driver.find_elements_by_xpath("//div[@class='makeFlex simpleow']")
    Flight_elements = []
    for row in fe:
        if len(row.text.split('\n')) >= 9:
            Flight_elements.append(row)
    # for row in Flight_elements:
    #     print(row.text.split('\n'))  
    #     print(len(row.text.split('\n')))
    Flight_elements = [x.text for x in Flight_elements]

    FlightName = [x.split('\n')[0] for x in Flight_elements]
    FlightName = pd.Series(FlightName)

    # Extracting all the prices using xpath
    Price = [x.split('\n')[len(x.split('\n'))-2] for x in Flight_elements]
    Price = pd.Series(Price)
    
    # Extracting all the From City details using xpath
    Fromcity = [x.split('\n')[2] for x in Flight_elements]
    Fromcity = pd.Series(Fromcity)
    
    # Extracting all the To City details using xpath
    Tocity = [x.split('\n')[6] for x in Flight_elements]
    Tocity = pd.Series(Tocity)
    
    # Extracting all the Duration details using xpath
    Duration = [x.split('\n')[3] for x in Flight_elements]
    Duration = pd.Series(Duration)
    
    # Extracting all the Departure time details using xpath
    Deptime = [x.split('\n')[1] for x in Flight_elements]
    Deptime = pd.Series(Deptime)
    
    #  # Extracting all the Arrival Time details using xpath
    Arrtime = [x.split('\n')[5] for x in Flight_elements]
    Arrtime = pd.Series(Arrtime)
    
    # Date_elements = driver.find_elements_by_xpath("//div[@class='item blue_active']")
    d = datetime.datetime(int(c), int(b), int(a))
    Date = [d] * len(Flight_elements)
    Date = pd.Series(Date)
            
    # Combining all the lists into a dataframe called 'df'
    df = pd.DataFrame({'Date':Date,"Airline":FlightName,"From City":Fromcity, "To City":Tocity, "Departure Time":Deptime,"Arrival Time":Arrtime,"Flight Duration":Duration,"Price":Price})
    
    # We will append the results obtained from every page into the empty datafram created earlier called 'Flight_Prices'
    Flight_Prices = Flight_Prices.append(df) 
    # print('Flight_Prices0', Flight_Prices)

Flight_Prices[Flight_Prices.Date==""] = np.NaN
Flight_Prices.Date = Flight_Prices.Date.fillna(method='ffill')
Flight_Prices.Price = Flight_Prices.Price.str.replace(",","").str.extract('(\d+)')
# print('Flight_Prices1', Flight_Prices)


Flight_Prices = Flight_Prices[Flight_Prices['Price'].notna()]
Flight_Prices['Price'] = pd.to_numeric(Flight_Prices['Price'])
# print('Flight_Prices3', Flight_Prices)


Flight_Prices = Flight_Prices.sort_values(['Price'], ascending=[True])
print('Flight_Prices', Flight_Prices.head(10))

driver.quit()
