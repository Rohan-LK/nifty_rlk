from selenium import webdriver
from time import sleep
import pandas as pd
from datetime import datetime, timedelta

def download_data(ticker_name, strt_time, end_time):
    url = 'https://finance.yahoo.com/quote/' +ticker_name+ '/history/?guccounter=1&frequency=1mo&period1=' +strt_time+ '&period2=' +end_time
    print('Ticker: ', ticker_name) #; print(url)
    
    driver = webdriver.Edge("edgedriver_win64/msedgedriver.exe")
    driver.get(url)

    download_button_xpath = '//*[@id="nimbus-app"]/section/section/section/article/div[1]/div[2]/div/a/span/span'
    sleep(5)
    # download_button = driver.find_element_by_xpath(download_button_xpath)
    # sleep(2)
    # download_button.click()
    # sleep(10)
    driver.quit()

def conv_func(date_string):
    date_object = datetime.strptime(date_string, "%d-%b-%Y")
    timestamp = (date_object + timedelta(days=1)).timestamp()
    # print(f"1-Jun-2024 timestamp: {int(timestamp)}")
    return str(timestamp)


if __name__ == '__main__':
    # strt_time = '1619827200'
    strt_time = conv_func( str(input('Please enter start datetime (e.g. 1-May-2023): ')) )
    # end_time = '1717210207'
    end_time = conv_func( str(input('Please enter end datetime (e.g. 1-June-2024): ')) )
    tickers_lst = list(pd.read_excel('tickers/tickers.xlsx')['Symbol'])

    for i in range(len(tickers_lst)):
        print(str(i+1) + '/' + str(len(tickers_lst)))
        ticker_name = tickers_lst[i]
        download_data(ticker_name, strt_time, end_time)
