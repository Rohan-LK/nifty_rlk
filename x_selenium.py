from selenium import webdriver
from time import sleep
import pandas as pd

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


if __name__ == '__main__':
    tickers_lst = list(pd.read_excel('tickers/tickers.xlsx')['Symbol'])
    strt_time = '1619827200'
    end_time = '1717210207'

    for i in range(len(tickers_lst)):
        print(str(i+1) + '/' + str(len(tickers_lst)))
        ticker_name = tickers_lst[i]
        download_data(ticker_name, strt_time, end_time)
