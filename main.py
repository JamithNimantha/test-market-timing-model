import time
import pandas as pd
from selenium import webdriver

all_data = []

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument("--log-level=3")
driver = webdriver.Chrome(options=chrome_options, executable_path=".\chromedriver.exe")
driver.implicitly_wait(20)
driver.get('https://etfdb.com/compare/volume/')

tab = driver.find_element_by_xpath('//*[@id="featured-wrapper"]/div[1]/div[2]/div[3]/div[2]/div[2]/table/tbody')

rows = tab.find_elements_by_tag_name('tr')
keywords = []
for row in rows:
    td = row.find_elements_by_tag_name('td')
    keywords.append(td[0].text)

combo = []

for i in range(1, 4):
    for j in range(1, 3):
        for k in range(2, 21):
            data = []
            data.append(i)
            data.append(j)
            data.append(k)
            combo.append(data)

c = 0
k = 0

for key in keywords:
    k += 1
    print('\n\n\nkey :', key)
    print('{} out of {}'.format(k, len(keywords)))
    for com in combo:
        try:
            c += 1
            driver.get(
                'https://www.portfoliovisualizer.com/test-market-timing-model'
            )
            driver.find_element_by_id('symbol').send_keys(key)
            time.sleep(2)
            try:
                driver.find_element_by_xpath('/html/body/div[2]/form/div[11]/div/div/span[2]').click()
                time.sleep(1)
                driver.find_element_by_xpath('/html/body/div[3]/div/div/div[3]/button[1]').click()
            except:
                continue
            driver.find_element_by_id('movingAverageType_chosen').click()
            time.sleep(1)
            driver.find_element_by_xpath('/html/body/div[2]/form/div[22]/div/div/div/ul/li[{}]'.format(com[0])).click()

            driver.find_element_by_id('movingAverageType_chosen').click()
            time.sleep(1)
            driver.find_element_by_xpath('/html/body/div[2]/form/div[22]/div/div/div/ul/li[{}]'.format(com[1])).click()

            driver.find_element_by_id('windowSize_chosen').click()
            time.sleep(1)
            driver.find_element_by_xpath('//*[@id="windowSize_chosen"]/div/ul/li[{}]'.format(com[2])).click()

            driver.find_element_by_id('submitButton').click()
        except:
            continue
        try:
            tab = driver.find_element_by_xpath('/html/body/div[2]/div[9]/div[1]/table/tbody')

            rows = tab.find_elements_by_tag_name('tr')
            rows.pop(1)
            for row in rows:
                td = row.find_elements_by_tag_name('td')
                data = []
                data.append(key)
                for i in td:
                    data.append(i.text)

                if len(data) == 12:
                    all_data.append(data)
                    print('Added')
                else:
                    print(len(data))
        except:
            continue
    df = pd.DataFrame(all_data, columns=['Key',
                                         'Portfolio', 'Initial Balance', 'Final Balance', 'CAGR', 'Stdev',
                                         'Best Year', 'Worst Year', 'Max. Drawdown', 'Sharpe Ratio',
                                         'Sortino Ratio', 'US Mkt Correlation'])
    df.to_excel('data.xlsx')
    print('Saved..')

df = pd.DataFrame(all_data, columns=['Key',
                                     'Portfolio', 'Initial Balance', 'Final Balance', 'CAGR', 'Stdev',
                                     'Best Year', 'Worst Year', 'Max. Drawdown', 'Sharpe Ratio',
                                     'Sortino Ratio', 'US Mkt Correlation'])
df.to_excel('data.xlsx')
driver.quit()
