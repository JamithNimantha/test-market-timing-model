import time
from datetime import date
import pandas as pd
from selenium import webdriver

today = date.today()
d1 = today.strftime("%d_%m_%Y")

all_data = []
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument("--log-level=3")
prefs = {"profile.managed_default_content_settings.images": 2}
chrome_options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=chrome_options, executable_path=".\chromedriver.exe")
driver.implicitly_wait(20)

combo = []

timing_model = str(input('Timing model: ')).split(',')
out = input('out of market asset: ').split(',')
timing_period = input('timing period:').split(',')
trading_frequency = input('trading frequency: ').split(',')
keywords = input('Ticker:').upper().split(',')

for i in timing_model:
    for j in out:
        for k in timing_period:
            for l in trading_frequency:
                data = []
                data.append(i)
                data.append(j)
                data.append(k)
                data.append(l)
                combo.append(data)

c = 0
k = 0
old = pd.read_excel('data.xlsx')
for key in keywords:
    k += 1
    print('\n\n\nkey :', key)
    print('{} out of {}'.format(k, len(keywords)))
    for com in combo:
        try:
            present_combo = com
            c += 1
            driver.get(
                'https://www.portfoliovisualizer.com/test-market-timing-model'
            )
            # timing model
            try:
                time.sleep(1)
                driver.find_element_by_id('timingModel_chosen').click()
                time.sleep(1)
                driver.find_element_by_xpath('//*[@id="timingModel_chosen"]/div/ul/li[{}]'.format(com[0])).click()
            except:
                pass

            # send tikcer
            try:
                driver.find_element_by_id('symbol').send_keys(key)
                try:
                    driver.find_element_by_xpath('/html/body/div[2]/form/div[11]/div/div/span[2]').click()
                    time.sleep(1)
                    driver.find_element_by_xpath('/html/body/div[3]/div/div/div[3]/button[1]').click()
                except:
                    continue
            except:
                driver.find_element_by_id('symbols').send_keys(key)
            time.sleep(2)

            # out of asset model

            try:
                driver.find_element_by_id('outOfMarketAssetType_chosen').click()
                time.sleep(1)
                driver.find_element_by_xpath('//*[@id="outOfMarketAssetType_chosen"]/div/ul/li[2]').click()
                driver.find_element_by_id('outOfMarketAsset').send_keys(com[1])
            except:
                pass
            # timing period
            try:
                driver.find_element_by_id('windowSize_chosen').click()
                time.sleep(1)
                driver.find_element_by_xpath('//*[@id="windowSize_chosen"]/div/ul/li[{}]'.format(com[2])).click()
            except:
                pass

            # trading frequency
            try:
                driver.find_element_by_id('rebalancePeriod_chosen').click()
                time.sleep(1)
                driver.find_element_by_xpath('//*[@id="rebalancePeriod_chosen"]/div/ul/li[{}]'.format(com[3])).click()
            except:
                pass

            time.sleep(2)
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
                data.append(present_combo[0])
                data.append(present_combo[1])
                data.append(present_combo[2])
                data.append(present_combo[3])
                data.append(key)
                for i in td:
                    data.append(i.text)

                if len(data) == 16:
                    data.append(d1)
                    all_data.append(data)
                    print('Added')
                else:
                    print(len(data))
        except:
            continue

    df = pd.DataFrame(all_data,
                      columns=['Timing model', 'out of market asset', 'timing period', 'trading frequency', 'Key',
                               'Portfolio', 'Initial Balance', 'Final Balance', 'CAGR', 'Stdev',
                               'Best Year', 'Worst Year', 'Max. Drawdown', 'Sharpe Ratio',
                               'Sortino Ratio', 'US Mkt Correlation', 'Date'])
    df.to_excel('data.xlsx', index=False)
    print('Saved..')

df = pd.DataFrame(all_data, columns=['Timing model', 'out of market asset', 'timing period', 'trading frequency', 'Key',
                                     'Portfolio', 'Initial Balance', 'Final Balance', 'CAGR', 'Stdev',
                                     'Best Year', 'Worst Year', 'Max. Drawdown', 'Sharpe Ratio',
                                     'Sortino Ratio', 'US Mkt Correlation', 'Date'])
df = df.append(old)
df.to_excel('data.xlsx', index=False)
driver.close()
driver.quit()
