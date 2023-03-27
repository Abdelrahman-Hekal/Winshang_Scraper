from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import undetected_chromedriver as uc
import time
import os
import re
from datetime import datetime
import pandas as pd
import warnings
import sys
import xlsxwriter
from multiprocessing import freeze_support
warnings.filterwarnings('ignore')

def initialize_bot(translate):

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    #chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    chrome_options.add_argument('--headless=new')
    chrome_options.page_load_strategy = 'eager'
    # disable location prompts & disable images loading
    if not translate:
        prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2}     
    else:
        prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2,   "translate_whitelists": {"zh-CN":"en"},"translate":{"enabled":"true"}}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    driver.set_page_load_timeout(300)

    return driver

def scrape_posts_mobile(driver, driver_en, output1, page, settings, month, year):

    print('-'*75)
    print(f'Scraping The Articles Links from: {page}')
    # getting the full posts list
    links = []
    nposts = 0
    #limit = settings["Number of Posts"]
    end = False
    prev_month = month - 1
    if prev_month == 0:
        prev_month = 12
    driver.get(page)
    for _ in range(80): 
        # scraping posts urls
        ul = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//ul[@id='bar_list_ul']")))    
        posts = wait(ul, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "li")))      
        for post in posts:
            try:
                date = wait(post, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.comment"))).get_attribute('textContent').split('-')[0]
                date = int(date)
                # scraping previous month data only
                if date == prev_month:       
                    nposts += 1
                    print(f'scraping the url for article {nposts}')
                    link = wait(post, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
                    links.append(link)
                #if nposts == limit:
                #    end = True
                #    break
            except:
                pass

        if end: break

        # moving to the next page
        try:
            url = wait(driver, 2).until(EC.presence_of_all_elements_located((By.XPATH, "//a[@class='s1']")))[-1].get_attribute('href')
            driver.get(url)
        except:
            break

    # scraping posts details
    print('-'*75)
    print('Scraping Articles Details...')
    print('-'*75)

    # reading previously scraped data for duplication checking
    scraped = []
    try:
        df = pd.read_excel(output1)
        scraped = df['unique_id'].values.tolist()
    except:
        pass


    n = len(links)
    data = pd.DataFrame()
    for i, link in enumerate(links):
        try:
            driver_en.get(link)
            driver.get(link)  
            time.sleep(1)

            try:
                htmlelement= wait(driver_en, 5).until(EC.presence_of_element_located((By.TAG_NAME, "html")))
                total_height = driver_en.execute_script("return document.body.scrollHeight")
                height = total_height/10
                new_height = 0
                for _ in range(10):
                    prev_hight = new_height
                    new_height += height             
                    driver_en.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                    time.sleep(0.5)
            except:
                pass

            details = {}

            # English article author and date
            en_author, date = '', ''             
            try:
                art = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//article[@class='art_title_op']")))
                date = wait(art, 2).until(EC.presence_of_element_located((By.TAG_NAME, "time"))).get_attribute('textContent').strip()
                text = art.get_attribute('textContent').strip()
                en_author = text.replace(date, '')
            except Exception as err:
                continue
            
            # checking if the article date is correct
            try:
                art_month = int(re.findall("\d+", date.split('-')[1])[0])    
                art_year = int(re.findall("\d+", date.split('-')[0])[-1])    
                art_day = int(re.findall("\d+", date.split('-')[2])[0]) 
                if art_month != prev_month or art_year != year: continue
                date = f'{art_month}/{art_day}/{art_year}'
            except:
                continue

            art_id = ''
            try:
                text = link.split('/')[-1]
                art_id = int(re.findall("\d+", text)[0])
            except:
                pass

            if art_id in scraped: continue

            details['sku'] = art_id
            details['unique_id'] = art_id
            details['articleurl'] = link

            print(f'Scraping the details of article {i+1}\{n}')
            # Chinese article title
            title = ''             
            try:
                title = wait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "h1"))).get_attribute('textContent').strip()
            except:
                print(f'Warning: failed to scrape the title for article: {link}')               
                
            details['articletitle'] = title            
            
            # Chinese article description
            des = ''             
            try:
                des = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@id='allcon']"))).get_attribute('textContent').strip()
            except:
                print(f'Warning: failed to scrape the description for article: {link}')               
                
            details['articledescription'] = des
                                    
            # English article title
            en_title = ''             
            try:
                en_title = wait(driver_en, 2).until(EC.presence_of_element_located((By.TAG_NAME, "h1"))).get_attribute('textContent').strip()
            except:
                print(f'Warning: failed to scrape the English title for article: {link}')               
                
            asian = re.findall(r'[\u3131-\ucb4c]+',en_title)
            if asian: continue
            details['articletitle in English'] = en_title          
            
            # English article description
            en_des = ''             
            try:
                en_des = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@id='allcon']"))).get_attribute('textContent').strip()
            except:
                print(f'Warning: failed to scrape the description for article: {link}')               
                
            asian = re.findall(r'[\u3131-\ucb4c]+',en_des)
            if asian: continue
            details['articledescription in English'] = en_des            
                                
            details['articleauthor'] = en_author
            details['articledatetime'] = date            
            
            # English article category
            en_cat = ''             
            try:
                en_cat = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='wshead']"))).get_attribute('textContent').strip()
            except:
                pass           
                
            asian = re.findall(r'[\u3131-\ucb4c]+',en_cat)
            if asian: continue
            details['articlecategory'] = en_cat

            # other columns
            details['domain'] = 'Winshang'
            details['hype'] = 0
            details['articletags'] = ''
            details['articleheader'] = ''
            details['articleimages'] = ''         
            details['articlecomment'] = ''

            # appending the output to the datafame       
            data = data.append([details.copy()])
        except Exception as err:
            pass
           
    # output to excel
    data['articledatetime'] = pd.to_datetime(data['articledatetime'])
    df1 = pd.read_excel(output1)
    df1 = df1.append(data)   
    df1 = df1.drop_duplicates()
    df1.to_excel(output1, index=False)
    
def scrape_posts_desktop(driver, driver_en, output1, page, settings, month, year):

    print('-'*75)
    print(f'Scraping The Articles Links from: {page}')
    # getting the full posts list
    links = []
    nposts = 0
    #limit = settings["Number of Posts"]
    end = False
    prev_month = month - 1
    if prev_month == 0:
        prev_month = 12
    driver.get(page)
    for _ in range(100): 
        # scraping posts urls
        ul = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='nlee']")))      
        posts = wait(ul, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "li")))      
        for post in posts:
            try:
                date = wait(post, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='win-nav-time']"))).get_attribute('textContent').split('-')[1]
                date = int(date)
                # scraping previous month data only
                if date == prev_month:       
                    link = wait(post, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
                    links.append(link)
                    nposts += 1
                    print(f'scraping the url for article {nposts}')
                #if nposts == limit:
                #    end = True
                #    break
            except:
                pass

        if end: break
        
        # moving to the next page
        try:
            url = wait(driver, 2).until(EC.presence_of_all_elements_located((By.XPATH, "//a[@class='s1']")))[-1].get_attribute('href')
            driver.get(url)
        except:
            break

    # scraping posts details
    print('-'*75)
    print('Scraping Articles Details...')
    print('-'*75)

    # reading previously scraped data for duplication checking
    scraped = []
    try:
        df = pd.read_excel(output1)
        scraped = df['unique_id'].values.tolist()
    except:
        pass

    n = len(links)
    data = pd.DataFrame()
    for i, link in enumerate(links):
        try:
            art_id = ''
            try:
                text = link.split('/')[-2] + link.split('/')[-1]
                art_id = int(re.findall("\d+", text)[0])
            except:
                pass

            if art_id in scraped: continue
            try:
                driver_en.get(link)
                driver.get(link)
                time.sleep(3)
            except:
                print(f'Warning: Failed to load the url: {link}')
                continue

            try:
                htmlelement= wait(driver_en, 5).until(EC.presence_of_element_located((By.TAG_NAME, "html")))
                total_height = driver_en.execute_script("return document.body.scrollHeight")
                height = total_height/10
                new_height = 0
                for _ in range(10):
                    prev_hight = new_height
                    new_height += height             
                    driver_en.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                    time.sleep(0.5)
            except:
                pass

            details = {}

            # English article author and date
            en_author, date = '', ''             
            try:
                div = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='left']")))
                elems = wait(div, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "span")))
                for elem in elems:
                    text = elem.get_attribute('textContent').strip()
                    if '-' in text:
                        date = text
                        num = re.findall("\d+", date.split('-')[0])[-1]
                        if en_author == '':
                            en_author = date.split(num)[0].strip()
                    elif ':' in text: continue
                    elif len(text) > 0:
                        en_author = text.strip()
            except Exception as err:
                pass
            
            # checking if the article date is correct
            try:
                art_month = int(re.findall("\d+", date.split('-')[1])[0])    
                art_year = int(re.findall("\d+", date.split('-')[0])[-1])    
                art_day = int(re.findall("\d+", date.split('-')[2])[0]) 
                if art_month != prev_month or art_year != year: continue
                date = f'{art_month}/{art_day}/{art_year}'
            except:
                continue    

            details['sku'] = art_id
            details['unique_id'] = art_id
            details['articleurl'] = link

            print(f'Scraping the details of article {i+1}\{n}')

            # Chinese article title
            title = ''             
            try:
                title = wait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "h1"))).get_attribute('textContent').strip()
            except:
                print(f'Warning: failed to scrape the title for article: {link}')               
                
            details['articletitle'] = title            
            
            # Chinese article description
            des = ''             
            try:
                des = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='win-news-content']"))).get_attribute('textContent').strip()
            except:
                pass              
                
            details['articledescription'] = des
                                    
            # English article title
            en_title = ''             
            try:
                en_title = wait(driver_en, 2).until(EC.presence_of_element_located((By.TAG_NAME, "h1"))).get_attribute('textContent').strip()
            except:
                print(f'Warning: failed to scrape the English title for article: {link}')               
                
            #asian = re.findall(r'[\u3131-\ucb4c]+',en_title)
            #if asian: continue
            details['articletitle in English'] = en_title          
            
            # English article description
            en_des = ''             
            try:
                en_des = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='win-news-content']"))).get_attribute('textContent').strip()
            except:
                pass    
                
            #asian = re.findall(r'[\u3131-\ucb4c]+',en_des)
            #if asian: continue                
            details['articledescription in English'] = en_des            
            #asian = re.findall(r'[\u3131-\ucb4c]+',en_author)
            #if asian: continue                            
            details['articleauthor'] = en_author
            details['articledatetime'] = date            
            
            # English article category
            en_cat = ''             
            try:
                en_cat = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='win-newsnav']"))).get_attribute('textContent').replace('\n', '').replace('                    ', '').strip()
            except:
                pass           
                
            asian = re.findall(r'[\u3131-\ucb4c]+',en_cat)
            if asian: continue 
            details['articlecategory'] = en_cat

            # article tags
            tags = ''
            try:
                div = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='win-news-key']")))
                elems = wait(div, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))
                for elem in elems:
                    try:
                        tags += elem.get_attribute('textContent').strip() + ', '
                    except:
                        pass
                tags = tags[:-2]
            except:
                pass

            # other columns
            details['domain'] = 'Winshang'
            details['hype'] = 0
            #asian = re.findall(r'[\u3131-\ucb4c]+',tags)
            #if asian: continue 
            details['articletags'] = tags
            details['articleheader'] = ''
            details['articleimages'] = ''
            details['articlecomment'] = ''

            # appending the output to the datafame       
            data = data.append([details.copy()])
        except Exception as err:
            pass
           
    # output to excel
    data['articledatetime'] = pd.to_datetime(data['articledatetime'])
    df1 = pd.read_excel(output1)
    df1 = df1.append(data)   
    df1 = df1.drop_duplicates()
    df1.to_excel(output1, index=False)
    
def scrape_posts(driver, driver_en, output1, page, settings, month, year):

    print('-'*75)
    print(f'Scraping The Articles Links from: {page}')
    # getting the full posts list
    links = []
    nposts = 0
    #limit = settings["Number of Posts"]
    end = False
    prev_month = month - 1
    if prev_month == 0:
        prev_month = 12
    driver.get(page)
    for _ in range(100): 
        # scraping posts urls
        ul = wait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//ul[@id='bar_list_ul']")))    
        posts = wait(ul, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "li")))      
        for post in posts:
            try:
                date = wait(post, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.comment"))).get_attribute('textContent').split('-')[0]
                date = int(date)
                # scraping previous month data only
                if date == prev_month:       
                    link = wait(post, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
                    # converting the links to the desktop version
                    if 'm.winshang' in link:
                        try:
                            text = link.split('/')[-1]
                            art_id = re.findall("\d+", text)[0]
                            link = 'https://news.winshang.com/html/0' + art_id[:2] + '/' + art_id[2:] + '.html'
                        except:
                            print(f'Warning: Failed to convert the following link to the desktop version: {link}')
                            continue
                    nposts += 1
                    print(f'scraping the url for article {nposts}')
                    links.append(link)
                #if nposts == limit:
                #    end = True
                #    break
            except:
                pass

        if end: break

        # moving to the next page
        try:
            url = wait(driver, 2).until(EC.presence_of_all_elements_located((By.XPATH, "//a[@class='s1']")))[-1].get_attribute('href')
            driver.get(url)
        except:
            break

    # scraping posts details
    print('-'*75)
    print('Scraping Articles Details...')
    print('-'*75)

    # reading previously scraped data for duplication checking
    scraped = []
    try:
        df = pd.read_excel(output1)
        scraped = df['unique_id'].values.tolist()
    except:
        pass

    n = len(links)
    data = pd.DataFrame()
    for i, link in enumerate(links):

        try:
            art_id = ''
            try:
                text = link.split('/')[-2] + link.split('/')[-1]
                art_id = int(re.findall("\d+", text)[0])
            except:
                pass
            if art_id in scraped: continue

            try:
                driver_en.get(link)
                driver.get(link)  
                time.sleep(2)
            except:
                print(f'Warning: Failed to load the url: {link}')
                continue

            try:
                htmlelement= wait(driver_en, 5).until(EC.presence_of_element_located((By.TAG_NAME, "html")))
                total_height = driver_en.execute_script("return document.body.scrollHeight")
                height = total_height/10
                new_height = 0
                for _ in range(10):
                    prev_hight = new_height
                    new_height += height             
                    driver_en.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                    time.sleep(0.5)
            except:
                pass

            details = {}

            # English article author and date
            en_author, date = '', ''             
            try:
                div = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='left']")))
                elems = wait(div, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "span")))
                for elem in elems:
                    text = elem.get_attribute('textContent').strip()
                    if '-' in text and '20' in text:
                        date = text
                        num = re.findall("\d+", date.split('-')[0])[-1]
                        if en_author == '':
                            en_author = date.split(num)[0].strip()
                    elif ':' in text: continue
                    elif len(text) > 0:
                        en_author = text.strip()
            except Exception as err:
                pass
            
            # checking if the article date is correct
            try:
                art_month = int(re.findall("\d+", date.split('-')[1])[0])    
                art_year = int(re.findall("\d+", date.split('-')[0])[-1])    
                art_day = int(re.findall("\d+", date.split('-')[2])[0])       
                if art_month != prev_month or art_year != year: continue
                date = f'{art_month}/{art_day}/{art_year}'
            except:
                print(f'Warning: Failed to extract the date for link: {link}')
                continue    

            details['sku'] = art_id
            details['unique_id'] = art_id
            details['articleurl'] = link

            print(f'Scraping the details of article {i+1}\{n}')

            # Chinese article title
            title = ''             
            try:
                title = wait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "h1"))).get_attribute('textContent').strip()
            except:
                print(f'Warning: failed to scrape the title for article: {link}')               
                
            details['articletitle'] = title            
            
            # Chinese article description
            des = ''             
            try:
                des = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='win-news-content']"))).get_attribute('textContent').strip()
            except:
                pass               
                
            details['articledescription'] = des
                                    
            # English article title
            en_title = ''             
            try:
                en_title = wait(driver_en, 2).until(EC.presence_of_element_located((By.TAG_NAME, "h1"))).get_attribute('textContent').strip()
            except:
                print(f'Warning: failed to scrape the English title for article: {link}')               
                
            #asian = re.findall(r'[\u3131-\ucb4c]+',en_title)
            #if asian: continue
            details['articletitle in English'] = en_title          
            
            # English article description
            en_des = ''             
            try:
                en_des = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='win-news-content']"))).get_attribute('textContent').strip()
            except:
                pass               
            #asian = re.findall(r'[\u3131-\ucb4c]+',en_des)
            #if asian: continue                
            details['articledescription in English'] = en_des

            # author is mentioned in the description
            if en_author == '':
                elems = wait(driver_en, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "strong")))
                for elem in elems:
                    text = elem.get_attribute('textContent').strip()
                    if 'Author' in text:
                        en_author = text.replace('Author', '')
                        break

            details['articleauthor'] = en_author
            details['articledatetime'] = date            
            
            # English article category
            en_cat = ''             
            try:
                en_cat = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='win-newsnav']"))).get_attribute('textContent').replace('\n', '').replace('                    ', '').strip()
            except:
                pass           
            asian = re.findall(r'[\u3131-\ucb4c]+',en_cat)
            if asian: continue                 
            details['articlecategory'] = en_cat

            # article tags
            tags = ''
            try:
                div = wait(driver_en, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='win-news-key']")))
                elems = wait(div, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))
                for elem in elems:
                    try:
                        text = elem.get_attribute('textContent').strip()
                        if text[-1] == ',':
                            text = text[:-1].strip()
                        tags += text + ', '
                    except:
                        pass
                tags = tags[:-2]
            except:
                pass

            # other columns
            details['domain'] = 'Winshang'
            details['hype'] = 0
            #asian = re.findall(r'[\u3131-\ucb4c]+',tags)
            #if asian: continue     
            details['articletags'] = tags
            details['articleheader'] = ''
            details['articleimages'] = ''
            details['articlecomment'] = ''

            # appending the output to the datafame       
            data = data.append([details.copy()])
        except Exception as err:
            print(f'Warning: the below error occurred while scraping the article: {link}')
            print(str(err))
           
    # output to excel
    data['articledatetime'] = pd.to_datetime(data['articledatetime'])
    df1 = pd.read_excel(output1)
    df1 = df1.append(data)   
    df1 = df1.drop_duplicates()
    df1.to_excel(output1, index=False)
 
def get_inputs():

    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '\\' in path:
        path += '\\winshang_settings.xlsx'
    else:
        path += '/winshang_settings.xlsx'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "winshang_settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        settings = {}
        df = pd.read_excel(path)
        cols = df.columns
        settings[cols[0]] = int(cols[1])
    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)

    # checking the settings dictionary
    keys = ["Number of Posts"]
    for key in keys:
        if key not in settings.keys():
            print(f"Warning: the setting '{key}' is not present in the settings file")
            settings[key] = 3000

    if settings["Number of Posts"] < 1:
        settings[key] = 3000

    return settings

def initialize_output():

    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\Scraped_Data\\' + stamp
    if os.path.exists(path):
        os.remove(path) 
    os.makedirs(path)

    file1 = f'Winshang_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
    else:
        output1 = path + "\\" + file1  

    # Create an new Excel file and add a worksheet.
    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()    

    return output1

def main():

    print('Initializing The Bot ...')
    freeze_support()
    start = time.time()
    #settings = get_inputs()
    settings = {}
    output1 = initialize_output()
    homepages = ["http://m.winshang.com/nlist12.html", "http://m.winshang.com/news.html", "http://m.winshang.com/nlist54.html"]
    month = datetime.now().month
    year = datetime.now().year
    try:
        driver = initialize_bot(False)
        driver_en = initialize_bot(True)
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()
    for page in homepages:
        try:
            scrape_posts(driver, driver_en, output1, page, settings, month, year)
        except Exception as err: 
            print(f'Warning: the below error occurred:\n {err}')
            driver.quit()
            time.sleep(5)
            driver = initialize_bot(False)
            driver_en = initialize_bot(True)

    try:
        scrape_posts_desktop(driver, driver_en, output1, 'http://m.winshang.com/newsletter.aspx', settings, month, year)
    except Exception as err: 
        print(f'Warning: the below error occurred:\n {err}')
        driver.quit()
        time.sleep(5)
        driver = initialize_bot(False)
        driver_en = initialize_bot(True)

    driver.quit()
    driver_en.quit()
    print('-'*75)
    elapsed_time = round(((time.time() - start)/60), 2)
    input(f'Process is completed in {elapsed_time} mins, Press any key to exit.')

if __name__ == '__main__':

    main()