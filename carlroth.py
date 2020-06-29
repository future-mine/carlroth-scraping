
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


from webdriver_manager.chrome import ChromeDriverManager
import os
import time
from datetime import datetime
import pytz
import shutil
import zipfile
import re
import csv


import PyPDF2
from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser


def initDriver():
    options = webdriver.ChromeOptions()
    download_dir = 'E:\\Studying_now\\futurestudying\\workIgor\\Erlend\\emdmillipore\\carlroth\\downloads'
    profile = {"plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}],
               "plugins.always_open_pdf_externally": True,
               "download.default_directory": download_dir}
    # profile = {"plugins.always_open_pdf_externally": True}
    options.add_experimental_option("prefs", profile)

    driver = webdriver.Chrome(options=options)

    return driver


driver = initDriver()


def writeUrls(item_desc):
    filename = "itemurls.csv"
    try:
        f = open(filename, "a")
        for itemurl in item_desc:
            f.write(itemurl + "\n")
        f.close()
    except:
        print("There was an error writing to the CSV data file.")


def writeProductsInfo(productsinfo):
    filename = "ProductsInfo.csv"

    # row_list = [
    #     ["SN", "Name", "Quotes"],
    #     [1, "Buddha", "What we think we become"],
    #     [2, "Mark Twain", "Never regret anything that made you smile"],
    #     [3, "Oscar Wilde", "Be yourself everyone else is already taken"]
    # ]

    # myData = [[1, 2, 3], ['Good Morning', 'Good Evening', 'Good Afternoon']]
    with open(filename, 'a', encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(productsinfo)

    # with open('quotes.csv', 'w', newline='') as file:
    #     writer = csv.writer(
    #         file, quoting=csv.QUOTE_NONNUMERIC, delimiter='|')
    #     writer.writerows(row_list)
        # try:
        #     f = open(filename, "a")
        #     for info in productsinfo:
        #         productinfo = ';'.join(info)
        #         print(productinfo)
        #         f.write(productinfo + "\n")
        #     f.close()
        # except:
        #     print("There was an error writing to the CSV data file.")


def getProductSdsInfo(file_path):
    output_string = StringIO()
    with open(file_path, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)

    content = output_string.getvalue()
    pattern0 = re.compile(r'[^\+]H[1-9][1-9][1-9][^\+]')
    pattern1 = re.compile(r'H[1-9][1-9][1-9]')
    pattern2 = re.compile(r'H[1-9][1-9][1-9][\+]H[1-9][1-9][1-9]')
    harzards_arr = []
    initarrforfirst = re.findall(pattern0, content)
    for init in initarrforfirst:
        final = re.findall(pattern1, init)[0]
        harzards_arr.append(final)
    secondarr = re.findall(pattern2, content)
    harzards_arr.extend(secondarr)
    # sds_pdf_Hazards_identification = list(set(sds_pdf_Hazards_identification))
    harzards_arr = list(set(harzards_arr))
    sds_pdf_Hazards_identification = ', '.join(harzards_arr)
    sds_pdf_product_name = ''
    sds_pdf_published_date = ''
    sds_pdf_revision_date = ''
    sds_pdf_manufacture_name = ''
    return [sds_pdf_product_name,	sds_pdf_published_date,	sds_pdf_revision_date,	sds_pdf_manufacture_name,	sds_pdf_Hazards_identification]


def compressToZip(filename):
    zipname = filename.split('.')[0]+'.zip'
    zf = zipfile.ZipFile(
        'downloads/'+zipname, "w", zipfile.ZIP_DEFLATED)
    zf.write('downloads/'+filename, filename)
    zf.close()
    return 'downloads/'+zipname


def getProductsUrl(site_links):
    productsurl_arr = []
    for link in site_links:
        link = 'https://www.carlroth.com/com/en/life-science/c/web_folder_260527'
        driver.get(link)
        while True:
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            rowdiv_arr = soup.find_all(
                "div", {"class": "row hidden-xs hidden-sm"})
            for rowdiv in rowdiv_arr:
                productdiv_arr = rowdiv.find_all("div", {"class": "col-md-4"})
                for productdiv in productdiv_arr:
                    producturl = 'https://www.carlroth.com' + \
                        productdiv.find(
                            "a", {"class": "btn btn-default btn-block"}).get('href')
                    productsurl_arr.append(producturl)
                    break
            try:
                # nextbut = driver.find_element_by_xpath('/html/body/main/div[3]/div[2]/div[2]/div/div/div[1]/div/div[2]/div/div[4]/ul/li[7]/a')
                nextbut = driver.find_element_by_class_name('pagination-next')
                nextbut.click()
                break
            except:
                break

        break

    return productsurl_arr


def getProductInfo(link):
    # driver = initDriver()
    driver.get(link)
    try:
        downloadplus = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "accessibletabscontent0-1"))
        )
        print(downloadplus)
        time.sleep(1)
    except:
        time.sleep(3)
    # driver.execute_script("window.scrollTo(0, 2000)")
    time.sleep(3)

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    title = soup.find('div', {'class': 'product-details'}).text
    brand = soup.find('div', {'class': 'brand-name'}).text

    # fordownloadbutt = []
    # fordownloadbutt.append(driver.find_element_by_id('accessibletabscontent0-0'))
    # fordownloadbutt.append(driver.find_element_by_id('accessibletabscontent0-1'))
    # for butit in fordownloadbutt:
    #   print(butit.text)
    #   if 'Downloads' in butit.text:
    #     print('Downloads  ', butit)
    #     butit.click()
    # driver.find_elements_by_xpath(
    #     "//h2[contains(text(), Downloads')]").click()
    # driver.find_element(By.XPATH, '//*[text()="Downloads"]').click()
    # print(driver.find_element_by_id('accessibletabscontent0-0').text)
    # if 'Downloads' in driver.find_element_by_id('accessibletabscontent0-0').text:
    #     print('click download')
    #     driver.find_element_by_id('accessibletabscontent0-0').click()
    # else:
    #     if 'Downloads' in driver.find_element_by_id('accessibletabscontent0-1').text:
    #         driver.find_element_by_id('accessibletabscontent0-1').click()

    # driver.find_element_by_id('accessibletabscontent0-0').click()
    # selectcountry = driver.find_element_by_id('secDataCountry')
    # dropdown = Select(selectcountry)
    # dropdown.select_by_visible_text("Norway")
    # # downloadbutt = driver.find_element_by_xpath("/html/body/main/div[3]/div[5]/div/div[4]/div[2]/div[41]/div/a")
    # downloadbutts = driver.find_elements_by_xpath(
    #     "//*[contains(text(), 'Norway')]")
    # downloadbutt = downloadbutts[len(downloadbutts)-1]
    # # print(downloadbutt)
    # pdfurl = downloadbutt.find_element_by_xpath("..").get_attribute('href')

    try:
        driver.find_element_by_id('accessibletabscontent0-1').click()
        selectcountry = driver.find_element_by_id('secDataCountry')
        dropdown = Select(selectcountry)
        dropdown.select_by_visible_text("Norway")
        # downloadbutt = driver.find_element_by_xpath("/html/body/main/div[3]/div[5]/div/div[4]/div[2]/div[41]/div/a")
        downloadbutts = driver.find_elements_by_xpath(
            "//*[contains(text(), 'Norway')]")
        downloadbutt = downloadbutts[len(downloadbutts)-1]
        # print(downloadbutt)
        pdfurl = downloadbutt.find_element_by_xpath("..").get_attribute('href')
        # soup = BeautifulSoup(driver.page_source, 'html.parser')
        # pdfurl = soup.find('div', {'data-component-uuiddata-component-uuiddata-component-uuid':'NO'}).find('a').['href']

        # print(pdfurl)
        downloadbutt.click()
        time.sleep(3)
        print('success download')
    except:
        return None

    soup = BeautifulSoup(driver.page_source, 'html.parser')

    pdffilename = pdfurl.split('/')[4].split('?')[0]
    Sds_info = getProductSdsInfo('downloads/' + pdffilename)
    zipfilename = compressToZip(pdffilename)
    print(pdffilename)
    print(pdfurl)
    source = 'carlroth.py'
    manufacturer_name = 'carlroth'
    product_url = link
    product_article_id = None
    sds_pdf = pdfurl
    sds_source = link
    sds_language = "Norwegian"
    product_name = title + brand

    sds_filename_in_zip = zipfilename

    # today = date.today()
    tz = pytz.timezone('Europe/Oslo')
    today = datetime.now(tz=tz)
    crawl_date = today.strftime("%d/%m/%Y")
    productinfo = [source, manufacturer_name, product_url,
                   product_article_id,	sds_pdf, sds_source,	sds_language,	product_name]
    productinfo.extend(Sds_info)
    productinfo.extend([sds_filename_in_zip,	crawl_date])
    return productinfo


site_links = ['https://www.carlroth.com/com/en/life-science/c/web_folder_260527',
              'https://www.carlroth.com/com/en/chemikalien/c/web_folder_260523',
              'https://www.carlroth.com/com/en/chemikalien/c/web_folder_260523',
              'https://www.carlroth.com/com/en/applications/c/web_folder_758652',
              'https://www.carlroth.com/com/en/performance-materials/c/web_folder_984827',
              ]
products_url_arr = getProductsUrl(site_links)
print(products_url_arr)
productsinfo_arr = []
for link in products_url_arr:
    product_info = getProductInfo(link)
    if product_info is None:
        print(link)
    else:
        productsinfo_arr.append(product_info)
print(len(productsinfo_arr))
writeProductsInfo(productsinfo_arr)

# driver.get('https://www.carlroth.com/com/en/general-reagents/dimethyl-sulphoxide-%28dmso%29/p/4720.3')
# ptr = driver.find_elements_by_xpath("//*[contains(text(), 'Subtotal:')]")
# print(ptr)
time.sleep(3)
driver.close()

# driver.close()
# source	manufacturer_name	product_url	product_article_id	sds_pdf	sds_source	sds_language	product_name	sds_pdf_product_name	sds_pdf_published_date	sds_pdf_revision_date	sds_pdf_manufacture_name	sds_pdf_Hazards_identification	sds_filename_in_zip	crawl_date

# Megaflis.py	Megaflis	https://megaflis.no/plumbo-rorvask-1-1kg.html	7.02E+12	https://megaflis.no/media/files/3857/PLUMBO_R_RVASK_1.pdf	https://megaflis.no/plumbo-rorvask-1-1kg.html	Norwegian	Plumbo rørvask 1.1kg	PLUMBO RØRVASK	11/8/2012	6/28/2017	KREFTING & CO. AS	H314,H318,H290	/megaflis/xyz00001.pdf	6/26/2020
