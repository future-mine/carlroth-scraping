
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



def getDownloadsFolderPath():
    currentpath = os.path.realpath(__file__)
    print(currentpath)
    # add_to_startup(currentpath)

    patharr = currentpath.split('\\')
    patharr[len(patharr)-1] = 'downloads'
    ptr = ('\\').join(patharr)
    return ptr

def initDriver():
    options = webdriver.ChromeOptions()
    download_dir = getDownloadsFolderPath()
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
    filename = "carlroth.csv"
    file_exists = os.path.isfile(filename)
    header = ['source',	'manufacturer_name', 'product_url',	'product_article_id',	'sds_pdf',	'sds_source',	'sds_language',	'product_name',	'sds_pdf_product_name',
              'sds_pdf_published_date',	'sds_pdf_revision_date',	'sds_pdf_manufacture_name',	'sds_pdf_Hazards_identification',	'sds_filename',	'crawl_date']
    with open(filename, 'a', encoding="utf-8-sig") as csvfile:
        # writer = csv.writer(csvfile,quotechar='|',  quoting=csv.QUOTE_ALL)
        writer = csv.DictWriter(csvfile, delimiter=',', lineterminator='\n',fieldnames=header)

        if not file_exists:
            writer.writeheader()  # file doesn't exist yet, write a header
        for row in productsinfo:
            writer.writerow({
              'source':row[0],	'manufacturer_name':row[1], 'product_url':row[2],	'product_article_id':row[3],	'sds_pdf':row[4],	'sds_source':row[5],	'sds_language':row[6],	'product_name':row[7],	'sds_pdf_product_name':row[8], 'sds_pdf_published_date':row[9],	'sds_pdf_revision_date':row[10],	'sds_pdf_manufacture_name':row[11],	'sds_pdf_Hazards_identification':row[12],	'sds_filename':row[13],	'crawl_date':row[14]
            })







# def writeProductsInfo(productsinfo):
#     filename = "carlroth.csv"

#     try:
#         open(filename, 'r', encoding="utf-8-sig")
#     except:
#         header = ['source',	'manufacturer_name', 'product_url',	'product_article_id',	'sds_pdf',	'sds_source',	'sds_language',	'product_name',	'sds_pdf_product_name',
#                   'sds_pdf_published_date',	'sds_pdf_revision_date',	'sds_pdf_manufacture_name',	'sds_pdf_Hazards_identification',	'sds_filename',	'crawl_date']
#         with open(filename, 'a', encoding="utf-8-sig", newline='') as csvfile:
#             # writer = csv.writer(csvfile,quotechar='',  quoting=csv.QUOTE_ALL)
#             writer = csv.writer(csvfile,delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            
#             writer.writerow(header)

#     with open(filename, 'a', encoding="utf-8-sig", newline='') as csvfile:
#         # writer = csv.writer(csvfile,quotechar='|',  quoting=csv.QUOTE_ALL)
#         writer = csv.writer(csvfile,delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
#         writer.writerows(productsinfo)

#     # with open('quotes.csv', 'w', newline='') as file:
#     #     writer = csv.writer(
#     #         file, quoting=csv.QUOTE_NONNUMERIC, delimiter='|')
#     #     writer.writerows(row_list)
#         # try:
#         #     f = open(filename, "a")
#         #     for info in productsinfo:
#         #         productinfo = ';'.join(info)
#         #         print(productinfo)
#         #         f.write(productinfo + "\n")
#         #     f.close()
#         # except:
#         #     print("There was an error writing to the CSV data file.")


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
    
    # print(content)
    try:
        # pattern = re.compile(r'^Frivillig sikkerhetsinformasjon basert på\nsikkerhetsbladformat iht. forordning (EF) nr. 1907/2006\n(REACH)\n[^n]\n')
        # ptr = re.findall(pattern, content)
        # ptr1 = ptr[0].split('\n')[1]
        product_num = file_path.split('-')[1]
        ptr2 = content.split('\n')
        ind = ptr2.index(product_num) - 1
        print(ind, ptr2[ind])
        for i in range(100):
            print(i,'  ', ptr2[i])
        sds_pdf_product_name = ptr2[ind]
    except:
        sds_pdf_product_name = None
    try:
        pattern = re.compile(r'dato for utarbeiding: [0-9][0-9].[0-9][0-9].[0-2][0-9][0-9][0-9]')
        ptr = re.findall(pattern, content)
        ptr1 = ptr[0].split(': ')[1]
        ptr2 = ptr1.split('.')
        print(ptr)
        sds_pdf_published_date = '{}/{}/{}'.format(ptr2[0], ptr2[1], ptr2[2])
    except:
        sds_pdf_published_date = None


    try:
        pattern = re.compile(r'Revidert: [0-9][0-9].[0-9][0-9].[0-2][0-9][0-9][0-9]')
        ptr = re.findall(pattern, content)
        print(ptr)
        ptr1 = ptr[0].split(': ')[1]
        ptr2 = ptr1.split('.')
        sds_pdf_revision_date = '{}/{}/{}'.format(ptr2[0], ptr2[1], ptr2[2])
    except:
        sds_pdf_revision_date = None


    pattern = re.compile(r'Opplysninger om leverandøren av sikkerhetsdatabladet\n[^\n]+\n')
    ptr = re.findall(pattern, content)
    ptr1 = ptr[0].split('\n')[1]
    print(ptr)
    sds_pdf_manufacture_name = ptr1

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
        driver.get(link)
        id = 0
        while True:
            if id > 1:
                break
            id += 1
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
          ####
                # break
            nextbut = driver.find_element_by_class_name('pagination-next')
            print(nextbut.get_attribute('class').split())
          
          ###
            # break
          
          
            if 'disabled' in nextbut.get_attribute('class').split():
                break
            nextbut.click()

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
        
    time.sleep(2)
    # driver.execute_script("window.scrollTo(0, 2000)")
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    try:    
        title = soup.find('div', {'class': 'product-details'}).find('strong').find('h1').text
    except:
        title = ''

    try:
        brand = soup.find('div', {'class': 'brand-name'}).text
    except:
        brand = ''
    
    print('title', title)

    time.sleep(3)
    

        

        
    # parent = driver.find_element_by_id('accessibletabscontent0-0')
    # try:
    #     selectcountry = parent.find_element_by_id('secDataCountry')
    #     if 'active' in parent.get_attribute('class').split():
    #         pass
    #     else:
    #         parent.click()
    # except:
    #     parent = driver.find_element_by_id('accessibletabscontent0-1')
    #     selectcountry = parent.find_element_by_id('secDataCountry')
    #     if 'active' in parent.get_attribute('class').split():
    #         pass
    #     else:
    #         parent.click()

    # print(type(selectcountry))
    # print(selectcountry is None)
    # if (selectcountry is None):
    #     print('init pass')
    #     parent = driver.find_element_by_id('accessibletabscontent0-0').click()
    #     selectcountry = parent.find_element_by_id('secDataCountry')
    #     if(selectcountry is None):
    #         print('click button')            
    #         parent = driver.find_element_by_id('accessibletabscontent0-1')
    #         selectcountry = parent.find_elements_by_id('secDataCountry')
    #         if(selectcountry is None):
    #             driver.find_element_by_id('accessibletabscontent0-1').click()
    

    try:
        test = driver.find_element_by_xpath('//*[@id="accessibletabscontent0-0"]/a')
        if 'Downloads' in test.text:
            element = driver.find_element_by_id('accessibletabscontent0-0')
            if 'active' in element.get_attribute('class').split():
                pass
            else:
                element.click()
        else:
            test = driver.find_element_by_xpath('//*[@id="accessibletabscontent0-1"]/a')
            if 'Downloads' in test.text:
                element = driver.find_element_by_id('accessibletabscontent0-1')
                if 'active' in element.get_attribute('class').split():
                    pass
                else:
                    element.click()
        
        selectcountry = driver.find_element_by_id('secDataCountry')
        dropdown = Select(selectcountry)
        dropdown.select_by_visible_text("Norway")
        downloadbutts = driver.find_elements_by_xpath(
            "//*[contains(text(), 'Norway')]")
        downloadbutt = downloadbutts[len(downloadbutts)-1]
        pdfurl = downloadbutt.find_element_by_xpath("..").get_attribute('href')
        downloadbutt.click()
        time.sleep(3)
        print('success download')
    except:
        return None



    pdffilename = pdfurl.split('/')[4].split('?')[0]
    print(pdfurl)
    file_path = 'downloads/' + pdffilename
    while True:
        try:
            open(file_path, 'r')
            break
        except:
            time.sleep(1)

    Sds_info = getProductSdsInfo(file_path)
    # zipfilename = compressToZip(pdffilename)
    source = 'carlroth.py'
    manufacturer_name = 'carlroth'
    product_url = link
    product_article_id = None
    sds_pdf = pdfurl
    sds_source = link
    sds_language = "Norwegian"

    product_name = title + ' ' + brand

    sds_filename_in_zip = file_path

    # today = date.today()
    tz = pytz.timezone('Europe/Oslo')
    today = datetime.now(tz=tz)
    crawl_date = today.strftime("%d/%m/%Y")
    productinfo = [source, manufacturer_name, product_url,
                   product_article_id,	sds_pdf, sds_source,	sds_language,	product_name]
    productinfo.extend(Sds_info)
    productinfo.extend([sds_filename_in_zip,	crawl_date])
    return productinfo


site_links = ['https://www.carlroth.com/com/en/performance-materials/c/web_folder_984827',
              'https://www.carlroth.com/com/en/life-science/c/web_folder_260527',
              'https://www.carlroth.com/com/en/chemikalien/c/web_folder_260523',
              'https://www.carlroth.com/com/en/chemikalien/c/web_folder_260523',
              'https://www.carlroth.com/com/en/applications/c/web_folder_758652',
              
              ]
products_url_arr = getProductsUrl(site_links)
print(products_url_arr)
productsinfo_arr = []
print("All products number",len(productsinfo_arr))
id = 0
for link in products_url_arr:
    product_info = getProductInfo(link)
    # if id > 2:
    #     break
    if product_info is None:
        print(link)
    else:
        productsinfo_arr.append(product_info)
        print(product_info)
    print('products number', id)
    id += 1
    

writeUrls(productsinfo_arr)
writeProductsInfo(productsinfo_arr)
time.sleep(3)
driver.close()

