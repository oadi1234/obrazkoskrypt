from docx import Document
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from PIL import Image
from lxml import etree
import zipfile
import re
import FreeSimpleGUI as sg
import os
import time
ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
pixel_offset = 312
startfrom=0
implicit_wait_time = 3
errornodes = []

layout = [
    [sg.Text("Plik Worda")],
    [sg.Input(size=(45, 1), key="File", readonly=True), sg.FileBrowse(target='File_1', key='Browse_1', enable_events=True), sg.In(visible=False, enable_events=True, key='File_1')],
    [sg.Text("Rozmiar Zrzutu Ekranu")],
    [sg.Text("50%, 40% i 33% działa tylko na pubmedach.")],
    [sg.Checkbox("100%", key='full', default = True), sg.Checkbox("50%", key='half', default = False), sg.Checkbox("40%", key='fourty', default = False), sg.Checkbox("33%", key='third', default = False), sg.Checkbox("mobile", key='mobile', default = False)],
    [sg.Submit()]
        ]
window = sg.Window("Ekstraktor Komentarzowa", layout)

try:
    with open('startfrom.txt', 'r') as f:
        for line in f:
            startfrom=int(line)
        print("Starting from comment: "+startfrom);
except Exception:
    pass

commentdata = []
def get_document_comments(filepath):
    comments_dict={}
    docxZip = zipfile.ZipFile(filepath)
    commentsXML = docxZip.read('word/comments.xml')
    et = etree.XML(commentsXML)
    comments = et.xpath('//w:comment',namespaces=ooXMLns)
    for c in comments:
        comment=c.xpath('string(.)',namespaces=ooXMLns)
        comment_id=c.xpath('@w:id',namespaces=ooXMLns)[0]
        comments_dict[comment_id]=comment
        commentdata.append(comment)
    return comments_dict
    
def comments_to_single_string_with_new_lines(commentdict):
    substr = "/http"
    newline = "\n"
    
    singlestring = ""
    for i in range(int(len(commentdict))):
        string = commentdict[str(i)]
        while True:
            try:
                idx = string.index(substr)
                string = string[:idx+1] + newline + string[idx+1:]
            except Exception:
                commentdict[str(i)] = string
                singlestring = singlestring + "\n" + string
                break
    return singlestring
    
def create_driver():
    chrome_options = Options()
    user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'
    chrome_options.add_argument(f'user-agent={user_agent}')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--allow-running-insecure-content')
    chrome_options.add_argument('--hide-scrollbars')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    chrome_options.add_argument('--log-level=3')
    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver
    
def screenshot_url_list(urls, driver, values) :
    i = 0
    cambridge_cookies_clicked = False
    ard_bmj_cookies_clicked = False
    biomedcentral_cookies_clicked = False
    for url in urls :
        try:
            i+=1
            if i>=startfrom:
                print(f"(Comment {i}) Making a screenshot of: "+url)
                driver.get(url)
                driver.execute_script("document.body.style.zoom='200%'")
                if "ard.bmj.com" in url:
                    ard_bmj_cookies_clicked = handle_ard_site(driver, ard_bmj_cookies_clicked)
                if "cambridge.org" in url:
                    cambridge_cookies_clicked = handle_cambridge_site(driver, cambridge_cookies_clicked)
                if "www.ncbi.nlm.nih.gov" in url:
                    handle_ncbi_site(driver)
                if "bmcpublichealth.biomedcentral" in url:
                    biomedcentral_cookies_clicked = handle_bmc_site(driver, biomedcentral_cookies_clicked)
                if "pubmed" in url:
                    handle_pubmed_site(driver)
                title = str(i)+ "_" + re.sub(r'[\W_]', '',(driver.title[:20]).replace(" ", "")) + ".png"
                make_screenshot(driver, values, url, title, i)
                if values['mobile']:
                    make_screenshot_mobile(driver, url, title, i)
                    
                cleanup(title)
        except Exception as ex:
            error = "Error at comment: "+str(i)
            errornodes.append(error)
            print(type(ex))
            print(ex.args)
            print(ex)
            print(error)
            print(f"Screenshot the above link marked as Comment({i}):\"{url}\" individually")
            continue
            

def make_screenshot(driver, values, url, title, i):

    if values['full']:
        driver.execute_script("document.body.style.zoom='200%'")
        driver.set_window_size(3840, 2160)
        driver.save_screenshot(title)
        im = Image.open(title)
        im = im.resize((3840, 2160))
        im.save(title[:-4] + "_full.png")
    if values['half'] and ("pubmed" in url or "www.ncbi.nlm.nih.gov" in url):
        driver.execute_script("document.body.style.zoom='220%'")
        driver.set_window_size(3840, 2160)
        driver.save_screenshot(title)
        im = Image.open(title)
        im = im.resize((3840, 2160))
        if "pubmed" in url:
            im = im.crop((3840*0.143, 0, 3840*0.643, 2160))
        else:
            im = im.crop((3840*0.12, 0, 3840*0.62, 2160))
        im.save(title[:-4] + "_half.png")
    if values['fourty'] and ("pubmed" in url or "www.ncbi.nlm.nih.gov" in url):
        driver.execute_script("document.body.style.zoom='180%'")
        driver.set_window_size(3840, 2160)
        driver.save_screenshot(title)
        im = Image.open(title)
        im = im.resize((3840, 2160))
        if "pubmed" in url:
            im = im.crop((3840*0.218, 0, 3840*0.618, 2160))
        else:
            im = im.crop((3840*0.201, 0, 3840*0.601, 2160))
        im.save(title[:-4] + "_fourty.png")
    if values['third'] and ("pubmed" in url or "www.ncbi.nlm.nih.gov" in url):
        driver.execute_script("document.body.style.zoom='160%'")
        driver.set_window_size(3840, 2160)
        driver.save_screenshot(title)
        im = Image.open(title)
        im = im.resize((3840, 2160))
        if "pubmed" in url:
            im = im.crop((3840*0.273, 0, 3840*0.603, 2160))
        else:
            im = im.crop((3840*0.26, 0, 3840*0.59, 2160))
        im.save(title[:-4] + "_third.png")
    
def make_screenshot_mobile(driver, url, title, i):
    driver.execute_script("document.body.style.zoom='100%'")
    driver.set_window_size(1080, 1920)
    driver.save_screenshot(title)
    im = Image.open(title)
    im.save(title[:-4] + "_mobile.png")
    
def cleanup(title):
    os.remove(title)
    
def handle_bmc_site(driver, biomedcentral_cookies_clicked):
    if not biomedcentral_cookies_clicked:
        driver.implicitly_wait(implicit_wait_time)
        cookie_elements = driver.find_elements("xpath", "/html/body/dialog/div/div/div[3]/button[1]")
        if len(cookie_elements) > 0:
            cookie_elements[0].click()
            biomedcentral_cookies_clicked = True
            time.sleep(1)
    elements = [None]*2
    elements[0] = driver.find_element("xpath", '//*[@id="top"]')
    elements[1] = driver.find_element("xpath", '/html/body/div[4]/aside')
    if len(elements) > 0:
        for element in elements:
            driver.execute_script("""
            var element = arguments[0];
            element.parentNode.removeChild(element);
            """, element)
    return biomedcentral_cookies_clicked
                
def handle_pubmed_site(driver):
    driver.implicitly_wait(implicit_wait_time)
    elements = [None]*2
    elements[0] = driver.find_element("xpath", '/html/body/header')
    elements[1] = driver.find_element("xpath", '/html/body/section[1]/div')
    if len(elements) > 0:
        for element in elements:
            driver.execute_script("""
            var element = arguments[0];
            element.parentNode.removeChild(element);
            """, element)
            
def handle_ncbi_site(driver):
    driver.execute_script("document.body.style.zoom='175%'")
    driver.implicitly_wait(implicit_wait_time)
    elements = driver.find_elements("xpath", '//*[@id="ui-ncbiexternallink-3"]/section[1]/div')
    if len(elements) > 0:
        driver.execute_script("""
        var element = arguments[0];
        element.parentNode.removeChild(element);
        """, elements[0])
    elements = driver.find_elements("xpath", '//*[@id="ui-ncbiexternallink-3"]/header[1]')
    if len(elements) > 0:
        driver.execute_script("""
        var element = arguments[0];
        element.parentNode.removeChild(element);
        """, elements[0])
    elements = driver.find_elements("xpath", '//*[@id="ui-ncbiexternallink-3"]/section[3]/div')
    if len(elements) > 0:
        driver.execute_script("""
        var element = arguments[0];
        element.parentNode.removeChild(element);
        """, elements[0])
    elements = driver.find_elements("xpath", '/html/body/header/div[1]')
    if len(elements) > 0:
        driver.execute_script("""
        var element = arguments[0];
        element.parentNode.removeChild(element);
        """, elements[0])
    elements = driver.find_elements("xpath", '/html/body/section[1]/div')
    if len(elements) > 0:
        driver.execute_script("""
        var element = arguments[0];
        element.parentNode.removeChild(element);
        """, elements[0])
        
def handle_cambridge_site(driver, cambridge_cookies_clicked):
    if not cambridge_cookies_clicked:
        driver.implicitly_wait(implicit_wait_time)
        cookie_elements = driver.find_elements("xpath", "/html/body/header/div[2]/div/div[2]/a")
        if len(cookie_elements) > 0:
            cookie_elements[0].click()
            cambridge_cookies_clicked = True
            time.sleep(1)
    elements = driver.find_elements("xpath", '/html/body/header/div[1]')
    if len(elements) > 0:
        driver.execute_script("""
        var element = arguments[0];
        element.parentNode.removeChild(element);
        """, elements[0])
    return cambridge_cookies_clicked
        
def handle_ard_site(driver, ard_bmj_cookies_clicked):
    if not ard_bmj_cookies_clicked:
        driver.implicitly_wait(implicit_wait_time)
        cookie_elements = driver.find_elements("xpath",'//*[@id="onetrust-accept-btn-handler"]')
        if len(cookie_elements) > 0:
            cookie_elements[0].click()
            ard_bmj_cookies_clicked = True
            time.sleep(1)
    element = driver.find_element("xpath", '//*[@id="page"]/section[1]/header')
    driver.execute_script("""
    var element = arguments[0];
    element.parentNode.removeChild(element);
    """, element)
    return ard_bmj_cookies_clicked
    
def write_error_nodes():
    file = open('errornodes.txt', 'w')
    for line in errornodes:
        file.write(line+"\n")
    file.close()
    
def document_name_prompt():
    print("Please provide .docx document name. Remember that it needs\nto be in the same directory as the application")
    text = input("(default text.docx if empty): ")
    if not text:
        return "text.docx"
    else:
        if not ".docx" in text :
            text = text + ".docx"
        return text
        
def try_read_document(filepath, values):
    try:
        document=filepath
        comments=get_document_comments(document)
        
        singlestring = comments_to_single_string_with_new_lines(comments)
        url_list = []

        urls = re.findall(r'(https?://\S+)', singlestring)
        
        driver = create_driver()
        screenshot_url_list(urls, driver, values)
    
        write_error_nodes()
        
        driver.quit()
    except Exception as ex:
        print(type(ex))
        print(ex.args)
        print(ex)
        print("An exception has occured")

if __name__=="__main__":
    filepath = "text.docx"
    while True:
        event, values = window.read()
        if event is None or event == "Cancel":
            break
        elif event == "File_1":
            filepath = values.get("File_1")
            window["File"].update(os.path.basename(filepath))
        elif event == "Submit":
            try_read_document(filepath, values)


