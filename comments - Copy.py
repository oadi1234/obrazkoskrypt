from docx import Document
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from PIL import Image
from lxml import etree
import zipfile
import re
ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
pixel_offset = 250
startfrom=0

try:
    with open('startfrom.txt', 'r') as f:
        for line in f:
            startfrom=int(line)
        print("Starting from comment: "+startfrom);
except Exception:
    pass

commentdata = []
def get_document_comments(docxFileName):
    comments_dict={}
    docxZip = zipfile.ZipFile(docxFileName)
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
    for i in range(int(len(com))):
        string = com[str(i)]
        while True:
            try:
                idx = string.index(substr)
                string = string[:idx+1] + newline + string[idx+1:]
            except Exception:
                com[str(i)] = string
                singlestring = singlestring + "\n" + string
                break
    return singlestring
    
def create_driver():
    chrome_options = Options()
    user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'
    chrome_options.add_argument(f'user-agent={user_agent}')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--allow-running-insecure-content')
    chrome_options.add_argument('--hide-scrollbars')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    chrome_options.add_argument('--log-level=3')
    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_window_size(3840, 2160+pixel_offset)
    return driver
    
def screenshot_url_list(urls, driver, errornodes) :
    i = 0
    for url in urls :
        try:
            i+=1
            if i>=startfrom:
                print(f"(Comment {i}) Making a screenshot of: "+url)
                driver.get(url)
                driver.execute_script("document.body.style.zoom='200%'")
                title = str(i)+"_"+re.sub(r'[\W_]', '',(driver.title).replace(" ", ""))+ ".png"
                driver.save_screenshot(title)
                im = Image.open(title)
                width, height = im.size
                if "PubMed" in title:
                    im = im.crop((0, pixel_offset, width, height+pixel_offset))
                else :
                    im = im.crop((0, 0, width, height))
                im.save(title)
        except Exception as ex:
            error = "Error at comment: "+str(i)
            errornodes.append(error)
            print(type(ex))
            print(ex.args)
            print(ex)
            print(error)
            print(f"Screenshot the above link marked as Comment({i}):\"{url}\" individually")
            continue
    
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

if __name__=="__main__":
    try:
        document= document_name_prompt()
        com=get_document_comments(document)
        
        singlestring = comments_to_single_string_with_new_lines(com)
        url_list = []

        urls = re.findall(r'(https?://\S+)', singlestring)
        
        driver = create_driver()
        errornodes = []
        screenshot_url_list(urls, driver, errornodes)
    
        write_error_nodes()
        
        driver.quit()
    except Exception as ex:
        print(type(ex))
        print(ex.args)
        print(ex)
        print("An exception has occured")
        input("Press any key to exit application...")