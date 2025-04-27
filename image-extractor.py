from selenium import webdriver
from oauth2client.service_account import ServiceAccountCredentials
import httplib2
import googleapiclient.discovery
from selenium.webdriver.chrome.service import Service
from PIL import Image
import os
import array
from selenium.webdriver.chrome.options import Options
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = os.getcwd()+"\creds.json"

pixel_offset = 250
url_list = []
credentials = ServiceAccountCredentials('creds.json',['https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
drive = googleapiclient.discovery.build('drive','v3')
with open('id.txt', 'r') as f:
    for line in f:
        fileId=line
    print("Google Docs file id: "+fileId)
startfrom=0
try:
    with open('startfrom.txt', 'r') as f:
        for line in f:
            startfrom=int(line)
        print("Starting from comment: "+startfrom);
except Exception:
    pass

res = []
page_token = None
while True:
    obj = drive.comments().list(fileId=fileId, pageToken=page_token, pageSize=100, fields="*").execute()
    if len(obj.get("comments", [])) > 0:
        res = [*res, *obj.get("comments", [])]
    page_token = obj.get("nextPageToken")
    if not page_token:
        break

print(res)
print(drive.comments().list(fileId=fileId, fields="*").execute())
for comment in res:
    url_list.append(comment['content'])
    
print(url_list)

chrome_options = Options()
user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'
chrome_options.add_argument(f'user-agent={user_agent}')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--window-size=1920,1080')
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--allow-running-insecure-content')
chrome_options.add_argument('--hide-scrollbars')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_argument('--log-level=3')
driver = webdriver.Chrome(options=chrome_options)
driver.set_window_size(3840, 2160+pixel_offset)

i = 0
errornodes = []
for url in url_list :
    try:
        i+=1
        if i>=startfrom:
            print("Making a screenshot of: "+url)
            driver.get(url)
            driver.execute_script("document.body.style.zoom='200%'")
            driver.save_screenshot(str(i)+'screenshot.png')
            im = Image.open(str(i)+'screenshot.png')
            im = im.crop((0, pixel_offset, 3840, 2160+pixel_offset))
            im.save(str(i)+'screenshot.png')
    except Exception:
        error = "Error at comment: "+i
        errornodes.append(error)
        print(error)
        continue

file = open('errornodes.txt', 'w')
for line in errornodes:
    file.write(line+"\n")
file.close()

driver.quit()
