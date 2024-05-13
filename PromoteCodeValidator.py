import os
import sys
import re
import requests
import time
import datetime
import csv
import pathlib
import subprocess
import pickle
from copy import deepcopy
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from webdriver_auto_update.chrome_app_utils import ChromeAppUtils
from webdriver_auto_update.webdriver_manager import WebDriverManager

strToday = datetime.datetime.today().strftime('%Y%m%d')
curDir = pathlib.Path(__file__).parent.resolve()

configDir = os.path.join(curDir, 'config')
reportDir = os.path.join(curDir, 'report', strToday)
inputURList = os.path.join(configDir, 'URL.csv')
outputHTML = os.path.join(reportDir, 'index.html')
outputMAIL = os.path.join(reportDir, 'mail.html')
outputSUB = os.path.join(reportDir, 'subject.txt')
mailPShell = os.path.join(configDir, 'mail.ps1')

pic_url = os.path.join(configDir, 'url.pickle')
pic_lightbox = os.path.join(configDir, 'lightbox.pickle')

MAX_LENGTH = 50
RETRY_TIME = 5

UsePickle = False
SharePoint = False
SendMail = False

strPASS = r'PASS'
strFAIL = r'FAIL'
strWARN = r'WARNING'
strMailAddr = ""

if len(sys.argv) > 1:
    strMailAddr = sys.argv[1]

for arg in sys.argv:
    if arg.lower() == "usepickle":
        UsePickle = True
    if arg.lower() == "sendmail":
        SharePoint = True

lstResult = []
strLightBox = ""

def CleanHTML(raw_html):
    CLEANR = re.compile('<.*?>|&([a-z0-9]+|#[0-9]{1,6}|#x[0-9a-f]{1,6});')
    cleantext = re.sub(CLEANR, '', raw_html)
    return cleantext

def HightlightResult(strHTML):
    flagColor = ""
    strHtmlColored = ""
    for line in strHTML.split('\n'):
        strClearText = CleanHTML(line).lstrip().rstrip()
        if strClearText == strFAIL:
            flagColor = r'<td bgColor="Pink">'
        elif strClearText == strWARN:
            flagColor = r'<td bgColor="LightYellow">'
        elif strClearText == strPASS:
            flagColor = ""
        if len(flagColor):
            strHtmlColored += line.replace(r'<td>', flagColor) + "\n"
        else:
            strHtmlColored += line + '\n'
    return strHtmlColored

def TrimURL(str, foldURL):
    ret = str
    if len(str) > MAX_LENGTH:
        if foldURL:
            ret = '<a href="%s" target="_blank">%s...</a>' % (ret, ret[0:MAX_LENGTH-3])
        else:
            ret = '<br/>'.join(str[i:i+MAX_LENGTH] for i in range(0, len(str), MAX_LENGTH))
    return ret  

def ReviseHTML(lstResult, foldURL=False):    
    iTotal = len(lstResult)
    iPass = 0
    iFail = 0
    iWarn = 0
    if len(lstResult):
        for dicRes in lstResult:
            trimedTransURL = ""
            if len(dicRes['Transiton']):
                for idx, tupTrans in enumerate(dicRes['Transiton']):
                    if dicRes["Result"] == strFAIL and tupTrans != dicRes['Transiton'][-1] and dicRes['ExpectedDestination'] in tupTrans[0]:
                        dicRes["Result"] = strWARN
                    trimedTransURL += "%d:%s %s <br/>%s<br/>" % (idx, tupTrans[1], tupTrans[2], tupTrans[0])
            
            dicRes['SourceURL'] = TrimURL(dicRes['SourceURL'], foldURL)
            dicRes['ExpectedDestination'] = TrimURL(dicRes['ExpectedDestination'], foldURL)
            dicRes['DestinationURL'] = TrimURL(dicRes['DestinationURL'], foldURL)
            
            dicRes['Transiton'] = trimedTransURL

            if dicRes["Result"] == strPASS:
                iPass += 1
            elif dicRes["Result"] == strWARN:
                iWarn += 1
            else:
                iFail += 1
    return lstResult, iTotal, iPass, iFail, iWarn

if UsePickle:
    if os.path.exists(pic_url):
        with open(pic_url, 'rb') as handle:
            lstResult = pickle.load(handle)
    if os.path.exists(pic_lightbox):
        with open(pic_lightbox, 'rb') as handle:
            strLightBox = pickle.load(handle)        

if len(lstResult) == 0:
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument("--window-size=960,540")
    options.add_argument("--hide-scrollbars")
    options.add_argument('ignore-certificate-errors')

    # Using ChromeAppUtils to inspect Chrome application version
    chrome_app_utils = ChromeAppUtils()
    chrome_app_version = chrome_app_utils.get_chrome_version()
    print("Chrome application version: ", chrome_app_version)

    # Create an instance of WebDriverManager
    driver_manager = WebDriverManager(curDir)

    # Call the main method to manage chromedriver
    driver_manager.main()

    driver = webdriver.Chrome(options=options)  
    
    with open(inputURList, newline='') as csvfile:
        rows = csv.DictReader(csvfile)
        idx = 0
        for row in rows:
            src_url = row['SourceURL']
            exp_url = row['ExpectedDestination']
            
            if src_url:
                r = None
                screenshot_filename = "%s_%s_%s.png" % (row['Component'], row['Touchpoint'], row['Tag'])
                screenshot_fullpath = os.path.join(reportDir, screenshot_filename)

                thumbnail_filename = "%s_%s_%s_thumb.png" % (row['Component'], row['Touchpoint'], row['Tag'])
                thumbnail_fullpath = os.path.join(reportDir, thumbnail_filename)

                if not os.path.exists(reportDir):
                    os.makedirs(reportDir)

                if r == None or r.status_code == 404:
                    nTry = 0
                    while nTry < RETRY_TIME:
                        try:
                            import urllib3
                            urllib3.disable_warnings()

                            r = requests.get(src_url, verify = False, timeout=10)
                            print("[Retry=%d: %d] requests.get(src_url=%s)" % (nTry, r.status_code, src_url) )

                            driver.get(src_url)
                            time.sleep(5)

                            dstURLexpected = False
                            if ' ' in exp_url:
                                MatchAll = True
                                for itm in exp_url.split():
                                    if not itm in driver.current_url:
                                        MatchAll = False
                                if MatchAll:
                                    dstURLexpected = True
                            elif exp_url in driver.current_url:
                                dstURLexpected = True

                            if not dstURLexpected:
                                for resp in r.history:
                                    if exp_url in resp.url:
                                        dstURLexpected = True
                                        break

                            if r and r.status_code == 200 and dstURLexpected:
                                break
                        except:
                            pass
                        nTry += 1
                                
                dicResult = deepcopy(row)
                dicResult["Result"] = strFAIL
                dicResult["DestinationURL"] = ""
                dicResult["Time"] = 0.0
                dicResult["Retry"] = nTry
                dicResult["Transiton"] = []
                iTotalTime = 0.0

                if r != None and driver.current_url:
                    driver.save_screenshot(screenshot_fullpath)
                    from PIL import Image
                    img = Image.open(screenshot_fullpath)
                    new_image = img.resize((240, 135))
                    img.close()
                    new_image.save(thumbnail_fullpath)
                    new_image.close()

                    iTotalTime += r.elapsed.microseconds
                    strLastTime =  "%.3fs" % (r.elapsed.microseconds / 1000000)

                    dicResult["DestinationURL"] = driver.current_url
                    dicResult["Code"] = r.status_code
                    dicResult["Thumbnail"] = "<a href='#%s'><img src='%s'></a>" % (idx, thumbnail_filename)
                    strLightBox += "<a href='javascript:history.back()' class='lightbox' id='%s'><img src='%s'></a>\n" % (idx, screenshot_filename)
                    idx += 1
                    
                    for resp in r.history:
                        iTotalTime += resp.elapsed.microseconds
                        strHisTime = "%.3fs" % (resp.elapsed.microseconds / 1000000)
                        dicResult["Transiton"].append((resp.url, resp.status_code, strHisTime))
                    
                    dicResult["Time"] = "%.3fs" % (iTotalTime / 1000000)
                    dicResult["Transiton"].append((r.url, r.status_code, strLastTime))

                    if ' ' in exp_url:
                        MatchAll = True
                        for itm in exp_url.split():
                            if not itm in driver.current_url:
                                MatchAll = False
                        if MatchAll:
                            dicResult["Result"] = strPASS
                    elif exp_url in driver.current_url:
                        dicResult["Result"] = strPASS

                    print("%s,%s,%s\n%s\n%s" % (dicResult["Component"],dicResult["Touchpoint"],dicResult["Tag"],dicResult["SourceURL"],dicResult["DestinationURL"]))                
                    time.sleep(5)
                
                lstResult.append(dicResult)
                   
            with open(pic_url, 'wb') as handle:
                pickle.dump(lstResult, handle, protocol=pickle.HIGHEST_PROTOCOL)
            with open(pic_lightbox, 'wb') as handle:
                pickle.dump(strLightBox, handle, protocol=pickle.HIGHEST_PROTOCOL)

    driver.close()

lstMailResult = deepcopy(lstResult)
lstResult, iTotal, iPass, iFail, iWarn = ReviseHTML(lstResult)
lstMailResult, iTotal, iPass, iFail, iWarn = ReviseHTML(lstMailResult, True)

for dicRes in lstMailResult:
    if "Thumbnail" in dicRes:
        del dicRes["Thumbnail"]
    if "Transiton" in dicRes:
        del dicRes["Transiton"]

import pandas as pd
df = pd.DataFrame(lstResult)
dfMail = pd.DataFrame(lstMailResult)

html_string = '''
<html>
    <head><title>URL Promote Code Validation Report</title></head>
    <style>
        body {{
            font-family: Monospace;
        }} 
        table {{
            font-size: 10pt; 
            font-family: Monospace;
            border-collapse: collapse; 
            border: 1px solid silver;
        }}
        td, th {{
            padding: 5px;
        }}
        tbody>tr:hover td {{
            background: silver;
        }}
		.lightbox {{
			position: absolute;
			top: 0;
			left: 0;
			width: 100%;
			height: 100%;
			display: flex;
			justify-content: center;
			align-items: center;
			background-color: #000000FF;
			z-index: -1;
			opacity: 0;
		}}
		.lightbox:target {{
			z-index: 1;
			opacity: 1;
		}}        
    </style>
    <body>
        <center><h1>URL Promote Code Validation Report</h1></center>
        <center>{time}</center>
        <center><h3>{brief}</h3></center>
{lightbox}
        <center>{table}</center>
    </body>
</html>
'''

mail_string = '''
<html>
    <head><title>URL Promote Code Validation Report</title></head>
    <style>
        table {{
            font-size: 10pt; 
            font-family: Monospace;
            border-collapse: collapse; 
            border: 1px solid silver;
        }}
        td, th {{
            padding: 5px;
        }}
    </style>
    <body>
        <center><h1>URL Promote Code Validation Report</h1></center>
        <center>{time}</center>        
        <center><h3>{brief}</h3></center>
        <center><h3><a href="{report}" target="_blank">Click to view full report with thumbnail and transition URLs</a></h3></center>
        <center>{table}</center>
    </body>
</html>
'''

strDateTime = datetime.datetime.today().strftime('%Y/%m/%d %H:%M:%S')
strBrief = r'Total: %d, <span style="color:green">Pass: %d</span>, <span style="color:Orange">Warning: %d</span>, <span style="color:red">Fail: %d</span>' % (iTotal, iPass, iWarn, iFail)
strReport = r'http://cpsdeployqa.twn.rd.hpicorp.net/promotecode/%s/' % strToday

strHTML = html_string.format(table=df.to_html(escape=False, justify='center', index=False), time=strDateTime, brief=strBrief, lightbox=strLightBox)
strHtmlColored = HightlightResult(strHTML)
with open(outputHTML, 'w') as f:
    f.write(strHtmlColored)

strMail = mail_string.format(table=dfMail.to_html(escape=False, justify='center', index=False), time=strDateTime, report=strReport, brief=strBrief)
strMailColored = HightlightResult(strMail)
with open(outputMAIL, 'w') as f:
    f.write(strMailColored)

strResult = "PASS" if iFail == 0 else "FAIL"
strMailSubject = "[%s] URL Promote Code Validation Report %s" % (strResult, strToday)

with open(outputSUB, 'w') as f:
    f.write(strMailSubject)

import shutil
file = reportDir + ".zip"
shutil.make_archive(reportDir, 'zip', reportDir)

if SendMail:
    subprocess.run(["powershell.exe", "-File", mailPShell, strMailAddr, outputSUB, outputMAIL], stdout=sys.stdout)