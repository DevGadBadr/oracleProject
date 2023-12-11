from win32com.client import Dispatch
import wget
import requests
import json
import zipfile
import os
import os
import shutil

def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
        
    jsondata = requests.get('https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json')
    jsonDic = json.loads(jsondata.text)
    versions = jsonDic['versions']

    version = version[:3]


    for item in versions:
        if item['version'][:3] == version:
            actual_version = version
            download_url = item['downloads']['chromedriver'][4]['url']
            break

    latest_driver_zip = wget.download(download_url,'chromedriver.zip')
    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall() # you can specify the destination folder path here
    # delete the zip file downloaded above
    os.remove(latest_driver_zip)

    shutil.move("./chromedriver-win64/chromedriver.exe", "./chromedriver.exe")

    return version

def getDriverfunc():
    content = os.listdir()
    if 'chromedriver.exe' in content:
        pass
    else:
        paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
        version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]
        print(version)
        get_version_via_com(version)

