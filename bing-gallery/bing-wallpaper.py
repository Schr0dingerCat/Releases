# -*- coding: utf-8 -*-
"""
Created on Thu Aug 10 20:12:28 2017
@author: lewiskoo
"""

import requests
import re
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

def updateurl(year,month,day,country):
    baseurl = 'http://www.istartedsomething.com/bingimages/?m=7&y='
    countries = ['au','ca','cn','de','fr','gb','uk','jp','nz','us']
    if int(month) <10:
        month = '0' + month
    if int(day) <10:
        day = '0' + day     
    url = baseurl + year + '#' + year + month + day + '-' + countries[country-1]
    para = year+month+day+countries[country-1]
    return url,para
	
def phantomjs_req(url):
    dcap = dict(DesiredCapabilities.PHANTOMJS)
    dcap["phantomjs.page.settings.userAgent"] = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; rv:25.0) Gecko/20100101 Firefox/25.0 ")
    d=webdriver.PhantomJS(desired_capabilities=dcap)
    d.get(url)
    html=d.page_source
    d.quit()
    return html
	
def parse_html(html):
    r=r"[a-zA-z]+://[^\s]*"
    re_jpg=re.compile(r)
    urllist = re.findall(re_jpg,html)
    img_url = urllist[-1]
    half_url=img_url[0:-18]
    max_rs=img_url[-18:-10]
    max_rs=str(max_rs)
    return half_url,max_rs
	
def download(url,max_rs,para):
    resolutions = ['1920x1200','1920x1080','1366x768','1280x768','1024x768','800x600','800x480','768x1280','720x1280','640x480','480x800','400x240','320x240','240x320']
    max_=resolutions.index(max_rs)
    total=len(resolutions)-max_
    print 'There are %s of resolutions. Which one do you like ? (%d means largest) ' %(str(total),max_)
    rs = input('I like No.')
    print "\n"
    weburl = url + resolutions[rs] +'.jpg'
    page = requests.get(weburl)
    code = page.status_code
    if code == 404:
        print "The resolution you want doesn't exist,we will return to default one.\n"
        weburl = url + max_rs + '.jpg'
        page = requests.get(weburl)
        img_name = para + '@' +  max_rs + '.jpg'
    else:
        img_name = para + '@' + resolutions[rs] + '.jpg'
    with open(img_name,'wb') as f:
        f.write(page.content)

    
def main(): 
    year = raw_input('What year do you want ? ')    
    month = raw_input('What month do you want ? ')
    day = raw_input('What day do you want ? ')
    country = input('Which country do you want ? (limited from 1 to 10) ')
    print "\n"
    url,para = updateurl(year,month,day,country)
    html = phantomjs_req(url)
    final_url,max_rs = parse_html(html)
    download(final_url,max_rs,para)
    print 'Download successful!'
    
if __name__=='__main__':
main()
