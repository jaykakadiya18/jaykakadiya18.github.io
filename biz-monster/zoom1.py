import csv
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from time import sleep




ser = Service("D:\chromedriver")
op = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=ser, options=op)

# headers = {
# 'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
# 'accept-encoding': 'gzip, deflate, br',
# 'accept-language': 'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7,gu;q=0.6',
# 'cache-control': 'max-age=0',
# 'cookie': 'OTZ=6326625_34_34__34_; SID=GAgiyDR2e5-XITbXKPr5EdQb1zmp5uqhiPBjUoblGuyVRhMor5XtCvOpJM9THuS57cAnAw.; __Secure-1PSID=GAgiyDR2e5-XITbXKPr5EdQb1zmp5uqhiPBjUoblGuyVRhMo06InVzft06MS-7EjyrGsWQ.; __Secure-3PSID=GAgiyDR2e5-XITbXKPr5EdQb1zmp5uqhiPBjUoblGuyVRhMorjVyDBVK5agkdUWwtKW9Sg.; HSID=A-N6mujC6KYLBnf2W; SSID=AgHicoGDjJVA4mNW5; APISID=WITXiw1WCgGFpAiz/A1etDr-BfNLT40ddr; SAPISID=44znKGqPGOE98ktl/Azxl7e6HNn1IgiODx; __Secure-1PAPISID=44znKGqPGOE98ktl/Azxl7e6HNn1IgiODx; __Secure-3PAPISID=44znKGqPGOE98ktl/Azxl7e6HNn1IgiODx; OGP=-19022622:; SEARCH_SAMESITE=CgQIzpQB; OGPC=19022622-1:19022552-2:; 1P_JAR=2022-02-01-04; NID=511=jeU6Q9J-C6yAD89ZPLDF6gTnpjwK_Ps26eSpdt9V8s10gNEuRvuPB7LGIKNjHRE7ftS6wbZ6V40uEMSPHgjtJqk2NepVeyQsQPeTkAxMRwwHGrHeEhQPz1rYK-1vzWH16dwweQ3Mcp3iXK2fed7xB2JZ0oCF-w0bo_AsppbGpcjgZDEQ7j51mZQj4MO8EtbRPWbNwYZ_VFXO19KhGHRl2hyhgmCrAzPP4byy7jRu2DDnf865TRW1XDi5_q7u6eL2wzrBwavHgRHY_0nL_DYDYweJOh9Imw5alS-9mSenZgL-93jwES831p0hqMm5Qrr1T2u6vdiVdtda81leUmsam6XDk3vsQA5Wkc1Pyh3eMnckaNU6NsuPyfcFhOrJUn6QxC0F7iAs46_Rq72Vlhb8mXEqEKzXe7URrKkhYuxnqj6kFWxS70Rm8FdjCakb9OCxXyphS1PzAKOt5pB583JJBC4; DV=Ux-wECmbyVZGACm1QWzNqJ34vB066xeh720zclIccQAAAMBLUjyomhJKWgAAAIgGqRZ74VWaPgAAAA; SIDCC=AJi4QfEvOOjSdvwg04661XVo_eVP_cuyhE4agB70xeXsvjUKAaH_4o9KF2xgYl-EdPuXR8uP2UAk; __Secure-3PSIDCC=AJi4QfGpxEYTm1ILZKP-nOJhHZVdS34wMJAVRpSge4awIggPHozw6L7453GTxgMJwqWYPvkbaWo_',
# 'sec-ch-ua': '" Not;A Brand";v="99", "Google Chrome";v="97", "Chromium";v="97"',
# 'sec-ch-ua-arch': "x86",
# 'sec-ch-ua-bitness': "64",
# 'sec-ch-ua-full-version': "97.0.4692.99",
# 'sec-ch-ua-mobile': '?0',
# 'sec-ch-ua-model': "",
# 'sec-ch-ua-platform': "Windows",
# 'sec-ch-ua-platform-version': "10.0.0",
# 'sec-fetch-dest': 'document',
# 'sec-fetch-mode': 'navigate',
# 'sec-fetch-site': 'same-origin',
# 'sec-fetch-user': '?1',
# 'upgrade-insecure-requests': '1',
# 'user-agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36",
# 'x-client-data': "CI22yQEIo7bJAQjBtskBCKmdygEIq/LKAQjr8ssBCJ75ywEI54TMAQjDk8wBCIiVzAEI5ZXMAQjulcwBCM6WzAEIoZfMAQjCl8wBGI6eywE="
# }
# proxies = {
#     "https": 'https://103.214.109.67:80',
#     "http": 'https://103.214.109.67:80'
# }
###############################################3
# rows = []
# with open(rb"C:\Users\jayka\PycharmProjects\pythonProject\test\Book1.csv") as csvfile:  # reading csv file
#     csvreader = csv.reader(csvfile)  # creating a csv reader object
#
#     fields = next(csvreader)  # extracting field names through first row"
#
#     for row in csvreader:  # extracting each data row one by one
#         rows.append(row)
#
#     lines = csvreader.line_num  # get total number of rows
#     print("Total no. of rows: %d" % (lines))
#
#     f = open('links.csv', 'w')  # title line
#     f.write("School Name,School url, Info Link\n")  # state name in csv
#
#     for i in rows[:1]: #lines
#         urll = i[1].replace("https://",'').replace('http://','') #0 is school name 1 is school url
#
#         url = 'https://google.com/search?q=' + urll +"%20inurl:zoominfo.com" #find data with school URL
#         print("------>" + url)
#
#         r = requests.get(url)
#         soup = BeautifulSoup(r.content, 'html.parser')
#         refresh = soup.find_all('meta', attrs={'http-equiv': 'refresh'})
#         soup.find_all('meta', attrs={'http-equiv': 'refresh'})
#
#         print(soup.prettify())
#         try:
#             div = soup.find_all('div', {"class": "kCrYT"})
#             tag = div[0].find('a', href=True)
#             # print(tag)
#             link = tag['href'].replace('/url?q=', '').split('&', 1)[0]
#             print(link)
#             if "/c/" in link:
#                 f.write(i[0].replace(',', '') + ',')  # 1st line write in csv
#                 f.write(i[1].replace(',', '') + ',')  # 2nd line write in csv
#                 f.write(link.replace(',', '').replace('/c/', '/pic/') + '\n')  # last line write in csv
#             else:
#                 f.write(i[0].replace(',', '') + ',')  # 1st line write in csv
#                 f.write(i[1].replace(',', '') + '\n')  # 2nd line write in csv
#         except:
#             f.write(i[0].replace(',', '') + ',')  # 1st line write in csv
#             f.write(i[1].replace(',', '') + '\n')  # 2nd line write in csv
#
#     f.close()

#book1.csv ===> links.csv


rows = []
with open(rb"C:\Users\jayka\PycharmProjects\pythonProject\test\Book1.csv") as csvfile:  # reading csv file
    csvreader = csv.reader(csvfile)  # creating a csv reader object

    fields = next(csvreader)  # extracting field names through first row"

    for row in csvreader:  # extracting each data row one by one
        rows.append(row)

    lines = csvreader.line_num  # get total number of rows
    print("Total no. of rows: %d" % (lines))

    f = open('links.csv', 'w')  # title line
    f.write("School Name,School url, Info Link\n")  # state name in csv

    for i in rows[210:220]:  # lines
        urll = i[1].replace("https://", '').replace('http://', '')  # 0 is school name 1 is school url

        url = 'https://google.com/search?q=' + urll + "%20inurl:zoominfo.com"  # find data with school URL
        print("------>" + url)
        driver.get(url)
        if i.index == 0:
            sleep(10)
        page = driver.page_source
        # r = requests.get(url)
        soup = BeautifulSoup(page, 'html.parser')
        print(soup.prettify())


        try:
            div = soup.find_all('div', {"class": "yuRUbf"})
            tag = div[0].find('a', href=True)
            # print(tag)
            link = tag['href'].replace('/url?q=', '').split('&', 1)[0]
            print(link)

            if "/c/" in link:
                f.write(i[0].replace(',', '') + ',')  # 1st line write in csv
                f.write(i[1].replace(',', '') + ',')  # 2nd line write in csv
                f.write(link.replace(',', '').replace('/c/', '/pic/') + '\n')  # last line write in csv
            else:
                f.write(i[0].replace(',', '') + ',')  # 1st line write in csv
                f.write(i[1].replace(',', '') + ',')  # 2nd line write in csv
                f.write("none" + '\n')  # 2nd line write in csv

        except:
            f.write(i[0].replace(',', '') + ',')  # 1st line write in csv
            f.write(i[1].replace(',', '') + ',')  # 2nd line write in csv
            f.write("none" + '\n')  # 2nd line write in csv

    f.close()

driver.quit()


