import requests
import xlsxwriter
from bs4 import BeautifulSoup

head = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'en-US,en;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'DNT': '1',
    'cookie': 'skin=noskin; session-id=146-9729770-2115806; ubid-main=135-1399509-1473902; x-wl-uid=1+DlGU1/VczjVyoIjhBNyJ4phyLPVo6J0uJk7Oy/mFo6JSwfEh+tJiyhBCTCrMfCJrL48nKgncBo=; session-token=lkk+BsPje/i7GlUPtZJOaaEtDhoSIxjCMIQNEL81jzj8oquboySdxqDPX6ABHnADHjzcpm1i33KERxlE+3L1EH9LF5cYFzQHU140OJe2rWDXiL6Belf20+95N5WDTK8Xnhw+HqeC0LH3Ib6EjiXp3wDEH6KgkOiH3pIBHdxPA2ZCr031d1CeYLAd389z7Iql0GXkO33y1lc0nBx+Ytr/UQLCPj4hNXO7nWbJMJkHlXMYJbgjfIvWoxSoZPMP3nyW; session-id-time=2082787201l; ca=AFJQCAIpsgIUMhCFBAIIFwQ=; x-amz-captcha-1=1529068850160185; x-amz-captcha-2=vgy1USVfez6V5r8/M8Xpug==; amznacsleftnav-1e4dfe77-0d78-3527-b54a-f23cc2cb231e=1; csm-hit=tb:s-Q7A7AFZT4CKR8K7KMDKE|1529070715749&adb:adblk_yes'
}

cookies = dict(head)
mainPage = "https://www.amazon.com/"
page = "https://www.amazon.com/s/b/ref=sv_hg_fl_404458011?ie=UTF8&node=404458011"
workbook = xlsxwriter.Workbook("Amazon_Sales.xlsx")
worksheet = workbook.add_worksheet()

worksheet.write('A1', "Ссылка на товар")
worksheet.write('B1', "BSR")
worksheet.write('I1', "Кол-во продавцов")

r = requests.get(page, headers=cookies)
print(r)
text = BeautifulSoup(r.text, "html.parser")

row = 2
l = 1
for q in range(l):
    print("Page: ", q)
    link = text.findAll("div", {"class": "s-item-container"})
    textFirst = text

    for i in link:
        print(l)
        l += 1
        new_page = i.find("div", {"class": "a-row a-spacing-mini"}).find("a")['href']
        print(new_page)
        r = requests.get(new_page, cookies=cookies)
        print(r)
        text = BeautifulSoup(r.text, "html.parser")

        checkBSR = True
        CoolBSR = 0
        koll = 0
        bsr = text.find("class", {"class": "a-section table-padding"})
        exit()
        try:
            bsr = text.find("table", {"id": "productDetails_detailBullets_sections1"}).findAll("tr")
            print(bsr)
            for o in range(len(bsr)):
                if bsr[o].find("th", class_="a-color-secondary a-size-base prodDetSectionEntry").text.replace(' ', '').replace('\n', '') == "BestSellersRank":
                    koll = o
            bsr = bsr[koll].text
            bsrBest = ""
            position = 0
            bsr = bsr.replace(' ', '').replace('\n', '')
            for j in range(len(bsr)):
                if bsr[j - 1] == "#":
                    position = j
                    break
            for j in range(position, position + 7):
                if bsr[j] == 'i':
                    break
                bsrBest += bsr[j]
            bsrBest = int(bsrBest.replace(',', ''))
            CoolBSR = bsrBest
            if CoolBSR > 20000:
                checkBSR = False
        except:
            bsr = text.find("li", {"id": "SalesRank"}).text
            start = 0
            bsrWritting = ""
            for j in range(len(bsr)):
                if bsr[j - 1] == '#':
                    start = j
                    break
            for j in range(j, j + 7):
                if bsr[j] == ' ':
                    break
                bsrWritting += bsr[j]
            bsrWritting = int(bsrWritting.replace(',', ''))
            CoolBSR = bsrWritting
            if CoolBSR > 20000:
                checkBSR = False

        try:
            buyers = text.find("div", {"class": "a-section a-spacing-small a-spacing-top-small"}).find("a")['href']
        except:
            buyers = text.find("div", {"class": "a-text-center a-spacing-mini"}).find("a")['href']
        try:
            cellBuyers = text.find("div", {"class": "a-section a-spacing-small a-spacing-top-small"}).find("a").text
            word = ""
            index = 0
            for number in range(len(cellBuyers)):
                if cellBuyers[number] == '(':
                    index = number
            for number in range(index, len(cellBuyers)):
                word += cellBuyers[number]
                if cellBuyers[number] == ')':
                    break
        except:
            word = text.find("div", {"class": "a-text-center a-spacing-mini"}).find("a").text

        if checkBSR:
            print(new_page)
            print(CoolBSR)
            print(word)
            worksheet.write('A' + str(row), new_page)
            worksheet.write('B' + str(row), CoolBSR)
            worksheet.write('I' + str(row), word)
            row += 1
        print("========")
    r = requests.get(mainPage + textFirst.find("a", {"title": "Next Page"})['href'], headers=cookies)
    text = BeautifulSoup(r.text, "html.parser")
