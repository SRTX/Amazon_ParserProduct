import requests
import xlsxwriter
from bs4 import BeautifulSoup

head = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/11.1.1 Safari/605.1.15',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'DNT': '1'
}

cookies = dict(head)
mainPage = "https://www.amazon.com/"
page = "https://www.amazon.com/s/ref=sr_pg_1?rh=n%3A1055398%2Cn%3A%211063498%2Cn%3A284507%2Cn%3A289668%2Cp_36%3A600-99999999&ie=UTF8&qid=1531754555"
workbook = xlsxwriter.Workbook("Amazon_Sales.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write('A1', "Ссылка на товар")
worksheet.write('B1', "BSR")
worksheet.write('C1', "Price")
worksheet.write('D1', "Рейтинг")
worksheet.write('E1', "Кол-во отзывов")
worksheet.write('G1', "Присутствие Amazona")
worksheet.write('H1', "FBA")
worksheet.write('I1', "Кол-во продавцов")

row = 2
r = requests.get(page, headers=cookies)
print(r)
text = BeautifulSoup(r.text, "html.parser")
kollPage = 0

pangMore = text.find("span", class_="pagnMore")
if pangMore is not None:
    kollPage = int(text.find("span", class_="pagnDisabled").text)
else:
    kollPage = int(len(text.findAll("span", class_="pagnLink")))
l = 1

for q in range(50):
    print("PAGE: ", q)
    link = text.findAll("div", {"class": "s-item-container"})
    textFirst = text
    for i in link:
        print(l)
        l += 1
        checkAmazon = False
        try:
            new_page = i.find("a", {"class": "a-link-normal a-text-normal"})['href']
        except:
            try:
                new_page = i.find("a", {"class": "a-link-normal s-access-detail-page s-color-twister-title-link a-text-normal"})['href']
            except:
                print(i)
                print("===================")
                continue
        if new_page[0] == '/':
            new_page = mainPage + new_page
        r = requests.get(new_page, headers=cookies)
        print(new_page)
        text = BeautifulSoup(r.text, "html.parser")
        a = True
        price = ""
        try:
            price = text.find("span", {"class": "a-price"}).text.replace(' ', '').replace('\n', '')
        except:
            worksheet.write('C' + str(row), "Нет цены")
            a = False
        b = True
        rating = ""
        try:
            rating = text.find("i", {"class": "a-icon a-icon-star a-star-4"}).text.replace('\n', '')
        except:
            worksheet.write('D' + str(row), "Нет рейтинга")
            b = False
        c = True
        reviews = ""
        try:
            reviews = text.find("span", {"id": "acrCustomerReviewText"}).text.replace('\n', '')
        except:
            worksheet.write('E' + str(row), "Нет отзывов")
            c = False

        # BSR
        checkBSR = True
        CoolBSR = 0
        koll = 0
        try:
            bsr = text.find("table", {"class": "a-keyvalue prodDetTable"}).findAll("tr")
            for o in range(len(bsr)):
                if bsr[o].find("th", class_="a-color-secondary a-size-base prodDetSectionEntry").text.replace(
                        ' ', '').replace('\n', '') == "BestSellersRank":
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
            try:
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
            except:
                try:
                    bsr = text.findAll("table", {"class": "a-keyvalue prodDetTable"})[1].findAll("tr")
                    for o in range(len(bsr)):
                        if bsr[o].find("th", class_="a-color-secondary a-size-base prodDetSectionEntry").text.replace(
                                ' ', '').replace('\n', '') == "BestSellersRank":
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
                    CoolBSR = -1

       # Количество FullFillment By Amazon
        if checkBSR:
            try:
                buyers = text.find("div", {"class": "a-section a-spacing-small a-spacing-top-small"}).find("a")['href']
            except:
                try:
                    buyers = text.find("div", {"class": "a-text-center a-spacing-mini"}).find("a")['href']
                except:
                    try:
                        buyers = text.find("span", {"class": "a-size-small aok-float-right"}).find("a")['href']
                    except:
                        with open("file.html", "w") as file:
                            file.write(str(text))
                        print("Currently unavailable")
                        print("===================")
                        continue
            try:
                cellBuyers = text.find("div", {"class": "a-section a-spacing-small a-spacing-top-small"}).find("a").text
                word = ""
                index = 0
                for number in range(len(cellBuyers)):
                    if cellBuyers[number] == '(':
                        index = number
                for number in range(index + 1, len(cellBuyers)):
                    word += cellBuyers[number]
                    if cellBuyers[number + 1] == ')':
                        break
                word = int(word)
                if word == 1:
                    checkBSR = False
            except:
                try:
                    word = text.find("div", {"class": "a-text-center a-spacing-mini"}).find("a").text
                except:
                    word = text.find("span", {"class": "a-size-small aok-float-right"}).find("a").text
            suite = requests.get(mainPage + buyers, headers=cookies)
            text = BeautifulSoup(suite.text, "html.parser")
            chet = 0
            collPag = 0
            try:
                collPage = text.find("li", {"class": "a-selected"})
                collPage.span.decompose()
                collPage.span.decompose()
                collPage = int(collPage.text)
            except:
                collPage = collPag
            rows = text.findAll("div", {"class": "a-row a-spacing-mini olpOffer"})
            whoBuyer = text.find("div", {"id": "olpProductByline"}).text.replace(' ', '').replace('\n', '')
            for j in rows:
                x = j.find("a", {"class": "a-popover-trigger a-declarative olpFbaPopoverTrigger"})
                try:
                    y = j.find("h3", {"class": "a-spacing-none olpSellerName"}).find("a").text.replace(' ', '').replace('\n', '')
                    #print("Продавец: ", whoBuyer)
                    #print("Проверочка:", y)
                    #print("Концовочка:", whoBuyer in y)
                    if whoBuyer in y:
                        checkBSR = False
                except:
                    y = j.find("h3", {"class": "a-spacing-none olpSellerName"}).find("img")['src']
                    if y == "http://ecx.images-amazon.com/images/I/01dXM-J1oeL.gif" or y == "https://images-na.ssl-images-amazon.com/images/I/01dXM-J1oeL.gif":
                        checkBSR = False
                if x != None:
                    x = x.text.replace(' ', '')
                    x = x.replace('\n', '')
                    if x == "FulfillmentbyAmazon" and (
                    (j.find("div", {"class": "a-section a-spacing-small"}).text).replace(' ', '').replace('\n', ''))[0] != "U":
                        chet += 1

            for j in range(collPage):
                nextPage = text.find("li", {"class": "a-last"})
                nextPage = nextPage.find("a")['href']
                nextPage = requests.get(mainPage + nextPage, headers=cookies)
                text = BeautifulSoup(nextPage.text, "html.parser")
                rows = text.findAll("div", {"class": "a-row a-spacing-mini olpOffer"})
                for k in rows:
                    x = k.find("a", {"class": "a-popover-trigger a-declarative olpFbaPopoverTrigger"})
                    try:
                        y = k.find("h3", {"class": "a-spacing-none olpSellerName"}).find("a").text.replace(' ',
                                                                                                           '').replace(
                            '\n', '')
                        if whoBuyer in y:
                            checkBSR = False
                    except:
                        y = k.find("h3", {"class": "a-spacing-none olpSellerName"}).find("img")['src']
                        if y == "http://ecx.images-amazon.com/images/I/01dXM-J1oeL.gif" or y == "https://images-na.ssl-images-amazon.com/images/I/01dXM-J1oeL.gif":
                            checkBSR = False
                    if x != None:
                        x = x.text.replace(' ', '')
                        x = x.replace('\n', '')
                        if x == "FulfillmentbyAmazon" and ((k.find("div", {"class": "a-section a-spacing-small"}).text).replace(' ', '').replace('\n', ''))[0] != "U":
                            chet += 1
        #----
        if checkBSR:
            print(price)
            print(rating)
            print(reviews)
            print(CoolBSR)
            print(word)
            worksheet.write('A' + str(row), new_page)
            if not checkAmazon:
                worksheet.write('G' + str(row), "-")
            elif checkAmazon == 1:
                worksheet.write('G' + str(row), "+")
            elif checkAmazon == 2:
                worksheet.write('G' + str(row), "?")
            if a:
                worksheet.write('C' + str(row), price)
            if b:
                worksheet.write('D' + str(row), rating)
            if c:
                worksheet.write('E' + str(row), reviews)
            worksheet.write('B' + str(row), CoolBSR)
            worksheet.write('I' + str(row), word)
            worksheet.write('H' + str(row), chet)
            row += 1
        print(chet)
        print("===================")
    print(text.find("a", {"class": "pagnNext"}))
    r = requests.get(mainPage + textFirst.find("a", {"class": "pagnNext"})['href'], headers=cookies)
    text = BeautifulSoup(r.text, "html.parser")
