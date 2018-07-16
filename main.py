from bs4 import BeautifulSoup


first = "https://www.ebay.com/sch/i.html?&_nkw="
last = "&LH_PrefLoc=98&_sop=12"

def link(name):
    name = name.replace(" ", "%20")
    name = first + name + last
    return name

if __name__ == '__main__':
    print(link("Sirena Vacuum Exclusive Royal Line Pro Ultra Deluxe Bonus Package w/ 2 Exclusive Extra Air Purifiers"))
