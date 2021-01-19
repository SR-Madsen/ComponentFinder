import xlwt # Used for Excel
from selenium import webdriver # Used for opening Chromium
import bs4 as bs # Used for scraping html

print("\r\nThis Python program will scrape electronics stores to find the price of a selected components.")
print("These prices are added to an Excel sheet to ease generation of Bill of Materials.")
print("When finished, write DONE to quit the program.")
print("-----------------------------------------------------------------------------------------------")

# Declarations for Excel
site_declarations = ["Mouser", "RS Online", "Farnell"]
wb = xlwt.Workbook()
sheets = [wb.add_sheet('Mouser'), wb.add_sheet('RS Online'), wb.add_sheet('Farnell')]
style = xlwt.easyxf('font: height 360')

# Set up Excel sheets
for i in range(len(sheets)):
    sheets[i].write(0, 0, site_declarations[i], style)
    sheets[i].write(2, 0, "Component Name")
    sheets[i].write(2, 1, "Link")
    sheets[i].write(2, 2, "Description")
    sheets[i].write(2, 3, "Price Per Unit")
    sheets[i].col(0).width = 256*30
    sheets[i].col(1).width = 256*100
    sheets[i].col(2).width = 256*80
    sheets[i].col(3).width = 256*30
i = 0

# To keep count of amount of components
component_numbers = [0, 0, 0]

# Prompt user for component name
print("Write the component name: ")
terminal_input = input()

while terminal_input.upper() != "DONE":

    driver = webdriver.Chrome(executable_path=r"C:\Git\ComponentFinder\chromedriver.exe")

    # Scrape Mouser
    driver.get('https://www.mouser.dk/Search/Refine?N=4292906361&Keyword=' + terminal_input + '&Ns=Pricing|0')
    page_content = driver.page_source
    soup = bs.BeautifulSoup(page_content, 'html.parser')

    # Find the relevant data
    name = soup.find('a', attrs = {'id': 'lnkMfrPartNumber_1'})
    price = soup.find('span', attrs = {'id': 'lblPrice_1_1'})
    description = soup.find('td', attrs = {'class': 'column desc-column hide-xsmall'})
    if name == None or price == None or description == None:
        # Perhaps it has gone instantly into product page; search again with values from product page instead
        name = soup.find('span', attrs = {'id': 'spnManufacturerPartNumber'})
        price = soup.find('td', attrs = {'class': 'text-right ext-price-col'})
        description = soup.find('span', attrs = {'id': 'spnDescription'})
        if name == None or price == None or description == None:
            print("Component not found on Mouser, may be out of stock")
        else:
            component_numbers[0] += 1
            sheets[0].write(2+component_numbers[0], 0, name.text.strip())
            sheets[0].write(2+component_numbers[0], 1, "https://www.mouser.dk/Search/Refine?N=4292906361&Keyword=" + terminal_input + "&Ns=Pricing|0")
            sheets[0].write(2+component_numbers[0], 2, description.text.strip())
            sheets[0].write(2+component_numbers[0], 3, price.text.strip())
    else:
        component_numbers[0] += 1
        sheets[0].write(2+component_numbers[0], 0, name.text.strip())
        sheets[0].write(2+component_numbers[0], 1, "https://www.mouser.dk/Search/Refine?N=4292906361&Keyword=" + terminal_input + "&Ns=Pricing|0")
        sheets[0].write(2+component_numbers[0], 2, description.text.strip())
        sheets[0].write(2+component_numbers[0], 3, price.text.strip())

    # Scrape RS Online
    driver.get('https://dk.rs-online.com/web/c/?pn=1&r=t&searchTerm=' + terminal_input + '&sortBy=price&sortOrder=asc&sra=oss')
    page_content = driver.page_source
    soup = bs.BeautifulSoup(page_content, 'html.parser')

    # Find the relevant data
    name = soup.find('span', attrs = {'data-qa': 'value'})
    price = soup.find('span', attrs = {'data-qa': 'price'})
    description = soup.find('div', attrs = {'data-qa': 'description'})
    if name == None or price == None or description == None:
        # Perhaps it has gone instantly into product page; search again with values from product page instead
        name = soup.find('dt', attrs = {'data-testid': 'mpn'})
        price = soup.find('td', attrs = {'data-testid': 'price-breaks-price'})
        description = soup.find('h1', attrs = {'data-testid': 'long-description'})
        if name == None or price == None or description == None:
            # Perhaps it's shown as a list instead of a grid
            name = soup.find('span', attrs = {'class': 'small-link'})
            price = soup.find('span', attrs = {'class': 'col-xs-12 price text-left'})
            description = soup.find('a', attrs = {'class': 'product-name'})
            if name == None or price == None or description == None:
                print("Component not found on RS Online, may be out of stock")
            else:
                component_numbers[1] += 1
                sheets[1].write(2+component_numbers[1], 0, name.text.strip())
                sheets[1].write(2+component_numbers[1], 1, "https://dk.rs-online.com/web/c/?pn=1&r=t&searchTerm=" + terminal_input + "&sortBy=price&sortOrder=asc&sra=oss")
                sheets[1].write(2+component_numbers[1], 2, description.text.strip())
                sheets[1].write(2+component_numbers[1], 3, price.text.strip())
        else:
            component_numbers[1] += 1
            sheets[1].write(2+component_numbers[1], 0, name.text.strip())
            sheets[1].write(2+component_numbers[1], 1, "https://dk.rs-online.com/web/c/?pn=1&r=t&searchTerm=" + terminal_input + "&sortBy=price&sortOrder=asc&sra=oss")
            sheets[1].write(2+component_numbers[1], 2, description.text.strip())
            sheets[1].write(2+component_numbers[1], 3, price.text.strip())
    else:
        component_numbers[1] += 1
        sheets[1].write(2+component_numbers[1], 0, name.text.strip())
        sheets[1].write(2+component_numbers[1], 1, "https://dk.rs-online.com/web/c/?pn=1&r=t&searchTerm=" + terminal_input + "&sortBy=price&sortOrder=asc&sra=oss")
        sheets[1].write(2+component_numbers[1], 2, description.text.strip())
        sheets[1].write(2+component_numbers[1], 3, price.text.strip())

    # Scrape Farnell
    driver.get('https://dk.farnell.com/search/prl/results?range=inc-in-stock&st=' + terminal_input + '&sort=P_PRICE')
    page_content = driver.page_source
    soup = bs.BeautifulSoup(page_content, 'html.parser')

    # Find the relevant data
    name = soup.find('td', attrs = {'class': 'productImage mftrPart'})
    price = soup.find('span', attrs = {'data-price': 'products-0-price-listPrice'})
    description = soup.find('td', attrs = {'class': 'description enhanceDescClmn'})
    if name == None or price == None or description == None:
        # Perhaps it has gone instantly into product page; search again with values from product page instead
        name = soup.find('dd', attrs = {'class': 'ManufacturerPartNumber'})
        price = soup.find('span', attrs = {'data-loaded': 'main-0-price-priceList-0-priceFormatted'})
        description = soup.find('h2', attrs = {'class': 'pdpAttributesName'})
        if name == None or price == None or description == None:
            print("Component not found on Farnell, may be out of stock")
        else:
            component_numbers[2] += 1
            sheets[2].write(2+component_numbers[2], 0, name.text.strip())
            sheets[2].write(2+component_numbers[2], 1, "https://dk.farnell.com/search/prl/results?range=inc-in-stock&st=" + terminal_input + "&sort=P_PRICE")
            sheets[2].write(2+component_numbers[2], 2, description.text.strip())
            sheets[2].write(2+component_numbers[2], 3, price.text.strip())
    else:
        component_numbers[2] += 1
        sheets[2].write(2+component_numbers[2], 0, name.text.strip())
        sheets[2].write(2+component_numbers[2], 1, "https://dk.farnell.com/search/prl/results?range=inc-in-stock&st=" + terminal_input + "&sort=P_PRICE")
        sheets[2].write(2+component_numbers[2], 2, description.text.strip())
        sheets[2].write(2+component_numbers[2], 3, price.text.strip())

    driver.close()
    print("Write the component name: ")
    terminal_input = input()

else:
    driver.quit()
    wb.save('Bill of Materials.xls')
    print("The program has now finished.")