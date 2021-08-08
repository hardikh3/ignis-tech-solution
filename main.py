#!/usr/bin/env python3

import json
import requests
import xlsxwriter 

# Class for the products from DSW

class DswProductCatelog: 
    def __init__(self):
        self.pagination = 0
        self.workbook = None
        self.worksheet = None
        self.maxRecords = None

    def xlinit(self,name):
        self.workbook = xlsxwriter.Workbook(name)
        self.worksheet = self.workbook.add_worksheet()
        self.xlwrite(0, ["Product URL", "Product Id", "Product Title", "Price", "Color", "Size"])
        
    def xlwrite(self, row, arr):
        for i in range(6):
            self.worksheet.write(row, i, arr[i])

    def fetch(self, url):
        headers = {'User-Agent':'Mozilla/5.0 (iPhone; CPU iPhone OS 8_0 like Mac OS X) AppleWebKit/600.1.3 (KHTML, like Gecko) Version/8.0 Mobile/12A4345d Safari/600.1.4'}
        res = requests.get(url, headers=headers)
    
        while res.status_code != 200:
            print("Err. Retrying ...")
            res = requests.get(url, headers=headers)

        return res.json()
        

    def retrieve(self):
        # Initialize Spreadsheet
        self.xlinit("catalogue.xlsx")

        while True:
            self.extractRecords()
            # Exit loop when all records are fetched 
            if self.pagination >= self.maxRecords:
                break

        # Close Spreadsheet
        self.workbook.close()

    def extractRecords(self):
        url = f'https://www.dsw.com/api/v1/content/pages/_/N-1z141hwZ1z141ju?pagePath=/pages/DSW/category&skipHeaderFooterContent=true&No={self.pagination}&filter=gender&locale=en_US&pushSite=DSW&tier=GUEST'
        data = self.fetch(url)
        contents = data["pageContentItem"]["contents"][0]["mainContent"][7]["contents"][0]
        records = contents["records"]

        # The Max records change inconsistently, some issue with the dsw api.
        self.maxRecords = contents["totalNumRecs"]

        forSomeReasonThereAreAds = 0
        
        for r in records:
            productId = None
            r = r["attributes"]

            # Yes there are ads in this code. WOW!!!
            try:
                productId = r["product.repositoryId"][0]
            except:
                continue
        
            brandName = r["brand"][0].lower()
            productName = r["product.displayName"][0].lower()
            activeColor = r["product.defaultColorCode"][0]
            productUrl = "https://www.dsw.com/en/us/product/{}-{}/{}?activeColor={}".format(brandName.replace("'","").replace(" ","-"),productName.replace("'","").replace(" ","-"),productId,activeColor)
            price = r["product.originalPrice"][0]
    
            colors = []
            colorlist = r["product.colorNames"][0].split("|")
            for c in colorlist:
                colors.append(c.split("~")[1])
            colors = ", ".join(colors)

            productDetails = self.retrieveProduct(productId)        
            sizes = []
            sizelist = productDetails["Response"]["product"]["childSKUs"]
            for s in sizelist:
                sizes.append(s["size"]["displayName"])
            sizes.reverse()
            sizes = ", ".join(sizes)

            # Save to the Spreadsheet
            row = self.pagination + forSomeReasonThereAreAds + 1
            self.xlwrite(row,[productUrl,productId,productName.upper(),price,colors,sizes])

            # Print progress 
            print(f'Fetching {row} of {self.maxRecords} records ...', end="\r")

            forSomeReasonThereAreAds += 1
        
        # Increment pagination
        self.pagination += forSomeReasonThereAreAds
        
    def retrieveProduct(self,productId):
        url = f'https://www.dsw.com/api/v1/products/{productId}?locale=en_US&pushSite=DSW' 
        return self.fetch(url)


# Create Instance
scrapedsw = DswProductCatelog()
scrapedsw.retrieve()
