from requests import get
from json import loads
import pandas as pd

url = input("URL: ")

print('Tüm ürünleri çekmek için 0 bas')
Range = int(input("Kaç ürün çekmek istiyorsun: "))

pageCount= 1
productList = []

if ('q=') in url:
    keywords = url.split('q=')[-1].split("&qt")[0]
    dataUrl = f'https://public.trendyol.com/discovery-web-searchgw-service/v2/api/infinite-scroll/sr?q={keywords}&pi={pageCount}'
else:
    keyword = url.split("com/")[-1].split("?")[0]
    dataUrl = f'https://public.trendyol.com/discovery-web-searchgw-service/v2/api/infinite-scroll/{keyword}?pi={pageCount}'

while True:
    if len(productList) >= Range and Range != 0 or pageCount > 208:
        break
    datas = loads(get(dataUrl).text)
    productList.extend(datas["result"]["products"])

    if len(datas["result"]['products']) == 24:
        pageCount += 1
    else:
        break

productList = productList[:Range] if Range > 0 else productList

productName = [name["name"] for name in productList]
productLink = [link["url"] for link in productList]
productImg = [img["images"] for img in productList]
proudctPrice = [str(price["price"]["sellingPrice"]) + " TL" for price in productList]

urunListesi = []

for i in range(len(productName)):
    urunListesi.append([productName[i], "https://trendyol.com"+productLink[i], "https://cdn.dsmcdn.com"+productImg[i][0], proudctPrice[i]])

df1 = pd.DataFrame(urunListesi,columns=(["Urun İsmi", "Urun Linki", "Urun Resmi", "Urun Fiyatı"]))
df1.to_excel("urunler.xlsx", sheet_name="urun", index=False)





