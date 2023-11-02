from requests_html import HTMLSession
from bs4 import BeautifulSoup
import pandas as pd
import time


#requests, httpx 
#requests_html, ada limitasi tertentu, scrapy splash
def access_url(url):
    session = HTMLSession()
    r = session.get(url) #response [200]
    r.html.render(wait=2, scrolldown = 5, sleep = 2)
    text = r.html.html
    soup = BeautifulSoup(text, "html.parser")
    return soup

def write_html_file(text):
    with open(f"output.html", "w", encoding="UTF-8") as file:
        file.writelines(text)

def read_html_file(file):
    with open(f"{file}","r" ,encoding="UTF-8") as f:
        text = f.read()
    text = BeautifulSoup(text, "html.parser")
    return text

def read_input_xlsx(filename):
    datas = pd.read_excel(f"{filename}.xlsx", sheet_name="Sheet1").to_dict(orient='records')
    urls = []
    for data in datas:
        urls.append(data['Target URL'])
    return urls

def html_parse(html_file): #fungsinya buat apa, for parsing, the output or return value, should be a dictionary. ex = {"title":ÖSTANÖ} dst
    title = html_file.find("div", class_="d-flex flex-row").get_text().strip()
    price = html_file.find("p", class_="itemBTI display-6").get_text().replace("Rp", "").replace(".", "").strip()
    desc = html_file.find("span", class_="itemFacts font-weight-normal").get_text().replace("\n","")
    item_code = html_file.find("span",class_="item-code").get_text().strip()
    designer = html_file.find("div", id="good-to-know").find("div").find_all("div")[-1].get_text().strip()
    measure = html_file.find_all("table", class_="table table-line table-sm")[-1].find_all("td")
    length = measure[5].get_text().strip().replace(" cm", "")
    width = measure[7].get_text().strip().replace(" cm", "")
    height = measure[9].get_text().strip().replace(" cm", "")
    details = html_file.find("div", class_="product-desc-wrapper mb-4").p.get_text().strip()
    image_list = []
    image = html_file.find_all("div", class_="image-container slick-slide")
    for div in image:
        img = div.find("span").find("img")["data-lazy"]
        image_list.append(img)

    dict_result = {"Nama Produk":title,
                   "Harga":price,
                   "Deskripsi":desc,
                   "No. Artikel":item_code,
                   "Panjang":length,                  
                    "Lebar":width,
                    "Tinggi":height,
                    "Desainer":designer,
                    "Rincian Produk":details,
                    "Gambar":image_list[0]
                    }
    return dict_result


if __name__ == "__main__": 
    urls = read_input_xlsx("input")
    print(read_input_xlsx("input"))
    results = []
    #time.sleep(600)
    for url in urls:
        print(f"scraping product page : {url}")
        #for first time running code
        html_text = access_url(url)
        df = html_parse(html_text)
        results.append(df)
    print(results)    

    #practically still single data no need dataframe
    df = pd.DataFrame(df)
    df.to_excel("output.xlsx")

    """

#scraping

1. deskripsi produk
2. no. artikel
3. opsional =Bahan sampai rangka dasar
4. desainer
5. url gambar
    dict_result = {"title":title,
                   "price":price,
                   "desc":desc,
                   "Ukuran" : {"kedalaman" : "45cm",
                                "Tinggi" : "76cm dst"}}
"""