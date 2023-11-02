import httpx, json, re, time, os, urllib.request, urllib.error
from datetime import datetime
import pandas as pd
from fake_useragent import UserAgent

ua = UserAgent()

def read_xlsx_input(input):
    # Read the Excel file and get its Skip status, if 1, then it will not executed
    data1 = pd.read_excel(f"{input}.xlsx", sheet_name="Sheet1").to_dict(orient= 'records')
    urls = []
    for data in data1:
        if data["Skip"] == 1:
            urls.append(data["Shop ID"])
        else:
            pass
    return urls
    
    
def shopid(url): #getting shopID. ex = https://shopee.sg/shop/358958440/search --> 358958440
    shop_id = re.search(r'/shop/(\d+)/search', url).group(1)
    return shop_id 

def api_url_target(shop_id): #formatting url to get all products of a shop, NOT TO GET EACH INDIVIDUAL DATAS
    return f"https://shopee.sg/api/v4/recommend/recommend?bundle=shop_page_product_tab_main&limit=999999999&offset=0&section=shop_page_product_tab_main_sec&shopid={shop_id}"
#999999999
def parsing_individual_product_url(item_id, shop_id):
    return f"https://shopee.sg/api/v4/pdp/get?item_id={item_id}&limit=99&offset=0&shop_id={shop_id}" #formatted individual product url api data


def request_api_access(api_url): #TO GET INTO API DATA, will return the format in JSON 
    client = httpx.Client(timeout=None)
    try:
        while True:
            useragent = ua.random
            headers = {"user-agent": useragent}
            r = client.get(api_url, headers=headers)
            if r.status_code == 200:
                result = r.json()
                #save_json(result, "result.json")
                return result 
    except httpx.ConnectTimeout or httpx.ConnectError as e:
        print(f"got {e}, retrying...")


def check_page(url):
    try:
        response = urllib.request.urlopen(url)
        return response.getcode()
    except Exception as e:
        return str(e)
            

def save_json(json_data, filename): #FOR SAVING JSON, for debugging and development purposes
    with open(filename, "w", encoding="utf-8") as json_file:
        json.dump(json_data, json_file, ensure_ascii=False, indent=4)

"""
def open_json(json_file_name): #TO READ JSON DATA, for debugging and development purposes
    try:
        with open(json_file_name, "r", encoding="utf-8") as json_file:
            data = json.load(json_file)
            print("json data loaded successfuly")
        return data
    except FileNotFoundError:
        print(f"File not found: {json_file_name}")
        return None
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return None
"""

def date_now():
    current_date = datetime.now()
    return current_date.strftime("%d-%b-%Y") #get current date, %d = day %b = month %Y = Year, %H = hour


def parsing_price(price):#PARSING PRICE, to SGD, ex 200000 -> 2.0 $SGD 
    parsed_price = price / 100000
    fixed_price = round(parsed_price, 2)
    return fixed_price 


def split_string(input_string):
    result = re.split(r'[SML]\s', input_string)
    # Removing empty strings from the resulting list
    result = [part for part in result if part != '']
    return result


def parsing_var(product_detail, url): #THE PRODUCT_DETAIL MUST BE A DICTIONARY. to parse the model variation
    #print(url)
    response = check_page(url)
    if response == "HTTP Error 400: Bad Request":
        status = "Not Exist"
        return None, None, status
    else:
        try:
            product = product_detail['data']['item']
            preorder = product['is_pre_order']
            models = product["models"]
            cat_id = product['categories'][-1]['catid']
            results = []
            product_status = ''
            if len(models) > 1:
                for model in models:
                    options = split_string(model['name'])
                    result = {
                        'product url' : url,
                        'model id' : str(model['model_id']),
                        'option1' : options[0],
                        'option2' : None if len(options) == 1 else options[1], #WIP
                        'url comparison' : None,
                        'url price' : None,
                        'shopee price' : parsing_price(model['price']),
                        'delta' : None,
                        'catid' : cat_id,
                        'status' : 'unlisted' if product['status'] == 8 else None
                    }
                    results.append(result)
            else:
                result = None
            return preorder, results, product_status, cat_id #preorder = return preorder status between 1 and 0 (true or false) #results = sheet2 results...
        except TypeError:
            product_status = 'Not exist'
            return None, None, product_status, cat_id #product is gone


def filter_None_value(list): #filtered out none value inside a list
    return [item for item in list if item is not None]


def save_output_folder(shop_name, product_result, processed_data):
    #if not os.path.exists(output_folder):
    #    os.makedirs(output_folder)
    #output_file_path = os.path.join(output_folder, f"{shop_name}.xlsx")
    if product_result is not None:
        filtered_product_variation = filter_None_value(product_result)
    else:
        filtered_product_variation = product_result
    df2_value = False
    if filtered_product_variation is not None:
        for truth_value in filtered_product_variation:
            if truth_value:
                df2_value = True
                break
            else:
                print("the list is none")
    with pd.ExcelWriter(f"{shop_name}.xlsx", engine='openpyxl') as file:
        df1 = pd.DataFrame(processed_data)
        df1[['catid']] = df1[['catid']].astype(str).replace(".","")
        df1.to_excel(file, sheet_name="Sheet1", index=False)
        if df2_value:
            df2 = pd.DataFrame(filtered_product_variation)
            df2[['model id', 'catid']] = df2[['model id', 'catid']].astype(str)
            df2.to_excel(file, sheet_name="Sheet2", index=False)
        else:
            pass
    

def read_xlsx(file_name, output="output"):
    # Create the full file path by joining the current directory and the relative path
    #excel_file_path = os.path.join(os.getcwd(), output, file_name)

    # Read the Excel file and store it into a pandas DataFrame
    xlsx = pd.ExcelFile(file_name)
    data1 = pd.read_excel(file_name, sheet_name="Sheet1").to_dict(orient= 'records')
    if len(xlsx.sheet_names) > 1: #to check if the xlsx file has Sheet2 
        data2 = pd.read_excel(file_name, sheet_name="Sheet2").to_dict(orient='records')
        return data1, data2
    else:
        data2 = None
        return data1, data2 #TO READ xlsx files, data1 = sheet1 | data2 = sheet2


def update_variation(product_variation, old_product_variation):
    product_variation['url comparison'] = old_product_variation['url comparison']
    product_variation['url price'] = old_product_variation['url price']
    product_variation['delta'] = old_product_variation['delta']
    product_variation['catid'] = old_product_variation['catid']
    return product_variation

def product_and_price_changes_check(shopid, productid, original_list,shop_name, product_url, price_min, price_max, item):
    for data in original_list:
        if data[f"{shop_name}"] == product_url or shopid and productid in data[f"{shop_name}"]:
            dict_value = data #GETTING PRODUCT DATA FROM ORIGINAL XLSX
            break 
        else :
            dict_value = item
    try:
        index = original_list.index(dict_value)
    except ValueError:
        return dict_value, -1
    #dict_value = next((data for data in original_list if data[shop_name] == product_url), item) not used
    if dict_value is not None:
        try:
            dict_value[shop_name] = item[shop_name]
            dict_value['product added date'] = item['product added date']
            dict_value['pre order'] = item['pre order']
            dict_value['Qty on hand'] = item['Qty on hand']
            dict_value['Number of rating'] = item['Number of rating']
            dict_value['variety'] = item['variety']
            dict_value['url comparison'] = dict_value['url comparison']
            dict_value['url price'] = dict_value['url price']
            dict_value['delta'] = dict_value['delta']
            #dict_value['catid'] = str(item['catid'])
            dict_value["status"] = item['status']
            dict_value["shopee price lowest"] = price_min
            dict_value["shopee price highest"] = price_max
            dict_value[date_now()] = item[date_now()]
            if dict_value["shopee price lowest"] == item["shopee price lowest"] or dict_value["shopee price highest"] == item["shopee price highest"]:           
                return dict_value, index #None means no price change
            else: #if prices are not the same with previous data, it will generate "price changes"
                item['status'] = "price changes" 
                return dict_value, index
        except TypeError: #for development purpose
            item['status'] = "Not Exist"
            return dict_value, index
        

def is_value_in_dicts(shopid, itemid, value, dict_list, shop_name):
    keys = [shop_name, 'catid']
    if value is not None:
        for dictionary in dict_list:
            #print(dictionary)
            if all(dictionary[key] == value[key] for key in keys) and shopid in dictionary[shop_name] and itemid in dictionary[shop_name]:
                #print(True)
                return True #CHECK IF A VALUE in the dictionary
    else:
        print(False)
        return False 

def is_value_in_xlsx2(shopid, itemid, value, dict_list):
    value['model id'] = int(value['model id'])
    key = 'model id'
    val = int(value['model id'])
    for i, dictionary in enumerate(dict_list):
        #print(type(dictionary['model_id']))
        if dictionary.get(key) == val:
            index_number = i
            return True, dictionary, index_number # Return dictionary when found
    return False, value, -1 # Return False and None when not found


def parsing_item_dictionary(item, item_link, preorder,item_price_highest, status, cat_id):
    return {
        item['shop_name']: item_link,
        'product added date' : datetime.fromtimestamp(item['ctime']).strftime('%d-%b-%Y'),
        'pre order' : 1 if preorder == True else 0, 
        'Qty on hand' : item['stock'],
        'Number of rating' : item['item_rating']['rating_count'][0],
        'variety' : 1 if len(item['tier_variations'][0]["options"]) > 1 else 0,
        'url comparison' : None,
        'url price' : None,
        'shopee price lowest' : parsing_price(item['price_min']),
        'shopee price highest' : parsing_price(item_price_highest) if item_price_highest == item['price_max'] else None,
        'delta' : None,
        'catid' : cat_id,
        'status' : 'unlisted' if item['status'] == 8 else status,
        date_now() : item['historical_sold']
    }


def add_new_date(list_of_sheet1):
    new_data = []
    for sheet in list_of_sheet1:
        update_data = list(sheet.items())
        update_data.insert(13, (date_now(), None))
        item = dict(update_data)
        new_data.append(item)
    return new_data


def parsing(datas):  # <--dictionary expected as input
    urls = []
    processed_data = []
    new_data = []
    items = datas["data"]["sections"][0]["data"]["item"]  # Extracting data from the input dictionary
    product_result = []
    new_variant = []
    shop_name = items[0]['shop_name']
    update_datas = []
    #print(shop_name)
    if os.path.exists(os.path.join(os.getcwd(), f"{shop_name}.xlsx")):
        xlsx1, sheet_xlsx2 = read_xlsx(f"{shop_name}.xlsx")
        sheet_xlsx1 = add_new_date(xlsx1)
    for item in items:
        #save_json(item, "sash")
        name = item['name']
        link_name = name.replace(' ','-').replace('%-','%25')
        itemid = str(item['itemid'])
        shopid = str(item['shopid'])
        #print(item['historical_sold'])
        item_link = f"https://shopee.sg/{link_name}-i.{shopid}.{itemid}"
        urls.append(item_link)
        print(f"shopname : {shop_name} scraping product: {item_link}")  
        api_product_url = parsing_individual_product_url(itemid, shopid)
        go_to_product_detail = request_api_access(api_product_url)
        #save_json(go_to_product_detail, "cobacoba.json")
        preorder, product_detail, product_status, cat_id = parsing_var(go_to_product_detail, item_link) #product detail = model variation data/sheet2
        #item_sold = 
        if product_detail is not None:
            product_result.extend(product_detail)
        #print(type(product_detail))
        if item['price_min'] == item['price_max']:
            item_price_highest = None
        else:
            item_price_highest = item['price_max']
        price_check = None
        status = 'sold out' if item['stock'] == 0 else price_check
        if os.path.exists(os.path.join(os.getcwd(), f"{shop_name}.xlsx")):
            #print('benar')
            data, index = product_and_price_changes_check(shopid, itemid, sheet_xlsx1, shop_name, item_link, price_min=parsing_price(item['price_min']), price_max= parsing_price(item['price_max']) if item['price_max'] is not None else None, item=parsing_item_dictionary(item, item_link, preorder, item_price_highest, status if status is not None else product_status, cat_id))
            update_datas.append(data)
            #print(is_value_in_dicts(shopid, itemid,data, sheet_xlsx1, shop_name))
            #print(is_value_in_dicts(shopid, itemid,data, sheet_xlsx1, shop_name))
            if is_value_in_dicts(shopid, itemid,data, sheet_xlsx1, shop_name) == True:
                #print(is_value_in_dicts(data, sheet_xlsx1, shop_name))
                sheet_xlsx1.insert(index, data)
                sheet_xlsx1.pop(index+1)
            else:
                #print("tidak benar")
                new_data.append(data) #STORE NEW PRODUCT into temp list, which will be used to be inserted below from already datas
            if product_detail is not None:
                for product in product_detail:
                    check, old_var_data, index_var = is_value_in_xlsx2(shopid, itemid, product, sheet_xlsx2)
                    #print(check)
                    try:
                        if check:
                            updated_product_var = update_variation(product, old_var_data)
                            sheet_xlsx2.insert(index_var, updated_product_var)
                            sheet_xlsx2.pop(index_var+1)
                        elif check == False:
                            #print("ada varian baru")
                            new_variant.append(product)
                    except ValueError:
                        new_variant.append(product)
        else:
            try:
                data = {
                    item['shop_name']: item_link,
                    'product added date' : datetime.fromtimestamp(item['ctime']).strftime('%d-%b-%Y'),
                    'pre order' : 1 if preorder == True else 0, 
                    'Qty on hand' : item['stock'],
                    'Number of rating' : item['item_rating']['rating_count'][0],
                    'variety' : 1 if len(item['tier_variations'][0]["options"]) > 1 else 0,
                    'url comparison' : "",
                    'url price' : "",
                    'shopee price lowest' : parsing_price(item['price_min']),
                    'shopee price highest' : parsing_price(item_price_highest) if item_price_highest == item['price_max'] else None,
                    'delta' : "",
                    'catid' : cat_id,
                    'status' : 'unlisted' if item['status'] == 8 else status if status is not None else product_status
                    }
            except TypeError: #USED IF PRODUCT IS GONE(not exist) anymore
                data = {
                    item['shop_name']: item_link,
                    'status' : "Not Exist" 
                }
            add_date = list(data.items())
            add_date.insert(13,(date_now(), item['historical_sold']))
            data = dict(add_date)
            processed_data.append(data)
    if not os.path.exists(os.path.join(os.getcwd(), f"{shop_name}.xlsx")):
        save_output_folder(shop_name, product_result, processed_data) #product_result = sheet2 of a shop, processed_data = sheet1
        print(f"{item['shop_name']}.xlsx has been saved")       
        #save_json(processed_data, f"{shop_name}.json") 
        return shop_name, processed_data, product_result #processed data = sheet1, product_result = sheet2
    else: #UPDATE SHOPEE FUNCTION HERE
        unscraped_datas = [(index, url) for index, url in enumerate(sheet_xlsx1) if url[shop_name] not in [item[shop_name] for item in update_datas]]
        #print(unscraped_datas)
        if len(unscraped_datas) > 0:
            updated_sheet1, updated_sheet2 = scraping_leftover(unscraped_datas,shop_name, shopid, sheet_xlsx1, sheet_xlsx2)
            updated_sheet1.extend(new_data)
            if updated_sheet2 is not None:
                updated_sheet2.extend(new_variant)
            while True:
                try:
                    save_output_folder(shop_name, updated_sheet2, updated_sheet1)
                    break
                except PermissionError:
                    print(f"the {shop_name}.xlsx is still opened, please close the file. restarting in 5 second")
                    time.sleep(5)
        else:
            sheet_xlsx1.extend(new_data)
            if sheet_xlsx2 is not None:
                sheet_xlsx2.extend(new_variant)
            while True:
                try:
                    save_output_folder(shop_name, sheet_xlsx2, sheet_xlsx1)
                    break
                except PermissionError:
                    print(f"the {shop_name}.xlsx is still opened, please close the file. restarting in 5 second")
                    time.sleep(5)
        print(f"{item['shop_name']}.xlsx is already there, data xlsx has been updated")
        return shop_name, sheet_xlsx1, sheet_xlsx2


def scraping_leftover(unscraped_datas, shop_name, shopid, sheet_xlsx1, sheet_xlsx2):
    for data in unscraped_datas:
        print(f"shopname : {shop_name} scraping url : {data[1][shop_name]}")
        # Using regular expression to extract the number from the URL
        match = re.search(r'\.(\d+)$', data[1][shop_name])
        if match:
            extracted_number = match.group(1)
            api_url = parsing_individual_product_url(item_id=extracted_number, shop_id=shopid)
            result = request_api_access(api_url)
            try:
                if result['data']['item']['item_status'] == "banned":
                    sheet_xlsx1[data[0]]['status'] = 'not exist'
                    #sheet_xlsx1[data[0]]['catid'] = int(str(sheet_xlsx1[data[0]]['catid']).replace(".",""))
                    sheet_xlsx1[data[0]]['Qty on hand'] = result['data']['item']['stock']
                elif result['data']['item']['status'] == 8:
                    sheet_xlsx1[data[0]]['status'] = "unlisted"
                    sheet_xlsx1[data[0]]['shopee price lowest'] = parsing_price(result['data']['item']['price_min'])
                    sheet_xlsx1[data[0]]['shopee price highest'] = parsing_price(result['data']['item']['price_max']) if result['data']['item']['price_max'] != result['data']['item']['price_min'] else None
                    sheet_xlsx1[data[0]]['Qty on hand'] = result['data']['item']['stock']
                    sheet_xlsx1[data[0]]['Number of rating'] = result['data']['product_review']['rating_count'][0]
                    sheet_xlsx1[data[0]]['variety'] = 1 if len(result['data']['item']['models']) > 1 else 0
                    #sheet_xlsx1[data[0]]['catid'] = result['data']['item']['cat_id']
                elif result['data']['item']['stock'] == 0:
                    sheet_xlsx1[data[0]]['status'] = 'sold out'
                    sheet_xlsx1[data[0]]['Qty on hand'] = 0
                    sheet_xlsx1[data[0]]['shopee price lowest'] = parsing_price(result['data']['item']['price_min'])
                    #sheet_xlsx1[data[0]]['catid'] = int(str(sheet_xlsx1[data[0]]['catid']).replace(".",""))
                    sheet_xlsx1[data[0]]['shopee price highest'] = parsing_price(result['data']['item']['price_max']) if result['data']['item']['price_max'] != result['data']['item']['price_min'] else None
                elif result['data']['item']['stock'] > 0:
                    sheet_xlsx1[data[0]]['status'] = ''
                    sheet_xlsx1[data[0]]['shopee price lowest'] = parsing_price(result['data']['item']['price_min'])
                    sheet_xlsx1[data[0]]['shopee price highest'] = parsing_price(result['data']['item']['price_max']) if result['data']['item']['price_max'] != result['data']['item']['price_min'] else None
                    sheet_xlsx1[data[0]]['Qty on hand'] = result['data']['item']['stock']
                    sheet_xlsx1[data[0]]['Number of rating'] = result['data']['product_review']['rating_count'][0]
                    sheet_xlsx1[data[0]]['variety'] = 1 if len(result['data']['item']['models']) > 1 else 0
                    #sheet_xlsx1[data[0]]['catid'] = result['data']['item']['cat_id']
                sheet_xlsx1[data[0]][date_now()] = result['data']['product_review']['historical_sold'] if sheet_xlsx1[data[0]]['status'] != 'not exist' else None
                for model in result['data']['item']['models']:
                    for xlsx2 in sheet_xlsx2:
                        if str(xlsx2['model id']) == str(model['model_id']) or str(xlsx2['option1']) in model['name']:
                            xlsx2['shopee price'] = parsing_price(model['price'])
                            xlsx2['option1'] = model['name']
                            #xlsx2['model id'] = model['model_id']
                            #Change this to model ID
                for xlsx2 in sheet_xlsx2:
                    if str(extracted_number) in xlsx2['product url']:
                        xlsx2['status'] = sheet_xlsx1[data[0]]['status']
            except TypeError:
                sheet_xlsx1[data[0]]['status'] = 'not exist'
                for xlsx2 in sheet_xlsx2:
                    if str(extracted_number) in xlsx2['product url']:
                        xlsx2['status'] = sheet_xlsx1[data[0]]['status']                   
        else:
            continue
    return sheet_xlsx1, sheet_xlsx2


if __name__ == "__main__":
    time_start = time.time()
    shops_urls = read_xlsx_input("input") 
    print(f"scraping URL(s): {shops_urls}")

    shops_catalogues_results = [] 
    url_products = []
    for url in shops_urls:
        shop_id = shopid(url) 
        #print(shop_id)
        api_url = api_url_target(shop_id)
        api_data = request_api_access(api_url)
        shops_catalogues_results.append(api_data)
    #save_json(shops_catalogues_results, f"{shop_id}.json")
    for result in shops_catalogues_results:
        shop_name, sheet1, sheet2 = parsing(result)
    end_time = time.time()
    elapsed_time = end_time - time_start
    print(f"Script execution time: {elapsed_time:.2f} seconds")