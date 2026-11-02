from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date
import time
from playwright.sync_api import Error as PWError, TimeoutError as PWTimeout

laptops_list=[]
with sync_playwright() as p:
    def open_page(page, url, max_retries=3):
        for attempt in range(max_retries):
            try:
                response = page.goto(url, timeout=15000)
                if response is not None and response.status in (429, 500, 502, 503, 504):
                    time.sleep(2 ** attempt)          # خطأ خادم مؤقت ← أعد
                    continue
                if response is None or not response.ok:        # ok = الحالة 200-299
                    raise RuntimeError(f"فشل الفتح: {response.status if response is not None else 'بلا استجابة'}")
                return response
            except PWTimeout:
                time.sleep(2 ** attempt)              # بطء ← أعد
            except PWError as e:
                if "net::ERR_" in str(e):
                    time.sleep(2 ** attempt)          # شبكة ← أعد
                else:
                    raise   # أي خطأ آخر ← لا تعد، اكشفه
        raise RuntimeError(f"فشل فتح {url} بعد {max_retries} محاولات")
    try:
        URL="https://webscraper.io/test-sites/e-commerce/more/computers/laptops"
        browser=p.chromium.launch(headless=True)
        page=browser.new_page()
        response= open_page(page, URL)
        CARD = "div.product-wrapper.card-body"
        while True:
            count = page.locator(CARD).count()          # عدد المنتجات (بدل n)
            if count >= 200:                            # يصحح 4 و6
                break
            more = page.locator("a.btn-primary").filter(has_text="More")
            if not more.is_visible():
                break
            try:
                more.click(timeout=5000)                # يصحح 5
                page.wait_for_function(                 # يصحح 3
                    "([sel, n]) => document.querySelectorAll(sel).length > n",
                    arg=[CARD, count], timeout=10000)
            except PWTimeout:
                break
        products=page.locator(CARD).all()
        for laptop in products:

            title=laptop.locator("a.title")
            if title.count()>0:
                laptop_title=title.first.get_attribute("title")  
            else:
                laptop_title="No title"

            price=laptop.locator("h4.price.float-end.pull-right")
            if price.count()>0:
                laptop_price=price.first.inner_text()
            else:
                laptop_price=None
            
            description=laptop.locator("p.card-text.description")
            if description.count()>0:
                laptop_description=description.first.inner_text()
            else:
                laptop_description="No description"
            
            reviews=laptop.locator("p.review-count.float-end.pull-right > span")
            if reviews.count()>0:
                laptop_reviews=reviews.first.inner_text()
            else:
                laptop_reviews=0

            laptop_dict={"Laptop name": laptop_title , "Price": laptop_price , "Description": laptop_description , "Reviews": laptop_reviews}
            laptops_list.append(laptop_dict)
    finally:
        browser.close()

    # print(laptops_list[:5])
    # print(len(laptops_list))
    # print("************************")

    df=pd.DataFrame(laptops_list)
    # print(df.head(3))
    # print("----------")
    # df.info()
    # print("----------")
    # print(df.sample(8))
    df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
    df["Price"]=df["Price"].str.replace("$" , "")
    df = df.dropna(subset=["Price"])
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce")
    df["Reviews"] = pd.to_numeric(df["Reviews"], errors="coerce")
    df.rename(columns={"Price":"Price ($)"} , inplace=True)
    df.drop_duplicates(subset=["Laptop name" , "Price ($)" , "Reviews" , "Description"] , inplace=True)
    # print("------------")
    # print(f"Type of each column: {df.dtypes}")
    # print("---------")
    # df.info()    
    # print("-----")
    # print(df.sample(3))
    print(f"The average price for all laptops ({len(df)} laptops): {df["Price ($)"].mean()}")
    df=df.sort_values(["Price ($)" , "Laptop name"] , ascending=[True,True]).reset_index(drop=True)
    print(f"The five cheapest laptops: \n{df.head(5)}")
    print(f"The five most expensive laptops: \n{df.tail(5)}")
    df["Brand"]=df["Laptop name"].str.split().str[0]
    # print("^^^^^^^^^^^^^^^^^^")
    # print(df["Brand"].unique())
    a=["Pavilion" , "Hewlett" , "ProBook" , "Packard"]
    for i in a:
        df["Brand"]=df["Brand"].str.replace(i , "HP") 
    df["Brand"]=df["Brand"].str.replace("Aspire" , "Acer")
    df["Brand"]=df["Brand"].str.replace("Inspiron" , "Dell")
    df["Brand"]=df["Brand"].str.replace("ThinkPad" , "Lenovo")
    # print(df["Brand"].unique())
    average_of_each_brand=df.groupby("Brand")["Price ($)"].agg(["count" , "mean"]).round(2)
    # print(average_of_each_brand)
    max_mean=average_of_each_brand["mean"].max()
    min_mean=average_of_each_brand["mean"].min()
    print(f"The most expensive brand among all brands: {average_of_each_brand["mean"].idxmax() } , with average price {max_mean} $")
    print(f"The cheapest brand among all brands: {average_of_each_brand["mean"].idxmin() } , with average price {min_mean} $")
    df.drop(columns=["Brand"] , inplace=True)
    df=df.sort_values(["Reviews" , "Price ($)"] , ascending=[False,True]).reset_index(drop=True)
    print(f"The number of laptops: {len(df)}")

    today = date.today().strftime("%Y-%m-%d")
    with pd.ExcelWriter(f"laptops_{today}.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All laptops", index=False)
        average_of_each_brand.to_excel(writer, sheet_name="Avg price by brand")

        