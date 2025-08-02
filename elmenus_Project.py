from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time
import xlsxwriter
import re

# Get user input for location and area
location = input("Enter your governorate: ")
area = input("Enter your area: ")
URL = "https://www.elmenus.com/"

#Open The Elmenus page on Firefox
App = webdriver.Firefox()
App.get(URL)

# Maximize the browser window
App.maximize_window() 

wait = WebDriverWait(App, 10)
time.sleep(2)


# find the dropdown button and click it
wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "vs__dropdown-toggle"))).click()

time.sleep(1)

# find the dropdown input field and enter the location
dropdown_input = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "vs__search")))
dropdown_input.send_keys(location)
dropdown_input.send_keys(Keys.ENTER)

 
time.sleep(1)

wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.submit-btn.btn.btn-primary"))).click()

time.sleep(5)

# Find the location input field and enter the area
location_input = App.find_element(By.XPATH, "//input[@placeholder='Your Location eg. Degla, Maadi']")
location_input.send_keys(area)

time.sleep(2)

# Click the search button
App.find_element(By.CLASS_NAME, "address-btn").click()

time.sleep(10)

#scroll down to load more restaurants
for _ in range(2):
    App.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(5)
    try:
        App.find_element(By.CSS_SELECTOR, ".btn.btn-primary.load-more-btn").click()
        time.sleep(5)
    except:
        break

# Get the page source after the search
page_source = App.page_source
soup = BeautifulSoup(page_source, "html.parser")
restaurants = soup.find_all("div", class_="restaurant-card restaurant-delivery-card col-md-5 col-sm-8 col-xs-16")
print(f"Found {len(restaurants)} restaurants.")
App.quit()

# Start details driver
details_driver = webdriver.Firefox()

restaurants_data = []
for rest in restaurants:
        name = rest.find("h3").text.strip()

        # Extract delivery time
        footer = rest.find("div", class_="card-footer clickable-item")
        delivery_time = "N/A"
        if footer:
            delivery_spans = footer.find_all("span")
            for span in delivery_spans:
                if "mins" in span.text:
                    delivery_time = span.get_text(strip=True).replace("mins", "").strip()
                    break

        reviews = rest.find("span", class_="reviews-count").text.strip()

        # Extract link
        link_tag = rest.find("a", href=True)
        link = "https://www.elmenus.com" + link_tag['href'] if link_tag else "N/A"

        # Extract address and ratings from the restaurant's page
        address = "N/A"
        rates = "N/A"
        try:
            details_driver.get(link)
            time.sleep(4)
            page_source2 = details_driver.page_source
            soup2 = BeautifulSoup(page_source2, "html.parser")

            # Address
            address_tag = soup2.find("p", class_="info-value")
            if address_tag:
                raw_address = address_tag.text
                address = ' '.join(raw_address.split()).replace("'", "").strip()

            # Rating
            ratings_tag = soup2.find("span", class_="vue-star-rating-rating-text")
            if ratings_tag:
                rates = ratings_tag.text.strip()

            

        except Exception as e:
            print(f"Error getting details for {name}: {e}")

        restaurants_data.append({
            "Restaurant Name": name,
            "Address": address,
            "Delivery Time": delivery_time if delivery_time != "N/A" else "غير متوفر",
            "Rating": rates,
            "Rates": reviews.replace("(", "").replace(")", "").strip(),
            "Restaurant Link": link
        })

details_driver.quit()

# Save the data to an Excel file
df = pd.DataFrame(restaurants_data)

# تحويل التقييم ومدة التوصيل لأرقام للتنسيق الصحيح
df["Rating"] = pd.to_numeric(df["Rating"], errors='coerce')
df["Delivery Time"] = pd.to_numeric(df["Delivery Time"], errors='coerce')

# تصدير إلى Excel باستخدام XlsxWriter
with pd.ExcelWriter("restaurants_data.xlsx", engine="xlsxwriter") as writer:
    df.to_excel(writer, index=False, sheet_name="Restaurants")

    workbook = writer.book
    worksheet = writer.sheets["Restaurants"]

    # تنسيقات
    header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter',
        'border': 1, 'bg_color': '#D7E4BC', 'align': 'center'
    })

    center_format = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })

    alt_row_format = workbook.add_format({
        'bg_color': '#F9F9F9', 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })

    high_rating_format = workbook.add_format({
        'bg_color': '#C6EFCE', 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })

    fast_delivery_format = workbook.add_format({
        'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })

    link_format = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': 'blue', 'underline': True
    })

    # كتابة العناوين
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # ضبط أعمدة العرض
    worksheet.set_column("A:A", 25)
    worksheet.set_column("B:B", 110)
    worksheet.set_column("C:C", 15)
    worksheet.set_column("D:D", 12)
    worksheet.set_column("E:E", 15)
    worksheet.set_column("F:F", 25) 
    
    # تنسيق كل صف
    for row_num in range(1, len(df) + 1):
        row_format = center_format if row_num % 2 == 1 else alt_row_format

        for col_num in range(len(df.columns)):
            value = df.iloc[row_num - 1, col_num]

            if pd.isna(value):
                value = ""

            if col_num == 3 and isinstance(value, (int, float)) and value >= 4.5:
                worksheet.write(row_num, col_num, value, high_rating_format)

            elif col_num == 2:
                cell_value = str(value)
                match = re.search(r'\d+', cell_value) 
                if match and int(match.group()) <= 45:
                    worksheet.write(row_num, col_num, value, fast_delivery_format)
                else:
                    worksheet.write(row_num, col_num, value, row_format)

            elif col_num == 5:
                
                worksheet.write_url(row_num, col_num, df.iloc[row_num - 1, col_num], string="link", cell_format=link_format)

            else:
                worksheet.write(row_num, col_num, value, row_format)

