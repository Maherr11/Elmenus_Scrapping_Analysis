# Scrapping Restaurant Delivery & Rating Analysis

This project analyzes restaurant data collected from [Elmenus](https://www.elmenus.com/), a food delivery platform in Egypt.  
The aim is to explore trends in **delivery times**, **customer ratings**, and **number of reviews**, and visualize key insights to better understand restaurant performance.

---

## Project Highlights

- **Data Collection**: Data was scraped using `Selenium` and `BeautifulSoup`
- **Data Analysis**:
  - Delivery time and rating distribution
  - Correlation between ratings and number of reviews
  - Boxplots to analyze rating variations by delivery speed
- **Excel Summary**: Generated a summary report of average ratings and delivery statistics

---

## Files Structure

| File Name              | Description                                       |
|------------------------|---------------------------------------------------|
| `Elmenus_Project.py`    | Python script for Scrapping and creating files   |
| `restaurants_data.xlsx` | Cleaned dataset from Elmenus (XLSX format)       |
| `analysis.py`           | Python script for all visualizations             |
| `summary_report.xlsx`   | Exported summary including averages & counts     |

---

## 🛠️ Tech Stack

- **Python**
- **Pandas** – for data cleaning & manipulation  
- **Seaborn / Matplotlib** – for plotting insights  
- **Openpyxl** – for Excel writing  
- *(Data scraping part: Selenium + BeautifulSoup)*

---

## Sample Visualizations

Here are a few sample plots generated from the data:

- Distribution of Delivery Times
- Distribution of Rating
- Rating Distribution by Delivery Speed  
- Ratings and Number of Reviews

---

## How to Run

1. Clone this repo  
2. Make sure you have the required libraries:
   ```bash
   pip install pandas matplotlib seaborn openpyxl
