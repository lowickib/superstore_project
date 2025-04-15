## 📚 Table of Contents

1. [🧾 Summary](#-summary)
2. [📦 Project Description](#-project-description)
3. [📁 Dataset](#-dataset)
    - [Key Features](#key-features-include)
4. [🧹 Data Preparation (Power Query)](#-data-preparation-power-query)
    - [Superstore Table](#superstore-table)
    - [Customers Table](#customers-table)
    - [Products Table](#products-table)
    - [Orders Table](#orders-table)
    - [Monthly_Orders Table](#monthly_orders-table)
5. [📊 RFM Analysis – Superstore Project](#-rfm-analysis--superstore-project)
    - [Objective](#objective)
    - [Input Data](#input-data)
    - [RFM Analysis Steps](#rfm-analysis-steps)
    - [Customer Segmentation](#customer-segmentation)
    - [Outcome](#outcome)
6. [📊 Sales & Profit Insights](#-sales--profit-insights)
    - [Sales and Profit by Cities and States](#sales-and-profit-by-cities-and-states)
    - [Sales and Profit by Category and Sub-Category](#sales-and-profit-by-category-and-sub-category)
    - [Sales and Profit YTD (Year-to-Date)](#sales-and-profit-ytd-year-to-date)
    - [Sales and Profit MAT (Moving Annual Total)](#sales-and-profit-mat-moving-annual-total)
    - [Sales and Profit MAT by Region and Category](#sales-and-profit-mat-by-region-and-category)
    - [Sales and Profit MA (Moving Average)](#sales-and-profit-ma-moving-average)
7. [📉 Discount Insights](#-discount-insights)
    - [Discount Categories](#discount-categories)
8. [👥 Customer Insights](#-customer-insights)
    - [Sales and Profit by Customer Type](#sales-and-profit-by-customer-type)
    - [Customer Type vs. Product Category](#customer-type-vs-product-category)
    - [Customer Type vs. Region](#customer-type-vs-region)
    - [Customer Ordering Behavior](#customer-ordering-behavior)
9. [💡 What I Learned](#-what-i-learned)


## 🧾 Summary

This project is a comprehensive sales performance and customer segmentation analysis based on the *Superstore* dataset from Kaggle. Using Excel and Power Query, the dataset was cleaned, transformed, and enriched to support in-depth business intelligence insights. Key components include RFM analysis, discount strategy evaluation, product and regional performance breakdowns, and customer behavior insights. The final deliverables consist of a series of well-structured dashboards and visualizations aimed at supporting strategic decisions around marketing, sales, and operations.

## 📦 Project Description

The project focuses on turning raw transactional sales data into actionable business insights using Excel. Key steps included:

- **Data Preparation with Power Query**: Cleaning, enriching, and transforming the raw dataset into structured tables (Orders, Customers, Products, Monthly Orders).
- **RFM Segmentation**: Customers were scored and segmented based on recency, frequency, and monetary metrics to identify top performers and those at risk.
- **Sales Performance Analysis**: Sales and profit were analyzed across dimensions like region, product category, sub-category, and customer type.
- **Time-Based Trends**: Moving averages and moving annual totals (MAT) revealed long-term trends and seasonal patterns.
- **Discount Evaluation**: The impact of discount levels on sales and profitability was visualized and quantified.
- **Customer Behavior**: Purchasing frequency and order timing were studied to understand engagement levels per customer segment.

All insights were visualized clear dashboards, using Excel’s advanced charting capabilities.


## 📁 Dataset

The dataset used in this project was sourced from Kaggle:

🔗 [Superstore Dataset – Final](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final/data)

This dataset contains detailed information about sales transactions, customers, products, and regions for a fictional retail store. It is commonly used for sales analytics and data visualization practice.

### Key features include:
- **Order details**: Order ID, order date, ship date, delivery time
- **Customer information**: Customer ID, name, segment
- **Product details**: Product ID, name, category, sub-category
- **Geographic info**: City, state, postal code, region
- **Financials**: Sales, discount, profit, quantity
- **Derived metrics**: Discount value, delivery time, discount type

This dataset was imported and transformed using Power Query in Excel before performing further analysis such as RFM segmentation and sales performance reporting.


## 🧹 Data Preparation (Power Query)

The raw dataset `Sample - Superstore.csv` was cleaned and transformed using Power Query (M language). Below is a breakdown of the data preparation steps:


### Superstore Table

The base query loads the raw CSV file and applies the following transformations:
- Promoted headers.
- Converted column types (dates, currency, numbers).
- Added a **Discount Value** column: $Discount$ $Value = Discount * Sales$
- Calculated **Delivery Time** as the number of days between `Order Date` and `Ship Date`.
- Renamed `Subtraction` to `Delivery Time`.


### Customers Table

Derived from the `Superstore` table by:
- Removing all non-customer-related columns (e.g., product, order, sales).
- Removing duplicate rows to ensure one record per unique customer.

### Products Table

Derived from the `Superstore` table by:
- Keeping only product-related columns: `Product ID`, `Category`, `Sub-Category`, `Product Name`.
- Removing duplicates based on `Product ID`.

### Orders Table

Derived from the `Superstore` table by:
- Removing unnecessary columns (e.g., customer name, product name, category).
- Removing duplicates to keep unique order entries.
- Creating a new column **Discount Type** using logic based on discount brackets:
  - `D – 51%+`
  - `C – 31–50%`
  - `B – 11–30%`
  - `A – 0–10%`

### Monthly_Orders Table

Created to support time-based analysis:
- Extracted `Year` and `Month` from the `Order Date`.
- Grouped data by `Year`, `Month`, `Region`, and `Category`.
- Aggregated total `Sales` and `Profit` per group.

---

These cleaned and structured tables form the basis for further analysis, including RFM segmentation, customer behavior patterns, and time-based trends.


## 📊 RFM Analysis – Superstore Project

### Objective

The goal of the RFM (Recency, Frequency, Monetary) analysis is to segment customers based on their purchase history in order to:
- Identify the most valuable and loyal customers,
- Detect inactive or at-risk customers,
- Optimize marketing, retention, and sales strategies.

### Input Data

Data was sourced from the file `superstore_project_v1.xlsx`, specifically from the following sheets:
- `Orders` – transactional data,
- `Customers` – customer information,
- `RFM` – final RFM output.

### RFM Analysis Steps

#### 1. Data Collection
From the `Orders` sheet, the following fields were used:
- `Customer Name`
- `Order Date`
- `Sales`

#### 2. Data Aggregation
For each customer, the following metrics were calculated:
- **Last Order Date** – date of the most recent purchase,
- **Count of Orders** – total number of orders placed,
- **Average Sale** – average order value.

#### 3. RFM Scoring
Based on the aggregated data, the following scores were assigned:

| Metric           | Description |
|------------------|-------------|
| **Recency Score**   | How recently the customer made a purchase – more recent = higher score |
| **Frequency Score** | How often the customer purchases – more orders = higher score |
| **Monetary Score**  | Average value of purchases – higher spend = higher score |

#### 4. Total RFM Score
A combined score was calculated using: $RFM = Recency Score + Frequency Score + Monetary Score$


An additional simplified `RFM Score` (on a 1–10 scale) was created for easier segmentation.

### Customer Segmentation

Based on the `RFM Score`, customers were assigned to one of five segments:

| Segment | Description |
|---------|-------------|
| `A – Top Customer`          | Most valuable, loyal and active customers |
| `B – Loyal / Engaged`       | Regular, reliable customers |
| `C – Promising / Average`   | Average customers with potential |
| `D – At Risk / Sleeping`    | Customers who haven’t purchased in a while |
| `E – Churned / Inactive`    | Lost or inactive customers |


### Outcome

This RFM analysis allows for:
- Identification of high-value customers (segment A),
- Re-engagement of inactive customers (segments D and E),
- Personalized marketing campaigns tailored to customer type.

## 📊 Sales & Profit Insights

### Sales and Profit by Cities and States
![Sales and Profit by Cities and States](assets/Sales%20and%20Profit%20by%20Cities%20and%20States.png)

🛠️ **Excel Features:** Four horizontal bar charts displaying the top 10 states and cities by total sales and profit. A consistent color scheme and layout (Sales on the left, Profit on the right) enhance visual clarity and comparison.  
🎨 **Design Choice:** 2x2 grid layout for quick side-by-side comparisons between states and cities, and between sales and profit. Highlights regional performance patterns.  
📉 **Data Organization:** Data grouped by location (State, City) and visualized using two key metrics: total Sales and total Profit. Each chart sorted in descending order.  
💡 **Insights Gained:**  
- California is the top-performing state in both sales and profit, confirming its position as the most lucrative market.  
- New York City leads among cities for both sales and profit, making it a strategic hotspot.  
- Texas ranks high in sales but doesn't appear in the top 10 for profit, indicating possible margin issues.  
- Detroit, Lafayette, and Jackson appear in the profit rankings without being top in sales, suggesting strong profitability or efficient operations in those cities.

![Sales and Profit by Cities and States](assets/Sales%20and%20Profit%20by%20Cities%20and%20States%20Scatter.png)

🛠️ **Excel Features:** Scatter plot comparing total sales (X-axis) and total profit (Y-axis) by state. Data points are color-coded to distinguish positive (light blue) and negative (dark blue) profit. Key states are labeled for emphasis.  
🎨 **Design Choice:** Dual-axis design allows for simultaneous comparison of sales volume and profitability. Color-coded points enhance immediate understanding of financial performance.  
📉 **Data Organization:** Each data point represents a U.S. state, showing total sales and corresponding profit. Labels highlight states with extreme or notable values.  
💡 **Insights Gained:**  
- California and New York have the highest sales and profits, making them top strategic markets.  
- Washington shows strong profitability despite relatively moderate sales — an indicator of high efficiency or strong margins.  
- Texas and Pennsylvania generate high sales but suffer losses, pointing to potential issues like low margins or high operational costs.  
- Ohio reports both low sales and negative profit — a candidate for further analysis or resource reallocation.

### Sales and Profit by Category and Sub-Category

![Sales and Profit by Category and Sub-Category](assets/Sales%20and%20Profit%20by%20Category%20and%20Sub-Category.png)

🛠️ **Excel Features:** Four horizontal bar charts comparing total sales and profit by Category and Sub-Category. Sales charts include data labels for added clarity.  
🎨 **Design Choice:** Categories and sub-categories are displayed side-by-side, allowing layered analysis — from broad market segments to specific product types.  
📉 **Data Organization:**  
- Data is grouped first by three main categories: Technology, Office Supplies, and Furniture.  
- Sub-category charts are nested and ordered by sales/profit values within their parent categories.

💡 **Insights Gained:**  
- **Technology** leads in both total sales and profit, confirming its role as the most valuable category.  
- **Office Supplies** has solid sales and the second-highest profit, showing consistent performance across many sub-categories.  
- **Furniture** has high sales but the lowest profit — and even losses in several sub-categories, such as **Tables** and **Bookcases**.  
- At the sub-category level:  
  - **Phones** generate the highest sales, followed by **Chairs** and **Storage**.  
  - **Copiers** are the most profitable sub-category, despite relatively low sales — indicating excellent margins.  
  - Several sub-categories in Furniture (especially **Tables**) show negative profit despite moderate sales, signaling pricing or cost inefficiencies.

![Sales and Profit by Category and Sub-Category](assets/Sales%20and%20Profit%20by%20Category%20and%20Sub-Category%20Scatter.png)

🛠️ **Excel Features:** Scatter plot comparing sales and profit for each sub-category. Color coding distinguishes between main categories:  
- Light blue: Furniture  
- Medium blue: Office Supplies  
- Darker blue: Technology  

🎨 **Design Choice:** Plotting sub-categories on a dual-axis chart enables a detailed performance overview by combining revenue potential (Sales) with efficiency (Profit). Labels help highlight key product types.

📉 **Data Organization:** Each dot represents one sub-category and is positioned according to its total sales (X-axis) and total profit (Y-axis). Category grouping is visualized using color.

💡 **Insights Gained:**  
- **Phones** and **Chairs** lead in sales, but **Phones** generate significantly higher profit, signaling better margins.  
- **Copiers** stand out with the highest profit, despite moderate sales — excellent margin performance.  
- **Tables** suffer from both high sales and strong negative profit, suggesting a major cost or pricing issue.  
- Several Office Supplies sub-categories (e.g., **Labels**, **Supplies**, **Envelopes**) show either minimal or negative profit — candidates for deeper analysis.  
- **Paper**, **Accessories**, and **Binders** combine solid sales with high profitability, making them strong overall performers.


### Sales and Profit YTD (Year-to-Date)
![Sales and Profit YTD (Year-to-Date)](assets/Sales%20and%20Profit%20YTD.png)

🛠️ **Excel Features:** Combination of column and line charts displaying annual and monthly cumulative (YTD) sales and profit across four years (2014–2017). Data labels added for clarity on final totals.  
🎨 **Design Choice:**  
- Bar charts provide quick annual comparisons.  
- Line charts highlight monthly trends and growth dynamics across years.  
- Consistent color coding for each year enhances readability and comparison.  
📉 **Data Organization:**  
- Sales and Profit are displayed side-by-side for each year (top row).  
- YTD values are plotted by month for each year (bottom row), giving a time-based perspective of performance.

💡 **Insights Gained:**  
- 2017 recorded the highest total **Sales** ($733,215) and **Profit** ($93,439), reflecting strong year-over-year growth.  
- There is a steady upward trend in both sales and profit from 2014 to 2017.  
- Growth acceleration is visible in the YTD lines — especially after month 5, where 2017 begins to pull ahead of previous years.  
- Each year maintains consistent seasonal patterns, but the gap between years widens — suggesting sustained improvement in business performance.

### Sales and Profit MAT (Moving Annual Total)

![Sales and Profit MAT (Moving Annual Total)](assets/Sales%20and%20Profit%20MAT.png)

🛠️ **Excel Features:** Two line charts tracking Moving Annual Total (MAT) for Sales and Profit over time. X-axis shows months across multiple years, Y-axis represents annualized rolling totals.  
🎨 **Design Choice:** Using MAT smooths out seasonal fluctuations and highlights long-term trends. Line charts allow for continuous tracking of business momentum.  
📉 **Data Organization:**  
- Each data point reflects the total sales/profit over the trailing 12 months.  
- Time is shown as a continuous axis combining years and months for better trend visibility.

💡 **Insights Gained:**  
- Both **Sales MAT** and **Profit MAT** show a clear upward trajectory from 2014 to late 2017.  
- The sharpest growth occurred in 2014, with a plateau in 2015, followed by renewed acceleration starting mid-2016.  
- The peak in late 2017 marks the highest performance period, although the slight decline at the end suggests a potential slowdown.  
- Profit MAT mirrors sales trends but shows greater volatility, indicating variations in cost structure or pricing effectiveness over time.

### Sales and Profit MAT by Region and Category

![Sales and Profit MAT by Region and Category](assets/Sales%20and%20Profit%20MAT%20by%20Region%20and%20Category.png)

🛠️ **Excel Features:** Four line charts showing Moving Annual Total (MAT) for Sales and Profit, segmented by Region and Product Category. Separate color schemes distinguish each group within charts.  
🎨 **Design Choice:** Comparative line charts reveal long-term performance trends across multiple dimensions. MAT smooths out seasonality and emphasizes overall momentum.  
📉 **Data Organization:**  
- Top row: Sales and Profit MAT by Region (Central, East, South, West).  
- Bottom row: Sales and Profit MAT by Product Category (Furniture, Office Supplies, Technology).  
- Each line represents a 12-month rolling sum per dimension across time.

💡 **Insights Gained:**  
**By Region:**  
- The **West** region consistently leads in both sales and profit MAT, showing strong and sustained growth.  
- The **East** region showed early strength but declined toward the end of 2017, especially in profit.  
- **Central** and **South** regions lag behind, with South being the least profitable overall.

**By Category:**  
- **Technology** drives the highest MAT in both sales and profit, peaking at the end of 2017.  
- **Office Supplies** maintain steady growth in sales, with profit catching up in the latter part of the timeline.  
- **Furniture** shows the weakest profit MAT, staying relatively flat throughout the period despite reasonable sales — suggesting lower margins or cost inefficiencies.

### Sales and Profit MA (Moving Average)

![Sales and Profit MA (Moving Average)](assets/Sales%20and%20Profit%20MA.png)

🛠️ **Excel Features:** Two line charts showing raw monthly data overlaid with a moving average (MA) line for both Sales and Profit. The MA smooths short-term fluctuations and highlights the trend.  
🎨 **Design Choice:** Overlaying MA on top of raw data allows simultaneous visibility of seasonality and underlying trend. Thin line for actuals, bold line for trend enhances contrast and interpretation.  
📉 **Data Organization:**  
- Left: Monthly **Sales** and **Sales MA**.  
- Right: Monthly **Profit** and **Profit MA**.  
- Both charts use consistent time axis (Year and Month) and separate Y-axes for magnitude clarity.

💡 **Insights Gained:**  
- Sales and Profit show clear seasonal spikes throughout each year — especially in Q4.  
- Moving Averages (MA) smooth out volatility and indicate consistent upward trends despite temporary dips.  
- The strongest **Sales MA** growth is visible in 2017, while **Profit MA** also trends upward, though with more fluctuations.  
- Profit volatility is higher than sales, revealing sensitivity to cost, pricing, or external factors even during strong revenue months.

## 📉 Discount Insights

### Discount Categories

![Sales and Profit by Discount Categories](assets/Sales%20and%20Profit%20by%20Discount%20Categories.png)

🛠️ **Excel Features:** Two clustered column charts displaying the average sales and average profit per discount category. Data labels included for precise comparison.  
🎨 **Design Choice:** Simple side-by-side layout enables easy analysis of how increasing discount levels affect both sales and profitability.  
📉 **Data Organization:**  
- Categories are grouped by discount range:  
  - A = 0–10%  
  - B = 11–30%  
  - C = 31–50%  
  - D = 51%+  
- Left chart: Average sales per discount category  
- Right chart: Average profit per discount category

💡 **Insights Gained:**  
- **Category C (31–50%)** drives the highest average sales, but results in the **largest average loss** (–$156.28).  
- **Category D (51%+)** also leads to negative profit, suggesting deep discounts are not financially viable.  
- **Category A (0–10%)** yields the **highest average profit ($67.46)** despite moderate sales, highlighting the benefit of low-discount transactions.  
- A clear negative correlation exists between discount level and profitability — **higher discounts lead to lower or negative margins**, which could undermine business sustainability.

![Sales and Profit by Discount Categories](assets/Discount%20Categories%20–%20Transaction%20Volume%20and%20Total%20Value.png)

🛠️ **Excel Features:**  
- Clustered column chart showing the number of transactions by discount category  
- Combo chart comparing total sales (bar) and total profit (line) on dual axes  

🎨 **Design Choice:**  
- Transaction volume chart gives a sense of scale for each discount category  
- Combo chart enables simultaneous evaluation of revenue and profit, making it easier to spot imbalance between sales and profitability  

📉 **Data Organization:**  
- Left chart: Number of transactions per discount range (A–D)  
- Right chart: Total sales and total profit for each discount category, using primary and secondary axes  

💡 **Insights Gained:**  
- **Category A (0–10%)** not only has the highest number of transactions (4,892), but also delivers the highest total sales **and** profit — it’s the healthiest segment overall.  
- **Category B (11–30%)** shows a strong number of transactions and solid sales, but significantly lower profit.  
- **Category C (31–50%)** and **D (51%+)** have far fewer transactions, yet generate substantial losses, confirming earlier insight that **deep discounts are financially harmful**.  
- The negative profit trend is sharp and clearly correlated with increasing discount level — the higher the discount, the worse the business outcome.

## 👥 Customer Insights

### Sales and Profit by Customer Type

![Sales and Profit by Customer Type](assets/Sales%20and%20Profit%20by%20Customer%20Type.png)

🛠️ **Excel Features:** Three bar charts showing:  
- Number of customers per type  
- Average sales and profit per customer type  
- Percentage of total customers vs total profit by category  

🎨 **Design Choice:** The segmentation is based on customer value (RFM-style), allowing for a clear view of both volume and profitability per segment. Comparing customer count against profit highlights the hidden value of key groups.

📉 **Data Organization:**  
Customers are divided into five types:  
- A – Top Customer  
- B – Loyal & Valuable  
- C – Promising / Average  
- D – At Risk / Sleeping  
- E – Churned / Inactive  

The data presents:  
- Total number of customers per segment  
- Average sales and profit per customer  
- Share of each segment in overall customer base and total profit  

💡 **Insights Gained:**  
- **Segment E – Churned/Inactive** is the largest by volume (258 customers) but generates only **9.58% of total profit** — high volume, low value.  
- **Segment A – Top Customer** is the smallest group (80 customers) yet contributes **23.70% of total profit** — extremely high individual value.  
- **Segment B – Loyal & Valuable** is the most balanced group — contributing **the largest share of total profit (29.60%)** with 20.62% of the customer base.  
- There’s a clear gradient in value: **the higher the customer segment, the greater the average sales and profit per customer** (e.g., A: $325.30 in sales, $43.71 in profit).  
- Segments **C** and **D** show mid-range potential — with meaningful customer count but lower contribution — suitable for upselling or reactivation campaigns.


### Customer Type vs. Product Category

![Customer Type vs. Product Category](assets/Customer%20Type%20vs.%20Product%20Category.png)

🛠️ **Excel Features:**  
- Two clustered column charts showing average sales and profit by product category and customer type  
- One stacked bar chart visualizing the percentage distribution of orders across product categories for each customer type  

🎨 **Design Choice:**  
Segmented comparison across customer types helps identify how each segment interacts with product categories. Color coding enhances clarity by consistently assigning shades to categories (Furniture, Office Supplies, Technology).

📉 **Data Organization:**  
- X-axis: Customer Types (A to E)  
- Metrics: Average Sales, Average Profit, and Order Composition per product category  
- Categories: Furniture, Office Supplies, Technology  

💡 **Insights Gained:**  
- **Top Customers (A)** and **Loyal & Valuable (B)** customers spend and profit most from **Technology** products — it’s the key revenue and profit driver for high-value clients.  
- **At Risk (D)** and **Churned (E)** segments show lower overall average sales and profit across all categories — with especially weak performance in **Furniture**.  
- **Furniture** contributes the least profit across all customer types — especially noticeable with nearly **zero or negative profit** in the average profit chart.  
- In terms of order composition:  
  - **Technology** consistently dominates all segments, representing the largest share of orders.  
  - **Office Supplies** have a relatively stable contribution across types, but their **profit margin is significantly higher**, especially for segments A and B.  
  - **Furniture** is a minor part of the order mix and yields the lowest returns — both in sales and profit.

### Customer Type vs. Region

![Customer Type vs. Region](assets/Customer%20Type%20vs.%20Region.png)

🛠️ **Excel Features:**  
- Clustered column chart for absolute number of customers by type across regions  
- 100% stacked bar chart to show proportional distribution of customer types within each region  

🎨 **Design Choice:**  
This combination offers a dual perspective: actual volume of customers per region and relative structure within each. Color-coded segments make it easy to compare customer type distribution regionally.

📉 **Data Organization:**  
- X-axis: Regions (Central, East, South, West)  
- Categories: Customer types A–E (from Top Customers to Churned/Inactive)  
- Metrics: Total count of customers and their proportion within each region  

💡 **Insights Gained:**  
- The **West** region has the largest customer base overall, followed by **East** and **Central**.  
- **South** has the smallest customer population across all types.  
- All regions show a **similar proportional structure**, with roughly equal shares of Top (A), Loyal (B), and Promising (C) customers.  
- **Churned/Inactive (E)** and **At Risk (D)** customers make up over 40% of the customer base in each region — a strong indication of untapped retention opportunities.  
- **East** and **West** are strategic focus areas due to both high volume and healthy share of valuable customers (A + B types).

### Customer Ordering Behavior

![Customer Ordering Behavior](assets/Customer%20Ordering%20Behavior.png)

🛠️ **Excel Features:**  
Two horizontal bar charts:  
- One shows the **average number of orders** per customer type  
- The other displays the **average time between orders** (in days)  

🎨 **Design Choice:**  
Horizontal layout accommodates long labels and allows for a side-by-side behavioral comparison across segments. These two metrics complement each other to highlight purchase frequency and consistency.

📉 **Data Organization:**  
- Y-axis: Customer types from A (Top) to E (Churned)  
- X-axis (Chart 1): Average number of orders  
- X-axis (Chart 2): Average days between orders  

💡 **Insights Gained:**  
- **Top Customers (A)** have the highest engagement — nearly **20 orders on average** and the **shortest interval between purchases (~65 days)**.  
- As we move down the value ladder (from B to E), **the number of orders decreases** and **the time between purchases increases**, clearly showing reduced engagement.  
- **Churned/Inactive (E)** customers place the **fewest orders (~8)** and have the **longest average gap (~118 days)** — confirming their inactive status.  
- The behavioral pattern is strongly aligned with customer value segments — demonstrating that **frequency and recency are key indicators of loyalty and profitability**.

## 💡 What I Learned

Throughout the project, I gained valuable experience in:

- **Data modeling and cleaning** using Power Query (M language) in Excel.
- **Building a data model** from a single flat file, including table creation and relationship structuring.
- **Applying RFM analysis** to classify and understand customer behavior and value.
- **Designing effective dashboards** that communicate insights clearly using bar charts, scatter plots, line charts, combo charts, and more.
- **Analyzing discount strategies**, identifying unprofitable segments, and linking discounts to customer types.
- **Using advanced Excel functions** like `AVERAGEIFS`, `COUNTIF`, `SUMIFS`, and `VLOOKUP` to automate calculations and segmentation.
- **Deriving actionable insights** from data — such as identifying churned customers, unprofitable products, and high-potential customer groups.

This project strengthened my ability to combine data transformation, business thinking, and visualization in one coherent analytical workflow.