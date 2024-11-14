# Coffee Bean Sales VBA Automation

This project provides a set of VBA macros designed to automate common data processing tasks for a coffee bean sales dataset. The dataset includes information on orders, customers, and products, and these scripts streamline operations like sales reporting, customer loyalty tracking, and inventory management.

## Dataset

The dataset used for this project is sourced from Kaggle: [Coffee Bean Sales Raw Dataset](https://www.kaggle.com/datasets/saadharoon27/coffee-bean-sales-raw-dataset). It contains the following sheets:
- **orders**: Information on individual orders, including Order ID, Order Date, Customer ID, Product ID, and Quantity.
- **customers**: Customer details, such as Customer ID and location.
- **products**: Details on product IDs and prices.

## VBA Modules Overview

The following VBA macros are included in the `Coffee_Bean_Sales_VBA.vb` file:

### 1. **Calculate Total Sales per Order**
Calculates total sales per order in the "orders" sheet by multiplying the quantity by the product price from the "products" sheet.

### 2. **Highlight Large Orders**
Highlights rows with large order quantities (above a specified threshold) in the "orders" sheet.

### 3. **Filter and Copy Orders by Date**
Filters orders within a specific date range and copies the results to a new sheet called "Filtered Orders."

### 4. **Generate Customer Summary**
Summarizes the total quantity of products ordered by each customer and outputs the data in a new sheet, "Customer Summary."

### 5. **Automated Sales Report Generation**
Generates a report with total quantity sold and total sales amount for each product, saved in a new sheet, "Sales Report."

### 6. **Price Update Automation**
Prompts the user to update product prices in the "products" sheet by entering a Product ID and the new price.

### 7. **Customer Loyalty Program Automation**
Creates a list of loyal customers who have ordered more than a specified quantity threshold and outputs it to a new sheet, "Loyal Customers."

### 8. **Inventory Tracking and Low Stock Alerts**
Checks the stock levels in the "products" sheet and alerts the user if stock levels fall below a specified threshold.

### 9. **Monthly Sales Summary by Country**
Generates a monthly sales summary by country, outputting the data to a new sheet, "Monthly Sales by Country."

## Setup

1. Open the **Coffee Bean Sales** Excel workbook.
2. Open the VBA editor by pressing `ALT + F11`.
3. Import the `Coffee_Bean_Sales_VBA.vb` file into the VBA editor by navigating to **File > Import File...**.
4. Ensure the dataset's structure and sheet names match those described above.

## Usage

- **Run a Macro**: To execute a macro, go to the Excel ribbon, select **Developer > Macros**, choose the macro you want to run, and click **Run**.
- **Modify Thresholds and Parameters**: Some macros, like the Customer Loyalty Program and Low Stock Alerts, have predefined thresholds. Update these values directly in the VBA code as needed.

## Notes

- **Data Requirements**: Ensure the "products" sheet contains columns for product ID and price, "customers" contains customer details and locations, and "orders" includes all fields required by the macros.
- **Version**: This VBA project was developed with Excel 365. Compatibility with other versions of Excel has not been tested.
