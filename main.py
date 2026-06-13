# Simulate a dataset that might be found in an Excel spreadsheet
sales_data = [
    {"product": "Laptop", "category": "Electronics", "quantity": 1, "price": 1200, "region": "North"},
    {"product": "Mouse", "category": "Electronics", "quantity": 5, "price": 25, "region": "North"},
    {"product": "Keyboard", "category": "Electronics", "quantity": 3, "price": 75, "region": "South"},
    {"product": "Desk Chair", "category": "Furniture", "quantity": 2, "price": 150, "region": "West"},
    {"product": "Monitor", "category": "Electronics", "quantity": 1, "price": 300, "region": "East"},
    {"product": "Coffee Table", "category": "Furniture", "quantity": 1, "price": 100, "region": "North"},
    {"product": "Webcam", "category": "Electronics", "quantity": 2, "price": 50, "region": "South"},
    {"product": "Bookshelf", "category": "Furniture", "quantity": 1, "price": 80, "region": "West"},
    {"product": "External SSD", "category": "Electronics", "quantity": 2, "price": 90, "region": "East"},
    {"product": "Gaming PC", "category": "Electronics", "quantity": 1, "price": 1500, "region": "North"},
]

print("--- Simulated Sales Data Analysis (Excel-like) ---")
print(f"Total records: {len(sales_data)}\n")

# 1. Calculate Total Revenue (Equivalent to SUM function in Excel)
total_revenue = sum(item["quantity"] * item["price"] for item in sales_data)
print(f"1. Total Revenue: ${total_revenue:,.2f}")

# 2. Sales by Category (Equivalent to SUMIF or Pivot Table in Excel)
sales_by_category = {}
for item in sales_data:
    category = item["category"]
    revenue = item["quantity"] * item["price"]
    sales_by_category[category] = sales_by_category.get(category, 0) + revenue

print("\n2. Revenue by Category:")
for category, revenue in sales_by_category.items():
    print(f"   - {category}: ${revenue:,.2f}")

# 3. Find Top Selling Product (Based on single transaction revenue, similar to MAX/VLOOKUP in Excel)
top_product = None
max_revenue_item = 0
for item in sales_data:
    item_revenue = item["quantity"] * item["price"]
    if item_revenue > max_revenue_item:
        max_revenue_item = item_revenue
        top_product = item["product"]

print(f"\n3. Highest Revenue Single Item Sale: {top_product} (${max_revenue_item:,.2f})")

# 4. Filter Data (Equivalent to FILTER function or applying filters in Excel)
electronics_sales = [item for item in sales_data if item["category"] == "Electronics"]
total_electronics_revenue = sum(item["quantity"] * item["price"] for item in electronics_sales)
print(f"\n4. Total Revenue for 'Electronics' Category: ${total_electronics_revenue:,.2f}")

# 5. Average Transaction Value (Equivalent to AVERAGE function in Excel)
if sales_data:
    average_transaction_value = total_revenue / len(sales_data)
    print(f"\n5. Average Transaction Value: ${average_transaction_value:,.2f}")
else:
    print("\n5. No data to calculate average transaction value.")

# 6. Count of products sold by region (Equivalent to COUNTIF or Pivot Table in Excel)
products_by_region = {}
for item in sales_data:
    region = item["region"]
    products_by_region[region] = products_by_region.get(region, 0) + item["quantity"]

print("\n6. Total Products Sold by Region:")
for region, count in products_by_region.items():
    print(f"   - {region}: {count} units")
