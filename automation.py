import pandas as pd
import random
from datetime import datetime, timedelta


products = ['Apple', 'Banana', 'Orange', 'Mango', 'Grapes']
num_rows = 100  
start_date = datetime(2026, 1, 1)

data = []

for i in range(num_rows):
    date = start_date + timedelta(days=i)
    product = random.choice(products)
    quantity = random.randint(1, 20)   
    price_dict = {'Apple': 2, 'Banana': 1.5, 'Orange': 3, 'Mango': 2.5, 'Grapes': 4}
    price = price_dict[product]
    data.append([date.strftime('%Y-%m-%d'), product, quantity, price])

# --- Step 2: Create DataFrame ---
df = pd.DataFrame(data, columns=['Date', 'Product', 'Quantity', 'Price'])

# --- Step 3: Add Total column ---
df['Total'] = df['Quantity'] * df['Price']

# --- Step 4: Summary by product ---
summary = df.groupby('Product')['Total'].sum().reset_index()

# --- Step 5: Save to Excel ---
with pd.ExcelWriter('sales_data.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Sales Data')
    summary.to_excel(writer, index=False, sheet_name='Summary')

print(" Excel file 'sales_data.xlsx' created successfully!")
