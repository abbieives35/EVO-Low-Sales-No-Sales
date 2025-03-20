import pandas as pd
from datetime import timedelta

# Load the Excel file into a pandas DataFrame
df = pd.read_excel(r'\\ISCFS01\RedirectedFolders\aives\Downloads\LowSalesNoSalesCriteriaAbbie2Results135item number fix.xls.xlsx')

# Strip any leading/trailing whitespace from column names and values
df.columns = df.columns.str.strip()
df['Customer Name'] = df['Customer Name'].str.strip()
df['Item'] = df['Item'].astype(str).str.strip()  # Ensure 'Item' is treated as text and trimmed
df['Class'] = df['Class'].str.strip()  # Ensure 'Class' column exists and is trimmed

# Convert 'Date' to datetime format
df['Date'] = pd.to_datetime(df['Date'])  # Ensure 'Date' is in datetime format

# Step 1: Summarize relevant fields for each customer-item relationship
customer_item_summary = df.groupby(['Customer Name', 'Item', 'Class']).agg(
    MostRecentItemOrderDate=('Date', 'max'),                 # Most recent order date for the specific item
    ItemOrderCount=('Date', 'count'),                        # Total orders for this item by this customer
    TotalAmount=('Amount', 'sum'),                           # Sum of Amount for each customer-item
    TotalQuantity=('Quantity', 'sum'),                       # Sum of Quantity for each customer-item
    Description=('Description', 'first'),                    # Item description
    SpecialistAssigned=('Specialist Assigned', 'first'),     # Specialist assigned
    SalesRep=('Sales Rep', 'first')                          # Sales rep
).reset_index()

# Step 2: Calculate the typical item order pattern in days, counting only one order per day
dedup_item_orders = df.drop_duplicates(subset=['Customer Name', 'Item', 'Class', 'Date'])
dedup_item_orders['ItemOrderInterval'] = dedup_item_orders.groupby(['Customer Name', 'Item', 'Class'])['Date'].diff().dt.days
typical_item_order_cycle = dedup_item_orders.groupby(['Customer Name', 'Item', 'Class'])['ItemOrderInterval'].mean().reset_index()
typical_item_order_cycle = typical_item_order_cycle.rename(columns={'ItemOrderInterval': 'TypicalItemOrderPattern'})

# Merge the typical item order pattern back into customer_item_summary
customer_item_summary = customer_item_summary.merge(typical_item_order_cycle, on=['Customer Name', 'Item', 'Class'], how='left')

# Calculate the next expected item order date
customer_item_summary['NextExpectedItemOrderDate'] = customer_item_summary['MostRecentItemOrderDate'] + \
                                                     pd.to_timedelta(customer_item_summary['TypicalItemOrderPattern'], unit='D')

# Step 3: Calculate typical class order pattern in days, counting only one order per day for the class
dedup_class_orders = df.drop_duplicates(subset=['Customer Name', 'Class', 'Date'])
dedup_class_orders['ClassOrderInterval'] = dedup_class_orders.groupby(['Customer Name', 'Class'])['Date'].diff().dt.days
typical_class_order_cycle = dedup_class_orders.groupby(['Customer Name', 'Class'])['ClassOrderInterval'].mean().reset_index()
typical_class_order_cycle = typical_class_order_cycle.rename(columns={'ClassOrderInterval': 'TypicalClassOrderPattern'})

# Merge the typical class order pattern and most recent class order date back into customer_item_summary
class_order_summary = df.groupby(['Customer Name', 'Class']).agg(
    MostRecentClassOrderDate=('Date', 'max')                 # Most recent order date for any item in the class
).reset_index()
class_order_summary = class_order_summary.merge(typical_class_order_cycle, on=['Customer Name', 'Class'], how='left')
customer_item_summary = customer_item_summary.merge(class_order_summary, on=['Customer Name', 'Class'], how='left')

# Calculate the next expected class order date
customer_item_summary['NextExpectedClassOrderDate'] = customer_item_summary['MostRecentClassOrderDate'] + \
                                                      pd.to_timedelta(customer_item_summary['TypicalClassOrderPattern'], unit='D')

# Step 4: Determine if the customer-item relationship is "late"
current_date = pd.Timestamp.now()
customer_item_summary['IsLate'] = customer_item_summary.apply(
    lambda x: (90 <= (current_date - x['NextExpectedItemOrderDate']).days <= 365) and
              (90 <= (current_date - x['NextExpectedClassOrderDate']).days <= 365),
    axis=1
)

# Ensure 'Item' column is formatted as text
customer_item_summary['Item'] = customer_item_summary['Item'].astype(str)

# Step 5: Save the updated DataFrame to a new Excel file with date formatting
output_file_path = r'\\ISCFS01\RedirectedFolders\aives\Downloads\Customer_Item_Summary_Final.xlsx'

# Use Excel writer to apply date format
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    customer_item_summary.to_excel(writer, index=False, sheet_name='Summary')
    workbook = writer.book
    worksheet = writer.sheets['Summary']

    # Define date format
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    
    # Apply the date format to relevant date columns
    date_columns = ['MostRecentItemOrderDate', 'NextExpectedItemOrderDate', 'MostRecentClassOrderDate', 'NextExpectedClassOrderDate']
    for col in date_columns:
        col_idx = customer_item_summary.columns.get_loc(col)  # Get the column index
        worksheet.set_column(col_idx + 1, col_idx + 1, 15, date_format)  # Apply format

print(f"\nData has been successfully updated with the requested columns and saved to the new Excel file: {output_file_path}")
