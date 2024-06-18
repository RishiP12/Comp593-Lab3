import sys
import os
import pandas as pd
from datetime import datetime

def validate_arguments():
    if len(sys.argv) != 2:
        print("Error: Please provide the path to the sales data CSV file as a command line argument.")
        sys.exit(1)
    
    csv_path = sys.argv[1]
    if not os.path.isfile(csv_path):
        print(f"Error: The file '{csv_path}' does not exist.")
        sys.exit(1)
    
    return csv_path

def create_orders_directory(csv_path):
    today = datetime.today().strftime('%Y-%m-%d')
    orders_dir = os.path.join(os.path.dirname(csv_path), f"Orders_{today}")
    if not os.path.exists(orders_dir):
        os.makedirs(orders_dir)
    return orders_dir

def process_sales_data(csv_path, orders_dir):
    
    sales_data = pd.read_csv(csv_path)

    
    print("Column names in the CSV file:", sales_data.columns.tolist())

    
    sales_data['TOTAL PRICE'] = sales_data['ITEM QUANTITY'] * sales_data['ITEM PRICE']

    
    grouped = sales_data.groupby('ORDER ID')
    
    for order_id, group in grouped:
        
        group = group.sort_values('ITEM NUMBER')
        
        
        order_file = os.path.join(orders_dir, f'Order_{order_id}.xlsx')
        with pd.ExcelWriter(order_file, engine='xlsxwriter') as writer:
            
            print(f"Processing order {order_id} with columns: {group.columns.tolist()}")
            try:
                group.to_excel(writer, sheet_name='Order', index=False, columns=[
                    'ORDER ID', 'ORDER DATE', 'ITEM NUMBER', 'PRODUCT LINE', 'PRODUCT CODE', 'ITEM QUANTITY', 'ITEM PRICE', 'TOTAL PRICE'
                ])
            except KeyError as e:
                print(f"Error: {e}")
                print("Please check the column names in the CSV file.")
                sys.exit(1)

            
            workbook = writer.book
            worksheet = writer.sheets['Order']

            
            money_format = workbook.add_format({'num_format': '$#,##0.00'})

            
            worksheet.set_column('H:H', None, money_format)

            
            grand_total = group['TOTAL PRICE'].sum()
            worksheet.write(len(group) + 1, 6, 'GRAND TOTAL')
            worksheet.write(len(group) + 1, 7, grand_total, money_format)

            
            worksheet.set_column('A:A', 10)
            worksheet.set_column('B:B', 12)
            worksheet.set_column('C:C', 25)
            worksheet.set_column('D:D', 14)
            worksheet.set_column('E:E', 12)
            worksheet.set_column('F:F', 12)
            worksheet.set_column('G:G', 12)
            worksheet.set_column('H:H', 12)

if __name__ == '__main__':
    csv_path = validate_arguments()
    orders_dir = create_orders_directory(csv_path)
    process_sales_data(csv_path, orders_dir)
    print(f"Order files have been successfully created in the directory: {orders_dir}")
