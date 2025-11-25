import pandas as pd
import numpy as np
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).parent
MAPPING_FILE = SCRIPT_DIR / 'tiktok_item_mapping.xlsx'
ORIGINAL_SHEET_NAME = 'OrderSKUList'

def load_mapping() -> pd.DataFrame:
    mapping_file_path = MAPPING_FILE
    mapping_type_dict = {
        'platform_item_id': str,
        'platform_item_name': str,
        'stock_item_id': str,
        'stock_item_name': str,
        'multiplier': np.int64,
    }
    mapping_df = pd.read_excel(mapping_file_path, sheet_name='Item Mapping', skiprows=1, dtype=mapping_type_dict)
    mapping_df.dropna(subset=['platform_item_id'], inplace=True)
    return mapping_df

def load_main_df(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name='OrderSKUList')
    
    df.drop(df.index[0], inplace=True)
    df.reset_index(inplace=True)
    columns= ['Order ID', 'SKU ID', 'Product Name', 'Quantity', 'SKU Unit Original Price', 'SKU Subtotal Before Discount', 'SKU Seller Discount', 'SKU Subtotal After Discount']
    df = df[columns]
    
    df['Order ID'] = df['Order ID'].astype(str)
    df['Quantity'] = df['Quantity'].astype(np.int64)
    df['SKU Unit Original Price'] = df['SKU Unit Original Price'].astype(np.float64)
    df['SKU Subtotal Before Discount'] = df['SKU Subtotal Before Discount'].astype(np.float64)
    df['SKU Seller Discount'] = df['SKU Seller Discount'].astype(np.float64)
    df['SKU Subtotal After Discount'] = df['SKU Subtotal After Discount'].astype(np.float64)

    # read canceled sheets
    try :
        canceled_df = pd.read_excel(file_path, sheet_name='canceled_orders', dtype={'canceled_orders_sn': str})
        canceled_order_sns = canceled_df['canceled_orders_sn'].dropna().unique()
        df = df[~df['Order ID'].isin(canceled_order_sns)]
        df.reset_index(inplace=True)
    # ValueError occurs when sheet does not exist
    except (ValueError):
        print('No canceled orders sheet found. Continuing without excluding any orders.')
        pass

    return df

# merge with mapping
def merge_mapping(df, mapping_df) -> pd.DataFrame:
    df_merged = pd.merge(df, mapping_df, left_on='SKU ID', right_on='platform_item_id', how='left')
    df_merged['จำนวนรวม'] = df_merged['Quantity'] * df_merged['multiplier']
    return df_merged

# calculate invoice 
def calculate_invoice(merge_df) -> pd.DataFrame:
    invoice_df = merge_df.groupby('stock_item_id').agg({
        'stock_item_name': 'first',
        'จำนวนรวม': 'sum',
        'SKU Subtotal Before Discount': 'sum',
        'SKU Seller Discount': 'sum'
    }).reset_index()
    invoice_df.loc['TOTAL'] = [
        'TOTAL',
        '', 
        '', 
        invoice_df['SKU Subtotal Before Discount'].sum(), 
        invoice_df['SKU Seller Discount'].sum()]
    return invoice_df

def export_to_excel(invoice_df: pd.DataFrame, n_unique_orders, output_file: str, input_file: str) -> None:
    # Read file for get Canceled orders sheet if exists
    # canceled_df = None
    # if 'canceled_orders' in pd.ExcelFile(output_file).sheet_names:
    try:
        canceled_df = pd.read_excel(input_file, sheet_name='canceled_orders', dtype={'canceled_orders_sn': str})
    except (ValueError):
        canceled_df = pd.DataFrame(columns=['canceled_orders_sn'])
    
    original_df = pd.read_excel(input_file, sheet_name=ORIGINAL_SHEET_NAME)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet 1: Original orders 
        original_df.to_excel(writer, sheet_name=ORIGINAL_SHEET_NAME, index=False)
        
        # Sheet 2: invoice
        invoice_df.to_excel(writer, sheet_name=f'invoice_{n_unique_orders}_orders', index=False)
        
        # Last sheet 1: Canceled orders (ensure string format)
        canceled_df.to_excel(writer, sheet_name='canceled_orders', index=False)

def main():
    print("Starting TikTok item mapping script...")
    # Read file from argument
    if len(sys.argv) < 2:
        print('Usage: python ecom_admin_tj.shopee <shopee20251118_sample.xlsx>')
        sys.exit(1)
    input_file: str = sys.argv[1]
    print(f'Reading input file: {input_file}')
    
    
    mapping_df = load_mapping()
    main_df = load_main_df(input_file)
    order_sn_unique = main_df['Order ID'].nunique()
    print(f'Total unique orders: {order_sn_unique}')
    merged_df = merge_mapping(main_df, mapping_df)
    invoice_df = calculate_invoice(merged_df)
    
    # Export to Excel
    if sys.argv[1].endswith('_output.xlsx'):
        output_file = input_file
    else:
        output_file = input_file.replace('.xlsx', '_output.xlsx')
    print(f'Exporting to Excel file: {output_file}')
    export_to_excel(invoice_df, order_sn_unique, output_file, input_file)

if __name__ == "__main__":
    main()
