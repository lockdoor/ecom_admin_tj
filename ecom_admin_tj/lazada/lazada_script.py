import pandas as pd
import numpy as np
import sys
import warnings

# Suppress openpyxl warning about default style
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

MAPPING_FILE = "./ecom_admin_tj/lazada/lazada_item_mapping.xlsx"
ORIGINAL_SHEET_NAME = "sheet1"

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
    
    columns= ['orderItemId', 'orderNumber', 'invoiceNumber', 
              'paidPrice', 'unitPrice', 'sellerDiscountTotal', 'itemName', 'lazadaSku']
    dtype_dict = {
        'orderItemId': str,
        'orderNumber': str,
        'invoiceNumber': str,
        'paidPrice': np.float64,
        'unitPrice': np.float64,
        'sellerDiscountTotal': np.float64,
        'itemName': str,
        'lazadaSku': str,
    }
    df = pd.read_excel(
        file_path, 
        sheet_name=ORIGINAL_SHEET_NAME, 
        dtype=dtype_dict, 
        usecols=columns, 
        engine='openpyxl',
        )
    
    df.fillna({'sellerDiscountTotal': 0}, inplace=True)
    df['lazadaSku'] = df['lazadaSku'].map(lambda x: x.split('_')[0])

    # read canceled sheets
    try :
        canceled_df = pd.read_excel(file_path, sheet_name='canceled_orders', dtype={'canceled_orders_sn': str})
        canceled_order_sns = canceled_df['canceled_orders_sn'].dropna().unique()
        df = df[~df['orderItemId'].isin(canceled_order_sns)]
        df.reset_index(inplace=True)
    # ValueError occurs when sheet does not exist
    except (ValueError):
        print('No canceled orders sheet found. Continuing without excluding any orders.')
        pass

    return df

# merge with mapping
def merge_mapping(df, mapping_df) -> pd.DataFrame:
    df_merged = pd.merge(df, mapping_df, left_on='lazadaSku', right_on='platform_item_id', how='left')
    return df_merged

def calculate_invoice(merge_df) -> pd.DataFrame:
    invoice_df = merge_df.groupby('stock_item_id').agg({
        'stock_item_name': 'first',
        'multiplier': 'sum',
        'paidPrice': 'sum',
        'unitPrice': 'sum',
        'sellerDiscountTotal': 'sum'
    }).reset_index()
    invoice_df.loc['TOTAL'] = [
        'TOTAL',
        '', 
        '', 
        invoice_df['paidPrice'].sum(),
        invoice_df['unitPrice'].sum(),
        invoice_df['sellerDiscountTotal'].sum()
        ]
    invoice_df.columns = ['stock_item_id', 'stock_item_name', 'จำนวนรวม', 'ลูกค้าจ่าย', 'ราคาสุทธิ', 'ส่วนลดรวม']
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
    order_sn_unique = main_df['orderNumber'].nunique()
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
