import numpy as np
import pandas as pd
import sys
from pathlib import Path

# Get the directory where this script is located
SCRIPT_DIR = Path(__file__).parent
MAPPING_FILE = SCRIPT_DIR / 'shopee_item_mapping.xlsx'
SHIPPING_FEE_ITEM_ID = '00-0000-00'
TOTAL = 'TOTAL'

# setup mapping_df
def load_mapping_df() -> pd.DataFrame:
    mapping_type_dict = {
        'platform_sku': str,
        'platform_item_name': str,
        'stock_item_id': str,
        'stock_item_name': str,
        'multiplier': np.int64,
        'ratio': np.float64,
    }
    mapping_df = pd.read_excel(MAPPING_FILE, sheet_name='Item Mapping', skiprows=1, dtype=mapping_type_dict)
    return mapping_df

def merge_mapping(df, mapping_df) -> pd.DataFrame:
    df_merged = pd.merge(df, mapping_df, left_on='เลขอ้างอิง Parent SKU', right_on='platform_sku', how='left')
    df_merged['จำนวนรวม'] = df_merged['จำนวน'] * df_merged['multiplier']
    return df_merged

def split_with_ratio(df) -> tuple[pd.DataFrame, pd.DataFrame]:
    ratio_1_df = df[df['ratio'] == 1]
    ratio_not_1_df = df[df['ratio'] != 1]
    return ratio_1_df, ratio_not_1_df

def calculate_invoice(merge_df, buyer_shipping_fee) -> pd.DataFrame:
    '''Use calculate_invoice to generate invoice dataframe from order dataframe
    Before using this function dataframe must be merged with mapping dataframe
    
    Args:
        merge_df (pd.DataFrame): Merged dataframe with mapping information
        buyer_shipping_fee (float): Shipping fee paid by buyer to be added to invoice
    '''
    ratio_1_df, ratio_not_1_df = split_with_ratio(merge_df)
    invoice_df: pd.DataFrame = ratio_1_df.groupby('stock_item_id').agg({
        'stock_item_name': 'first', 
        'จำนวนรวม': 'sum', 
        'ราคาขายสุทธิ': 'sum', 
        })
    
    for _, row in ratio_not_1_df.iterrows():
        stock_item_id = row['stock_item_id']
        stock_item_name = row['stock_item_name']
        quantity = row['multiplier']
        total_price = row['ราคาขายสุทธิ']
        ratio = row['ratio']
        
        adj_total_price = total_price * ratio
        
        if stock_item_id in invoice_df.index:
            # print(f'Processing stock_item_id: {stock_item_name}, ratio: {ratio}, quantity: {quantity}, adj_total_price: {adj_total_price}')
            invoice_df.at[stock_item_id, 'จำนวนรวม'] += quantity
            invoice_df.at[stock_item_id, 'ราคาขายสุทธิ'] += adj_total_price
        else:
            # print(f'stock_item_id: {stock_item_id}, ratio: {ratio}, quantity: {quantity}, adj_total_price: {adj_total_price}')
            invoice_df.loc[stock_item_id] = [stock_item_name, quantity, adj_total_price]
    
    # Add buyer shipping fee row
    invoice_df.loc[SHIPPING_FEE_ITEM_ID] = ['ค่าจัดส่งที่ชำระโดยผู้ซื้อ', 1, buyer_shipping_fee]
    
    # Add total row
    invoice_df.loc[TOTAL] = ['รวมทั้งหมด', 1, invoice_df['ราคาขายสุทธิ'].sum()]
    
    return invoice_df

def load_main_df(input_file: str) -> pd.DataFrame:
    cols = ['หมายเลขคำสั่งซื้อ', 'เลขอ้างอิง Parent SKU',  'ชื่อสินค้า', 
            'ราคาตั้งต้น', 'ราคาขาย', 'จำนวน', 'ราคาขายสุทธิ', 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ', 
            'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ', 'ผู้ซื้อร้องขอใบกำกับภาษี', 'วันที่คาดว่าจะทำการจัดส่งสินค้า']
    ori_df = pd.read_excel(input_file, sheet_name='orders', usecols=cols)
    ori_df.dropna(subset=['หมายเลขคำสั่งซื้อ'], inplace=True)
    ori_df['ราคาขายสุทธิ'] = ori_df['ราคาขายสุทธิ'].astype(np.float64)
    ori_df['วันที่คาดว่าจะทำการจัดส่งสินค้า'] = pd.to_datetime(ori_df['วันที่คาดว่าจะทำการจัดส่งสินค้า'], errors='coerce')
    # today is first row in df
    today = ori_df['วันที่คาดว่าจะทำการจัดส่งสินค้า'].iloc[0]
    today_df = ori_df[ori_df['วันที่คาดว่าจะทำการจัดส่งสินค้า'] == today]
    # read canceled sheets
    try :
        canceled_df = pd.read_excel(input_file, sheet_name='canceled_orders')
        canceled_order_sns = canceled_df['canceled_orders_sn'].dropna().unique()
        today_df = today_df[~today_df['หมายเลขคำสั่งซื้อ'].isin(canceled_order_sns)]
    # ValueError occurs when sheet does not exist
    except (FileNotFoundError, ValueError):
        pass
    return today_df

def export_to_excel(to_day_orders_df: pd.DataFrame, invoice_dict: dict[str, pd.DataFrame], 
                    deduct_stock_df: pd.DataFrame, output_file: str, input_file: str) -> None:
    """Export original orders and invoices to Excel with multiple sheets
    
    Args:
        original_df: Original orders dataframe
        invoice_dict: Dictionary of invoice dataframes {group_key: invoice_df}
        output_file: Output Excel filename
    """
    # Read file for get Canceled orders sheet if exists
    # canceled_df = None
    # if 'canceled_orders' in pd.ExcelFile(output_file).sheet_names:
    try:
        canceled_df = pd.read_excel(input_file, sheet_name='canceled_orders')
    except (FileNotFoundError, ValueError):
        canceled_df = pd.DataFrame(columns=['canceled_orders_sn'])
    
    original_df = pd.read_excel(input_file, sheet_name='orders')

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet 1: Original orders 
        original_df.to_excel(writer, sheet_name='orders', index=False)
        
        # Sheet 2: To day orders
        to_day_orders_df.to_excel(writer, sheet_name='to_day_orders', index=False)
        
        # Sheet 3+: Each invoice
        for group_key, invoice_df in invoice_dict.items():
            # Sanitize sheet name (Excel has max 31 chars and no special chars)
            sheet_name = str(group_key).replace('/', '_')[:31]
            invoice_df.to_excel(writer, sheet_name=sheet_name, index=True)
        
        # Last sheet 1: Stock deduction summary
        deduct_stock_df.to_excel(writer, sheet_name='Stock Deduction', index=True)
        
        # Last sheet 2: Canceled orders
        canceled_df.to_excel(writer, sheet_name='canceled_orders', index=False)

def main() -> None:
    '''Main function to process Shopee orders and generate invoices
    1. Read Shopee today orders from Excel
    2. Load item mapping
    3. Merge orders with mapping
    4. Group orders by VAT request and calculate invoices
    5. Calculate stock deduction summary
    6. Export all data to Excel
    '''
    # Read file from argument
    if len(sys.argv) < 2:
        print('Usage: python lab_script.py <shopee_orders_sample.xlsx>')
        sys.exit(1)
    input_file: str = sys.argv[1]
    print(f'Reading input file: {input_file}')
    to_day_orders_df: pd.DataFrame = load_main_df(input_file)
    
    # Load mapping dataframe
    mapping_df: pd.DataFrame = load_mapping_df()
    
    # Merge with mapping
    merged_df: pd.DataFrame = merge_mapping(to_day_orders_df, mapping_df)
    
    # Create invoice dict
    invoice_group_dict: dict[str, pd.DataFrame] = {}
    # Group by No VAT requested
    no_vat_order_df: pd.DataFrame = merged_df[merged_df['ผู้ซื้อร้องขอใบกำกับภาษี'] == 'No']
    number_of_no_vat_orders: int = no_vat_order_df['หมายเลขคำสั่งซื้อ'].nunique()
    invoice_group_dict[f'no_vat_{number_of_no_vat_orders}_orders'] = no_vat_order_df
    print(f'Number of No VAT requested orders: {number_of_no_vat_orders}')
    # Group by VAT requested
    df_vat: pd.DataFrame = merged_df[merged_df['ผู้ซื้อร้องขอใบกำกับภาษี'] == 'Yes']
    for order_sn in df_vat['หมายเลขคำสั่งซื้อ'].unique():
        invoice_group_dict[order_sn] = df_vat[df_vat['หมายเลขคำสั่งซื้อ'] == order_sn].copy()
    # Calculate invoices
    for group_key, group_df in invoice_group_dict.items():
        print(f'\nProcessing group: {group_key}')
        buyer_shipping_fee: float = group_df['ค่าจัดส่งที่ชำระโดยผู้ซื้อ'].sum()
        order_invoice_df = calculate_invoice(group_df, buyer_shipping_fee)
        invoice_group_dict[group_key] = order_invoice_df
        
    # Calculate amount of items to deduct from stock
    deduct_stock_df = pd.DataFrame(columns=['stock_item_name', 'quantity'])
    deduct_stock_df.index.name = 'stock_item_id'
    for group_key, invoice_df in invoice_group_dict.items():
        for stock_item_id, row in invoice_df.iterrows():
            # Skip shipping and total rows
            if stock_item_id in [SHIPPING_FEE_ITEM_ID, TOTAL]:
                continue
            
            if stock_item_id in deduct_stock_df.index:
                deduct_stock_df.at[stock_item_id, 'quantity'] += row['จำนวนรวม']
            else:
                deduct_stock_df.loc[stock_item_id] = {
                    'stock_item_name': row['stock_item_name'],
                    'quantity': row['จำนวนรวม']
                }
        
    # Export to Excel
    if sys.argv[1].endswith('_output.xlsx'):
        output_file = sys.argv[1]
    else:
        output_file = sys.argv[1].replace('.xlsx', '_output.xlsx')
    print(f'Exporting to Excel file: {output_file}')
    export_to_excel(to_day_orders_df, invoice_group_dict, deduct_stock_df, output_file, sys.argv[1])
    
if __name__ == '__main__':
    main()
