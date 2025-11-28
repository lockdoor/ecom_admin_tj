from ..common.base import Base
import pandas as pd
import numpy as np
from pathlib import Path

class Shopee(Base):
    
    SHIPPING_FEE_ITEM_ID = '00-0000-00'
    TOTAL = 'TOTAL'
    invoice_group_dict: dict[str, pd.DataFrame] = {}
    deduct_stock_df: pd.DataFrame | None = None
    
    def __init__(self, input_file: str, output_file: str = None, shipping_date = None):
        """Initialize Shopee processor with specific settings"""
        super().__init__(input_file, output_file, shipping_date)
        
        # Set Shopee-specific attributes
        self.SCRIPT_DIR = Path(__file__).parent
        self.MAPPING_FILE = self.SCRIPT_DIR / "shopee_item_mapping.xlsx"
        self.ORIGINAL_SHEET_NAME = "orders"
        self.merge_left = 'เลขอ้างอิง Parent SKU'
        self.merge_right = 'platform_sku'
    
    def load_mapping(self) -> pd.DataFrame:
        """Load item mapping specific to Shopee"""
        mapping_type_dict = {
                'platform_item_id': str,
                'platform_sku': str,
                'platform_item_name': str,
                'stock_item_id': str,
                'stock_item_name': str,
                'multiplier': np.int64,
                'ratio': np.float64,
            }
        self.mapping_df = pd.read_excel(
            self.MAPPING_FILE, sheet_name='Item Mapping', 
            skiprows=1, dtype=mapping_type_dict)
        return self.mapping_df
    
    def merge_mapping(self) -> pd.DataFrame:
        """Merge main dataframe with Shopee mapping"""
        super().merge_mapping()
        self.merged_df['จำนวนรวม'] = self.merged_df['จำนวน'] * self.merged_df['multiplier']
        return self.merged_df

    def load_main_df(self) -> pd.DataFrame:
        """Load main data from Shopee input file"""
        
        # Required columns
        required_cols = ['หมายเลขคำสั่งซื้อ', 'เลขอ้างอิง Parent SKU',  'ชื่อสินค้า', 
                        'ราคาตั้งต้น', 'ราคาขาย', 'จำนวน', 'ราคาขายสุทธิ', 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ', 
                        'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ', 'ผู้ซื้อร้องขอใบกำกับภาษี', 'วันที่คาดว่าจะทำการจัดส่งสินค้า']
        
        # Try to read with cancellation reason column, if not exists, read without it
        if self.original_df is None:
            try:
                self.original_df = pd.read_excel(
                    self.input_file, sheet_name=self.ORIGINAL_SHEET_NAME, 
                    usecols=required_cols + ['เหตุผลในการยกเลิกคำสั่งซื้อ'])
                has_cancel_reason = True
            except ValueError:
                # Column doesn't exist, read without it
                self.original_df = pd.read_excel(
                    self.input_file, sheet_name=self.ORIGINAL_SHEET_NAME, 
                    usecols=required_cols)
                has_cancel_reason = False
        
        self.main_df = self.original_df.dropna(subset=['หมายเลขคำสั่งซื้อ']).copy()
        self.main_df['ราคาขายสุทธิ'] = self.original_df['ราคาขายสุทธิ'].astype(np.float64)
        self.main_df['วันที่คาดว่าจะทำการจัดส่งสินค้า'] = pd.to_datetime(self.original_df['วันที่คาดว่าจะทำการจัดส่งสินค้า'], errors='coerce')

        # today is first row in df
        if self.shipping_date is not None:
            today = self.shipping_date
        else:
            today = self.main_df['วันที่คาดว่าจะทำการจัดส่งสินค้า'].iloc[0]
        # check only date equal (ignore time part)
        self.main_df = self.main_df[self.main_df['วันที่คาดว่าจะทำการจัดส่งสินค้า'].dt.date == today.date()]
        
        # Filter out canceled orders based on cancellation reason (only if column exists)
        if has_cancel_reason:
            self.main_df = self.main_df[self.main_df['เหตุผลในการยกเลิกคำสั่งซื้อ'].isna()]
        
        # Load canceled orders from separate sheet and exclude them
        self.load_canceled_orders()
        canceled_order_sns = self.canceled_orders_df['canceled_orders_sn'].dropna().unique()
        self.main_df = self.main_df[~self.main_df['หมายเลขคำสั่งซื้อ'].isin(canceled_order_sns)]
        
        # count unique order numbers
        self.order_sn_unique = self.main_df['หมายเลขคำสั่งซื้อ'].nunique()
        
        return self.main_df
    
    #TODO
    def calculate_invoice(self, merge_df: pd.DataFrame, buyer_shipping_fee: float=0.0) -> pd.DataFrame:
        '''Use calculate_invoice to generate invoice dataframe from order dataframe
        Before using this function dataframe must be merged with mapping dataframe
        
        Args:
            merge_df (pd.DataFrame): Merged dataframe with mapping information
            buyer_shipping_fee (float): Shipping fee paid by buyer to be added to invoice
        '''
        def split_with_ratio(df) -> tuple[pd.DataFrame, pd.DataFrame]:
            ratio_1_df = df[df['ratio'] == 1]
            ratio_not_1_df = df[df['ratio'] != 1]
            return ratio_1_df, ratio_not_1_df
        
        if merge_df is None:
            raise ValueError("merged_df is None. Please merge mapping before calculating invoice.")
        
        ratio_1_df, ratio_not_1_df = split_with_ratio(merge_df)
        invoice_df: pd.DataFrame = ratio_1_df.groupby('stock_item_id').agg({
            'stock_item_name': 'first', 
            'จำนวนรวม': 'sum', 
            'ราคาขายสุทธิ': 'sum', 
            })
    
        for _, row in ratio_not_1_df.iterrows():
            stock_item_id = row['stock_item_id']
            stock_item_name = row['stock_item_name']
            quantity = row['จำนวนรวม']
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
        # debug print
        # print(f'Processed stock_item_id: {stock_item_name}, ratio: {ratio}, quantity: {quantity}, adj_total_price: {adj_total_price}')
    
        # Add buyer shipping fee row
        invoice_df.loc[self.SHIPPING_FEE_ITEM_ID] = ['ค่าจัดส่งที่ชำระโดยผู้ซื้อ', 1, buyer_shipping_fee]
        
        # Add total row
        invoice_df.loc[self.TOTAL] = ['รวมทั้งหมด', 1, invoice_df['ราคาขายสุทธิ'].sum()]
        
        return invoice_df

    def export_excel(self) -> None:
        """Export original orders and invoices to Excel with multiple sheets
        
        Args:
            invoice_dict (dict[str, pd.DataFrame]): Dictionary of invoice dataframes keyed by group
            deduct_stock_df (pd.DataFrame): Dataframe summarizing stock deductions
        """

        with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
            # Sheet 1: Original orders 
            self.original_df.to_excel(writer, sheet_name='orders', index=False)
            
            # Sheet 2: To day orders
            self.main_df.to_excel(writer, sheet_name='to_day_orders', index=False)
            
            # Sheet 3+: Each invoice
            for group_key, invoice_df in self.invoice_group_dict.items():
                # Sanitize sheet name (Excel has max 31 chars and no special chars)
                sheet_name = str(group_key).replace('/', '_')[:31]
                invoice_df.to_excel(writer, sheet_name=sheet_name, index=True)
            
            # Last sheet 1: Stock deduction summary
            self.deduct_stock_df.to_excel(writer, sheet_name='Stock Deduction', index=True)
            
            # Last sheet 2: Canceled orders
            self.canceled_orders_df.to_excel(writer, sheet_name='canceled_orders', index=False)
    
    def calculate_group_invoice(self) -> None:
        '''Group by No VAT requested and VAT requested orders
        Then calculate invoices for each group and store in invoice_group_dict
        '''
        # Group by No VAT requested
        no_vat_order_df: pd.DataFrame = self.merged_df[self.merged_df['ผู้ซื้อร้องขอใบกำกับภาษี'] == 'No']
        number_of_no_vat_orders: int = no_vat_order_df['หมายเลขคำสั่งซื้อ'].nunique()
        self.invoice_group_dict[f'no_vat_{number_of_no_vat_orders}_orders'] = no_vat_order_df
        print(f'Number of No VAT requested orders: {number_of_no_vat_orders}')
        # Group by VAT requested
        df_vat: pd.DataFrame = self.merged_df[self.merged_df['ผู้ซื้อร้องขอใบกำกับภาษี'] == 'Yes']
        for order_sn in df_vat['หมายเลขคำสั่งซื้อ'].unique():
            self.invoice_group_dict[order_sn] = df_vat[df_vat['หมายเลขคำสั่งซื้อ'] == order_sn].copy()
        # Calculate invoices
        for group_key, group_df in self.invoice_group_dict.items():
            print(f'Processing group: {group_key}')
            buyer_shipping_fee: float = group_df['ค่าจัดส่งที่ชำระโดยผู้ซื้อ'].sum()
            order_invoice_df = self.calculate_invoice(group_df, buyer_shipping_fee)
            self.invoice_group_dict[group_key] = order_invoice_df
            
    def calculate_total_deduct_stock(self) -> pd.DataFrame:
        """
        Calculate amount of items to deduct from stock
        """
        self.deduct_stock_df = pd.DataFrame(columns=['stock_item_name', 'quantity'])
        self.deduct_stock_df.index.name = 'stock_item_id'
        for _, invoice_df in self.invoice_group_dict.items():
            for stock_item_id, row in invoice_df.iterrows():
                # Skip shipping and total rows
                if stock_item_id in [self.SHIPPING_FEE_ITEM_ID, self.TOTAL]:
                    continue
                
                if stock_item_id in self.deduct_stock_df.index:
                    self.deduct_stock_df.at[stock_item_id, 'quantity'] += row['จำนวนรวม']
                else:
                    self.deduct_stock_df.loc[stock_item_id] = {
                        'stock_item_name': row['stock_item_name'],
                        'quantity': row['จำนวนรวม']
                    }
       
    def process(self) -> None:
        '''Main function to process Shopee orders and generate invoices
        1. Read Shopee today orders from Excel
        2. Load item mapping
        3. Merge orders with mapping
        4. Group orders by VAT request and calculate invoices
        5. Calculate stock deduction summary
        6. Export all data to Excel
        '''

        print(f"Starting {self.__class__.__name__}...")
        print(f'Reading input file: {self.input_file}')
        print(f'Processing date: {self.shipping_date.strftime("%Y-%m-%d") if self.shipping_date else "Not specified"}')

        # Load data
        self.mapping_df = self.load_mapping()
        self.main_df = self.load_main_df()

        # Process
        self.merged_df = self.merge_mapping()
        self.calculate_group_invoice()
        self.calculate_total_deduct_stock()
        
        # Export
        print(f'Exporting to Excel file: {self.output_file}')
        self.export_excel()
        
        print("Process completed successfully!")
