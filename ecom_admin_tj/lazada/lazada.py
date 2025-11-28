from ..common.base import Base
import pandas as pd
import numpy as np
from pathlib import Path


class Lazada(Base):
    
    def __init__(self, input_file: str, output_file: str = None, shipping_date = None):
        """Initialize Lazada processor with specific settings
                
        Args:
            input_file: Path to input Excel file
            output_file: Optional custom output file path
            shipping_date: Not used in Lazada processing (kept for interface compatibility)
        """
        # Pass None for shipping_date since Lazada doesn't use it
        if shipping_date is not None:
            print('Warning: shipping_date parameter is not used in Lazada processing.')
        super().__init__(input_file, output_file, shipping_date = None)
        
        # Set Lazada-specific attributes
        self.SCRIPT_DIR = Path(__file__).parent
        self.MAPPING_FILE = self.SCRIPT_DIR / "lazada_item_mapping.xlsx"
        self.ORIGINAL_SHEET_NAME = "sheet1"
        self.merge_left = 'lazadaSku'
        self.merge_right = 'platform_item_id'
    
    def load_mapping(self) -> pd.DataFrame:
        """Load item mapping specific to Lazada"""
        mapping_file_path = self.MAPPING_FILE
        mapping_type_dict = {
            'platform_item_id': str,
            'platform_item_name': str,
            'stock_item_id': str,
            'stock_item_name': str,
            'multiplier': np.int64,
        }
        self.mapping_df = pd.read_excel(mapping_file_path, sheet_name='Item Mapping', skiprows=1, dtype=mapping_type_dict)
        self.mapping_df.dropna(subset=['platform_item_id'], inplace=True)
        return self.mapping_df

    def load_main_df(self) -> pd.DataFrame:
        """Load main data from Lazada input file"""
        
        # read original sheet
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
        self.original_df = pd.read_excel(self.input_file, sheet_name=self.ORIGINAL_SHEET_NAME)
        self.main_df = pd.read_excel(
            self.input_file, 
            sheet_name=self.ORIGINAL_SHEET_NAME, 
            dtype=dtype_dict, 
            usecols=columns, 
            engine='openpyxl',
            )
        self.main_df.fillna({'sellerDiscountTotal': 0}, inplace=True)
        self.main_df['lazadaSku'] = self.main_df['lazadaSku'].map(lambda x: x.split('_')[0])
        
        # read canceled sheets    
        self.load_canceled_orders()
        canceled_order_sns = self.canceled_orders_df['canceled_orders_sn'].dropna().unique()
        self.main_df = self.main_df[~self.main_df['orderItemId'].isin(canceled_order_sns)]
        
        # count unique order numbers
        self.order_sn_unique = self.main_df['orderNumber'].nunique()

        return self.main_df

    def calculate_invoice(self) -> pd.DataFrame:
        """Calculate invoice specific to Lazada"""
        self.invoice_df = self.merged_df.groupby('stock_item_id').agg({
            'stock_item_name': 'first',
            'multiplier': 'sum',
            'paidPrice': 'sum',
            'unitPrice': 'sum',
            'sellerDiscountTotal': 'sum'
        }).reset_index()
        self.invoice_df.loc['TOTAL'] = [
            'TOTAL',
            '', 
            '', 
            self.invoice_df['paidPrice'].sum(),
            self.invoice_df['unitPrice'].sum(),
            self.invoice_df['sellerDiscountTotal'].sum()
            ]
        self.invoice_df.columns = ['stock_item_id', 'stock_item_name', 'จำนวนรวม', 'ลูกค้าจ่าย', 'ราคาสุทธิ', 'ส่วนลดรวม']
        return self.invoice_df

    def export_excel(self) -> None:
        """Export Lazada invoice to Excel file"""
        with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
            # Sheet 1: Original orders 
            self.original_df.to_excel(writer, sheet_name=self.ORIGINAL_SHEET_NAME, index=False)
            
            # Sheet 2: invoice
            self.invoice_df.to_excel(writer, sheet_name=f'invoice_{self.order_sn_unique}_orders', index=False)
            
            # Last sheet 1: Canceled orders (ensure string format)
            self.canceled_orders_df.to_excel(writer, sheet_name='canceled_orders', index=False)
