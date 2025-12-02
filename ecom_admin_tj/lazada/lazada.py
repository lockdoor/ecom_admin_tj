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
            'lazadaId': str,
            'orderNumber': str,
            'invoiceNumber': str,
            'paidPrice': np.float64,
            'unitPrice': np.float64,
            'sellerDiscountTotal': np.float64,
            'itemName': str,
            'lazadaSku': str,
        }
        self.original_df = pd.read_excel(
            self.input_file, 
            sheet_name=self.ORIGINAL_SHEET_NAME,
            dtype=dtype_dict)
        self.main_df = pd.read_excel(
            self.input_file, 
            sheet_name=self.ORIGINAL_SHEET_NAME, 
            dtype=dtype_dict, 
            usecols=columns)
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

    def calculate_finance_df(self) -> pd.DataFrame:
        """Calculate finance dataframe specific to Lazada"""
        if self.merged_df is None:
            raise ValueError("merged_df is not loaded. Please run merge_mapping() first.")
        self.finance_df = self.merged_df.groupby('orderNumber', sort=False).agg({
            'paidPrice': 'sum',
            'unitPrice': 'sum',
            'sellerDiscountTotal': 'sum',
        }).reset_index()
        
        # Add footer row with totals
        total_row = {
            'orderNumber': 'TOTAL',
            'paidPrice': self.finance_df['paidPrice'].sum(),
            'unitPrice': self.finance_df['unitPrice'].sum(),
            'sellerDiscountTotal': self.finance_df['sellerDiscountTotal'].sum(),
        }
        self.finance_df.loc[len(self.finance_df)] = total_row
        
        return self.finance_df

    def export_excel(self) -> None:
        """Export Lazada invoice to Excel file"""

        from openpyxl.worksheet.worksheet import Worksheet
        
        with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
            # Sheet 1: Original orders 
            self.original_df.to_excel(writer, sheet_name=self.ORIGINAL_SHEET_NAME, index=False)
            original_sheet: Worksheet = writer.sheets[self.ORIGINAL_SHEET_NAME]
            self._formating_header(original_sheet)
            
            # Sheet 2: invoice_{order_sn}_orders
            self.invoice_df.to_excel(writer, sheet_name=f'invoice_{self.order_sn_unique}_orders', index=False)
            invoice_sheet: Worksheet = writer.sheets[f'invoice_{self.order_sn_unique}_orders']
            invoice_sheet.column_dimensions['A'].width = 18  # stock_item_id
            invoice_sheet.column_dimensions['B'].width = 48  # stock_item_name
            invoice_sheet.column_dimensions['C'].width = 14  # จำนวนรวม
            invoice_sheet.column_dimensions['D'].width = 14  # ลูกค้าจ่าย
            invoice_sheet.column_dimensions['E'].width = 14  # ราคาสุทธิ
            invoice_sheet.column_dimensions['F'].width = 14  # ส่วนลดรวม
            self._formating_header(sheet=invoice_sheet)
            self._formatting_body(sheet=invoice_sheet, start_row=2, end_row=len(self.invoice_df), start_col=1, end_col=6)
            self._formatting_footer(sheet=invoice_sheet, footer_row=len(self.invoice_df)+1)
            
            # Canceled orders (ensure string format)
            self.canceled_orders_df.to_excel(writer, sheet_name='canceled_orders', index=False)
            self._cancel_orders_to_excel(writer)
            
            # Finance summary
            self.finance_df.to_excel(writer, sheet_name='Finance Summary', index=False)
            finance_sheet: Worksheet = writer.sheets['Finance Summary']
            finance_sheet.column_dimensions['A'].width = 24  # orderNumber
            finance_sheet.column_dimensions['B'].width = 14  # paidPrice
            finance_sheet.column_dimensions['C'].width = 14  # unitPrice
            finance_sheet.column_dimensions['D'].width = 30  # sellerDiscountTotal
            self._formating_header(finance_sheet)
            self._formatting_body(
                sheet=finance_sheet, 
                start_row=2, 
                end_row=len(self.finance_df), 
                start_col=1, 
                end_col=4)
            self._formatting_footer(sheet=finance_sheet, footer_row=len(self.finance_df)+1)
