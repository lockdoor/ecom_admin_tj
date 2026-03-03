import warnings
import pandas as pd
import numpy as np
import re
from pathlib import Path
from datetime import datetime
from tqdm import tqdm
from ...common.reconciliation_mixin import ReconciliationMixin, Worksheet

class ShopeeFinanceMixin(ReconciliationMixin):
    """
    Finance related methods for Shopee admin
    
    Extends ReconciliationMixin to provide Shopee-specific finance operations
    including report cleaning, reconciliation, and Excel formatting.
    """

    # Shopee-specific report type dict (inherits from parent but can override)
    report_type_dict = {
            'วันที่': str, 
            'ประเภทการทำธุรกรรม': str, 
            'คำอธิบาย': str, 
            'รหัสคำสั่งซื้อ': str,
            'รูปแบบธุรกรรม': str, 
            'จำนวนเงิน': np.float64, 
            'สถานะ': str, 
            'ยอดเงินหลังทำธุรกรรมเสร็จสิ้น': np.float64,
            'admin_record_file': 'string',
            'ราคาขายสุทธิ': np.float64,
            'ค่าจัดส่งที่ชำระโดยผู้ซื้อ': np.float64,
            'รวมชำระ': np.float64
        }

    @classmethod
    def make_finance_report_df(cls, original_report_file: str) -> pd.DataFrame:
        """
        Create a cleaned finance report from the original Shopee report file
        
        Uses parent's generic method with Shopee-specific parameters.
        """
        return super().make_finance_report_df(
            original_report_file=original_report_file,
            sheet_name='Transaction Report',
            header_row=17,
            report_columns=[
                'วันที่', 'ประเภทการทำธุรกรรม', 'รหัสคำสั่งซื้อ',
                'จำนวนเงิน', 'สถานะ', 'admin_record_file',
                'ราคาขายสุทธิ', 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ', 'รวมชำระ'
            ]
        )
    
    def _report_sheet_format_width_column(self, sheet: Worksheet):
        sheet.column_dimensions['A'].width = 20  # วันที่
        sheet.column_dimensions['B'].width = 20  # ประเภทการทำธุรกรรม
        sheet.column_dimensions['C'].width = 20  # รหัสคำสั่งซื้อ
        sheet.column_dimensions['D'].width = 12  # จำนวนเงิน
        sheet.column_dimensions['E'].width = 15  # สถานะ
        sheet.column_dimensions['F'].width = 30  # admin_record_file
        sheet.column_dimensions['G'].width = 12  # ราคาขายสุทธิ
        sheet.column_dimensions['H'].width = 12  # ค่าจัดส่งที่ชำระโดยผู้ซื้อ
        sheet.column_dimensions['I'].width = 12  # รวมชำระ
    
    @classmethod
    def make_finance_report(cls, original_report_file: str, output_file: str = None, auto_rename: bool = True) -> str:
        """
        Create a cleaned finance report from the original Shopee report file
        
        Args:
            original_report_file: Path to original Shopee finance report
            output_file: Optional custom output filename
            auto_rename: Auto-rename if file exists
            
        Returns:
            Path to the created output file
        """
        report_df = cls.make_finance_report_df(original_report_file)

        if output_file is None:
            # Extract date range from input filename (e.g., 20260112_20260118)
            input_filename = Path(original_report_file).stem
            date_match = re.search(r'(\d{8}_\d{8})', input_filename)
            
            if date_match:
                date_range = date_match.group(1)
                output_file = f'shopee_cleaned_finance_report_{date_range}.xlsx'
            else:
                output_file = 'shopee_cleaned_finance_report.xlsx'

        output_path = Path(output_file)

        # Auto-rename if file exists
        if output_path.exists() and auto_rename:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            stem = output_path.stem
            suffix = output_path.suffix
            output_file = f"{stem}_{timestamp}{suffix}"
            print(f"⚠️  File exists. Saving as: {output_file}")

        # Save cleaned report to output_file
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            report_df.to_excel(excel_writer=writer, sheet_name='Transaction Report', index=False)
            report_sheet = writer.sheets['Transaction Report']
            cls()._report_sheet_format_width_column(sheet=report_sheet)
            cls()._formating_header(sheet=report_sheet)
            print(f"✅ Saved to: {output_file}")

        return output_file

    # Note: admin_check, draw_progress_bar, and finance_check methods
    # are now inherited from ReconciliationMixin parent class.
    # They can be overridden here if Shopee-specific behavior is needed.
