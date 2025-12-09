import warnings
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from tqdm import tqdm

class ShopeeFinanceMixin:
    """Finance related methods for Shopee admin"""

    @classmethod
    def make_finance_report_df(cls, original_report_file: str) -> pd.DataFrame:
        """Create a cleaned finance report from the original report file"""

        # Suppress openpyxl UserWarnings about reported_file
        warnings.filterwarnings(
            "ignore", 
            category=UserWarning, 
            module='openpyxl'
        )

        report_type_dict = {
            'วันที่': str, 
            'ประเภทการทำธุรกรรม': str, 
            'คำอธิบาย': str, 
            'รหัสคำสั่งซื้อ': str,
            'รูปแบบธุรกรรม': str, 
            'จำนวนเงิน': np.float64, 
            'สถานะ': str, 
            'ยอดเงินหลังทำธุรกรรมเสร็จสิ้น': np.float64
        }
        report_df = pd.read_excel(
            original_report_file, 
            sheet_name='Transaction Report', 
            header=17, 
            dtype=report_type_dict)

        # Add column ['admin_record_file': str, 'ราคาขายสุทธิ': np.float64, 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ': np.float64, 'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ': np.float64]
        # Initialize with NaN values
        report_df['admin_record_file'] = np.nan

        return report_df

    @classmethod
    def make_finance_report(cls, original_report_file: str, output_file: str = None, auto_rename: bool = True) -> str:
        """Create a cleaned finance report from the original report file"""

        report_df = cls.make_finance_report_df(
            original_report_file)

        if output_file is None:
            output_file = 'cleaned_finance_report.xlsx'

        output_path = Path(output_file)

        # ถ้าไฟล์มีอยู่แล้วและต้องการ rename อัตโนมัติ
        if output_path.exists() and auto_rename:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            stem = output_path.stem
            suffix = output_path.suffix
            output_file = f"{stem}_{timestamp}{suffix}"
            print(f"⚠️  File exists. Saving as: {output_file}")

        # Save cleaned report to output_file
        report_df.to_excel(output_file, index=False)
        print(f"✅ Saved to: {output_file}")

        return output_file

    @classmethod
    def finance_check(cls, reported_file: str, admin_file: str, inplace=False) -> pd.DataFrame:
        """Compare reported finance file with calculated finance file
        Args:
            reported_file (str): Path to the cleaned reported finance file
            admin_file (str): Path to the admin finance file
            inplace (bool): Whether to update the reported file in place
        Returns:
            pd.DataFrame: Merged DataFrame after reconciliation
        """

        try:
            reported_df = pd.read_excel(reported_file, dtype=str)
        except ValueError as e:
            raise ValueError(f"❌ Error reading reported file '{reported_file}': {e}")

        # Visualize matched orders with progress bar
        number_of_nan_admin_record = reported_df['admin_record_file'].isna().sum()
        matched_orders = len(reported_df) - number_of_nan_admin_record
        total_orders = len(reported_df)
        
        # Determine color based on match percentage
        match_percentage = (matched_orders / total_orders * 100) if total_orders > 0 else 0
        if match_percentage >= 80:
            color = '\033[92m'  # Green
        elif match_percentage >= 50:
            color = '\033[93m'  # Yellow
        else:
            color = '\033[91m'  # Red
        reset_color = '\033[0m'
        
        with tqdm(total=total_orders, desc=f"{color}Matched Orders{reset_color}", unit="order", ncols=80, 
                  bar_format='{desc}: {percentage:3.1f}%|{bar}| {n_fmt}/{total_fmt}',
                  colour='green' if match_percentage >= 80 else 'yellow' if match_percentage >= 50 else 'red') as pbar:
            pbar.update(matched_orders)

        if admin_file is None:
            print("⚠️  No admin file provided. Exiting finance check.")
            return reported_df
        
        try:
            admin_df = pd.read_excel(admin_file, dtype=str, sheet_name='Finance Summary')
        except ValueError as e:
            raise ValueError(f"❌ Error reading admin file '{admin_file}': {e}")
        print("Number of orders in admin file:", len(admin_df))

        # Check if any order IDs from admin_file already exist in reported_df (excluding NaN records)
        already_matched = reported_df[reported_df['admin_record_file'].notna()]
        if not already_matched.empty:
            duplicate_orders = admin_df[admin_df['หมายเลขคำสั่งซื้อ'].isin(already_matched['รหัสคำสั่งซื้อ'])]
            if not duplicate_orders.empty:
                duplicate_ids = duplicate_orders['หมายเลขคำสั่งซื้อ'].tolist()
                admin_filename = Path(admin_file).name
                raise ValueError(f"❌ Found {len(duplicate_ids)} order IDs in '{admin_filename}' that were already matched: {duplicate_ids[:5]}{'...' if len(duplicate_ids) > 5 else ''}")

        # Determine which columns from admin_df should be merged
        # For first merge: all columns except key
        # For subsequent merges: only columns that don't exist yet, plus ensure data columns are updated
        reported_cols = set(reported_df.columns)
        
        # Columns that should always be synced from admin_df (data columns)
        data_columns = ['ราคาขายสุทธิ', 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ', 'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ']
        
        # Get new columns (not in reported_df) plus data columns (to update) plus key column
        admin_cols_to_merge = ['หมายเลขคำสั่งซื้อ'] + [
            col for col in admin_df.columns 
            if col not in reported_cols or col in data_columns
        ]
        
        # Remove duplicates while preserving order
        admin_cols_to_merge = list(dict.fromkeys(admin_cols_to_merge))
        
        # Select columns from admin_df
        admin_df_filtered = admin_df[admin_cols_to_merge].copy()

        # Merge with indicator to track which rows matched
        merged_df = reported_df.merge(
            admin_df_filtered, 
            left_on='รหัสคำสั่งซื้อ',
            right_on='หมายเลขคำสั่งซื้อ',
            how='left',
            indicator=True,
            suffixes=('', '_new')
        )

        admin_filename = Path(admin_file).name
        matched_count: int = merged_df[merged_df['_merge'] == 'both'].shape[0]
        print(f"✅ Matched {matched_count} orders with {admin_filename}")
        
        # อัปเดต admin_record_file สำหรับ rows ที่ merge สำเร็จ
        merged_df.loc[merged_df['_merge'] == 'both', 'admin_record_file'] = admin_filename
        
        # Update data columns for matched rows (handle _new suffix from merge)
        for col in data_columns:
            new_col = f'{col}_new'
            if new_col in merged_df.columns:
                # Update only matched rows with new values
                merged_df.loc[merged_df['_merge'] == 'both', col] = merged_df.loc[merged_df['_merge'] == 'both', new_col]
                merged_df = merged_df.drop(columns=[new_col])
        
        # ลบ column _merge และ หมายเลขคำสั่งซื้อ (duplicate)
        merged_df = merged_df.drop(columns=['_merge', 'หมายเลขคำสั่งซื้อ'])
        
        # แสดงผลสรุป
        print(f"⚠️  Remaining unmatched: {merged_df['admin_record_file'].isna().sum()}")

        if inplace:
            # บันทึกผลลัพธ์กลับไปยัง reported_file
            merged_df.to_excel(reported_file, index=False)
            print(f"✅ Updated reported file saved to: {reported_file}")
        
        return merged_df
