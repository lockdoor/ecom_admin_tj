import warnings
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from tqdm import tqdm
from ...common.excel_format_mixin import ExcelFormatMixin, Worksheet

class ShopeeFinanceMixin(ExcelFormatMixin):
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
        report_df['admin_record_file'] = pd.NA
        report_df['ราคาขายสุทธิ'] = np.nan
        report_df['ค่าจัดส่งที่ชำระโดยผู้ซื้อ'] = np.nan
        report_df['ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ'] = np.nan
        # Set up dtypes
        report_df = report_df.astype({
            'admin_record_file': 'string',
            'ราคาขายสุทธิ': 'float64',
            'ค่าจัดส่งที่ชำระโดยผู้ซื้อ': 'float64',
            'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ': 'float64'
        })

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
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            report_df.to_excel(excel_writer=writer, sheet_name='Transaction Report', index=False)
            report_sheet = writer.sheets['Transaction Report']
            report_sheet.column_dimensions['A'].width = 20  # วันที่
            report_sheet.column_dimensions['B'].width = 30  # ประเภทการทำธุรกรรม
            report_sheet.column_dimensions['C'].width = 50  # คำอธิบาย
            report_sheet.column_dimensions['D'].width = 20  # รหัสคำสั่งซื้อ
            report_sheet.column_dimensions['E'].width = 25  # รูปแบบธุรกรรม
            report_sheet.column_dimensions['F'].width = 15  # จำนวนเงิน
            report_sheet.column_dimensions['G'].width = 15  # สถานะ
            report_sheet.column_dimensions['H'].width = 25  # ยอดเงินหลังทำธุรกรรมเสร็จสิ้น
            report_sheet.column_dimensions['I'].width = 30  # admin_record_file
            cls()._formating_header(sheet=report_sheet)
            print(f"✅ Saved to: {output_file}")

        return output_file

    @classmethod
    def admin_check(
            cls,
            matched_df: pd.DataFrame, 
            admin_file: str,
            admin_df: pd.DataFrame,
            dry_run: bool=True,
            allow_replace: bool=False) -> pd.DataFrame:
        """Mark received orders in admin finance summary file
        Args:
            matched_df (pd.DataFrame): DataFrame with order IDs that were matched
            admin_file (str): Path to the admin file
            admin_df (pd.DataFrame): Admin DataFrame to update
            dry_run (bool): Whether to update the admin file in place
            allow_replace (bool): Allow replacing existing reconciliation records
        Returns:
            pd.DataFrame: Updated admin DataFrame
        """

        print("📋 Checking admin file for payment reconciliation...")
        
        # Check if any order IDs from matched_df already exist in admin_df (excluding NaN records)
        if 'reported_file' in admin_df.columns:
            print("Column 'reported_file' exists in admin file. Checking for duplicates...")
            already_matched = admin_df[admin_df['reported_file'].notna()]
            if not already_matched.empty:
                duplicate_orders = matched_df[matched_df['หมายเลขคำสั่งซื้อ'].isin(already_matched['หมายเลขคำสั่งซื้อ'])]
                if not duplicate_orders.empty:
                    duplicate_ids = duplicate_orders['หมายเลขคำสั่งซื้อ'].tolist()
                    reported_filename = matched_df['reported_file'].iloc[0] if 'reported_file' in matched_df.columns else 'unknown'
                    if not allow_replace:
                        raise ValueError(f"❌ Found {len(duplicate_ids)} order IDs from '{reported_filename}' that were already reconciled in admin file: {duplicate_ids[:5]}{'...' if len(duplicate_ids) > 5 else ''}")
                    else:
                        print(f"⚠️  Found {len(duplicate_ids)} duplicate order IDs. Updating existing records...")
                        # Update reported_file for these order IDs instead of removing
                        admin_df.loc[admin_df['หมายเลขคำสั่งซื้อ'].isin(duplicate_ids), 'reported_file'] = reported_filename

        # Merge matched orders into admin_df
        merged_df = admin_df.merge(
            matched_df, 
            left_on='หมายเลขคำสั่งซื้อ',
            right_on='หมายเลขคำสั่งซื้อ',
            how='left',
            indicator=True,
            suffixes=('', '_reported')
        )
        
        # Update reported_file for matched rows
        if 'reported_file' in matched_df.columns:
            reported_filename = matched_df['reported_file'].iloc[0]
            # Initialize column if it doesn't exist
            if 'reported_file' not in admin_df.columns:
                merged_df['reported_file'] = ""
            # Update only matched rows
            merged_df.loc[merged_df['_merge'] == 'both', 'reported_file'] = reported_filename
            
            matched_count = merged_df[merged_df['_merge'] == 'both'].shape[0]
            print(f"✅ Marked {matched_count} orders as received in admin file from {reported_filename}")
        
        # Drop _merge indicator and any duplicate columns from merge
        columns_to_drop = ['_merge']
        # Drop any _reported suffix columns that were added during merge
        reported_cols = [col for col in merged_df.columns if col.endswith('_reported')]
        columns_to_drop.extend(reported_cols)
        merged_df = merged_df.drop(columns=columns_to_drop)
        
        if not dry_run:
            # Save updated admin file
            with pd.ExcelWriter(admin_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # add 1 footer row
                # Add footer row with totals
                total_row = {
                    'หมายเลขคำสั่งซื้อ': 'TOTAL',
                    'ราคาขายสุทธิ': merged_df['ราคาขายสุทธิ'].sum(),
                    'ค่าจัดส่งที่ชำระโดยผู้ซื้อ': merged_df['ค่าจัดส่งที่ชำระโดยผู้ซื้อ'].sum(),
                    'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ': merged_df['ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ'].sum(),
                }
                merged_df.loc[len(merged_df)] = total_row
                merged_df.to_excel(writer, sheet_name='Finance Summary', index=False)
                # self.finance_df.to_excel(writer, sheet_name='Finance Summary', index=False)
                finance_sheet: Worksheet = writer.sheets['Finance Summary']
                finance_sheet.column_dimensions['A'].width = 25  # หมายเลขคำสั่งซื้อ
                finance_sheet.column_dimensions['B'].width = 15  # ราคาขายสุทธิ
                finance_sheet.column_dimensions['C'].width = 15  # ค่าจัดส่งที่ชำระโดยผู้ซื้อ
                finance_sheet.column_dimensions['D'].width = 20  # ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ
                cls()._formating_header(finance_sheet)
                cls()._formatting_footer(sheet=finance_sheet, footer_row=len(merged_df)+1)
                print(f"✅ Updated admin file saved to: {admin_file}")
        else:
            print(f"🔍 Dry-run mode: Admin file not updated")
        
        return merged_df

    @classmethod
    def draw_progress_bar(cls, reported_df: pd.DataFrame):
        # Visualize matched orders with progress bar
        number_of_nan_admin_record: int = reported_df['admin_record_file'].isna().sum()
        matched_orders: int = len(reported_df) - number_of_nan_admin_record
        total_orders: int = len(reported_df)
        
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

    @classmethod
    def finance_check(cls, reported_file: str, admin_file: str, dry_run=False, allow_replace=False) -> pd.DataFrame:
        """Compare reported finance file with calculated finance file
        Args:
            reported_file (str): Path to the cleaned reported finance file
            admin_file (str): Path to the admin finance file
            dry_run (bool): Whether to update the reported file in place
            allow_replace (bool): Allow replacing existing matched records
        Returns:
            pd.DataFrame: Merged DataFrame after reconciliation
        """

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
            'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ': np.float64
        }

        try:
            reported_df = pd.read_excel(reported_file, dtype=report_type_dict, sheet_name='Transaction Report')
        except ValueError as e:
            raise ValueError(f"❌ Error reading reported file '{reported_file}': {e}")

        # Before processing, show initial progress
        cls.draw_progress_bar(reported_df)

        if admin_file is None:
            print("=============== ⚠️ No admin file provided. Exiting finance check. ===============")
            return reported_df
        
        admin_type_dict = {
            'หมายเลขคำสั่งซื้อ': str,
            'ราคาขายสุทธิ': np.float64,
            'ค่าจัดส่งที่ชำระโดยผู้ซื้อ': np.float64,
            'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ': np.float64,
            'reported_file': str
        }
        try:
            admin_df = pd.read_excel(admin_file, dtype=admin_type_dict, sheet_name='Finance Summary', skipfooter=1)
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
                if not allow_replace:
                    raise ValueError(f"❌ Found {len(duplicate_ids)} order IDs in '{admin_filename}' that were already matched: {duplicate_ids[:5]}{'...' if len(duplicate_ids) > 5 else ''}")
                else:
                    print(f"⚠️  Found {len(duplicate_ids)} duplicate order IDs. Replacing existing records...")
                    # Remove old matched data for these order IDs
                    reported_df.loc[reported_df['รหัสคำสั่งซื้อ'].isin(duplicate_ids), 'admin_record_file'] = pd.NA
                    # Also clear data columns for re-matching
                    data_columns = ['ราคาขายสุทธิ', 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ', 'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ']
                    for col in data_columns:
                        if col in reported_df.columns:
                            reported_df.loc[reported_df['รหัสคำสั่งซื้อ'].isin(duplicate_ids), col] = pd.NA

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

        admin_filename: str = Path(admin_file).name
        matched_count: int = merged_df[merged_df['_merge'] == 'both'].shape[0]
        print(f"✅ Matched {matched_count} orders with {admin_filename}")
        if matched_count == 0:
            print("=============== ⚠️  No matched orders found for reconciliation. ===============")
            return merged_df
        
        # อัปเดต admin_record_file สำหรับ rows ที่ merge สำเร็จ
        merged_df.loc[merged_df['_merge'] == 'both', 'admin_record_file'] = admin_filename
        
        # Update data columns for matched rows (handle _new suffix from merge)
        for col in data_columns:
            new_col = f'{col}_new'
            if new_col in merged_df.columns:
                # Update only matched rows with new values
                merged_df.loc[merged_df['_merge'] == 'both', col] = merged_df.loc[merged_df['_merge'] == 'both', new_col]
                merged_df = merged_df.drop(columns=[new_col])

        # keep orderIDs as dataframe for merge marking received
        matched_df: pd.DataFrame = merged_df.loc[merged_df['_merge'] == 'both', ['หมายเลขคำสั่งซื้อ']].copy()
        matched_df['reported_file'] = Path(reported_file).name
        
        # ลบ column _merge และ หมายเลขคำสั่งซื้อ (duplicate)
        try:
            merged_df = merged_df.drop(columns=['_merge', 'หมายเลขคำสั่งซื้อ', 'reported_file'])
        except KeyError:
            merged_df = merged_df.drop(columns=['_merge', 'หมายเลขคำสั่งซื้อ'])
        cls.draw_progress_bar(merged_df)
        
        # แสดงผลสรุป
        print(f"⚠️  Remaining unmatched: {merged_df['admin_record_file'].isna().sum()}")

        cls.admin_check(
            matched_df=matched_df,
            admin_file=admin_file,
            admin_df=admin_df,
            dry_run=dry_run,
            allow_replace=allow_replace
        )

        if not dry_run:
            # บันทึกผลลัพธ์กลับไปยัง reported_file
            # merged_df.to_excel(reported_file, index=False)

            # Save cleaned report to output_file
            with pd.ExcelWriter(reported_file, engine='openpyxl') as writer:
                merged_df.to_excel(excel_writer=writer, sheet_name='Transaction Report', index=False)
                report_sheet = writer.sheets['Transaction Report']
                report_sheet.column_dimensions['A'].width = 20  # วันที่
                report_sheet.column_dimensions['B'].width = 30  # ประเภทการทำธุรกรรม
                report_sheet.column_dimensions['C'].width = 50  # คำอธิบาย
                report_sheet.column_dimensions['D'].width = 20  # รหัสคำสั่งซื้อ
                report_sheet.column_dimensions['E'].width = 25  # รูปแบบธุรกรรม
                report_sheet.column_dimensions['F'].width = 15  # จำนวนเงิน
                report_sheet.column_dimensions['G'].width = 15  # สถานะ
                report_sheet.column_dimensions['H'].width = 25  # ยอดเงินหลังทำธุรกรรมเสร็จสิ้น
                report_sheet.column_dimensions['I'].width = 30  # admin_record_file
                report_sheet.column_dimensions['J'].width = 15  # ราคาขายสุทธิ
                report_sheet.column_dimensions['K'].width = 15  # ค่าจัดส่งที่ชำระโดยผู้ซื้อ
                report_sheet.column_dimensions['L'].width = 15  # ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ
                cls()._formating_header(sheet=report_sheet)
                print(f"✅ Updated reported file saved to: {reported_file}")
        
        print("===============🏁 Finance check completed.===============")
        return merged_df
