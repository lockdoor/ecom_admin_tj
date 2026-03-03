"""
Finance Reconciliation Mixin

Provides advanced reconciliation features for matching reported finance
data with admin records. Can be used by any platform that needs these features.
"""
import warnings
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict
from tqdm import tqdm
from .finance_base_mixin import FinanceBaseMixin, Worksheet


class ReconciliationMixin(FinanceBaseMixin):
    """
    Advanced reconciliation features for finance processing
    
    Provides:
    - Report cleaning (make_finance_report)
    - Finance reconciliation (finance_check)
    - Admin file updates (admin_check)
    - Progress visualization
    """
    
    # Default report type dict - can be overridden by subclasses
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
    def make_finance_report_df(
        cls,
        original_report_file: str,
        sheet_name: str = 'Transaction Report',
        header_row: int = 17,
        report_columns: list = None
    ) -> pd.DataFrame:
        """
        Create a cleaned finance report DataFrame from the original report file
        
        Args:
            original_report_file: Path to the original finance report Excel file
            sheet_name: Name of the sheet containing the report
            header_row: Row number where headers start
            report_columns: List of columns to keep (if None, uses default)
            
        Returns:
            Cleaned DataFrame with standardized columns
        """
        # Suppress openpyxl UserWarnings
        warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')
        
        report_df = pd.read_excel(
            original_report_file,
            sheet_name=sheet_name,
            header=header_row,
            dtype=cls().report_type_dict
        )
        
        # Default columns if not specified
        if report_columns is None:
            report_columns = [
                'วันที่', 'ประเภทการทำธุรกรรม', 'รหัสคำสั่งซื้อ',
                'จำนวนเงิน', 'สถานะ', 'admin_record_file',
                'ราคาขายสุทธิ', 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ', 'รวมชำระ'
            ]
        
        # Initialize new columns with appropriate defaults
        report_df['admin_record_file'] = pd.NA
        report_df['ราคาขายสุทธิ'] = np.nan
        report_df['ค่าจัดส่งที่ชำระโดยผู้ซื้อ'] = np.nan
        report_df['รวมชำระ'] = np.nan
        
        # Set up dtypes
        report_df = report_df.astype({
            'admin_record_file': 'string',
            'ราคาขายสุทธิ': 'float64',
            'ค่าจัดส่งที่ชำระโดยผู้ซื้อ': 'float64',
            'รวมชำระ': 'float64'
        })
        
        return report_df[report_columns]
    
    def _report_sheet_format_width_column(self, sheet: Worksheet):
        """Format column widths for report sheet"""
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
    def draw_progress_bar(cls, reported_df: pd.DataFrame):
        """Visualize matched orders with progress bar"""
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
        
        with tqdm(
            total=total_orders,
            desc=f"{color}Matched Orders{reset_color}",
            unit="order",
            ncols=80,
            bar_format='{desc}: {percentage:3.1f}%|{bar}| {n_fmt}/{total_fmt}',
            colour='green' if match_percentage >= 80 else 'yellow' if match_percentage >= 50 else 'red'
        ) as pbar:
            pbar.update(matched_orders)
    
    @classmethod
    def admin_check(
        cls,
        matched_df: pd.DataFrame,
        admin_file: str,
        admin_df: pd.DataFrame,
        dry_run: bool = True,
        allow_replace: bool = False,
        column_mapping: Dict[str, str] = None
    ) -> pd.DataFrame:
        """
        Mark received orders in admin finance summary file
        
        Args:
            matched_df: DataFrame with order IDs that were matched
            admin_file: Path to the admin file
            admin_df: Admin DataFrame to update
            dry_run: Whether to update the admin file in place
            allow_replace: Allow replacing existing reconciliation records
            column_mapping: Dict mapping standard names to actual column names
            
        Returns:
            Updated admin DataFrame
        """
        # Default column mapping for Shopee-like structure
        if column_mapping is None:
            column_mapping = {
                'order_id': 'หมายเลขคำสั่งซื้อ',
                'total': 'รวม',
                'net_price': 'ราคาขายสุทธิ',
                'buyer_shipping': 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ',
                'platform_shipping': 'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ',
                'reported_file': 'reported_file'
            }
        
        order_col = column_mapping['order_id']
        print("📋 Checking admin file for payment reconciliation...")
        
        # Check for duplicates
        if 'reported_file' in admin_df.columns:
            print("Column 'reported_file' exists in admin file. Checking for duplicates...")
            already_matched = admin_df[admin_df['reported_file'].notna()]
            if not already_matched.empty:
                duplicate_orders = matched_df[matched_df[order_col].isin(already_matched[order_col])]
                if not duplicate_orders.empty:
                    duplicate_ids = duplicate_orders[order_col].tolist()
                    reported_filename = matched_df['reported_file'].iloc[0] if 'reported_file' in matched_df.columns else 'unknown'
                    if not allow_replace:
                        raise ValueError(
                            f"❌ Found {len(duplicate_ids)} order IDs from '{reported_filename}' "
                            f"that were already reconciled in admin file: {duplicate_ids[:5]}"
                            f"{'...' if len(duplicate_ids) > 5 else ''}"
                        )
                    else:
                        print(f"⚠️  Found {len(duplicate_ids)} duplicate order IDs. Updating existing records...")
                        admin_df.loc[admin_df[order_col].isin(duplicate_ids), 'reported_file'] = reported_filename
        
        # Merge matched orders into admin_df
        merged_df = admin_df.merge(
            matched_df,
            left_on=order_col,
            right_on=order_col,
            how='left',
            indicator=True,
            suffixes=('', '_reported')
        )
        
        # Update reported_file for matched rows
        if 'reported_file' in matched_df.columns:
            reported_filename = matched_df['reported_file'].iloc[0]
            if 'reported_file' not in admin_df.columns:
                merged_df['reported_file'] = ""
            merged_df.loc[merged_df['_merge'] == 'both', 'reported_file'] = reported_filename
            
            matched_count = merged_df[merged_df['_merge'] == 'both'].shape[0]
            print(f"✅ Marked {matched_count} orders as received in admin file from {reported_filename}")
        
        # Drop merge indicators and duplicate columns
        columns_to_drop = ['_merge']
        reported_cols = [col for col in merged_df.columns if col.endswith('_reported')]
        columns_to_drop.extend(reported_cols)
        merged_df = merged_df.drop(columns=columns_to_drop)
        
        if not dry_run:
            # Save updated admin file
            with pd.ExcelWriter(admin_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Add footer row with totals
                total_row = {order_col: 'TOTAL'}
                for col in [column_mapping['total'], column_mapping['net_price'],
                           column_mapping['buyer_shipping'], column_mapping['platform_shipping']]:
                    if col in merged_df.columns:
                        total_row[col] = merged_df[col].sum()
                
                merged_df.loc[len(merged_df)] = total_row
                merged_df.to_excel(writer, sheet_name='Finance Summary', index=False)
                finance_sheet: Worksheet = writer.sheets['Finance Summary']
                
                # Set column widths
                finance_sheet.column_dimensions['A'].width = 25
                finance_sheet.column_dimensions['B'].width = 15
                finance_sheet.column_dimensions['C'].width = 15
                finance_sheet.column_dimensions['D'].width = 15
                finance_sheet.column_dimensions['E'].width = 20
                
                cls()._formating_header(finance_sheet)
                cls()._formatting_footer(sheet=finance_sheet, footer_row=len(merged_df) + 1)
                print(f"✅ Updated admin file saved to: {admin_file}")
        else:
            print(f"🔍 Dry-run mode: Admin file not updated")
        
        return merged_df
    
    @classmethod
    def finance_check(
        cls,
        reported_file: str,
        admin_file: str,
        dry_run: bool = False,
        allow_replace: bool = False,
        column_mapping: Dict[str, str] = None
    ) -> pd.DataFrame:
        """
        Compare reported finance file with calculated finance file
        
        Args:
            reported_file: Path to the cleaned reported finance file
            admin_file: Path to the admin finance file
            dry_run: Whether to update the reported file in place
            allow_replace: Allow replacing existing matched records
            column_mapping: Dict mapping standard names to actual column names
            
        Returns:
            Merged DataFrame after reconciliation
        """
        # Default column mapping for Shopee-like structure
        if column_mapping is None:
            column_mapping = {
                'order_id': 'หมายเลขคำสั่งซื้อ',
                'reported_order_id': 'รหัสคำสั่งซื้อ',
                'total': 'รวม',
                'net_price': 'ราคาขายสุทธิ',
                'buyer_shipping': 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ',
                'platform_shipping': 'ค่าจัดส่งที่ Shopee ออกให้โดยประมาณ',
                'admin_record_file': 'admin_record_file',
                'total_payment': 'รวมชำระ'
            }
        
        try:
            reported_df = pd.read_excel(
                reported_file,
                dtype=cls().report_type_dict,
                sheet_name='Transaction Report'
            )
        except ValueError as e:
            raise ValueError(f"❌ Error reading reported file '{reported_file}': {e}")
        
        # Show initial progress
        cls.draw_progress_bar(reported_df)
        
        if admin_file is None:
            print("=============== ⚠️ No admin file provided. Exiting finance check. ===============")
            return reported_df
        
        # Load admin file
        admin_type_dict = {
            column_mapping['order_id']: str,
            column_mapping['net_price']: np.float64,
            column_mapping['buyer_shipping']: np.float64,
            column_mapping['platform_shipping']: np.float64,
            'reported_file': str
        }
        
        try:
            admin_df = pd.read_excel(
                admin_file,
                dtype=admin_type_dict,
                sheet_name='Finance Summary',
                skipfooter=1
            )
            
            # Calculate 'รวม' if not present
            if column_mapping['total'] not in admin_df.columns:
                print(f"Column '{column_mapping['total']}' not found. Calculating from price columns.")
                admin_df[column_mapping['total']] = (
                    admin_df[column_mapping['net_price']] +
                    admin_df[column_mapping['buyer_shipping']]
                )
                
        except ValueError as e:
            raise ValueError(f"❌ Error reading admin file '{admin_file}': {e}")
        
        print(f"Number of orders in admin file: {len(admin_df)}")
        
        # Check for duplicate matches
        already_matched = reported_df[reported_df[column_mapping['admin_record_file']].notna()]
        if not already_matched.empty:
            duplicate_orders = admin_df[
                admin_df[column_mapping['order_id']].isin(already_matched[column_mapping['reported_order_id']])
            ]
            if not duplicate_orders.empty:
                duplicate_ids = duplicate_orders[column_mapping['order_id']].tolist()
                admin_filename = Path(admin_file).name
                if not allow_replace:
                    raise ValueError(
                        f"❌ Found {len(duplicate_ids)} order IDs in '{admin_filename}' "
                        f"that were already matched: {duplicate_ids[:5]}"
                        f"{'...' if len(duplicate_ids) > 5 else ''}"
                    )
                else:
                    print(f"⚠️  Found {len(duplicate_ids)} duplicate order IDs. Replacing existing records...")
                    reported_df.loc[
                        reported_df[column_mapping['reported_order_id']].isin(duplicate_ids),
                        column_mapping['admin_record_file']
                    ] = pd.NA
        
        # Merge reported with admin
        data_columns = [
            column_mapping['total'],
            column_mapping['net_price'],
            column_mapping['buyer_shipping'],
            column_mapping['platform_shipping']
        ]
        
        admin_cols_to_merge = [column_mapping['order_id']] + [
            col for col in admin_df.columns
            if col not in reported_df.columns or col in data_columns
        ]
        admin_cols_to_merge = list(dict.fromkeys(admin_cols_to_merge))
        
        admin_df_filtered = admin_df[admin_cols_to_merge].copy()
        
        merged_df = reported_df.merge(
            admin_df_filtered,
            left_on=column_mapping['reported_order_id'],
            right_on=column_mapping['order_id'],
            how='left',
            indicator=True,
            suffixes=('', '_new')
        )
        
        matched_count: int = merged_df[merged_df['_merge'] == 'both'].shape[0]
        admin_filename: str = Path(admin_file).name
        print(f"✅ Matched {matched_count} orders with {admin_filename}")
        
        if matched_count == 0:
            print("=============== ⚠️  No matched orders found for reconciliation. ===============")
            return merged_df
        
        # Update matched records
        merged_df.loc[merged_df['_merge'] == 'both', column_mapping['admin_record_file']] = admin_filename
        
        # Update data columns
        for col in data_columns:
            new_col = f'{col}_new'
            if new_col in merged_df.columns:
                merged_df.loc[merged_df['_merge'] == 'both', col] = merged_df.loc[
                    merged_df['_merge'] == 'both', new_col
                ]
                merged_df = merged_df.drop(columns=[new_col])
        
        # Prepare matched DataFrame for admin_check
        matched_df: pd.DataFrame = merged_df.loc[
            merged_df['_merge'] == 'both',
            [column_mapping['order_id']]
        ].copy()
        matched_df['reported_file'] = Path(reported_file).name
        
        # Clean up merge columns
        try:
            merged_df = merged_df.drop(
                columns=['_merge', column_mapping['order_id'], 'reported_file',
                        column_mapping['platform_shipping']]
            )
        except KeyError:
            merged_df = merged_df.drop(
                columns=['_merge', column_mapping['order_id'], column_mapping['platform_shipping']]
            )
        
        cls.draw_progress_bar(merged_df)
        
        # Calculate total payment
        merged_df[column_mapping['total_payment']] = (
            merged_df[column_mapping['net_price']] +
            merged_df[column_mapping['buyer_shipping']]
        )
        
        # Sort by admin_record_file
        merged_df.sort_values(column_mapping['admin_record_file'], inplace=True)
        
        print(f"⚠️  Remaining unmatched: {merged_df[column_mapping['admin_record_file']].isna().sum()}")
        
        # Update admin file
        cls.admin_check(
            matched_df=matched_df,
            admin_file=admin_file,
            admin_df=admin_df,
            dry_run=dry_run,
            allow_replace=allow_replace,
            column_mapping=column_mapping
        )
        
        if not dry_run:
            # Save updated reported file
            with pd.ExcelWriter(reported_file, engine='openpyxl') as writer:
                merged_df.to_excel(excel_writer=writer, sheet_name='Transaction Report', index=False)
                report_sheet = writer.sheets['Transaction Report']
                cls()._report_sheet_format_width_column(sheet=report_sheet)
                cls()._formating_header(sheet=report_sheet)
                print(f"✅ Updated reported file saved to: {reported_file}")
        
        print("===============🏁 Finance check completed.===============")
        return merged_df
