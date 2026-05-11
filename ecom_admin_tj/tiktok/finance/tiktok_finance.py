import warnings
import pandas as pd
import numpy as np
import re
from pathlib import Path
from datetime import datetime
from tqdm import tqdm
from ...common.reconciliation_mixin import ReconciliationMixin, Worksheet

class TikTokFinanceMixin(ReconciliationMixin):
    """
    Finance related methods for TikTok admin
    
    Extends ReconciliationMixin to provide TikTok-specific finance operations
    including report cleaning, reconciliation, and Excel formatting.
    
    TikTok Finance Logic:
        - Order ID from finance report matches with admin 'Order ID'
        - Total calculated: SKU Subtotal Before Discount - SKU Seller Discount = รวม
        - Admin file contains: Order ID, รวม, SKU Subtotal Before/After Discount, SKU Seller Discount
    """

    # TikTok-specific report type dict
    report_type_dict = {
        'Order/adjustment ID  ': str,  # Note: has trailing spaces in raw report
        'Order created time': str,
        'Total Revenue': np.float64,  # Keep original value from report
        'admin_record_file': 'string',
        'รวม': np.float64,  # From admin: SKU Subtotal Before Discount - SKU Seller Discount
    }

    @classmethod
    def make_finance_report_df(cls, original_report_file: str) -> pd.DataFrame:
        """
        Create a cleaned finance report from the original TikTok report file
        
        Reads TikTok finance report and adds columns for reconciliation with admin data.
        The report will include columns for matching with TikTok admin's Finance Summary.
        
        Args:
            original_report_file: Path to original TikTok finance report
            
        Returns:
            DataFrame with cleaned TikTok finance data including reconciliation columns
        """
        # Suppress openpyxl warnings
        warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')
        
        # Read TikTok finance report
        report_df = pd.read_excel(
            original_report_file,
            sheet_name='Order details',
            header=0,  # TikTok report has headers at row 0
        )
        
        # Select and rename key columns
        # Note: 'Order/adjustment ID  ' has trailing spaces in raw report
        # Note: 20260508 'Order/adjustment ID  ' change to 'Order/Adjustment ID'
        header_order_names = ['Order/adjustment ID  ', 'Order/Adjustment ID']
        header_order_name = None
        for name in header_order_names:
            if name in report_df.columns:
                header_order_name = name
                break

        if header_order_name is None:
            raise ValueError(f'Order/adjustment ID column not found in report. Found: {report_df.columns}')

        report_df = report_df[[
            header_order_name,
            'Order created time',
            'Total Revenue'
        ]].copy()

        # change header name to old header name  Order/adjustment ID
        report_df = report_df.rename(columns={header_order_name: header_order_names[0]})
        
        # Add reconciliation columns (will be populated during finance_check)
        report_df['admin_record_file'] = pd.NA
        report_df['รวม'] = np.nan  # Will be populated from admin
        
        # Set proper dtypes
        report_df = report_df.astype({
            'Order/adjustment ID  ': str,
            'Order created time': str,
            'Total Revenue': np.float64,  # Keep original value
            'admin_record_file': 'string',
            'รวม': np.float64  # From admin calculation
        })
        
        return report_df
    
    def _report_sheet_format_width_column(self, sheet: Worksheet):
        """
        Format column widths for TikTok finance report
        
        Args:
            sheet: Worksheet object to format
        """
        sheet.column_dimensions['A'].width = 25  # Order/adjustment ID
        sheet.column_dimensions['B'].width = 20  # Order created time
        sheet.column_dimensions['C'].width = 15  # Total Revenue
        sheet.column_dimensions['D'].width = 12  # รวม
        sheet.column_dimensions['E'].width = 20  # SKU Subtotal Before Discount
        sheet.column_dimensions['F'].width = 18  # SKU Seller Discount
        sheet.column_dimensions['G'].width = 30  # admin_record_file
    
    @classmethod
    def make_finance_report(cls, original_report_file: str, output_file: str = None, auto_rename: bool = True) -> str:
        """
        Create a cleaned finance report from the original TikTok report file
        
        Args:
            original_report_file: Path to original TikTok finance report
            output_file: Optional custom output filename
            auto_rename: Auto-rename if file exists
            
        Returns:
            Path to the created output file
        """
        report_df = cls.make_finance_report_df(original_report_file)

        if output_file is None:
            # Extract date range from input filename (e.g., income_20260204_20260210)
            input_filename = Path(original_report_file).stem
            date_match = re.search(r'(\d{8}_\d{8})', input_filename)
            
            if date_match:
                date_range = date_match.group(1)
                output_file = f'tiktok_cleaned_finance_report_{date_range}.xlsx'
            else:
                output_file = 'tiktok_cleaned_finance_report.xlsx'

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
            report_df.to_excel(excel_writer=writer, sheet_name='Order details', index=False)
            report_sheet = writer.sheets['Order details']
            cls()._report_sheet_format_width_column(sheet=report_sheet)
            cls()._formating_header(sheet=report_sheet)
            print(f"✅ Saved to: {output_file}")

        return output_file

    @classmethod
    def finance_check(
        cls,
        reported_file: str,
        admin_file: str,
        dry_run: bool = False,
        allow_replace: bool = False,
        column_mapping: dict = None
    ) -> pd.DataFrame:
        """
        TikTok-specific finance reconciliation
        
        Overrides parent method to use TikTok-specific column mapping
        and sheet name ('Order details' instead of 'Transaction Report').
        Matches 'Order/adjustment ID  ' from report with 'Order ID' from admin.
        
        TikTok Finance Logic:
        - Total Revenue: Keep original value from report (DO NOT recalculate)
        - รวม: From admin (calculated as: SKU Subtotal Before Discount - SKU Seller Discount)
        
        Args:
            reported_file: Path to cleaned TikTok finance report
            admin_file: Path to TikTok admin file
            dry_run: Preview mode without saving
            allow_replace: Allow replacing existing matches
            column_mapping: Custom column mapping (uses TikTok defaults if None)
            
        Returns:
            Reconciled DataFrame
        """
        # TikTok-specific column mapping
        if column_mapping is None:
            column_mapping = {
                'order_id': 'Order ID',  # Admin Finance Summary column
                'reported_order_id': 'Order/adjustment ID  ',  # Report column (has trailing spaces!)
                'total': 'รวม',  # Admin calculated total (will be recalculated correctly)
                'net_price': 'SKU Subtotal Before Discount',  # From admin
                'buyer_shipping': 'SKU Seller Discount',  # From admin
                'platform_shipping': 'SKU Subtotal After Discount',  # From admin
                'admin_record_file': 'admin_record_file',
                'total_payment': '_temp_total_payment'  # Temporary - will be dropped
            }
        
        # Call parent's finance_check with dry_run=True to prevent parent from saving
        # We'll save ourselves after fixing the columns
        merged_df = super().finance_check(
            reported_file=reported_file,
            admin_file=admin_file,
            dry_run=True,  # Prevent parent from saving - we'll save after fixes
            allow_replace=allow_replace,
            column_mapping=column_mapping,
            sheet_name='Order details'  # TikTok uses 'Order details' instead of 'Transaction Report'
        )
        
        # Recalculate 'รวม' with TikTok formula: Before Discount - Seller Discount
        # (Parent calculated it as: Before Discount + Seller Discount, which is wrong for TikTok)
        if 'SKU Subtotal Before Discount' in merged_df.columns and 'SKU Seller Discount' in merged_df.columns:
            # Calculate รวม correctly for matched rows only
            merged_df.loc[merged_df['admin_record_file'].notna(), 'รวม'] = (
                merged_df.loc[merged_df['admin_record_file'].notna(), 'SKU Subtotal Before Discount'] - 
                merged_df.loc[merged_df['admin_record_file'].notna(), 'SKU Seller Discount']
            )
        
        # Prepare matched DataFrame and call our TikTok-specific admin_check
        # (Parent's admin_check was called with dry_run=True, so it didn't save)
        matched_orders = merged_df.loc[
            merged_df['admin_record_file'].notna(),
            ['Order/adjustment ID  ', 'admin_record_file']
        ].copy()
        
        if not matched_orders.empty:
            # Read admin file to update with our TikTok-specific logic
            admin_type_dict = {
                column_mapping['order_id']: str,
                column_mapping['total']: np.float64,
                column_mapping['net_price']: np.float64,
                column_mapping['buyer_shipping']: np.float64,
                column_mapping['platform_shipping']: np.float64,
                'reported_file': str,  # For tracking which report matched this admin record
            }
            
            admin_df = pd.read_excel(
                admin_file,
                dtype=admin_type_dict,
                sheet_name='Finance Summary',
                skipfooter=1
            )
            
            # Prepare matched_df for admin_check
            matched_df_for_admin = pd.DataFrame({
                column_mapping['order_id']: matched_orders['Order/adjustment ID  '].values,
                'reported_file': Path(reported_file).name
            })
            
            # Call our TikTok-specific admin_check with the user's dry_run parameter
            cls.admin_check(
                matched_df=matched_df_for_admin,
                admin_file=admin_file,
                admin_df=admin_df,
                dry_run=dry_run,  # Use the dry_run parameter from the user
                allow_replace=allow_replace,
                column_mapping=column_mapping
            )
        
        # Drop unwanted columns - keep only the 7 required columns
        columns_to_keep = [
            'Order/adjustment ID  ',
            'Order created time',
            'Total Revenue',
            'รวม',
            'SKU Subtotal Before Discount',
            'SKU Seller Discount',
            'admin_record_file',        
        ]
        
        # Drop any extra columns that were merged from admin
        cols_to_drop = [col for col in merged_df.columns if col not in columns_to_keep]
        if cols_to_drop:
            merged_df = merged_df.drop(columns=cols_to_drop)
        
        # Reorder columns to ensure correct order
        merged_df = merged_df[columns_to_keep]
        
        # Sort by admin_record_file
        merged_df.sort_values('admin_record_file', inplace=True)
        
        # Now save the corrected file ourselves (if not in dry_run mode)
        if not dry_run:
            with pd.ExcelWriter(reported_file, engine='openpyxl') as writer:
                merged_df.to_excel(excel_writer=writer, sheet_name='Order details', index=False)
                report_sheet = writer.sheets['Order details']
                cls()._report_sheet_format_width_column(sheet=report_sheet)
                cls()._formating_header(sheet=report_sheet)
                print(f"✅ Updated reported file saved to: {reported_file}")
        else:
            print(f"🔍 Dry-run mode: Reported file not updated")
        
        print("===============🏁 Finance check completed.===============")
        return merged_df
    
    @classmethod
    def admin_check(
        cls,
        matched_df: pd.DataFrame,
        admin_file: str,
        admin_df: pd.DataFrame,
        dry_run: bool = False,
        allow_replace: bool = False,
        column_mapping: dict = None
    ) -> pd.DataFrame:
        """
        TikTok-specific admin file update
        
        Overrides parent method to use TikTok-specific calculation logic.
        For TikTok: รวม = SKU Subtotal Before Discount - SKU Seller Discount
        (Parent uses addition, but TikTok requires subtraction)
        
        Args:
            matched_df: DataFrame with matched order IDs
            admin_file: Path to TikTok admin file
            admin_df: Admin DataFrame to update
            dry_run: Preview mode without saving
            allow_replace: Allow replacing existing matches
            column_mapping: Custom column mapping (uses TikTok defaults if None)
            
        Returns:
            Updated admin DataFrame
        """
        from openpyxl.worksheet.worksheet import Worksheet
        from pathlib import Path
        
        # TikTok-specific column mapping
        if column_mapping is None:
            column_mapping = {
                'order_id': 'Order ID',
                'total': 'รวม',
                'net_price': 'SKU Subtotal Before Discount',
                'buyer_shipping': 'SKU Seller Discount',
                'platform_shipping': 'SKU Subtotal After Discount',
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
        
        # Recalculate 'รวม' with TikTok formula: Before Discount - Seller Discount
        # (This ensures correct values before merging and saving)
        if column_mapping['net_price'] in admin_df.columns and column_mapping['buyer_shipping'] in admin_df.columns:
            admin_df[column_mapping['total']] = (
                admin_df[column_mapping['net_price']] - 
                admin_df[column_mapping['buyer_shipping']]
            )
            print(f"✅ Recalculated 'รวม' column using TikTok formula (subtraction)")
        
        # Reorder columns to match TikTok structure (same as tiktok.py line 151-155)
        order_columns = [
            column_mapping['order_id'],  # Order ID
            column_mapping['total'],  # รวม
            column_mapping['net_price'],  # SKU Subtotal Before Discount
            column_mapping['buyer_shipping'],  # SKU Seller Discount
            column_mapping['platform_shipping'],  # SKU Subtotal After Discount
        ]
        
        # Add reported_file if it exists
        if 'reported_file' in admin_df.columns:
            order_columns.append('reported_file')
        
        # Keep only existing columns in the specified order
        existing_columns = [col for col in order_columns if col in admin_df.columns]
        admin_df = admin_df[existing_columns]
        
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
            # Save updated admin file with TikTok-specific calculation
            with pd.ExcelWriter(admin_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Add footer row with totals (calculate TOTAL['รวม'] correctly)
                total_row = {order_col: 'TOTAL'}
                
                # For TikTok: TOTAL['รวม'] = sum(Before Discount) - sum(Seller Discount)
                # NOT sum('รวม') to avoid propagating calculation errors
                for col in [column_mapping['net_price'], column_mapping['buyer_shipping'], 
                           column_mapping['platform_shipping']]:
                    if col in merged_df.columns:
                        total_row[col] = merged_df[col].sum()
                
                # Calculate TOTAL for 'รวม' using TikTok formula
                total_row[column_mapping['total']] = (
                    total_row.get(column_mapping['net_price'], 0) - 
                    total_row.get(column_mapping['buyer_shipping'], 0)
                )
                
                merged_df.loc[len(merged_df)] = total_row
                merged_df.to_excel(writer, sheet_name='Finance Summary', index=False)
                finance_sheet: Worksheet = writer.sheets['Finance Summary']
                
                # Set column widths matching TikTok structure
                finance_sheet.column_dimensions['A'].width = 25  # Order ID
                finance_sheet.column_dimensions['B'].width = 12  # รวม
                finance_sheet.column_dimensions['C'].width = 18  # SKU Subtotal Before Discount
                finance_sheet.column_dimensions['D'].width = 18  # SKU Seller Discount
                finance_sheet.column_dimensions['E'].width = 18  # SKU Subtotal After Discount
                
                cls()._formating_header(finance_sheet)
                cls()._formatting_footer(sheet=finance_sheet, footer_row=len(merged_df) + 1)
                print(f"✅ Updated admin file saved to: {admin_file}")
        else:
            print(f"🔍 Dry-run mode: Admin file not updated")
        
        return merged_df

    # Note: admin_check, draw_progress_bar, and finance_check methods
    # are now inherited from ReconciliationMixin parent class.
    # They can be overridden here if TikTok-specific behavior is needed.
