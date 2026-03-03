"""
Shopee New Report (Clean Report) Entry Point

Creates a cleaned finance report from original Shopee finance report.
This extracts and formats the essential transaction data into a clean
Excel file with proper column widths and formatting.

Usage:
    # Basic usage - auto-generate output filename
    python -m ecom_admin_tj.shopee.finance.new_report input_report.xlsx
    
    # Specify custom output filename
    python -m ecom_admin_tj.shopee.finance.new_report input_report.xlsx -o cleaned.xlsx
    
    # Overwrite existing file (no auto-rename)
    python -m ecom_admin_tj.shopee.finance.new_report input_report.xlsx --no-auto-rename
"""
import sys
from ....common.cli.report_cleaner_cli import ReportCleanerCLI
from ..shopee_finance import ShopeeFinanceMixin

def main():
    """Main entry point for Shopee report cleaner"""
    cli = ReportCleanerCLI(
        platform_name='Shopee',
        finance_mixin_class=ShopeeFinanceMixin
    )
    exit_code = cli.run()
    sys.exit(exit_code)

if __name__ == "__main__":
    main()
