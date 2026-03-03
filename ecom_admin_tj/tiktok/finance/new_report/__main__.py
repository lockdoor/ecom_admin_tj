"""
TikTok New Report (Clean Report) Entry Point

Creates a cleaned finance report from original TikTok finance report.
Extracts key columns: Order/adjustment ID, Order created time, and Total Revenue.

Usage:
    # Basic usage - auto-generate output filename
    python -m ecom_admin_tj.tiktok.finance.new_report input_report.xlsx
    
    # Specify custom output filename
    python -m ecom_admin_tj.tiktok.finance.new_report input_report.xlsx -o cleaned.xlsx
    
    # Overwrite existing file (no auto-rename)
    python -m ecom_admin_tj.tiktok.finance.new_report input_report.xlsx --no-auto-rename
"""
import sys
from ....common.cli.report_cleaner_cli import ReportCleanerCLI
from ..tiktok_finance import TikTokFinanceMixin

def main():
    """Main entry point for TikTok report cleaner"""
    cli = ReportCleanerCLI(
        platform_name='TikTok',
        finance_mixin_class=TikTokFinanceMixin
    )
    exit_code = cli.run()
    sys.exit(exit_code)

if __name__ == "__main__":
    main()
