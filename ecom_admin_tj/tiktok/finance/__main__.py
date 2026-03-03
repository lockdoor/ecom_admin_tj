"""
TikTok Finance Reconciliation Entry Point

Processes TikTok finance reports and reconciles them with admin records.
Matches Order IDs from finance report with TikTok admin Finance Summary sheet.

Usage:
    # Single file reconciliation
    python -m ecom_admin_tj.tiktok.finance report.xlsx -a admin_file.xlsx
    
    # Batch processing from directory
    python -m ecom_admin_tj.tiktok.finance report.xlsx -d admin_dir/ \\
        --date-from 2026-01-01 --date-to 2026-01-31
    
    # Dry-run mode (preview without updating)
    python -m ecom_admin_tj.tiktok.finance report.xlsx -a admin_file.xlsx --dry-run
"""
import sys
from ...common.cli.finance_cli import FinanceCLI
from .tiktok_finance import TikTokFinanceMixin

def main():
    """Main entry point for TikTok finance CLI"""
    cli = FinanceCLI(
        platform_name='TikTok',
        finance_checker_class=TikTokFinanceMixin
    )
    exit_code = cli.run()
    sys.exit(exit_code)

if __name__ == "__main__":
    main()
