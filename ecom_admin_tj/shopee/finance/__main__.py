"""
Shopee Finance Reconciliation Entry Point

Processes Shopee finance reports and reconciles them with admin records.

Usage:
    # Single file reconciliation
    python -m ecom_admin_tj.shopee.finance report.xlsx -a admin_file.xlsx
    
    # Batch processing from directory
    python -m ecom_admin_tj.shopee.finance report.xlsx -d admin_dir/ \\
        --date-from 2026-01-01 --date-to 2026-01-31
    
    # Dry-run mode (preview without updating)
    python -m ecom_admin_tj.shopee.finance report.xlsx -a admin_file.xlsx --dry-run
"""
import sys
from ...common.cli.finance_cli import FinanceCLI
from .shopee_finance import ShopeeFinanceMixin

def main():
    """Main entry point for Shopee finance CLI"""
    cli = FinanceCLI(
        platform_name='Shopee',
        finance_checker_class=ShopeeFinanceMixin
    )
    exit_code = cli.run()
    sys.exit(exit_code)

if __name__ == "__main__":
    main()
