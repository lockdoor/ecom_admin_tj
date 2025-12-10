import argparse
import re
from pathlib import Path
from datetime import datetime
from .shopee_finance import ShopeeFinanceMixin

def create_argument_parser() -> argparse.ArgumentParser:
    """
    Create argument parser for Shopee finance processing
    Returns:
        argparse.ArgumentParser: Configured argument parser
    """
    
    parser = argparse.ArgumentParser(
        description=f'Process Shopee finance reports',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument(
        'report_file',
        type=str,
        help='Path to the original finance report file'
    )

    parser.add_argument(
        '-a', '--admin',
        type=str,
        help='Path to the admin finance file for reconciliation',
        dest='admin_file',
        required=False
    )

    parser.add_argument(
        '-d', '--admin-dir',
        type=str,
        help='Directory containing multiple admin files (*_output.xlsx)',
        dest='admin_dir',
        required=False
    )

    parser.add_argument(
        '--date-from',
        type=str,
        help='Start date for filtering admin files (YYYY-MM-DD)',
        dest='date_from',
        required=False
    )

    parser.add_argument(
        '--date-to',
        type=str,
        help='End date for filtering admin files (YYYY-MM-DD)',
        dest='date_to',
        required=False
    )

    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview changes without updating the reported file',
        dest='dry_run',
        default=False,
        required=False
    )
    
    parser.add_argument(
        '--allow-replace',
        action='store_true',
        help='Allow replacing existing matched/reconciled records',
        dest='allow_replace',
        default=False,
        required=False
    )
    return parser

def extract_date_from_filename(filename: str) -> str | None:
    """Extract date from filename pattern like shopee20251208_output.xlsx"""
    match = re.search(r'(\d{8})_output\.xlsx$', filename)
    if match:
        date_str = match.group(1)
        # Convert YYYYMMDD to YYYY-MM-DD
        return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
    return None

def filter_admin_files(admin_dir: str, date_from: str = None, date_to: str = None) -> list:
    """Get list of admin files filtered by date range"""
    admin_path = Path(admin_dir)
    if not admin_path.exists():
        raise ValueError(f"Admin directory not found: {admin_dir}")
    
    admin_files = []
    for file in admin_path.glob("*_output.xlsx"):
        file_date = extract_date_from_filename(file.name)
        if file_date:
            # Check date range
            if date_from and file_date < date_from:
                continue
            if date_to and file_date > date_to:
                continue
            admin_files.append((str(file), file_date))
    
    # Sort by date
    admin_files.sort(key=lambda x: x[1])
    return admin_files

def main():
    parser = create_argument_parser()
    parsed_args = parser.parse_args()
    
    try:
        # Check if using single file or directory mode
        if parsed_args.admin_dir:
            # Multiple files mode
            admin_files = filter_admin_files(
                parsed_args.admin_dir,
                parsed_args.date_from,
                parsed_args.date_to
            )
            
            if not admin_files:
                print("⚠️  No admin files found matching the criteria.")
                return
            
            print(f"📁 Found {len(admin_files)} admin file(s) to process")
            for admin_file, file_date in admin_files:
                print(f"\n{'='*80}")
                print(f"Processing: {Path(admin_file).name} (Date: {file_date})")
                print(f"{'='*80}")
                
                ShopeeFinanceMixin.finance_check(
                    reported_file=parsed_args.report_file,
                    admin_file=admin_file,
                    dry_run=parsed_args.dry_run,
                    allow_replace=parsed_args.allow_replace
                )
        else:
            # Single file mode
            ShopeeFinanceMixin.finance_check(
                reported_file=parsed_args.report_file,
                admin_file=parsed_args.admin_file,
                dry_run=parsed_args.dry_run,
                allow_replace=parsed_args.allow_replace
            )
    except FileNotFoundError as e:
        print(f"❌ File not found: {e.filename}")
    except ValueError as e:
        print(f"❌ Value error: {e}")
    

if __name__ == "__main__":
    main()
