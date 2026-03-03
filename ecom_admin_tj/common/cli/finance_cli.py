"""
Finance CLI - Generic command-line interface for finance reconciliation

Provides standardized CLI for finance checking and reconciliation
that can be used by any platform.
"""
import argparse
import re
import sys
from pathlib import Path
from typing import Type, Optional


class FinanceCLI:
    """
    Generic CLI for finance reconciliation operations
    
    Provides standardized argument parsing and execution flow for:
    - Single file reconciliation
    - Batch processing from directory
    - Date range filtering
    - Dry-run mode
    """
    
    def __init__(
        self,
        platform_name: str,
        finance_checker_class: Type,
        filename_pattern: str = r'(\d{8})_output\.xlsx$'
    ):
        """
        Initialize Finance CLI
        
        Args:
            platform_name: Name of the platform (e.g., 'Shopee', 'Tiktok')
            finance_checker_class: Class that implements finance_check method
            filename_pattern: Regex pattern to extract date from filename
        """
        self.platform_name = platform_name
        self.finance_checker_class = finance_checker_class
        self.filename_pattern = filename_pattern
        self.parser = self._create_parser()
    
    def _create_parser(self) -> argparse.ArgumentParser:
        """Create standard argument parser for finance processing"""
        parser = argparse.ArgumentParser(
            description=f'Process {self.platform_name} finance reports',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog=f"""
Examples:
  # Single file reconciliation
  python -m ecom_admin_tj.{self.platform_name.lower()}.finance \\
      report.xlsx -a admin_file.xlsx
  
  # Batch processing from directory
  python -m ecom_admin_tj.{self.platform_name.lower()}.finance \\
      report.xlsx -d admin_dir/ --date-from 2026-01-01 --date-to 2026-01-31
  
  # Dry-run mode (preview without updating)
  python -m ecom_admin_tj.{self.platform_name.lower()}.finance \\
      report.xlsx -a admin_file.xlsx --dry-run
            """
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
            help='Preview changes without updating files',
            dest='dry_run',
            default=False
        )
        
        parser.add_argument(
            '--allow-replace',
            action='store_true',
            help='Allow replacing existing matched records',
            dest='allow_replace',
            default=False
        )
        
        return parser
    
    def extract_date_from_filename(self, filename: str) -> Optional[str]:
        """
        Extract date from filename using regex pattern
        
        Args:
            filename: Filename to extract date from
            
        Returns:
            Date string in YYYY-MM-DD format, or None if not found
        """
        match = re.search(self.filename_pattern, filename)
        if match:
            date_str = match.group(1)
            # Convert YYYYMMDD to YYYY-MM-DD
            return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
        return None
    
    def filter_admin_files(
        self,
        admin_dir: str,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None
    ) -> list:
        """
        Get list of admin files filtered by date range
        
        Args:
            admin_dir: Directory containing admin files
            date_from: Start date (YYYY-MM-DD) or None
            date_to: End date (YYYY-MM-DD) or None
            
        Returns:
            List of (filepath, date) tuples sorted by date
        """
        admin_path = Path(admin_dir)
        if not admin_path.exists():
            raise ValueError(f"Admin directory not found: {admin_dir}")
        
        admin_files = []
        for file in admin_path.glob("*_output.xlsx"):
            file_date = self.extract_date_from_filename(file.name)
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
    
    def run(self, args: Optional[list] = None) -> int:
        """
        Execute finance reconciliation process
        
        Args:
            args: Command-line arguments (None = use sys.argv)
            
        Returns:
            Exit code (0 = success, 1 = error)
        """
        try:
            parsed_args = self.parser.parse_args(args)
        except SystemExit as e:
            return e.code if e.code else 0
        
        try:
            # Validate arguments
            if not parsed_args.admin_file and not parsed_args.admin_dir:
                print("⚠️  No admin file or directory specified. Run with -h for help.")
                return 0
            
            if parsed_args.admin_file and parsed_args.admin_dir:
                print("❌ Cannot specify both --admin and --admin-dir")
                return 1
            
            # Check if using single file or directory mode
            if parsed_args.admin_dir:
                return self._run_batch_mode(parsed_args)
            else:
                return self._run_single_mode(parsed_args)
                
        except FileNotFoundError as e:
            print(f"❌ File not found: {e.filename if hasattr(e, 'filename') else e}")
            return 1
        except ValueError as e:
            print(f"❌ Value error: {e}")
            return 1
        except KeyboardInterrupt:
            print("\n⚠️  Process interrupted by user")
            return 130
        except Exception as e:
            print(f"❌ Unexpected error: {e}")
            import traceback
            traceback.print_exc()
            return 1
    
    def _run_single_mode(self, parsed_args) -> int:
        """Run reconciliation with single admin file"""
        self.finance_checker_class.finance_check(
            reported_file=parsed_args.report_file,
            admin_file=parsed_args.admin_file,
            dry_run=parsed_args.dry_run,
            allow_replace=parsed_args.allow_replace
        )
        return 0
    
    def _run_batch_mode(self, parsed_args) -> int:
        """Run reconciliation with multiple admin files from directory"""
        admin_files = self.filter_admin_files(
            parsed_args.admin_dir,
            parsed_args.date_from,
            parsed_args.date_to
        )
        
        if not admin_files:
            print("⚠️  No admin files found matching the criteria.")
            return 0
        
        print(f"📁 Found {len(admin_files)} admin file(s) to process")
        
        for admin_file, file_date in admin_files:
            print(f"\n{'='*80}")
            print(f"Processing: {Path(admin_file).name} (Date: {file_date})")
            print(f"{'='*80}")
            
            try:
                self.finance_checker_class.finance_check(
                    reported_file=parsed_args.report_file,
                    admin_file=admin_file,
                    dry_run=parsed_args.dry_run,
                    allow_replace=parsed_args.allow_replace
                )
            except Exception as e:
                print(f"❌ Error processing {admin_file}: {e}")
                if not parsed_args.dry_run:
                    # In non-dry-run mode, stop on error
                    return 1
                # In dry-run mode, continue with other files
        
        return 0


def create_finance_cli(platform_name: str, finance_checker_class: Type) -> FinanceCLI:
    """
    Factory function to create FinanceCLI instance
    
    Args:
        platform_name: Name of the platform (e.g., 'Shopee', 'Tiktok')
        finance_checker_class: Class that implements finance_check method
        
    Returns:
        Configured FinanceCLI instance
    """
    return FinanceCLI(platform_name, finance_checker_class)
