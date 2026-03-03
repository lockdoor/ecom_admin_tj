"""
Generic CLI for cleaning platform finance reports

This module provides a reusable command-line interface for cleaning
and formatting finance reports from various e-commerce platforms.

Example:
    >>> from ecom_admin_tj.common.cli.report_cleaner_cli import ReportCleanerCLI
    >>> from ecom_admin_tj.shopee.finance.shopee_finance import ShopeeFinanceMixin
    >>> 
    >>> cli = ReportCleanerCLI(
    ...     platform_name='Shopee',
    ...     finance_mixin_class=ShopeeFinanceMixin
    ... )
    >>> cli.run()
"""
import sys
import argparse
from pathlib import Path
from typing import Type


class ReportCleanerCLI:
    """
    Generic CLI for cleaning platform finance reports
    
    This class provides a consistent command-line interface for cleaning
    and formatting finance reports across different e-commerce platforms.
    
    The finance_mixin_class must implement:
        - make_finance_report(original_report_file, output_file, auto_rename) -> str
    
    Attributes:
        platform_name: Name of the platform (e.g., 'Shopee', 'TikTok', 'Lazada')
        finance_mixin_class: Finance mixin class with make_finance_report() method
        parser: ArgumentParser instance for parsing command-line arguments
    
    Usage:
        cli = ReportCleanerCLI(
            platform_name='Shopee',
            finance_mixin_class=ShopeeFinanceMixin
        )
        exit_code = cli.run()
        sys.exit(exit_code)
    """
    
    def __init__(self, platform_name: str, finance_mixin_class: Type):
        """
        Initialize Report Cleaner CLI
        
        Args:
            platform_name: Name of platform (e.g., 'Shopee', 'TikTok', 'Lazada')
            finance_mixin_class: Finance mixin class with make_finance_report() classmethod
        """
        self.platform_name = platform_name
        self.finance_mixin_class = finance_mixin_class
        self.parser = self._create_parser()
    
    def _create_parser(self) -> argparse.ArgumentParser:
        """
        Create argument parser with platform-specific help text
        
        Returns:
            Configured ArgumentParser instance
        """
        parser = argparse.ArgumentParser(
            description=f'Create cleaned {self.platform_name} finance report',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog=f"""
Examples:
  # Basic usage - auto-generate output filename
  python -m ecom_admin_tj.{self.platform_name.lower()}.finance.new_report input_report.xlsx
  
  # Specify custom output filename
  python -m ecom_admin_tj.{self.platform_name.lower()}.finance.new_report input_report.xlsx -o cleaned.xlsx
  
  # Overwrite existing file (no auto-rename)
  python -m ecom_admin_tj.{self.platform_name.lower()}.finance.new_report input_report.xlsx --no-auto-rename

Output:
  Creates a cleaned Excel file with properly formatted columns and headers.
  If output filename is not specified, it will be auto-generated based on
  the input filename's date range pattern.
"""
        )
        
        parser.add_argument(
            'original_file',
            type=str,
            help=f'Path to original {self.platform_name} finance report file'
        )
        
        parser.add_argument(
            '-o', '--output',
            type=str,
            dest='output_file',
            help='Path to save the cleaned report (optional, auto-generated if not provided)',
            required=False
        )
        
        parser.add_argument(
            '--no-auto-rename',
            action='store_false',
            dest='auto_rename',
            help='Disable auto-rename when output file exists (will overwrite)',
            default=True
        )
        
        return parser
    
    def run(self) -> int:
        """
        Run the report cleaner CLI
        
        Parses command-line arguments, validates input, and calls the
        finance mixin's make_finance_report method.
        
        Returns:
            Exit code (0 for success, 1 for error)
        """
        try:
            args = self.parser.parse_args()
            
            # Validate input file exists
            input_path = Path(args.original_file)
            if not input_path.exists():
                print(f"❌ Error: Input file not found: {args.original_file}")
                return 1
            
            # Display processing information
            print(f"🧹 Cleaning {self.platform_name} finance report...")
            print(f"📄 Input: {args.original_file}")
            if args.output_file:
                print(f"📄 Output: {args.output_file}")
            else:
                print(f"📄 Output: Auto-generated based on input filename")
            
            # Call make_finance_report classmethod
            output_path = self.finance_mixin_class.make_finance_report(
                original_report_file=args.original_file,
                output_file=args.output_file,
                auto_rename=args.auto_rename
            )
            
            print(f"✅ Successfully created cleaned report: {output_path}")
            return 0
            
        except FileNotFoundError as e:
            print(f"❌ Error: File not found - {e}")
            return 1
        except KeyboardInterrupt:
            print(f"\n⚠️  Operation cancelled by user")
            return 1
        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
            return 1
