from abc import ABC, abstractmethod
import pandas as pd
from pathlib import Path
import argparse
from datetime import date, datetime


class Base(ABC):
    SCRIPT_DIR: Path | None = None
    MAPPING_FILE: Path | None = None
    ORIGINAL_SHEET_NAME: str | None = None
    
    def __init__(self, input_file: str, output_file: str = None, shipping_date: datetime = None):
        """
        Initialize with input file path
        
        Args:
            input_file: Path to the input Excel file
            output_file: Optional custom output file path
            date: Optional date for filtering/processing
        """
        self.input_file: str = input_file
        self.output_file: str = output_file or self._generate_output_filename()
        self.shipping_date: datetime | None = shipping_date
        self.mapping_df: pd.DataFrame | None = None
        self.original_df: pd.DataFrame | None = None
        self.main_df: pd.DataFrame | None = None
        self.merged_df: pd.DataFrame | None = None
        self.invoice_df: pd.DataFrame | None = None
        self.canceled_orders_df: pd.DataFrame | None = None
        self.order_sn_unique: int = 0
        self.merge_left : str | None = None
        self.merge_right : str | None = None
    
    def _generate_output_filename(self) -> str:
        """Generate output filename from input filename"""
        if self.input_file.endswith('_output.xlsx'):
            return self.input_file
        return self.input_file.replace('.xlsx', '_output.xlsx')
    
    @abstractmethod
    def load_mapping(self) -> pd.DataFrame:
        """Load item mapping from Excel file"""
        pass
    
    @abstractmethod
    def load_main_df(self) -> pd.DataFrame:
        """Load main data from input file"""
        pass

    def load_canceled_orders(self) -> pd.DataFrame:
        """Load canceled orders from input file if exists"""
        if self.canceled_orders_df is None:
            try :
                self.canceled_orders_df = pd.read_excel(self.input_file, sheet_name='canceled_orders')
            # ValueError occurs when sheet does not exist
            except (ValueError):
                print('No canceled orders sheet found. Continuing without excluding any orders.')
                self.canceled_orders_df = pd.DataFrame(columns=['canceled_orders_sn'])
        return self.canceled_orders_df
    
    def merge_mapping(self) -> pd.DataFrame:
        """Merge main dataframe with mapping"""
        
        if self.merge_left is None or self.merge_right is None:
            raise ValueError("merge_left and merge_right attributes must be set before merging.")
        
        self.merged_df = pd.merge(
            self.main_df, 
            self.mapping_df, 
            left_on=self.merge_left, 
            right_on=self.merge_right, 
            how='left')
        return self.merged_df
    
    @abstractmethod
    def calculate_invoice(self) -> pd.DataFrame:
        """Calculate invoice from merged dataframe"""
        pass
    
    @abstractmethod
    def export_excel(self) -> None:
        """Export invoice to Excel file"""
        pass
    
    def process(self) -> None:
        """Main execution flow - template method pattern"""
        print(f"Starting {self.__class__.__name__}...")
        print(f'Reading input file: {self.input_file}')
        print(f'Processing date: {self.shipping_date.strftime("%Y-%m-%d") if self.shipping_date else "Not specified"}')
        
        # Load data
        self.mapping_df = self.load_mapping()
        self.main_df = self.load_main_df()
        
        # Process
        self.merged_df = self.merge_mapping()
        self.invoice_df = self.calculate_invoice()
        
        print(f'Unique order numbers processed: {self.order_sn_unique}')
        
        # Export
        print(f'Exporting to Excel file: {self.output_file}')
        self.export_excel()
        
        print("Process completed successfully!")
    
    @classmethod
    def create_argument_parser(cls) -> argparse.ArgumentParser:
        """
        Create argument parser with common arguments
        Subclasses can override to add custom arguments
        
        Returns:
            Configured ArgumentParser instance
        """
        platform_name = cls.__module__.split('.')[-2]
        
        parser = argparse.ArgumentParser(
            description=f'Process {platform_name.capitalize()} orders and generate invoice',
            formatter_class=argparse.RawDescriptionHelpFormatter
        )
        
        parser.add_argument(
            'input_file',
            type=str,
            help='Path to input Excel file'
        )
        
        parser.add_argument(
            '-o', '--output',
            type=str,
            dest='output_file',
            help='Path to output Excel file (default: input_file_output.xlsx)'
        )
        
        parser.add_argument(
            '-d', '--shipping_date',
            type=str,
            dest='shipping_date',
            help='Processing date in YYYY-MM-DD format (default: None)'
        )
        
        return parser
    
    @classmethod
    def from_args(cls, args: list[str] = None):
        """
        Factory method to create instance from command line arguments
        
        Args:
            args: Command line arguments (defaults to sys.argv[1:])
        
        Returns:
            Instance of the class
        """
        parser = cls.create_argument_parser()
        parsed_args = parser.parse_args(args)
        
        # Parse date if provided
        shipping_date = None
        if parsed_args.shipping_date:
            try:
                shipping_date = datetime.strptime(parsed_args.shipping_date, '%Y-%m-%d')
            except ValueError:
                parser.error(f'Invalid date format: {parsed_args.shipping_date}. Use YYYY-MM-DD')
        
        return cls(
            input_file=parsed_args.input_file,
            output_file=parsed_args.output_file,
            shipping_date=shipping_date
        )
