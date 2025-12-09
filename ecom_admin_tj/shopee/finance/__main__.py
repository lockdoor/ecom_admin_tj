import argparse
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
        '--inplace',
        action='store_true',
        help='Update the reported file in place after reconciliation',
        dest='inplace',
        default=True,
        required=False
    )
    return parser

def main():
    parser = create_argument_parser()
    parsed_args = parser.parse_args()
    
    ShopeeFinanceMixin.finance_check(
        reported_file=parsed_args.report_file,
        admin_file=parsed_args.admin_file,
        inplace=parsed_args.inplace
    )
    

if __name__ == "__main__":
    main()
