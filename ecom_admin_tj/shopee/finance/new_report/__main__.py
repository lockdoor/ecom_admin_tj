from ..shopee_finance import ShopeeFinanceMixin
import argparse

def create_argument_parser() -> argparse.ArgumentParser:
    
    parser = argparse.ArgumentParser(
        description='Create new Shopee finance clean report',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    parser.add_argument(
        'original_file',
        type=str,
        help='Path to the original Shopee finance report file',
    )

    parser.add_argument(
        '-o', '--output',
        type=str,
        help='Path to save the new cleaned finance report',
        dest='output_file',
        required=False
    )

    return parser

def main():
    parser = create_argument_parser()
    parsed_args = parser.parse_args()
    
    ShopeeFinanceMixin.make_finance_report(
        original_report_file=parsed_args.original_file,
        output_file=parsed_args.output_file
    )

if __name__ == "__main__":
    main()
