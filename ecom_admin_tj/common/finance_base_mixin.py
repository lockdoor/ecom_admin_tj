"""
Base Finance Operations Mixin

Provides common finance calculation and formatting functionality
that can be used across all platforms (Shopee, Lazada, Tiktok).
"""
from abc import ABC, abstractmethod
import pandas as pd
import numpy as np
from typing import Dict, List, Callable
from .excel_format_mixin import ExcelFormatMixin, Worksheet


class FinanceBaseMixin(ExcelFormatMixin):
    """
    Base mixin for common finance operations
    
    Provides:
    - Generic finance summary calculation
    - Excel export formatting for finance sheets
    - Column width configuration
    - Footer totals generation
    """
    
    @abstractmethod
    def _get_finance_column_mapping(self) -> Dict[str, str]:
        """
        Get platform-specific column name mapping
        
        Each platform must implement this to map standard names to their columns.
        
        Standard keys should include:
        - 'order_id': Order/transaction ID column
        - 'total': Total amount column (calculated)
        - Other platform-specific price columns
        
        Returns:
            Dict mapping standard names to actual column names
            
        Example:
            {
                'order_id': 'หมายเลขคำสั่งซื้อ',
                'net_price': 'ราคาขายสุทธิ',
                'buyer_shipping': 'ค่าจัดส่งที่ชำระโดยผู้ซื้อ',
            }
        """
        raise NotImplementedError("Subclasses must implement _get_finance_column_mapping()")
    
    def calculate_finance_summary(
        self,
        df: pd.DataFrame,
        groupby_col: str,
        agg_dict: Dict[str, str],
        total_formula: Callable[[pd.DataFrame], pd.Series],
        total_col_name: str = 'รวม'
    ) -> pd.DataFrame:
        """
        Generic finance summary calculation
        
        Args:
            df: Source DataFrame
            groupby_col: Column to group by (usually order ID)
            agg_dict: Aggregation dictionary for groupby
            total_formula: Function to calculate total column
            total_col_name: Name for the total column
            
        Returns:
            DataFrame with finance summary
        """
        finance_df = df.groupby(groupby_col, sort=False).agg(agg_dict).reset_index()
        finance_df[total_col_name] = total_formula(finance_df)
        
        return finance_df
    
    def add_finance_footer(
        self,
        finance_df: pd.DataFrame,
        order_col: str,
        numeric_cols: List[str],
        footer_label: str = 'TOTAL'
    ) -> pd.DataFrame:
        """
        Add footer row with totals to finance DataFrame
        
        Args:
            finance_df: Finance DataFrame
            order_col: Order ID column name
            numeric_cols: List of numeric columns to sum
            footer_label: Label for footer row
            
        Returns:
            DataFrame with footer row added
        """
        total_row = {order_col: footer_label}
        
        for col in numeric_cols:
            if col in finance_df.columns:
                total_row[col] = finance_df[col].sum()
        
        finance_df.loc[len(finance_df)] = total_row
        return finance_df
    
    def _format_finance_sheet(
        self,
        sheet: Worksheet,
        column_widths: Dict[str, int],
        data_row_count: int,
        num_columns: int
    ) -> None:
        """
        Apply standard formatting to finance sheet
        
        Args:
            sheet: Excel worksheet
            column_widths: Dict mapping column letters to widths
            data_row_count: Number of data rows (excluding header)
            num_columns: Total number of columns
        """
        # Set column widths
        for col_letter, width in column_widths.items():
            sheet.column_dimensions[col_letter].width = width
        
        # Format header
        self._formating_header(sheet)
        
        # Format body (if there are data rows)
        if data_row_count > 0:
            self._formatting_body(
                sheet=sheet,
                start_row=2,
                end_row=data_row_count + 1,  # +1 for header
                start_col=1,
                end_col=num_columns
            )
        
        # Format footer (last row)
        self._formatting_footer(sheet=sheet, footer_row=data_row_count + 2)  # +2 for header + footer
    
    def export_finance_sheet(
        self,
        writer,
        finance_df: pd.DataFrame,
        sheet_name: str,
        column_widths: Dict[str, int]
    ) -> None:
        """
        Export finance DataFrame to Excel with formatting
        
        Args:
            writer: ExcelWriter instance
            finance_df: Finance DataFrame to export
            sheet_name: Name for the sheet
            column_widths: Dict mapping column letters to widths
        """
        finance_df.to_excel(writer, sheet_name=sheet_name, index=False)
        finance_sheet: Worksheet = writer.sheets[sheet_name]
        
        data_row_count = len(finance_df) - 1  # Exclude footer from data rows
        num_columns = len(finance_df.columns)
        
        self._format_finance_sheet(
            sheet=finance_sheet,
            column_widths=column_widths,
            data_row_count=data_row_count,
            num_columns=num_columns
        )
