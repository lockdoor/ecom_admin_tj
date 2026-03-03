"""
Shopee Admin Processing Entry Point

Processes Shopee orders and generates invoices, finance summaries,
and stock deduction reports.
"""
from ..common.cli.platform_runner import PlatformRunner
from .shopee import Shopee

if __name__ == "__main__":
    PlatformRunner.run(Shopee)
