ORIGINAL_FINANCE_REPORT_FILE=./report_file/my_balance_transaction_report.shopee.20251124_20251130.xlsx

CLEAN_FINANCE_REPORT_FILE=./report_file/shopee_cleaned_finance_report.xlsx

ADMIN_FILE=./admin_file/shopee20251122_output.xlsx

DATE_FROM=2025-11-22

# shopee new cleaned finance report file
snf:
	python -m ecom_admin_tj.shopee.finance.new_report -o$(CLEAN_FINANCE_REPORT_FILE) $(ORIGINAL_FINANCE_REPORT_FILE)

# shopee finance check
sfc:
	python -m ecom_admin_tj.shopee.finance --admin $(ADMIN_FILE) $(CLEAN_FINANCE_REPORT_FILE)

# shopee finance check without admin file
sfcna:
	python -m ecom_admin_tj.shopee.finance $(CLEAN_FINANCE_REPORT_FILE)

# shopee process admin file
sp:
	python -m ecom_admin_tj.shopee -d $(DATE_FROM) $(ADMIN_FILE)
