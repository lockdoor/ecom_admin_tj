ORIGINAL_FINANCE_REPORT_FILE=./report_file/my_balance_transaction_report.shopee.20251201_20251207.xlsx

CLEAN_FINANCE_REPORT_FILE=./report_file/shopee_cleaned_finance_report_20251201_20251207.xlsx

# ADMIN_FILE_PATH=~/Library/CloudStorage/SynologyDrive-ecom_admin/e_commerce_shopee
ADMIN_FILE_PATH=./admin_file

ADMIN_FILE=$(ADMIN_FILE_PATH)/shopee20251117_output.xlsx

DATE_FROM=2025-11-04
DATE_TO=2025-12-31

# Helper function to extract date from filename and process admin files
define process_admin_files
	@for file in $(ADMIN_FILE_PATH)/*_output.xlsx; do \
		date=$$(basename "$$file" _output.xlsx | grep -oE '[0-9]{8}$$'); \
		formatted_date=$$(echo $$date | sed 's/\([0-9]\{4\}\)\([0-9]\{2\}\)\([0-9]\{2\}\)/\1-\2-\3/'); \
		if [ -n "$(1)" ] && [ "$$formatted_date" \< "$(1)" ]; then \
			continue; \
		fi; \
		if [ -n "$(2)" ] && [ "$$formatted_date" \> "$(2)" ]; then \
			continue; \
		fi; \
		echo "Processing file: $$file (Date: $$formatted_date)"; \
		python -m ecom_admin_tj.shopee -d $$formatted_date "$$file"; \
	done
endef

# shopee new cleaned finance report file
snf:
	python -m ecom_admin_tj.shopee.finance.new_report -o$(CLEAN_FINANCE_REPORT_FILE) $(ORIGINAL_FINANCE_REPORT_FILE)

# shopee finance check
sfc:
	python -m ecom_admin_tj.shopee.finance --admin $(ADMIN_FILE) $(CLEAN_FINANCE_REPORT_FILE)

# shopee finance check with multiple admin files by date range
# Usage: make sfcm DATE_FROM=2025-11-01 DATE_TO=2025-11-30
sfcm:
	python -m ecom_admin_tj.shopee.finance -d $(ADMIN_FILE_PATH) --date-from $(DATE_FROM) --date-to $(DATE_TO) $(CLEAN_FINANCE_REPORT_FILE)

# shopee finance check with multiple admin files and allow replace
# Usage: make sfcmr DATE_FROM=2025-11-01 DATE_TO=2025-11-30
sfcmr:
	python -m ecom_admin_tj.shopee.finance -d $(ADMIN_FILE_PATH) --date-from $(DATE_FROM) --date-to $(DATE_TO) --allow-replace $(CLEAN_FINANCE_REPORT_FILE)

# shopee finance check without admin file
sfcna:
	python -m ecom_admin_tj.shopee.finance $(CLEAN_FINANCE_REPORT_FILE)

# shopee process admin file
sp:
	python -m ecom_admin_tj.shopee -d $(DATE_FROM) '$(ADMIN_FILE)'

# shopee process all admin files
spm:
	$(call process_admin_files,$(DATE_FROM),$(DATE_TO))

# shopee process all admin files with custom date range
# Usage: make spmd DATE_FROM=2025-11-01 DATE_TO=2025-11-30
spmd:
	$(call process_admin_files,$(DATE_FROM),$(DATE_TO))
