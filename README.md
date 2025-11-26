# ECOM_ADMIN_TJ
This project help admin to calculate orders to invoices, it decreate numbers of invoices. Before this admin make invoice per order, the program aggregate order group by buyer request TAX

## Version 0.0.1
Support shopee only

## Version 0.0.2
Support tiktok

## Version 0.0.3
Support lazada

## To build package
```
python -m pip install --upgrade build
python -m build
```

## To install package
```
pip install ./dist/ecom_admin_tj-0.0.1.tar.gz
```

## install from github
```
pip install git+https://github.com/lockdoor/ecom_admin_tj.git
pip install git+https://github.com/lockdoor/ecom_admin_tj.git@v0.0.3
```

## To use package for shopee
```
python -m ecom_admin_tj.shopee [example_file.xlsx] optional[YYYY-MM-DD]
```

รองรับ excel 2 อย่าง
- ปกติ มีการสร้าง excel ก่อนขนส่งมารับ
- ไม่ปกติ ขนส่งมารับสร้าง excel ทีหลัง

## Requirement
- ecom_admin_tj.common.stock_items.csv
- ecom_admin_tj.shopee.shopee_item_mapping.xlsx
- ecom_admin_tj.tiktok.tiktok_item_mapping.xlsx
- ecom_admin_tj.tiktok.mapping.tiktok_product_local_products_list.json
- ecom_admin_tj.lazada.lazada_item_mapping.xlsx
- ecom_admin_tj.lazada.mapping.lazada_products.xlsx

## Windows Lib path
```
Users\[username]\AppData\Local\Python\[python version]\Lib\site-packages
```
