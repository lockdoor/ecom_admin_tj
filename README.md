# ECOM_ADMIN_TJ
This project help admin to calculate orders to invoices, it decreate numbers of invoices. Before this admin make invoice per order, the program aggregate order group by buyer request TAX

## Version 0.0.1
Support shopee only

## To build package
```
python -m pip install --upgrade build
python -m build
```

## To install package
```
pip install ./dist/ecom_admin_tj-0.0.1.tar.gz
```

## To use package for shopee
```
python -m ecom_admin_tj.shopee [example_file.xlsx] optional[YYYY-MM-DD]
```

รองรับ excel 2 อย่าง
- ปกติ มีการสร้าง excel ก่อนขนส่งมารับ
- ไม่ปกติ ขนส่งมารับสร้าง excel ทีหลัง