import sys

def cmd_shopee():
    from ecom_admin_tj.shopee.shopee_script import main as shopee_main
    # ปรับ sys.argv ให้ shopee_script อ่านได้ถูกต้อง
    # เปลี่ยนจาก ['manage.py', 'shopee', 'file.xlsx'] 
    # เป็น ['shopee_script.py', 'file.xlsx']
    sys.argv = ['shopee_script.py'] + sys.argv[2:]
    shopee_main()

def print_help():
    help_text = """
    Usage: python manage.py <command>
    
    Available commands:
        shopee     Run the Shopee management script.
        help       Show this help message.
    """
    print(help_text)

if __name__ == "__main__":
    """Main entry point for manage.py script."""
    if len(sys.argv) < 2:
        print_help()
        sys.exit(1)
        
    commands = {
        "shopee": cmd_shopee,
        "help": print_help,
    }
    
    command = sys.argv[1].lower()
    
    if command not in commands:
        print(f"Unknown command: {command}")
        print_help()
        sys.exit(1)
        
    try:
        commands[command]()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(130)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
