from .shopee import Shopee

if __name__ == "__main__":
    try:
        shopee = Shopee.from_args()
        shopee.process()
    except ValueError as e:
        print(f"Error: {e}")
