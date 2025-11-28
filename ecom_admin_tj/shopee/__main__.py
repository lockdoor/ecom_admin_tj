from .shopee import Shopee

if __name__ == "__main__":
    shopee = Shopee.from_args()
    shopee.process()
