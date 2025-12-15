from .tiktok import Tiktok

if __name__ == '__main__':
    try:
        tiktok = Tiktok.from_args()
        tiktok.process()
    except ValueError as e:
        print(f"Error: {e}")
