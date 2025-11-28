from .tiktok import Tiktok

if __name__ == '__main__':
    tiktok = Tiktok.from_args()
    tiktok.process()