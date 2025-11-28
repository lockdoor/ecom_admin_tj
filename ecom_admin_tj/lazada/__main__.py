from .lazada import Lazada

if __name__ == "__main__":
    lazada = Lazada.from_args()
    lazada.process()
