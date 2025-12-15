from .lazada import Lazada
import warnings
import openpyxl

if __name__ == "__main__":
    warnings.filterwarnings("ignore", category=UserWarning)
    try:
        lazada = Lazada.from_args()
        lazada.process()
    except ValueError as e:
        print(f"Error: {e}")
