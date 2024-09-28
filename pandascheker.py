try:
    import pandas as pd
    print("Pandas is installed.")
    print("Pandas version:", pd.__version__)
except ImportError:
    print("Pandas is not installed.")
