from datareadr.datareadr import datareadr
import warnings

def load_ui(path):
    datareadr().__load_ui__(path)

warnings.filterwarnings("ignore",
                                "Pandas doesn't allow columns to be created via a new attribute name - see https://pandas.pydata.org/pandas-docs/stable/indexing.html#attribute-access",
                                UserWarning)
