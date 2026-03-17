"""Basic tests for excel-utils"""

import pandas as pd
from excel_utils import ExcelMerger, ExcelConverter, DataCleaner


def test_merger_init():
    merger = ExcelMerger()
    assert merger is not None
    assert '.xlsx' in merger.supported_exts


def test_converter_to_json():
    # Create a simple dataframe
    df = pd.DataFrame({
        'name': ['Alice', 'Bob', 'Charlie'],
        'age': [25, 30, 35]
    })
    
    converter = ExcelConverter()
    json_str = converter.to_json(df)
    assert len(json_str) > 0
    assert 'Alice' in json_str


def test_data_cleaner():
    # Create dataframe with empty rows
    df = pd.DataFrame({
        'A': [1, None, 3, None, 5],
        'B': [None, None, None, None, None]
    })
    
    cleaner = DataCleaner(df)
    cleaner.remove_empty_columns(threshold=1)
    cleaned = cleaner.get()
    
    # Should remove column B which is all empty
    assert list(cleaned.columns) == ['A']
    assert len(cleaned.dropna()) == 3
