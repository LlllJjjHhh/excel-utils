import pandas as pd
from typing import List, Optional


class DataCleaner:
    """Clean Excel data - remove empty rows/columns, drop duplicates, etc."""
    
    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
    
    def remove_empty_rows(self, threshold: int = 1) -> 'DataCleaner':
        """Remove rows that are all or mostly empty"""
        self.df = self.df.dropna(thresh=threshold)
        return self
    
    def remove_empty_columns(self, threshold: int = 1) -> 'DataCleaner':
        """Remove columns that are all or mostly empty"""
        self.df = self.df.dropna(axis=1, thresh=threshold)
        return self
    
    def drop_duplicates(self, subset: Optional[List[str]] = None) -> 'DataCleaner':
        """Drop duplicate rows"""
        self.df = self.df.drop_duplicates(subset=subset)
        return self
    
    def fill_missing(self, value: str = '') -> 'DataCleaner':
        """Fill missing values"""
        self.df = self.df.fillna(value)
        return self
    
    def trim_whitespace(self) -> 'DataCleaner':
        """Trim whitespace from string columns"""
        for col in self.df.select_dtypes(include=['object']).columns:
            self.df[col] = self.df[col].str.strip()
        return self
    
    def get(self) -> pd.DataFrame:
        """Get the cleaned DataFrame"""
        return self.df
    
    def save(self, output_file: str, **kwargs) -> None:
        """Save cleaned data to Excel"""
        self.df.to_excel(output_file, index=False, **kwargs)
