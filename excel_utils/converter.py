import pandas as pd
import json
from typing import List, Dict, Union, Optional


class ExcelConverter:
    """Convert Excel to other formats (CSV, JSON)"""
    
    @staticmethod
    def to_csv(
        input_file: Union[str, pd.DataFrame],
        output_file: str,
        sheet_name: Optional[Union[str, int]] = 0,
        index: bool = False,
        **kwargs
    ) -> None:
        """Convert Excel to CSV"""
        if isinstance(input_file, str):
            df = pd.read_excel(input_file, sheet_name=sheet_name)
        else:
            df = input_file
            
        df.to_csv(output_file, index=index, **kwargs)
    
    @staticmethod
    def to_json(
        input_file: Union[str, pd.DataFrame],
        output_file: Optional[str] = None,
        orient: str = "records",
        indent: int = 2,
        **kwargs
    ) -> Union[str, None]:
        """Convert Excel to JSON"""
        if isinstance(input_file, str):
            df = pd.read_excel(input_file)
        else:
            df = input_file
            
        json_str = df.to_json(orient=orient, indent=indent, **kwargs)
        
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(json_str)
            return None
        else:
            return json_str
    
    @staticmethod
    def from_csv(csv_file: str, output_excel: str, **kwargs) -> None:
        """Convert CSV back to Excel"""
        df = pd.read_csv(csv_file, **kwargs)
        df.to_excel(output_excel, index=False)
