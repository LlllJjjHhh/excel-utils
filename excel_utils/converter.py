import pandas as pd
import json
import os
from pathlib import Path
from typing import List, Dict, Union, Optional


class ExcelConverter:
    """Convert Excel to other formats (CSV, JSON)"""
    
    def __init__(self):
        self.supported_exts = ['.xlsx', '.xls']
    
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
    
    def batch_convert_to_csv(
        self,
        input_dir: Union[str, Path],
        output_dir: Union[str, Path],
        keep_original_name: bool = True,
        **kwargs
    ) -> List[Dict]:
        """Batch convert all Excel files in a directory to CSV"""
        input_dir = Path(input_dir)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        results = []
        
        for ext in self.supported_exts:
            for excel_file in input_dir.glob(f"*{ext}"):
                if keep_original_name:
                    output_file = output_dir / f"{excel_file.stem}.csv"
                else:
                    output_file = output_dir / f"{excel_file.name}.csv"
                
                try:
                    self.to_csv(excel_file, output_file, **kwargs)
                    results.append({
                        'input': str(excel_file),
                        'output': str(output_file),
                        'status': 'success',
                        'error': None
                    })
                except Exception as e:
                    results.append({
                        'input': str(excel_file),
                        'output': str(output_file),
                        'status': 'error',
                        'error': str(e)
                    })
        
        return results
    
    def batch_convert_to_json(
        self,
        input_dir: Union[str, Path],
        output_dir: Union[str, Path],
        keep_original_name: bool = True,
        **kwargs
    ) -> List[Dict]:
        """Batch convert all Excel files in a directory to JSON"""
        input_dir = Path(input_dir)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        results = []
        
        for ext in self.supported_exts:
            for excel_file in input_dir.glob(f"*{ext}"):
                if keep_original_name:
                    output_file = output_dir / f"{excel_file.stem}.json"
                else:
                    output_file = output_dir / f"{excel_file.name}.json"
                
                try:
                    self.to_json(excel_file, output_file, **kwargs)
                    results.append({
                        'input': str(excel_file),
                        'output': str(output_file),
                        'status': 'success',
                        'error': None
                    })
                except Exception as e:
                    results.append({
                        'input': str(excel_file),
                        'output': str(output_file),
                        'status': 'error',
                        'error': str(e)
                    })
        
        return results
