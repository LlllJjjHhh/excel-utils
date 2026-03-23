import pandas as pd
from pathlib import Path
from typing import List, Dict, Union


class ExcelMerger:
    """Merge multiple Excel files into one"""
    
    def __init__(self):
        self.supported_exts = ['.xlsx', '.xls']
    
    def find_excel_files(self, input_dir: Union[str, Path]) -> List[Path]:
        """Find all Excel files in directory"""
        input_dir = Path(input_dir)
        excel_files = []
        
        for ext in self.supported_exts:
            excel_files.extend(input_dir.glob(f"*{ext}"))
            excel_files.extend(input_dir.glob(f"*{ext.upper()}"))
            
        return sorted(excel_files)
    
    def merge_files(
        self,
        input_dir: Union[str, Path],
        output_file: Union[str, Path],
        sheet_name: str = "Merged",
        include_source: bool = False
    ) -> Dict:
        """
        Merge all Excel files from a directory
        
        Args:
            input_dir: Directory containing Excel files
            output_file: Output merged Excel file
            sheet_name: Sheet name for the output
            include_source: Whether to include a column for source filename
        
        Returns:
            Dictionary with merge statistics
        """
        excel_files = self.find_excel_files(input_dir)
        
        if not excel_files:
            raise ValueError(f"No Excel files found in {input_dir}")
        
        all_dfs = []
        total_rows = 0
        
        for file in excel_files:
            df = pd.read_excel(file)
            
            if include_source:
                df['_source_file'] = file.name
                
            all_dfs.append(df)
            total_rows += len(df)
        
        merged_df = pd.concat(all_dfs, ignore_index=True)
        merged_df.to_excel(output_file, index=False, sheet_name=sheet_name)
        
        return {
            'files_processed': len(excel_files),
            'total_rows': total_rows,
            'output_file': str(output_file)
        }
