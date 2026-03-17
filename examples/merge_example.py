#!/usr/bin/env python3
"""
Example: Merge multiple Excel files
"""

import os
import sys
sys.path.append('..')

from excel_utils import ExcelMerger

def main():
    # Create sample directory if it doesn't exist
    sample_dir = "sample_excels"
    os.makedirs(sample_dir, exist_ok=True)
    
    # Check if we have sample files
    if len(os.listdir(sample_dir)) == 0:
        print(f"Put your Excel files into '{sample_dir}' directory and run this example again.")
        return
    
    # Merge them
    merger = ExcelMerger()
    try:
        result = merger.merge_files(
            input_dir=sample_dir,
            output_file="merged_output.xlsx",
            include_source=True
        )
        print(f"✅ Success!")
        print(f"   Files processed: {result['files_processed']}")
        print(f"   Total rows: {result['total_rows']}")
        print(f"   Output saved to: {result['output_file']}")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()
