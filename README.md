# Excel Utils

A handy Python toolset for batch processing Excel files. Supports merging multiple Excel files, converting formats, cleaning data, and more.

## Features

- ✨ Batch merge multiple Excel files into one
- 🔄 Convert Excel to CSV/JSON formats (and vice versa)
- 🔄 Batch convert multiple files at once
- 🧹 Clean empty rows/columns, remove duplicates, trim whitespace automatically
- 📊 Comprehensive data statistics and analysis
- 🔍 Outlier detection
- 📈 Correlation analysis
- 📊 Group by analysis
- 📄 Export statistics to HTML report
- 🎯 Support for .xlsx, .xls files

## Installation

```bash
git clone https://github.com/LlllJjjHhh/excel-utils.git
cd excel-utils
pip install -r requirements.txt
```

## Usage

### Merge multiple Excel files

```python
from excel_utils import ExcelMerger

merger = ExcelMerger()
result = merger.merge_files(
    input_dir="./excels",
    output_file="merged_output.xlsx",
    include_source=True  # Add source filename column
)
print(f"Merged {result['total_rows']} rows saved to {result['output_file']}")
print(f"Processed {result['files_processed']} files")
```

### Convert Excel to CSV

```python
from excel_utils import ExcelConverter

converter = ExcelConverter()
# Single file conversion
converter.to_csv("input.xlsx", "output.csv")

# Batch convert all Excel files in a directory
results = converter.batch_convert_to_csv(
    input_dir="./excels",
    output_dir="./csv_output"
)
for result in results:
    print(f"{result['status']}: {result['input']}")
```

### Convert Excel to JSON

```python
from excel_utils import ExcelConverter

converter = ExcelConverter()
# Single file conversion
converter.to_json("input.xlsx", "output.json")

# Batch convert
results = converter.batch_convert_to_json(
    input_dir="./excels",
    output_dir="./json_output"
)

# Convert CSV back to Excel
ExcelConverter.from_csv("input.csv", "output.xlsx")
```

### Clean Data

```python
import pandas as pd
from excel_utils import DataCleaner

df = pd.read_excel("input.xlsx")
cleaner = DataCleaner(df)

cleaned_df = (cleaner
    .remove_empty_rows(threshold=1)
    .remove_empty_columns(threshold=1)
    .drop_duplicates()
    .fill_missing("N/A")
    .trim_whitespace()
    .get()
)

cleaner.save("cleaned_output.xlsx")
```

### Data Statistics and Analysis

```python
import pandas as pd
from excel_utils import DataStats

df = pd.read_excel("data.xlsx")
stats = DataStats(df)

# Basic summary
summary = stats.basic_summary()
print(summary)

# Numeric column statistics
numeric_stats = stats.numeric_stats()
print(numeric_stats)

# Frequency analysis for categorical columns
freq = stats.frequency_analysis("category_column", top_n=10)
print(freq)

# Correlation matrix
corr = stats.correlation_matrix()
print(corr)

# Find outliers
outliers_iqr = stats.find_outliers("price", method="iqr")
outliers_zscore = stats.find_outliers("price", method="zscore")
print(f"Found {len(outliers_iqr)} outliers using IQR method")

# Group by analysis
grouped = stats.group_by_analysis("category", "price", agg_func="mean")
print(grouped)

# Export HTML report
stats.export_report("stats_report.html")
```

## Example

Check the `examples` folder for more usage examples.

## Requirements

- pandas
- openpyxl
- xlrd
- numpy

## License

MIT

---

If this project helped you, buy me a coffee ☕

<img src="Qrcode.jpg" alt="Buy Me A Coffee" width="300">
