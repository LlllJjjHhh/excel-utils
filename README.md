# Excel Utils

A handy Python toolset for batch processing Excel files. Supports merging multiple Excel files, converting formats, cleaning data, and more.

## Features

- ✨ Batch merge multiple Excel files into one
- 🔄 Convert Excel to CSV/JSON formats
- 🧹 Clean empty rows/columns automatically
- 📊 Basic data statistics and analysis
- 🎯 Support for .xlsx, .xls files

## Installation

```bash
git clone https://github.com/<your-username>/excel-utils.git
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
    output_file="merged_output.xlsx"
)
print(f"Merged {result['total_rows']} rows saved to {result['output_file']}")
```

### Convert Excel to CSV

```python
from excel_utils import ExcelConverter

converter = ExcelConverter()
converter.to_csv("input.xlsx", "output.csv")
```

## Example

Check the `examples` folder for more usage examples.

## License

MIT

---

If this project helped you, buy me a coffee ☕

<img src="Qrcode.jpg" alt="Buy Me A Coffee" width="300">
