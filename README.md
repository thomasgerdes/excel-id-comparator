# Excel ID-Based Comparator

A robust Python tool for comparing Excel files based on unique identifiers, highlighting actual data changes rather than positional differences. Perfect for tracking changes in datasets where rows may be inserted, deleted, or reordered.

**Author**: Thomas Gerdes  
**Version**: 1.0.0  
**License**: MIT

## Features

- **ID-Based Comparison**: Compares records by unique identifiers, not by row position
- **Smart Auto-Detection**: Automatically detects sheet names and ID columns with fallback options
- **Visual Change Highlighting**: 
  - 🔴 Red text for modified values (with comments showing previous values)
  - 🟢 Green highlighting for new records
  - 🟠 Separate sheet for deleted records
- **Robust Error Handling**: Continues processing even when individual cells have issues
- **Cross-Platform**: Works on Windows, macOS, Linux, Google Colab, and Jupyter Notebook
- **Flexible Configuration**: Customizable sheet names, ID columns, case sensitivity
- **Detailed Reporting**: Comprehensive statistics and change summaries
- **Command Line Interface**: Easy to use from terminal or integrate into workflows

## Requirements

- Python 3.7+
- pandas
- openpyxl
- numpy

## Installation & Usage

### Method 1: Direct Execution (Recommended)

```bash
# Clone or download the repository
git clone https://github.com/thomasgerdes/excel-id-comparator.git
cd excel-id-comparator

# Install dependencies
pip install pandas openpyxl numpy

# Run the tool
python excel_id_comparator.py old_file.xlsx new_file.xlsx
```

### Method 2: Google Colab (No Installation Required)

Perfect for users who prefer not to install Python locally:

1. Open [Google Colab](https://colab.research.google.com)
2. Create a new notebook
3. Copy the entire `excel_id_comparator.py` code into a cell
4. Run the cell - it automatically detects Colab and provides file upload interface
5. Upload your Excel files and follow the prompts

### Method 3: Jupyter Notebook

```bash
# Install Jupyter if not already installed
pip install jupyter

# Start Jupyter
jupyter notebook

# Create new notebook and run the comparison script
```

## Usage Examples

```bash
# Simple comparison (auto-detection)
python excel_id_comparator.py original.xlsx updated.xlsx

# With custom output and settings
python excel_id_comparator.py file1.xlsx file2.xlsx -o report.xlsx --sheet "Data" --id-column "A"

# Case-insensitive comparison
python excel_id_comparator.py file1.xlsx file2.xlsx --case-insensitive
```

### Python API
```python
from excel_id_comparator import ExcelIDComparator

# Basic usage
comparator = ExcelIDComparator()
report_path = comparator.compare_files('old_file.xlsx', 'new_file.xlsx')

# With configuration
config = {'sheet_name': 'Data', 'id_column': 'A', 'case_sensitive': False}
comparator = ExcelIDComparator(config)
report_path = comparator.compare_files('file1.xlsx', 'file2.xlsx')
```

## Output

The tool generates an Excel report with:
- **Main Sheet**: Your data with red text for changes, green highlighting for new records
- **Summary Sheet**: Statistics and configuration details  
- **Deleted Records Sheet**: Removed records (if any)

## Use Cases

Data auditing, version control, quality assurance, compliance documentation, collaboration reviews, database migrations, master data management.

## Version History

- **1.0.0**: Initial release with core comparison functionality
  - ID-based comparison engine
  - Auto-detection of sheet structure
  - Visual change highlighting
  - Command line interface
  - Comprehensive error handling

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

**Thomas Gerdes**
- GitHub: [@thomasgerdes](https://github.com/thomasgerdes)
- Website: [https://thomasgerdes.de](https://thomasgerdes.de)

## Development

This tool was created with AI assistance and follows Python best practices and coding standards. The code is designed to be maintainable, well-documented, and cross-platform compatible.

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.
