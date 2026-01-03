# excel-merger-tool
Merge multiple Excel files, remove duplicates, and filter data with one click. Desktop application for batch Excel processing.


# 📊 Excel Merger & Processor Tool

Merge multiple Excel files, remove duplicates, and filter data with one click. Perfect for batch processing and data consolidation.

![Python](https://img.shields.io/badge/Python-3.14%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)

## ✨ Features

- **Batch Processing**: Merge unlimited Excel files at once
- **Duplicate Removal**: Automatically detect and remove duplicate rows
- **Advanced Filtering**: Filter data by column and text
- **Source Tracking**: Add source filename column to track data origin
- **Multiple Formats**: Export to XLSX or CSV
- **Real-time Statistics**: See row counts, column info, and memory usage
- **Activity Log**: Track all processing steps

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/excel-merger-tool.git

# Install dependencies
pip install -r requirements.txt

# Run the tool
python excel_merger_tool.py
```

### Requirements

- Python 3.14+
- pandas
- openpyxl
- tkinter (usually included with Python)

## 📖 How to Use

1. **Add Files**: Click "Add Files" and select multiple Excel files
2. **Set Options**:
   - ✅ Merge all sheets
   - ✅ Remove duplicates
   - ✅ Add source filename column
3. **Optional Filter**: Specify column name and filter text
4. **Process**: Click "Process Files" button
5. **Export**: Save the merged result as XLSX or CSV

### Example Workflow

```
Input:
- sales_jan.xlsx (100 rows)
- sales_feb.xlsx (150 rows)
- sales_mar.xlsx (200 rows)

Output:
- merged_result.xlsx (450 rows, all combined)
```

## 🎯 Use Cases

- Consolidate monthly sales reports
- Merge inventory data from multiple locations
- Combine survey responses
- Aggregate financial data
- Unify customer databases

## ⚠️ IMPORTANT DISCLAIMER

**This tool is provided "AS IS" without warranty of any kind.**

### Data Safety
- **Always backup your original files** before processing
- The tool does not modify source files, but data loss can still occur
- **No guarantee** of 100% accuracy in merging or duplicate detection
- Complex Excel features (formulas, charts, formatting) may be lost

### Technical Limitations
- Large files (100MB+) may cause memory issues or crashes
- Processing time depends on file size and computer specs
- Special characters or date formats may not convert perfectly
- Merged output is data-only (no formulas or formatting preserved)

### No Warranty
- **Not responsible for data loss** or business decisions made with this tool
- No guarantee of successful processing for all file types
- May not work with password-protected or corrupted files
- No liability for any damages or losses

### Your Responsibility
- Verify output data accuracy before use
- Test with small files first
- Keep original files as backup
- Understand that duplicate detection is based on exact row matching

## 🛠️ Troubleshooting

**Problem**: Program freezes during processing
- **Solution**: Files may be too large. Try with fewer files or smaller datasets.

**Problem**: "Error reading file"
- **Solution**: Ensure files are valid Excel format (.xlsx or .xls). Check if files are corrupted or password-protected.

**Problem**: Duplicate removal not working as expected
- **Solution**: Duplicates are detected by comparing entire rows. Partial matches won't be removed.

**Problem**: Date or number formatting looks wrong
- **Solution**: Excel formatting is not preserved. You may need to reformat the output file.

## 💡 Tips

- Start with a few files to test
- Use the filter feature to reduce output size
- Check statistics panel to verify results
- Export to CSV if you only need raw data

## 📝 License

MIT License - Use at your own risk. See LICENSE file for details.

## 🤝 Support

- This is a standalone tool delivered as-is
- For bugs or issues, please open a GitHub issue
- No ongoing support or maintenance included
- File-specific issues cannot be individually debugged

## ⚡ Version

**Version 1.0.0** - Initial Release

### Known Limitations
- Does not preserve Excel formulas
- Does not preserve cell formatting (colors, fonts)
- Does not preserve charts or images
- Maximum recommended: 50 files or 500MB total

---

**Built by Dumok Data Lab**

*Remember: Always backup your data before processing.*
