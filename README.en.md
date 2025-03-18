# Excel Sheet Segmentation and Consolidation Tool

#### Project Overview
An automation tool based on xlwings library, providing following core features:
- Batch processing Excel files in specified directory
- Intelligent worksheet structure recognition
- Merge data by specified sheet index
- Automatic source file identification
- Generate timestamped consolidated files

#### System Architecture
- Core library: xlwings 0.30.12+
- Language: Python 3.8+
- Dependency management: pip
- Directory structure:
  ```
  ├── data/         # Raw data storage
  ├── output/       # Result output 
  ├── main_v1.5.py  # Main program
  └── utils.py      # Path handling utilities
  ```

#### Installation Guide
1. Requirements
   - Windows 10/11
   - Microsoft Office 2016+
   - Python 3.8+

2. Dependency Installation
   ```powershell
   pip install xlwings==0.30.12
   pip install pywin32==306
   ```

3. Configuration
   - Create data folder in project root for input files
   - Ensure all Excel files are in xlsx format

#### User Manual
1. File Preparation
   - Place Excel files in data directory
   - Maintain consistent column structure across files

2. Execution
   ```powershell
   python main_v1.5.py
   ```

3. Workflow
   - Program displays detected worksheet list
   - Input target sheet index (numeric)
   - Wait for completion notification (~1 file/sec)
   - Result file generated in project root

4. Output Example
   ```
   Accounts_Payable_Summary_20240318.xlsx
   └── Sheet1
       ├── A1:H100  Source data
       └── Column I Source file tags
   ```

#### Important Notes
1. File Specifications
   - Single file recommended <500k rows
   - Total rows after merging < Excel limit (1,048,576 rows)
   - Use English filenames recommended

2. Troubleshooting
   - Check following if program interrupts:
     - Excel files not being used by other programs
     - Correct sheet index input
     - data directory exists and not empty

3. Performance Optimization
   - Close other Excel instances when processing large files
   - Contact developer for enterprise version for big data processing

#### Version History
v1.5 Updates:
- Added progress indicators
- Optimized path handling
- Fixed multi-sheet recognition issues
- Enhanced input validation
