# Data Cleaning & Reporting Automation - Setup Guide

## System Requirements

- **Python**: 3.8 or higher
- **OS**: Windows, macOS, or Linux
- **RAM**: Minimum 2GB (4GB recommended)
- **Disk Space**: ~300MB for dependencies and outputs

## Step-by-Step Installation

### Step 1: Clone/Download Repository

```bash
git clone https://github.com/yourusername/data-cleaning-automation.git
cd data-cleaning-automation
```

### Step 2: Create Virtual Environment

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**macOS/Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### Step 3: Upgrade pip

```bash
python -m pip install --upgrade pip
```

### Step 4: Install Dependencies

```bash
pip install -r requirements.txt
```

**Installed Packages:**
- pandas - Data manipulation
- numpy - Numerical operations
- matplotlib - Visualization
- seaborn - Statistical graphics
- openpyxl - Excel file handling
- scikit-learn - Machine learning utilities

### Step 5: Verify Installation

```bash
python data_cleaning_automation.py
```

**Expected Output:**
```
======================================================================
DATA CLEANING & REPORTING AUTOMATION PIPELINE
======================================================================

📊 Step 1: Generating sample dataset...
   Generated 525 records

🔍 Step 2: Assessing data quality...
   ✗ Missing values: 300
   ✗ Duplicates: 25
   
...

✅ DATA CLEANING COMPLETE!
📁 OUTPUT FILES GENERATED:
   ✓ data_quality_dashboard.png
   ✓ Data_Cleaning_Report.xlsx
   ✓ cleaned_data.csv
   ✓ Detailed_Report.txt
```

## Project Structure

After installation:
```
data-cleaning-automation/
├── data_cleaning_automation.py   # Main script
├── requirements.txt              # Dependencies
├── README.md                     # Documentation
├── outputs/                      # Generated files
│   ├── data_quality_dashboard.png
│   ├── Data_Cleaning_Report.xlsx
│   ├── cleaned_data.csv
│   └── Detailed_Report.txt
└── venv/                        # Virtual environment
```

## Running the Project

### Basic Execution

```bash
python data_cleaning_automation.py
```

### With Output Logging

```bash
python data_cleaning_automation.py > output.log 2>&1
```

### Check Results

Generated files appear in current directory:
- `data_quality_dashboard.png` - 9-panel visualization
- `Data_Cleaning_Report.xlsx` - Excel report
- `cleaned_data.csv` - Clean dataset
- `Detailed_Report.txt` - Text report

## Troubleshooting

### Issue: Python Not Found

**Windows:**
```bash
# Try full path
C:\Python39\python.exe data_cleaning_automation.py

# Or use py launcher
py -3 data_cleaning_automation.py
```

**macOS/Linux:**
```bash
python3 data_cleaning_automation.py
```

### Issue: Module Not Found

**Activate Virtual Environment:**
```bash
# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

**Reinstall Dependencies:**
```bash
pip install --force-reinstall -r requirements.txt
```

### Issue: Permission Denied (Linux/macOS)

```bash
chmod +x data_cleaning_automation.py
python3 data_cleaning_automation.py
```

### Issue: Out of Memory

- The script uses ~200MB RAM
- Close unnecessary applications
- For larger datasets, modify `n_records` parameter in code

## Customizing the Project

### Modify Input Data Size

Edit `data_cleaning_automation.py`:
```python
# Change this line (line ~575):
df_original = generate_dirty_dataset(n_records=1000)  # Was 500
```

### Add Custom Cleaning Rules

In `DataQualityReporter.clean_data()` method, add before the return statement:
```python
# Custom cleaning rule
if 'YourColumn' in self.df.columns:
    # Your logic here
    self.df['YourColumn'] = self.df['YourColumn'].apply(your_function)
    self.cleaning_log.append("Applied custom cleaning rule")
```

### Use Your Own Data

Replace `generate_dirty_dataset()` with your data:
```python
# Instead of:
df_original = generate_dirty_dataset(n_records=500)

# Use:
df_original = pd.read_csv('your_data.csv')
```

## Performance Optimization

### For Large Datasets (>10MB)

1. **Reduce data loading:**
```python
df = pd.read_csv('file.csv', nrows=100000)  # Load subset
```

2. **Use efficient data types:**
```python
dtypes = {
    'ID': 'int32',
    'Name': 'category',
    'Amount': 'float32'
}
df = pd.read_csv('file.csv', dtype=dtypes)
```

3. **Process in chunks:**
```python
for chunk in pd.read_csv('file.csv', chunksize=10000):
    # Process chunk
    pass
```

## Development Setup

### Using Jupyter Notebook

```bash
pip install jupyter
jupyter notebook
```

Create new notebook, import and use:
```python
from data_cleaning_automation import DataQualityReporter
import pandas as pd

# Load your data
df = pd.read_csv('your_data.csv')

# Create reporter and clean
reporter = DataQualityReporter(df)
df_cleaned = reporter.clean_data()
```

### Using IDE

Recommended editors:
- **VS Code** (Free, lightweight)
- **PyCharm** (Professional, feature-rich)
- **Sublime Text** (Fast, simple)

### Git Configuration

```bash
git config --global user.name "Your Name"
git config --global user.email "your.email@example.com"

# Create development branch
git checkout -b feature/improvement
```

## System-Specific Notes

### Windows

- Use PowerShell or Command Prompt
- Backslash for paths: `C:\Users\Name\project`
- May need to enable long paths for large datasets

### macOS

- Install Xcode Command Line Tools if needed:
```bash
xcode-select --install
```
- Use `python3` instead of `python`

### Linux

Ubuntu/Debian:
```bash
sudo apt-get install python3-pip python3-venv
python3 -m venv venv
```

Fedora/CentOS:
```bash
sudo dnf install python3-pip
python3 -m venv venv
```

## Updating Dependencies

Check for outdated packages:
```bash
pip list --outdated
```

Update specific package:
```bash
pip install --upgrade pandas
```

Update all packages:
```bash
pip install --upgrade -r requirements.txt
```

## Uninstalling

Remove virtual environment:
```bash
# Windows
deactivate
rmdir /s venv

# macOS/Linux
deactivate
rm -rf venv
```

## Docker Setup (Optional)

Create `Dockerfile`:
```dockerfile
FROM python:3.9-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .
CMD ["python", "data_cleaning_automation.py"]
```

Build and run:
```bash
docker build -t data-cleaning .
docker run data-cleaning
```

## Next Steps

1. ✅ Installation complete
2. 📖 Read `README.md` for project overview
3. ▶️ Run `python data_cleaning_automation.py`
4. 📊 Check generated outputs
5. 🔧 Customize for your data
6. 📤 Upload to GitHub

## Getting Help

- Review code comments and docstrings
- Check error messages in console
- Verify all dependencies installed
- Test with sample data first
- Post issues on GitHub

## Quick Reference

```bash
# Activate environment
source venv/bin/activate  # macOS/Linux
venv\Scripts\activate     # Windows

# Install packages
pip install -r requirements.txt

# Run project
python data_cleaning_automation.py

# Deactivate environment
deactivate

# View generated files
ls -la  # macOS/Linux
dir     # Windows
```

---

**Estimated Setup Time**: 10-15 minutes  
**Difficulty Level**: Beginner-friendly  
**Maintenance**: Zero (except updates)

Good luck! 🚀
