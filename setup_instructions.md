# Setup Instructions for Job Market Explorer

## Prerequisites

1. **Microsoft Excel** with Python in Excel feature enabled
2. **Microsoft 365 subscription** (required for Python in Excel)
3. **Python in Excel add-in** installed

## Installation Steps

### 1. Enable Python in Excel

1. Open Microsoft Excel
2. Go to **Insert** > **Get Add-ins**
3. Search for "Python in Excel" and install it
4. Enable Python in Excel in your Excel settings
5. Restart Excel if prompted

### 2. Install Python Dependencies

Open a terminal/command prompt and run:

```bash
pip install -r requirements.txt
```

### 3. Create Excel Workbook Structure

1. Create a new Excel workbook
2. Create 5 worksheets with these names:
   - **Data Input** - For loading job data
   - **Filter Controls** - For filter parameters
   - **Analysis Dashboard** - For summary statistics
   - **Visualizations** - For charts and graphs
   - **Export** - For exporting results

### 4. Add Python-Enabled Cells

Copy the Python code from `excel_template.html` into Python-enabled cells in Excel:

#### Sheet 1: Data Input
- **Cell A1**: Load job data
```python
import pandas as pd
from python_functions import load_job_data

df = load_job_data('sample_data/jobs_sample.csv')
df.head(10)
```

#### Sheet 2: Filter Controls
- **Cells B2-B8**: Filter parameters (manual input)
- **Cell C10**: Apply filters
```python
from python_functions import filter_jobs

filtered_df = filter_jobs(
    df,
    job_title=xl("B2"),
    location=xl("B3"),
    min_salary=xl("B4"),
    max_salary=xl("B5"),
    min_experience=xl("B6"),
    max_experience=xl("B7"),
    keyword=xl("B8")
)
filtered_df
```

#### Sheet 3: Analysis Dashboard
- **Cell D2**: Job summary
- **Cell D5**: Salary statistics
- **Cell D8**: Top job titles
- **Cell D12**: Top locations

#### Sheet 4: Visualizations
- **Cell E2**: Salary chart
- **Cell E15**: Location chart
- **Cell E28**: Experience chart

#### Sheet 5: Export
- **Cell F2**: Export filtered data
- **Cell F5**: Export summary report

### 5. Load Sample Data

1. Use the provided sample data in `sample_data/jobs_sample.csv`
2. Or load your own job data in the same format

## Usage

1. **Load Data**: The data will automatically load in Sheet 1
2. **Set Filters**: Adjust filter parameters in Sheet 2
3. **View Analysis**: Check the dashboard in Sheet 3
4. **Explore Charts**: View visualizations in Sheet 4
5. **Export Results**: Use Sheet 5 to export filtered data

## Troubleshooting

### Python in Excel Not Working
- Ensure you have Microsoft 365 subscription
- Check that Python in Excel add-in is properly installed
- Restart Excel and try again

### Import Errors
- Make sure `python_functions.py` is in the same directory as your Excel file
- Check that all required packages are installed: `pip install -r requirements.txt`

### Data Loading Issues
- Verify the CSV file path is correct
- Check that the CSV file has the required columns: Job Title, Company, Location, Experience, Salary, Job Description

## File Structure

```
pythoninexcel/
├── README.md
├── requirements.txt
├── python_functions.py
├── excel_template.html
├── setup_instructions.md
├── sample_data/
│   └── jobs_sample.csv
└── job_market_explorer.xlsx (your Excel workbook)
```
