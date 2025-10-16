# Python in Excel - Job Market Explorer

An interactive Excel workbook powered by Python in Excel for exploring and analyzing job market data with filtering, visualization, and export capabilities.

## 🎯 Project Overview

This project provides a comprehensive tool for job seekers and recruiters to explore job positions through an interactive Excel workbook powered by Python in Excel. Users can filter, visualize, and analyze job market trends with real-time data exploration capabilities directly within Excel.

## ✨ Features

### 📊 Interactive Data Exploration
- **Dynamic Filtering**: Filter jobs by role, location, experience level, and salary range
- **Keyword Search**: Search through job descriptions for specific skills or requirements
- **Real-time Updates**: All visualizations and tables update dynamically based on filters

### 📈 Data Visualizations
- **Salary Analysis**: Bar charts showing average salary by job title
- **Geographic Distribution**: Pie charts displaying job distribution by location
- **Experience Requirements**: Histograms showing experience level requirements
- **Interactive Charts**: Built with Plotly for enhanced user interaction

### 💾 Data Management
- **Multiple Formats**: Support for CSV and Excel file inputs
- **Export Options**: Save filtered results to Excel or CSV formats
- **Data Preview**: Comprehensive table view with key job information

## 🛠️ Technology Stack

- **Python Libraries** (via Python in Excel):
  - `pandas` - Data manipulation and analysis
  - `matplotlib` - Data visualization and charts
  - `seaborn` - Statistical data visualization
  - `numpy` - Numerical computing
  - `openpyxl` - Excel file operations

- **Data Formats**: CSV, Excel (.xlsx)
- **Environment**: Microsoft Excel with Python in Excel enabled

## 📋 Data Structure

The application expects job data with the following columns:
- **Job Title** - Position name
- **Company** - Employer name
- **Location** - Geographic location
- **Experience** - Required experience level
- **Salary** - Compensation information
- **Job Description** - Detailed job requirements and responsibilities

## 🚀 Getting Started

### Prerequisites
- Microsoft Excel (with Python in Excel feature enabled)
- Microsoft 365 subscription (required for Python in Excel)
- Python in Excel add-in installed

### Installation
1. Clone this repository
2. Open the Excel workbook in Microsoft Excel
3. Ensure Python in Excel is enabled in your Excel settings
4. The workbook contains Python-enabled cells for data analysis

### Usage
1. Load your job data into the designated data sheet
2. Use the Python-powered filtering cells to analyze jobs by criteria
3. View automatically generated charts and visualizations
4. Export filtered results using the built-in Python functions

## 📁 Project Structure

```
pythoninexcel/
├── README.md           # This file
├── note.txt           # Project planning notes
├── job_market_explorer.xlsx  # Main Excel workbook with Python integration
└── sample_data/       # Sample job data files
    ├── jobs_sample.csv
    └── jobs_sample.xlsx
```

## 🔧 Features in Detail

### Interactive Filters (Python-powered Excel cells)
- **Role Filter**: Python function to filter by specific job titles
- **Location Filter**: Python function to filter by geographic regions
- **Experience Filter**: Python function to filter by experience requirements
- **Salary Range Filter**: Python function to filter by compensation range
- **Keyword Search**: Python function for text-based search in job descriptions

### Visualizations (Python-generated charts in Excel)
- **Salary Trends**: Python-generated charts showing average salary by role
- **Geographic Analysis**: Python-generated charts showing job distribution by location
- **Experience Patterns**: Python-generated histograms of experience requirements
- **Market Insights**: Python-powered interactive charts embedded in Excel cells

## 📊 Sample Use Cases

- **Job Seekers**: Find positions matching their skills and salary expectations
- **Recruiters**: Analyze market trends and competitive positioning
- **Career Counselors**: Understand industry requirements and salary ranges
- **HR Professionals**: Benchmark compensation and requirements

## 🤝 Contributing

This is a practice project for learning Python in Excel data analysis and interactive visualization techniques. Feel free to fork and experiment with additional Python-powered Excel features!

## 📝 License

This project is for educational and practice purposes.

---

*Built with Python in Excel, pandas, matplotlib, and interactive data visualization tools for exploring the job market directly within Excel.*
