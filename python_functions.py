"""
Python functions for Job Market Explorer in Excel
These functions will be used in Python-enabled Excel cells
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from typing import List, Dict, Any

def load_job_data(file_path: str) -> pd.DataFrame:
    """
    Load job data from CSV or Excel file
    """
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        
        # Clean and standardize data
        df['Salary'] = pd.to_numeric(df['Salary'].str.replace('$', '').str.replace(',', ''), errors='coerce')
        df['Experience_Years'] = df['Experience'].str.extract(r'(\d+)').astype(float)
        
        return df
    except Exception as e:
        return pd.DataFrame({'Error': [str(e)]})

def filter_jobs(df: pd.DataFrame, 
                job_title: str = None, 
                location: str = None, 
                min_salary: float = None, 
                max_salary: float = None,
                min_experience: float = None,
                max_experience: float = None,
                keyword: str = None) -> pd.DataFrame:
    """
    Filter jobs based on multiple criteria
    """
    filtered_df = df.copy()
    
    if job_title and job_title != 'All':
        filtered_df = filtered_df[filtered_df['Job Title'].str.contains(job_title, case=False, na=False)]
    
    if location and location != 'All':
        filtered_df = filtered_df[filtered_df['Location'].str.contains(location, case=False, na=False)]
    
    if min_salary is not None:
        filtered_df = filtered_df[filtered_df['Salary'] >= min_salary]
    
    if max_salary is not None:
        filtered_df = filtered_df[filtered_df['Salary'] <= max_salary]
    
    if min_experience is not None:
        filtered_df = filtered_df[filtered_df['Experience_Years'] >= min_experience]
    
    if max_experience is not None:
        filtered_df = filtered_df[filtered_df['Experience_Years'] <= max_experience]
    
    if keyword and keyword.strip():
        filtered_df = filtered_df[filtered_df['Job Description'].str.contains(keyword, case=False, na=False)]
    
    return filtered_df

def get_unique_values(df: pd.DataFrame, column: str) -> List[str]:
    """
    Get unique values from a column for dropdown options
    """
    return ['All'] + sorted(df[column].dropna().unique().tolist())

def calculate_salary_stats(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Calculate salary statistics
    """
    if df.empty:
        return {'mean': 0, 'median': 0, 'min': 0, 'max': 0, 'count': 0}
    
    return {
        'mean': round(df['Salary'].mean(), 2),
        'median': round(df['Salary'].median(), 2),
        'min': round(df['Salary'].min(), 2),
        'max': round(df['Salary'].max(), 2),
        'count': len(df)
    }

def create_salary_chart(df: pd.DataFrame) -> plt.Figure:
    """
    Create salary distribution chart
    """
    if df.empty:
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.text(0.5, 0.5, 'No data to display', ha='center', va='center', transform=ax.transAxes)
        return fig
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    
    # Salary by Job Title
    salary_by_title = df.groupby('Job Title')['Salary'].mean().sort_values(ascending=True)
    salary_by_title.plot(kind='barh', ax=ax1, color='skyblue')
    ax1.set_title('Average Salary by Job Title')
    ax1.set_xlabel('Salary ($)')
    ax1.tick_params(axis='y', labelsize=8)
    
    # Salary Distribution
    ax2.hist(df['Salary'], bins=10, color='lightgreen', alpha=0.7, edgecolor='black')
    ax2.set_title('Salary Distribution')
    ax2.set_xlabel('Salary ($)')
    ax2.set_ylabel('Frequency')
    
    plt.tight_layout()
    return fig

def create_location_chart(df: pd.DataFrame) -> plt.Figure:
    """
    Create location distribution chart
    """
    if df.empty:
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.text(0.5, 0.5, 'No data to display', ha='center', va='center', transform=ax.transAxes)
        return fig
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    
    # Jobs by Location
    location_counts = df['Location'].value_counts()
    location_counts.plot(kind='pie', ax=ax1, autopct='%1.1f%%')
    ax1.set_title('Job Distribution by Location')
    ax1.set_ylabel('')
    
    # Average Salary by Location
    salary_by_location = df.groupby('Location')['Salary'].mean().sort_values(ascending=True)
    salary_by_location.plot(kind='barh', ax=ax2, color='orange')
    ax2.set_title('Average Salary by Location')
    ax2.set_xlabel('Salary ($)')
    
    plt.tight_layout()
    return fig

def create_experience_chart(df: pd.DataFrame) -> plt.Figure:
    """
    Create experience level chart
    """
    if df.empty:
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.text(0.5, 0.5, 'No data to display', ha='center', va='center', transform=ax.transAxes)
        return fig
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Experience distribution
    experience_counts = df['Experience'].value_counts()
    experience_counts.plot(kind='bar', ax=ax, color='purple', alpha=0.7)
    ax.set_title('Job Distribution by Experience Level')
    ax.set_xlabel('Experience Level')
    ax.set_ylabel('Number of Jobs')
    ax.tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    return fig

def export_filtered_data(df: pd.DataFrame, filename: str = 'filtered_jobs.xlsx') -> str:
    """
    Export filtered data to Excel file
    """
    try:
        df.to_excel(filename, index=False)
        return f"Data exported successfully to {filename}"
    except Exception as e:
        return f"Export failed: {str(e)}"

def get_job_summary(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Get comprehensive job market summary
    """
    if df.empty:
        return {'total_jobs': 0, 'unique_companies': 0, 'unique_locations': 0, 'avg_salary': 0}
    
    return {
        'total_jobs': len(df),
        'unique_companies': df['Company'].nunique(),
        'unique_locations': df['Location'].nunique(),
        'avg_salary': round(df['Salary'].mean(), 2),
        'salary_range': f"${df['Salary'].min():,.0f} - ${df['Salary'].max():,.0f}",
        'top_locations': df['Location'].value_counts().head(3).to_dict(),
        'top_companies': df['Company'].value_counts().head(3).to_dict()
    }
