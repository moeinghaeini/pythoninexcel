"""
Demo script to showcase the Job Market Explorer functionality
This demonstrates how the Python functions work with the sample data
"""

import pandas as pd
import matplotlib.pyplot as plt
from python_functions import *

def run_demo():
    """Run a comprehensive demo of the Job Market Explorer"""
    
    print("🚀 Job Market Explorer Demo")
    print("=" * 50)
    
    # Load sample data
    print("\n📊 Loading sample job data...")
    df = load_job_data('sample_data/jobs_sample.csv')
    print(f"Loaded {len(df)} job records")
    print(f"Columns: {list(df.columns)}")
    
    # Show data preview
    print("\n📋 Data Preview:")
    print(df.head())
    
    # Show unique values for filters
    print("\n🔍 Available Filter Options:")
    print(f"Job Titles: {get_unique_values(df, 'Job Title')}")
    print(f"Locations: {get_unique_values(df, 'Location')}")
    
    # Demo filtering
    print("\n🔎 Demo Filtering:")
    print("Filtering for Software Engineer positions in San Francisco...")
    
    filtered_df = filter_jobs(
        df,
        job_title="Software Engineer",
        location="San Francisco",
        min_salary=100000,
        max_salary=150000
    )
    
    print(f"Found {len(filtered_df)} matching jobs")
    if not filtered_df.empty:
        print(filtered_df[['Job Title', 'Company', 'Location', 'Salary']])
    
    # Demo salary analysis
    print("\n💰 Salary Analysis:")
    salary_stats = calculate_salary_stats(df)
    print(f"Average Salary: ${salary_stats['mean']:,.2f}")
    print(f"Median Salary: ${salary_stats['median']:,.2f}")
    print(f"Salary Range: ${salary_stats['min']:,.2f} - ${salary_stats['max']:,.2f}")
    
    # Demo job summary
    print("\n📈 Job Market Summary:")
    summary = get_job_summary(df)
    for key, value in summary.items():
        print(f"{key.replace('_', ' ').title()}: {value}")
    
    # Demo keyword search
    print("\n🔍 Keyword Search Demo:")
    print("Searching for jobs containing 'Python'...")
    python_jobs = filter_jobs(df, keyword="python")
    print(f"Found {len(python_jobs)} jobs mentioning Python")
    if not python_jobs.empty:
        print(python_jobs[['Job Title', 'Company', 'Job Description']].head())
    
    # Demo export functionality
    print("\n💾 Export Demo:")
    export_result = export_filtered_data(filtered_df, 'demo_filtered_jobs.xlsx')
    print(export_result)
    
    # Create visualizations
    print("\n📊 Creating Visualizations...")
    
    # Salary chart
    fig1 = create_salary_chart(df)
    fig1.suptitle('Job Market Explorer - Salary Analysis', fontsize=16)
    plt.tight_layout()
    plt.savefig('salary_analysis.png', dpi=300, bbox_inches='tight')
    print("✅ Salary chart saved as 'salary_analysis.png'")
    plt.close()
    
    # Location chart
    fig2 = create_location_chart(df)
    fig2.suptitle('Job Market Explorer - Location Analysis', fontsize=16)
    plt.tight_layout()
    plt.savefig('location_analysis.png', dpi=300, bbox_inches='tight')
    print("✅ Location chart saved as 'location_analysis.png'")
    plt.close()
    
    # Experience chart
    fig3 = create_experience_chart(df)
    fig3.suptitle('Job Market Explorer - Experience Analysis', fontsize=16)
    plt.tight_layout()
    plt.savefig('experience_analysis.png', dpi=300, bbox_inches='tight')
    print("✅ Experience chart saved as 'experience_analysis.png'")
    plt.close()
    
    print("\n🎉 Demo completed successfully!")
    print("\nNext steps:")
    print("1. Open 'job_market_explorer.xlsx' in Microsoft Excel")
    print("2. Enable Python in Excel")
    print("3. Copy the Python code from the Excel cells")
    print("4. Run the code to see interactive results")

if __name__ == "__main__":
    run_demo()
