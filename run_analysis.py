#!/usr/bin/env python3
"""
Simple runner for financial analysis - works in Cursor/interactive environments
"""

import sys
import subprocess

def install_requirements():
    """Install required packages if not available"""
    required_packages = ['pandas', 'numpy', 'matplotlib', 'seaborn', 'openpyxl']
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

def main():
    """Main function to run the analysis"""
    # Install requirements if needed
    print("Checking and installing required packages...")
    install_requirements()
    
    # Now import the analysis module
    from financial_analysis import run_analysis
    
    # Run the analysis
    print("\nStarting financial analysis...")
    analyzer, summary = run_analysis()
    
    if analyzer and summary:
        print("\n🎉 Analysis completed successfully!")
        print("📊 Charts should be displayed above")
        print("💾 Dashboard saved as 'financial_analysis_dashboard.png'")
    else:
        print("❌ Analysis failed. Please check the error messages above.")

if __name__ == "__main__":
    main()