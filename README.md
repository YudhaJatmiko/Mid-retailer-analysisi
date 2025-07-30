Checking and installing required packages...
Installing openpyxl...
Collecting openpyxl
  Using cached openpyxl-3.1.5-py2.py3-none-any.whl.metadata (2.5 kB)
Collecting et-xmlfile (from openpyxl)
  Using cached et_xmlfile-2.0.0-py3-none-any.whl.metadata (2.7 kB)
Using cached openpyxl-3.1.5-py2.py3-none-any.whl (250 kB)
Using cached et_xmlfile-2.0.0-py3-none-any.whl (18 kB)
Installing collected packages: et-xmlfile, openpyxl
   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2/2 [openpyxl]
Successfully installed et-xmlfile-2.0.0 openpyxl-3.1.5
All libraries imported successfully!

Starting financial analysis...
Step 1: Extracting financial data...
================================================================================
EXTRACTING FINANCIAL STATEMENTS DATA
================================================================================
Income Statement:
         Revenue      COGS Gross Profit     SG&A Other  EBITDA Depreciation  \
Year 1  112640.0  -97429.0      15211.0 -10899.0 -63.0  4249.0      -1029.0   
Year 2  116199.0  -99938.0      16261.0 -11445.0 -65.0  4751.0      -1127.0   
Year 3  118719.0 -101646.0      17073.0 -12068.0 -78.0  4927.0      -1255.0   
Year 4  129025.0 -110512.0      18513.0 -12950.0 -82.0  5481.0      -1370.0   
Year 5  141576.0 -121715.0      19861.0 -13876.0 -68.0  5917.0      -1437.0   
...
================================================================================
CREATING VISUALIZATIONS
================================================================================
Dashboard saved as 'financial_analysis_dashboard.png'
Output is truncated. View as a scrollable element or open in a text editor. Adjust cell output settings...


Step 5: Generating summary report...

================================================================================
FINANCIAL ANALYSIS SUMMARY REPORT
================================================================================
GROWTH ANALYSIS:
• Revenue CAGR (8 years): 8.23%
• Net Income CAGR (8 years): 13.54%

PROFITABILITY ANALYSIS (8-year averages):
• Average ROE: 22.79%
• Average ROA: 7.40%
• Average Gross Margin: 14.01%
• Average Net Profit Margin: 2.19%

EFFICIENCY ANALYSIS:
• Average Asset Turnover: 3.39x

DUPONT ANALYSIS INSIGHTS:
• Latest ROE: 28.51%
• Driven by: Net Margin (2.56%), Asset Turnover (3.31x), Equity Multiplier (3.37x)

TREND ANALYSIS:
• ROE is increasing over the period
...

🎉 Analysis completed successfully!
📊 Charts should be displayed above
💾 Dashboard saved as 'financial_analysis_dashboard.png'
Output is truncated. View as a scrollable element or open in a text editor. Adjust cell output settings
