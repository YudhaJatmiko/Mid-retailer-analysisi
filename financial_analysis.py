#!/usr/bin/env python3
"""
Mid Retailer Financial Analysis
Comprehensive financial analysis following Excel template structure
"""

# Import required libraries
try:
    import pandas as pd
    import numpy as np
    import matplotlib.pyplot as plt
    import seaborn as sns
    import warnings
    warnings.filterwarnings('ignore')
    
    # Set display options for better output
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', None)
    
    print("All libraries imported successfully!")
    
except ImportError as e:
    print(f"Import Error: {e}")
    print("Please install required libraries:")
    print("pip install pandas numpy matplotlib seaborn openpyxl")
    exit(1)

class MidRetailerAnalysis:
    def __init__(self, file_path):
        self.file_path = file_path
        self.financial_data = {}
        
    def extract_financial_statements(self):
        """Extract and structure financial statement data"""
        print("="*80)
        print("EXTRACTING FINANCIAL STATEMENTS DATA")
        print("="*80)
        
        # Load the Financial Statements sheet
        df = pd.read_excel(self.file_path, sheet_name='Financial Statements', header=None)
        
        # Define years (columns 5-12 represent years 1-8)
        years = [f'Year {i}' for i in range(1, 9)]
        
        # Extract Income Statement data
        income_statement = {
            'Revenue': df.iloc[7, 5:13].values,
            'COGS': df.iloc[8, 5:13].values,
            'Gross Profit': df.iloc[9, 5:13].values,
            'SG&A': df.iloc[12, 5:13].values,
            'Other': df.iloc[13, 5:13].values,
            'EBITDA': df.iloc[14, 5:13].values,
            'Depreciation': df.iloc[17, 5:13].values,
            'EBIT': df.iloc[18, 5:13].values,
            'Interest Expense': df.iloc[21, 5:13].values,
            'Interest Income': df.iloc[22, 5:13].values,
            'EBT': df.iloc[23, 5:13].values,
            'Taxes': df.iloc[26, 5:13].values,
            'Net Income': df.iloc[29, 5:13].values
        }
        
        # Extract Balance Sheet data (Assets)
        assets = {
            'Cash': df.iloc[75, 5:13].values,
            'Accounts Receivable': df.iloc[76, 5:13].values,
            'Inventory': df.iloc[77, 5:13].values,
            'Current Assets': df.iloc[78, 5:13].values,
            'PP&E': df.iloc[80, 5:13].values,
            'Other Assets': df.iloc[81, 5:13].values,
            'Total Assets': df.iloc[83, 5:13].values
        }
        
        # Extract Balance Sheet data (Liabilities & Equity)
        liabilities_equity = {
            'Accounts Payable': df.iloc[88, 5:13].values,
            'Current Liabilities': df.iloc[90, 5:13].values,
            'Long Term Debt': df.iloc[92, 5:13].values,
            'Total Liabilities': df.iloc[93, 5:13].values,
            'Common Equity': df.iloc[98, 5:13].values,
            'Retained Earnings': df.iloc[99, 5:13].values,
            'Total Equity': df.iloc[100, 5:13].values,
            'Total Liab & Equity': df.iloc[103, 5:13].values
        }
        
        # Create DataFrames
        self.income_statement = pd.DataFrame(income_statement, index=years)
        self.balance_sheet_assets = pd.DataFrame(assets, index=years)
        self.balance_sheet_liab_equity = pd.DataFrame(liabilities_equity, index=years)
        
        # Store all financial data
        self.financial_data = {
            'income_statement': self.income_statement,
            'assets': self.balance_sheet_assets,
            'liabilities_equity': self.balance_sheet_liab_equity
        }
        
        print("Income Statement:")
        print(self.income_statement)
        print("\nBalance Sheet - Assets:")
        print(self.balance_sheet_assets)
        print("\nBalance Sheet - Liabilities & Equity:")
        print(self.balance_sheet_liab_equity)
        
        return self.financial_data
    
    def calculate_ratios(self):
        """Calculate financial ratios following the Excel template structure"""
        print("\n" + "="*80)
        print("CALCULATING FINANCIAL RATIOS")
        print("="*80)
        
        # Profitability Ratios
        profitability_ratios = {}
        
        # Return on Equity = Net Income / Total Equity
        profitability_ratios['Return on Equity'] = (
            self.income_statement['Net Income'] / 
            self.balance_sheet_liab_equity['Total Equity']
        )
        
        # Return on Assets = Net Income / Total Assets
        profitability_ratios['Return on Assets'] = (
            self.income_statement['Net Income'] / 
            self.balance_sheet_assets['Total Assets']
        )
        
        # Gross Margin = Gross Profit / Revenue
        profitability_ratios['Gross Margin'] = (
            self.income_statement['Gross Profit'] / 
            self.income_statement['Revenue']
        )
        
        # SG&A % of Revenue = SG&A / Revenue
        profitability_ratios['SG&A % of Revenue'] = (
            abs(self.income_statement['SG&A']) / 
            self.income_statement['Revenue']
        )
        
        # EBITDA Margin = EBITDA / Revenue
        profitability_ratios['EBITDA Margin'] = (
            self.income_statement['EBITDA'] / 
            self.income_statement['Revenue']
        )
        
        # EBIT Margin = EBIT / Revenue
        profitability_ratios['EBIT Margin'] = (
            self.income_statement['EBIT'] / 
            self.income_statement['Revenue']
        )
        
        # Net Profit Margin = Net Income / Revenue
        profitability_ratios['Net Profit Margin'] = (
            self.income_statement['Net Income'] / 
            self.income_statement['Revenue']
        )
        
        # Efficiency Ratios
        efficiency_ratios = {}
        
        # Asset Turnover = Revenue / Total Assets
        efficiency_ratios['Asset Turnover'] = (
            self.income_statement['Revenue'] / 
            self.balance_sheet_assets['Total Assets']
        )
        
        # Working Capital Turnover = Revenue / (Current Assets - Current Liabilities)
        working_capital = (self.balance_sheet_assets['Current Assets'] - 
                          self.balance_sheet_liab_equity['Current Liabilities'])
        efficiency_ratios['Working Capital Turnover'] = (
            self.income_statement['Revenue'] / working_capital
        )
        
        # Cash Turnover = Revenue / Cash
        efficiency_ratios['Cash Turnover'] = (
            self.income_statement['Revenue'] / 
            self.balance_sheet_assets['Cash']
        )
        
        # A/R Turnover = Revenue / Accounts Receivable
        efficiency_ratios['A/R Turnover'] = (
            self.income_statement['Revenue'] / 
            self.balance_sheet_assets['Accounts Receivable']
        )
        
        # Inventory Turnover = COGS / Inventory
        efficiency_ratios['Inventory Turnover'] = (
            abs(self.income_statement['COGS']) / 
            self.balance_sheet_assets['Inventory']
        )
        
        # Leverage Ratios
        leverage_ratios = {}
        
        # Debt to Equity = Total Liabilities / Total Equity
        leverage_ratios['Debt to Equity'] = (
            self.balance_sheet_liab_equity['Total Liabilities'] / 
            self.balance_sheet_liab_equity['Total Equity']
        )
        
        # Debt to Assets = Total Liabilities / Total Assets
        leverage_ratios['Debt to Assets'] = (
            self.balance_sheet_liab_equity['Total Liabilities'] / 
            self.balance_sheet_assets['Total Assets']
        )
        
        # Equity Multiplier (Total Asset to Equity) = Total Assets / Total Equity
        leverage_ratios['Equity Multiplier'] = (
            self.balance_sheet_assets['Total Assets'] / 
            self.balance_sheet_liab_equity['Total Equity']
        )
        
        # Store ratios
        self.profitability_ratios = pd.DataFrame(profitability_ratios)
        self.efficiency_ratios = pd.DataFrame(efficiency_ratios)
        self.leverage_ratios = pd.DataFrame(leverage_ratios)
        
        print("Profitability Ratios:")
        print(self.profitability_ratios.round(4))
        print("\nEfficiency Ratios:")
        print(self.efficiency_ratios.round(4))
        print("\nLeverage Ratios:")
        print(self.leverage_ratios.round(4))
        
        return {
            'profitability': self.profitability_ratios,
            'efficiency': self.efficiency_ratios,
            'leverage': self.leverage_ratios
        }
    
    def dupont_3_step_analysis(self):
        """Perform 3-Step DuPont Analysis: ROE = Net Profit Margin × Asset Turnover × Equity Multiplier"""
        print("\n" + "="*80)
        print("3-STEP DUPONT ANALYSIS")
        print("="*80)
        
        # Calculate components
        net_profit_margin = (self.income_statement['Net Income'] / 
                           self.income_statement['Revenue'])
        
        asset_turnover = (self.income_statement['Revenue'] / 
                         self.balance_sheet_assets['Total Assets'])
        
        equity_multiplier = (self.balance_sheet_assets['Total Assets'] / 
                           self.balance_sheet_liab_equity['Total Equity'])
        
        # Calculate ROE using DuPont decomposition
        roe_dupont = net_profit_margin * asset_turnover * equity_multiplier
        
        # Direct ROE calculation for verification
        roe_direct = (self.income_statement['Net Income'] / 
                     self.balance_sheet_liab_equity['Total Equity'])
        
        dupont_3_step = pd.DataFrame({
            'Net Profit Margin': net_profit_margin,
            'Asset Turnover': asset_turnover,
            'Equity Multiplier': equity_multiplier,
            'ROE (DuPont)': roe_dupont,
            'ROE (Direct)': roe_direct,
            'Difference': roe_dupont - roe_direct
        })
        
        print("3-Step DuPont Analysis:")
        print(dupont_3_step.round(4))
        
        # Additional breakdown for better understanding
        print("\nBreakdown Components:")
        print(f"Net Profit Margin = Net Income / Revenue")
        print(f"Asset Turnover = Revenue / Total Assets")
        print(f"Equity Multiplier = Total Assets / Total Equity")
        print(f"ROE = Net Profit Margin × Asset Turnover × Equity Multiplier")
        
        self.dupont_3_step = dupont_3_step
        return dupont_3_step
    
    def dupont_5_step_analysis(self):
        """Perform 5-Step DuPont Analysis with additional decomposition"""
        print("\n" + "="*80)
        print("5-STEP DUPONT ANALYSIS")
        print("="*80)
        
        # Calculate components
        # Tax Burden = Net Income / EBT
        tax_burden = (self.income_statement['Net Income'] / 
                     self.income_statement['EBT'])
        
        # Interest Burden = EBT / EBIT
        interest_burden = (self.income_statement['EBT'] / 
                          self.income_statement['EBIT'])
        
        # EBIT Margin = EBIT / Revenue
        ebit_margin = (self.income_statement['EBIT'] / 
                      self.income_statement['Revenue'])
        
        # Asset Turnover = Revenue / Total Assets
        asset_turnover = (self.income_statement['Revenue'] / 
                         self.balance_sheet_assets['Total Assets'])
        
        # Equity Multiplier = Total Assets / Total Equity
        equity_multiplier = (self.balance_sheet_assets['Total Assets'] / 
                           self.balance_sheet_liab_equity['Total Equity'])
        
        # Calculate ROE using 5-step DuPont decomposition
        roe_dupont_5 = (tax_burden * interest_burden * ebit_margin * 
                       asset_turnover * equity_multiplier)
        
        # Direct ROE calculation for verification
        roe_direct = (self.income_statement['Net Income'] / 
                     self.balance_sheet_liab_equity['Total Equity'])
        
        dupont_5_step = pd.DataFrame({
            'Tax Burden': tax_burden,
            'Interest Burden': interest_burden,
            'EBIT Margin': ebit_margin,
            'Asset Turnover': asset_turnover,
            'Equity Multiplier': equity_multiplier,
            'ROE (5-Step DuPont)': roe_dupont_5,
            'ROE (Direct)': roe_direct,
            'Difference': roe_dupont_5 - roe_direct
        })
        
        print("5-Step DuPont Analysis:")
        print(dupont_5_step.round(4))
        
        # Additional breakdown for better understanding
        print("\nBreakdown Components:")
        print(f"Tax Burden = Net Income / EBT")
        print(f"Interest Burden = EBT / EBIT")
        print(f"EBIT Margin = EBIT / Revenue")
        print(f"Asset Turnover = Revenue / Total Assets")
        print(f"Equity Multiplier = Total Assets / Total Equity")
        print(f"ROE = Tax Burden × Interest Burden × EBIT Margin × Asset Turnover × Equity Multiplier")
        
        self.dupont_5_step = dupont_5_step
        return dupont_5_step
    
    def create_visualizations(self):
        """Create visualizations for the financial analysis"""
        print("\n" + "="*80)
        print("CREATING VISUALIZATIONS")
        print("="*80)
        
        try:
            # Set up the plotting style
            plt.style.use('default')  # Changed from 'seaborn-v0_8' to 'default'
            fig, axes = plt.subplots(2, 3, figsize=(18, 12))
            fig.suptitle('Mid Retailer Financial Analysis Dashboard', fontsize=16, fontweight='bold')
            
            # 1. Revenue and Net Income Trend
            axes[0, 0].plot(self.income_statement.index, self.income_statement['Revenue'], 
                           marker='o', linewidth=2, label='Revenue')
            axes[0, 0].plot(self.income_statement.index, self.income_statement['Net Income'], 
                           marker='s', linewidth=2, label='Net Income')
            axes[0, 0].set_title('Revenue vs Net Income Trend')
            axes[0, 0].set_ylabel('USD (thousands)')
            axes[0, 0].legend()
            axes[0, 0].tick_params(axis='x', rotation=45)
            
            # 2. Profitability Margins
            axes[0, 1].plot(self.profitability_ratios.index, self.profitability_ratios['Gross Margin'], 
                           marker='o', label='Gross Margin')
            axes[0, 1].plot(self.profitability_ratios.index, self.profitability_ratios['EBITDA Margin'], 
                           marker='s', label='EBITDA Margin')
            axes[0, 1].plot(self.profitability_ratios.index, self.profitability_ratios['Net Profit Margin'], 
                           marker='^', label='Net Profit Margin')
            axes[0, 1].set_title('Profitability Margins')
            axes[0, 1].set_ylabel('Ratio')
            axes[0, 1].legend()
            axes[0, 1].tick_params(axis='x', rotation=45)
            
            # 3. Return Ratios
            axes[0, 2].plot(self.profitability_ratios.index, self.profitability_ratios['Return on Equity'], 
                           marker='o', label='ROE')
            axes[0, 2].plot(self.profitability_ratios.index, self.profitability_ratios['Return on Assets'], 
                           marker='s', label='ROA')
            axes[0, 2].set_title('Return Ratios')
            axes[0, 2].set_ylabel('Ratio')
            axes[0, 2].legend()
            axes[0, 2].tick_params(axis='x', rotation=45)
            
            # 4. Asset Turnover and Efficiency
            axes[1, 0].plot(self.efficiency_ratios.index, self.efficiency_ratios['Asset Turnover'], 
                           marker='o', label='Asset Turnover')
            axes[1, 0].plot(self.efficiency_ratios.index, self.efficiency_ratios['Inventory Turnover'], 
                           marker='s', label='Inventory Turnover')
            axes[1, 0].set_title('Efficiency Ratios')
            axes[1, 0].set_ylabel('Ratio')
            axes[1, 0].legend()
            axes[1, 0].tick_params(axis='x', rotation=45)
            
            # 5. 3-Step DuPont Components
            axes[1, 1].plot(self.dupont_3_step.index, self.dupont_3_step['Net Profit Margin'], 
                           marker='o', label='Net Profit Margin')
            axes[1, 1].plot(self.dupont_3_step.index, self.dupont_3_step['Asset Turnover'], 
                           marker='s', label='Asset Turnover')
            axes[1, 1].plot(self.dupont_3_step.index, self.dupont_3_step['Equity Multiplier'], 
                           marker='^', label='Equity Multiplier')
            axes[1, 1].set_title('3-Step DuPont Components')
            axes[1, 1].set_ylabel('Ratio')
            axes[1, 1].legend()
            axes[1, 1].tick_params(axis='x', rotation=45)
            
            # 6. ROE Comparison
            axes[1, 2].plot(self.dupont_3_step.index, self.dupont_3_step['ROE (Direct)'], 
                           marker='o', linewidth=2, label='ROE (Direct)')
            axes[1, 2].plot(self.dupont_5_step.index, self.dupont_5_step['ROE (5-Step DuPont)'], 
                           marker='s', linewidth=2, label='ROE (5-Step DuPont)')
            axes[1, 2].set_title('ROE Analysis')
            axes[1, 2].set_ylabel('ROE')
            axes[1, 2].legend()
            axes[1, 2].tick_params(axis='x', rotation=45)
            
            plt.tight_layout()
            
            # Save the plot
            try:
                plt.savefig('financial_analysis_dashboard.png', dpi=300, bbox_inches='tight')
                print("Dashboard saved as 'financial_analysis_dashboard.png'")
            except Exception as save_error:
                print(f"Could not save plot: {save_error}")
            
            # Display the plot
            plt.show()
            
        except Exception as e:
            print(f"Error creating visualizations: {e}")
            print("Continuing without visualizations...")
    
    def generate_summary_report(self):
        """Generate a comprehensive summary report"""
        print("\n" + "="*80)
        print("FINANCIAL ANALYSIS SUMMARY REPORT")
        print("="*80)
        
        # Calculate growth rates
        revenue_growth = ((self.income_statement['Revenue'].iloc[-1] / 
                         self.income_statement['Revenue'].iloc[0]) ** (1/7)) - 1
        
        net_income_growth = ((self.income_statement['Net Income'].iloc[-1] / 
                            self.income_statement['Net Income'].iloc[0]) ** (1/7)) - 1
        
        # Average ratios over the period
        avg_roe = self.profitability_ratios['Return on Equity'].mean()
        avg_roa = self.profitability_ratios['Return on Assets'].mean()
        avg_gross_margin = self.profitability_ratios['Gross Margin'].mean()
        avg_net_margin = self.profitability_ratios['Net Profit Margin'].mean()
        avg_asset_turnover = self.efficiency_ratios['Asset Turnover'].mean()
        
        print(f"GROWTH ANALYSIS:")
        print(f"• Revenue CAGR (8 years): {revenue_growth:.2%}")
        print(f"• Net Income CAGR (8 years): {net_income_growth:.2%}")
        
        print(f"\nPROFITABILITY ANALYSIS (8-year averages):")
        print(f"• Average ROE: {avg_roe:.2%}")
        print(f"• Average ROA: {avg_roa:.2%}")
        print(f"• Average Gross Margin: {avg_gross_margin:.2%}")
        print(f"• Average Net Profit Margin: {avg_net_margin:.2%}")
        
        print(f"\nEFFICIENCY ANALYSIS:")
        print(f"• Average Asset Turnover: {avg_asset_turnover:.2f}x")
        
        print(f"\nDUPONT ANALYSIS INSIGHTS:")
        latest_3_step = self.dupont_3_step.iloc[-1]
        print(f"• Latest ROE: {latest_3_step['ROE (Direct)']:.2%}")
        print(f"• Driven by: Net Margin ({latest_3_step['Net Profit Margin']:.2%}), Asset Turnover ({latest_3_step['Asset Turnover']:.2f}x), Equity Multiplier ({latest_3_step['Equity Multiplier']:.2f}x)")
        
        # Identify trends
        roe_trend = "increasing" if self.profitability_ratios['Return on Equity'].iloc[-1] > self.profitability_ratios['Return on Equity'].iloc[0] else "decreasing"
        margin_trend = "improving" if self.profitability_ratios['Net Profit Margin'].iloc[-1] > self.profitability_ratios['Net Profit Margin'].iloc[0] else "declining"
        
        print(f"\nTREND ANALYSIS:")
        print(f"• ROE is {roe_trend} over the period")
        print(f"• Net profit margins are {margin_trend} over the period")
        
        return {
            'revenue_cagr': revenue_growth,
            'net_income_cagr': net_income_growth,
            'avg_roe': avg_roe,
            'avg_roa': avg_roa,
            'avg_gross_margin': avg_gross_margin,
            'avg_net_margin': avg_net_margin,
            'avg_asset_turnover': avg_asset_turnover
        }

# Main execution function
def run_analysis():
    """Run the complete financial analysis"""
    file_path = "Mid Retailer Financial Analysis - Blank.xlsx"
    
    try:
        # Initialize the analysis
        analyzer = MidRetailerAnalysis(file_path)
        
        # Extract financial statements
        print("Step 1: Extracting financial data...")
        financial_data = analyzer.extract_financial_statements()
        
        # Calculate ratios
        print("\nStep 2: Calculating financial ratios...")
        ratios = analyzer.calculate_ratios()
        
        # Perform DuPont analysis
        print("\nStep 3: Performing DuPont analysis...")
        dupont_3 = analyzer.dupont_3_step_analysis()
        dupont_5 = analyzer.dupont_5_step_analysis()
        
        # Create visualizations
        print("\nStep 4: Creating visualizations...")
        analyzer.create_visualizations()
        
        # Generate summary report
        print("\nStep 5: Generating summary report...")
        summary = analyzer.generate_summary_report()
        
        print("\n" + "="*80)
        print("ANALYSIS COMPLETE!")
        print("="*80)
        
        return analyzer, summary
        
    except FileNotFoundError:
        print(f"Error: Could not find the Excel file '{file_path}'")
        print("Please make sure the file is in the same directory as this script.")
        return None, None
    except Exception as e:
        print(f"Error during analysis: {e}")
        return None, None

# Main execution
if __name__ == "__main__":
    analyzer, summary = run_analysis()