import pandas as pd
import glob
import os

def calculate_all_years():
    print("Calculating annual sums with updated filters...")
    print("-" * 60)
    print(f"{'Year':<10} | {'Original Sum (Approx)':<20} | {'New Sum (Net Executed)':<25} | {'Difference':<15}")
    print("-" * 60)

    all_files = glob.glob("tableau_BudgetData*.xlsx") + glob.glob("tableau_tableau_BudgetData*.xlsx")
    
    results = {}

    for filename in sorted(all_files):
        # Extract year
        year_str = ''.join(filter(str.isdigit, filename))[:4]
        year = int(year_str) if year_str else 0
        
        if year < 2015 or year > 2024:
            continue
            
        try:
            df = pd.read_excel(filename)
            
            # 1. Convert to numeric
            df['הוצאה נטו'] = pd.to_numeric(df['הוצאה נטו'], errors='coerce').fillna(0)
            raw_sum = df['הוצאה נטו'].sum()
            
            # 2. Filter: Expense
            if 'הוצאה/הכנסה' in df.columns:
                df = df[df['הוצאה/הכנסה'] == 'הוצאה']
            
            # 3. Filter: Executed
            if 'סוג תקציב' in df.columns:
                df = df[df['סוג תקציב'] == 'ביצוע']
                
            # 4. Filter: Exclude Debt Repayment (62, 35)
            if 'קוד רמה 2' in df.columns:
                df = df[~df['קוד רמה 2'].isin([62, 35])]
                
            # 5. Filter: Exclude Business Enterprises (New Filter)
            if 'שם סוג סעיף' in df.columns:
                df = df[df['שם סוג סעיף'] != 'מפעלים עסקיים']
                
            # 6. Filter: Positive values
            df = df[df['הוצאה נטו'] > 0]
            
            final_sum = df['הוצאה נטו'].sum()
            results[year] = final_sum
            
            # We don't have the "Original Sum" easily without re-running the old logic, 
            # but we can just show the final result which is what the user wants.
            print(f"{year:<10} | {'---':<20} | {final_sum:,.2f}")
            
        except Exception as e:
            print(f"Error processing {year}: {e}")

    print("-" * 60)
    return results

if __name__ == "__main__":
    calculate_all_years()
