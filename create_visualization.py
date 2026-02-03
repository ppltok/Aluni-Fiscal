"""
יצירת ויזואליזציה אינטראקטיבית של תקציב המדינה
עיצוב מודרני עם Dark Mode ו-UX פרימיום
"""

import pandas as pd
import glob
import json
import os
import webbrowser

# ============================
# NAME NORMALIZATION MAPPINGS
# ============================
# Some budget categories changed names over the years but represent the same item
NAME_MAPPINGS = {
    # תת-תחום mappings (שם רמה 2)
    'ביטוח לאומי': 'הקצבות ביטוח לאומי',
    'העברות ביטוח לאומי': 'הקצבות ביטוח לאומי',
    # משרד התחבורה - שם תת-תחום השתנה ב-2019
    'פיתוח התחבורה': 'תחבורה',
}

def normalize_name(name):
    """Normalize budget item names that changed across years"""
    if pd.isna(name):
        return name
    return NAME_MAPPINGS.get(str(name), str(name))

def load_all_budget_data():
    """טעינת כל קבצי התקציב"""
    all_files = glob.glob("tableau_BudgetData*.xlsx") + glob.glob("tableau_tableau_BudgetData*.xlsx")

    all_data = {}
    all_income = {}
    all_commitments = {}

    for filename in all_files:
        try:
            print(f"טוען {filename}...")
            df = pd.read_excel(filename)

            # חילוץ שנה
            year_str = ''.join(filter(str.isdigit, filename))[:4]
            year = int(year_str) if year_str else 0

            if year < 2015 or year > 2024:
                continue

            # יצירת מבנה נתונים להיררכיה
            hierarchy_cols = ['שם רמה 1', 'שם רמה 2', 'שם סעיף', 'שם תחום', 'שם תקנה', 'שם מיון רמה 1']

            # --- עיבוד הכנסות ---
            # Income = negative values (flipped to positive)
            # Sources:
            # 1. הכנסה rows with NEGATIVE values (actual income)
            # 2. הוצאה rows with NEGATIVE values (expenses that are actually income)
            income_items = []
            
            # 1. הכנסות מסומנות כ"הכנסה" - only NEGATIVE values are actual income
            if 'הוצאה/הכנסה' in df.columns:
                income_df = df[df['הוצאה/הכנסה'] == 'הכנסה'].copy()
                
                if 'סוג תקציב' in income_df.columns:
                    income_df = income_df[income_df['סוג תקציב'] == 'ביצוע']
                
                # Filter out state income (הכנסות category) - we only want ministry income
                if 'שם רמה 1' in income_df.columns:
                    income_df = income_df[income_df['שם רמה 1'] != 'הכנסות']

                income_df['הוצאה נטו'] = pd.to_numeric(income_df['הוצאה נטו'], errors='coerce').fillna(0)
                
                # Only include NEGATIVE values (actual income) and flip to positive
                income_df = income_df[income_df['הוצאה נטו'] < 0]
                
                for _, row in income_df.iterrows():
                    path = []
                    for col in hierarchy_cols:
                        raw_val = row.get(col, '')
                        if pd.notna(raw_val):
                            path.append(normalize_name(str(raw_val)))
                        else:
                            path.append('')
                    
                    value = row['הוצאה נטו']
                    
                    if path[0]:
                        # Flip negative to positive (income is stored as positive)
                        income_items.append({
                            'path': path,
                            'value': abs(float(value))
                        })
            
            # 2. הוצאות שליליות (גם הן הכנסות בפועל) - flip to positive
            if 'הוצאה/הכנסה' in df.columns:
                negative_exp_df = df[df['הוצאה/הכנסה'] == 'הוצאה'].copy()
                
                if 'סוג תקציב' in negative_exp_df.columns:
                    negative_exp_df = negative_exp_df[negative_exp_df['סוג תקציב'] == 'ביצוע']
                
                negative_exp_df['הוצאה נטו'] = pd.to_numeric(negative_exp_df['הוצאה נטו'], errors='coerce').fillna(0)
                negative_exp_df = negative_exp_df[negative_exp_df['הוצאה נטו'] < 0]
                
                for _, row in negative_exp_df.iterrows():
                    path = []
                    for col in hierarchy_cols:
                        raw_val = row.get(col, '')
                        if pd.notna(raw_val):
                            path.append(normalize_name(str(raw_val)))
                        else:
                            path.append('')
                    
                    value = row['הוצאה נטו']
                    
                    if path[0]:
                        # Flip negative to positive (income is stored as positive)
                        income_items.append({
                            'path': path,
                            'value': abs(float(value))
                        })
            
            all_income[year] = income_items

            # --- עיבוד הוצאות ---
            if 'הוצאה/הכנסה' in df.columns:
                df = df[df['הוצאה/הכנסה'] == 'הוצאה']

            if 'סוג תקציב' in df.columns:
                df = df[df['סוג תקציב'] == 'ביצוע']

            # סינון קוד רמה 2 = 62 (החזרי חוב קרן) ו-35
            if 'קוד רמה 2' in df.columns:
                df = df[~df['קוד רמה 2'].isin([62])]
            
            # Additional filtering: exclude specific section codes (קוד סעיף)
            if 'קוד סעיף' in df.columns and 'קוד מיון רמה 2' in df.columns:
                # Convert section code to 4-digit string for comparison
                seif_code = df['קוד סעיף'].astype(str).str.zfill(4)
                
                # Codes to exclude: 0000, 0089, 0091, 0093, 0094, 0095, 0098
                should_exclude = (
                    seif_code.str.startswith('0000') |
                    seif_code.str.startswith('0089') |
                    seif_code.str.startswith('0091') |
                    seif_code.str.startswith('0093') |
                    seif_code.str.startswith('0094') |
                    seif_code.str.startswith('0095') |
                    seif_code.str.startswith('0098')
                )
                
                # Special case: code 0084 is only excluded if קוד מיון רמה 2 != 266
                is_0084_to_exclude = seif_code.str.startswith('0084') & (df['קוד מיון רמה 2'] != 266)
                should_exclude = should_exclude | is_0084_to_exclude
                
                # Keep only rows that should NOT be excluded
                df = df[~should_exclude]

            # המרת הוצאה נטו למספר
            df['הוצאה נטו'] = pd.to_numeric(df['הוצאה נטו'], errors='coerce').fillna(0)
            
            # Note: We do NOT subtract יתרת התחיבויות (commitment balance)
            # This matches the approach in join_phases.py which includes commitment_balance
            # as part of the budget amounts without subtraction
            
            df = df[df['הוצאה נטו'] > 0]

            data_items = []
            commitment_items = []
            for _, row in df.iterrows():
                # Build path with name normalization
                path = []
                for col in hierarchy_cols:
                    raw_val = row.get(col, '')
                    if pd.notna(raw_val):
                        path.append(normalize_name(str(raw_val)))
                    else:
                        path.append('')
                value = row['הוצאה נטו']

                # בדיקה אם זה שכר
                miun_rama1 = str(row.get('שם מיון רמה 1', '')) if pd.notna(row.get('שם מיון רמה 1')) else ''
                is_salary = miun_rama1 == 'שכר'

                # Get commitment balance for this row
                commitment_value = 0
                if 'יתרת התחיבויות' in row.index:
                    commitment_value = pd.to_numeric(row['יתרת התחיבויות'], errors='coerce')
                    if pd.isna(commitment_value):
                        commitment_value = 0

                if path[0] and value > 0:  # רק אם יש רמה 1 וערך חיובי
                    data_items.append({
                        'name': path[-1] if path[-1] else path[-2] if path[-2] else path[0],
                        'path': path,
                        'value': float(value),
                        'isSalary': is_salary,
                        'miunRama1': miun_rama1,
                        'program': str(row.get('שם תכנית', '')) if pd.notna(row.get('שם תכנית')) else '',
                        'classification': str(row.get('שם מיון רמה 2', '')) if pd.notna(row.get('שם מיון רמה 2')) else ''
                    })
                    
                    # Add commitment item with same path structure
                    if commitment_value != 0:
                        commitment_items.append({
                            'path': path,
                            'value': float(commitment_value)
                        })

            all_data[year] = data_items
            all_commitments[year] = commitment_items
            print(f"  נטענו {len(data_items)} רשומות הוצאה, {len(income_items)} רשומות הכנסה ו-{len(commitment_items)} רשומות התחייבויות לשנת {year}")

        except Exception as e:
            print(f"שגיאה בטעינת {filename}: {e}")

    return all_data, all_income, all_commitments


def create_html_file(budget_data, income_data, commitment_data):
    """יצירת קובץ HTML עם הנתונים"""

    # קריאת התבנית
    template_path = os.path.join(os.path.dirname(__file__), 'budget_visualization.html')

    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()

    # המרת הנתונים ל-JSON
    json_data = json.dumps(budget_data, ensure_ascii=False)
    json_income = json.dumps(income_data, ensure_ascii=False)
    json_commitments = json.dumps(commitment_data, ensure_ascii=False)

    # החלפת placeholder
    final_html = html_template.replace('__BUDGET_DATA_PLACEHOLDER__', json_data)
    final_html = final_html.replace('__INCOME_DATA_PLACEHOLDER__', json_income)
    final_html = final_html.replace('__COMMITMENT_DATA_PLACEHOLDER__', json_commitments)

    # שמירה
    output_path = os.path.join(os.path.dirname(__file__), 'budget_interactive.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_html)

    return output_path


def create_time_series_file(budget_data):
    """יצירת קובץ HTML להשוואה שנתית"""

    template_path = os.path.join(os.path.dirname(__file__), 'time_series_template.html')

    if not os.path.exists(template_path):
        print("  תבנית time_series_template.html לא נמצאה, מדלג...")
        return None

    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()

    json_data = json.dumps(budget_data, ensure_ascii=False)
    final_html = html_template.replace('__BUDGET_DATA_PLACEHOLDER__', json_data)

    output_path = os.path.join(os.path.dirname(__file__), 'time_series.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_html)

    return output_path


def create_salary_percentage_file(budget_data):
    """יצירת קובץ HTML לניתוח אחוזי שכר"""

    template_path = os.path.join(os.path.dirname(__file__), 'salary_percentage_template.html')

    if not os.path.exists(template_path):
        print("  תבנית salary_percentage_template.html לא נמצאה, מדלג...")
        return None

    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()

    json_data = json.dumps(budget_data, ensure_ascii=False)
    final_html = html_template.replace('__BUDGET_DATA_PLACEHOLDER__', json_data)

    output_path = os.path.join(os.path.dirname(__file__), 'salary_percentage.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_html)

    return output_path


def create_ministry_overview_file(budget_data):
    """יצירת קובץ HTML לסקירת משרדים"""

    template_path = os.path.join(os.path.dirname(__file__), 'ministry_overview_template.html')

    if not os.path.exists(template_path):
        print("  תבנית ministry_overview_template.html לא נמצאה, מדלג...")
        return None

    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()

    json_data = json.dumps(budget_data, ensure_ascii=False)
    final_html = html_template.replace('__BUDGET_DATA_PLACEHOLDER__', json_data)

    output_path = os.path.join(os.path.dirname(__file__), 'ministry_overview.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_html)

    return output_path


def create_sunburst_file(budget_data):
    """יצירת קובץ HTML לתרשים Sunburst"""

    template_path = os.path.join(os.path.dirname(__file__), 'sunburst_template.html')

    if not os.path.exists(template_path):
        print("  תבנית sunburst_template.html לא נמצאה, מדלג...")
        return None

    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()

    json_data = json.dumps(budget_data, ensure_ascii=False)
    final_html = html_template.replace('__BUDGET_DATA_PLACEHOLDER__', json_data)

    output_path = os.path.join(os.path.dirname(__file__), 'sunburst_budget.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_html)

    return output_path


def create_five_pillars_file(budget_data):
    """יצירת קובץ HTML ל-5 עמודי התקציב"""

    template_path = os.path.join(os.path.dirname(__file__), 'five_pillars_template.html')

    if not os.path.exists(template_path):
        print("  תבנית five_pillars_template.html לא נמצאה, מדלג...")
        return None

    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()

    json_data = json.dumps(budget_data, ensure_ascii=False)
    final_html = html_template.replace('__BUDGET_DATA_PLACEHOLDER__', json_data)

    output_path = os.path.join(os.path.dirname(__file__), 'five_pillars.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_html)

    return output_path


def main():
    print("=" * 60)
    print("יצירת ויזואליזציה אינטראקטיבית של תקציב המדינה")
    print("=" * 60)

    # טעינת נתונים
    print("\n[1/7] טוען נתונים מכל השנים...")
    budget_data, income_data, commitment_data = load_all_budget_data()

    print(f"\nנטענו נתונים ל-{len(budget_data)} שנים: {sorted(budget_data.keys())}")
    
    # Print total expenditure summary
    print("\nסיכום הוצאות לפי שנה (לאחר סינון):")
    for year in sorted(budget_data.keys()):
        year_total = sum(item['value'] for item in budget_data[year])
        print(f"  {year}: {year_total:,.2f}")
    grand_total = sum(sum(item['value'] for item in budget_data[year]) for year in budget_data.keys())
    print(f"  סה\"כ כל השנים: {grand_total:,.2f}")
    
    # Print commitment balance summary
    print("\nסיכום יתרת התחיבויות לפי שנה:")
    for year in sorted(commitment_data.keys()):
        year_commitment_total = sum(item['value'] for item in commitment_data[year])
        print(f"  {year}: {year_commitment_total:,.2f}")
    commitment_grand_total = sum(sum(item['value'] for item in commitment_data[year]) for year in commitment_data.keys())
    print(f"  סה\"כ כל השנים: {commitment_grand_total:,.2f}")

    # יצירת HTML ראשי
    print("\n[2/7] יוצר קובץ HTML ראשי...")
    output_path = create_html_file(budget_data, income_data, commitment_data)
    print(f"נשמר: {output_path}")

    # יצירת HTML להשוואה שנתית
    print("\n[3/7] יוצר קובץ השוואה שנתית...")
    time_series_path = create_time_series_file(budget_data)
    if time_series_path:
        print(f"נשמר: {time_series_path}")

    # יצירת HTML לאחוזי שכר
    print("\n[4/7] יוצר קובץ אחוזי שכר...")
    salary_path = create_salary_percentage_file(budget_data)
    if salary_path:
        print(f"נשמר: {salary_path}")

    # יצירת HTML לסקירת משרדים
    print("\n[5/7] יוצר קובץ סקירת משרדים...")
    ministry_path = create_ministry_overview_file(budget_data)
    if ministry_path:
        print(f"נשמר: {ministry_path}")

    # יצירת HTML לתרשים Sunburst
    print("\n[6/8] יוצר קובץ Sunburst...")
    sunburst_path = create_sunburst_file(budget_data)
    if sunburst_path:
        print(f"נשמר: {sunburst_path}")

    # יצירת HTML ל-5 עמודי התקציב
    print("\n[7/8] יוצר קובץ 5 עמודי התקציב...")
    five_pillars_path = create_five_pillars_file(budget_data)
    if five_pillars_path:
        print(f"נשמר: {five_pillars_path}")

    # פתיחה בדפדפן
    print("\n[8/8] פותח בדפדפן...")
    webbrowser.open('file://' + os.path.abspath(output_path))

    print("\n" + "=" * 60)
    print("הסתיים בהצלחה!")
    print("")
    print("שימוש:")
    print("  • גרור את הסליידר לבחירת שנה")
    print("  • לחץ על כל קופסא לכניסה פנימה")
    print("  • לחץ 'חזרה' או על הנתיב למעלה לחזור")
    print("  • לחץ על כפתור הניווט לויזואליזציות נוספות")
    print("=" * 60)


if __name__ == "__main__":
    main()
