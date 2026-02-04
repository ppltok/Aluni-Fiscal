"""
יצירת ויזואליזציה אינטראקטיבית של תקציב המדינה
עיצוב מודרני עם Dark Mode ו-UX פרימיום
"""

import pandas as pd
import glob
import json
import os
import webbrowser
import re

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

                # Get budget code (קוד תקנה) for matching with paid supports
                budget_code = None
                if 'קוד תקנה' in row.index:
                    code_val = row['קוד תקנה']
                    if pd.notna(code_val):
                        budget_code = int(code_val)

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
                        'code': budget_code,  # קוד תקנה for matching with paid supports
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


def load_paid_supports_data(budget_data):
    """
    Load paid supports data from CSV and match to budget codes.
    Returns data structured by year with:
    - totalPaid: total amount paid
    - recipientCount: number of unique recipients
    - byCode: aggregated data by budget code (קוד תקנה)
    - recipients: list of individual recipients for table display
    - orphanRecords/orphanAmount/orphanCodes: unmatched records stats
    - flowData: hierarchical data for Sankey diagram (רמה 1 → רמה 2 → סעיף → תקנה)
    - recipientsByCode: recipients grouped by budget code for drill-down
    """
    csv_path = os.path.join(os.path.dirname(__file__), 'table_of_paid_supports.csv')
    
    if not os.path.exists(csv_path):
        print(f"  קובץ תמיכות לא נמצא: {csv_path}")
        return {}
    
    print(f"  טוען נתוני תמיכות מ-{csv_path}...")
    df = pd.read_csv(csv_path)
    
    # Extract budget code from תקנה column (8-digit code at start)
    df['קוד_תקנה'] = df['תקנה'].str.extract(r'^(\d{8})')[0].astype(float).astype('Int64')
    
    # Extract תקנה name (text after the code)
    df['שם_תקנה'] = df['תקנה'].str.replace(r'^\d{8}\s*', '', regex=True)
    
    # Build set of valid codes per year from budget data
    # Also build hierarchy info for Sankey
    budget_codes_by_year = {}
    budget_info_by_year = {}  # code -> {name, path, value}
    
    for year, items in budget_data.items():
        budget_codes_by_year[year] = set()
        budget_info_by_year[year] = {}
        for item in items:
            if 'code' in item and item['code']:
                code = item['code']
                budget_codes_by_year[year].add(code)
                if code not in budget_info_by_year[year]:
                    budget_info_by_year[year][code] = {
                        'name': item.get('name', ''),
                        'path': item.get('path', []),
                        'value': 0
                    }
                budget_info_by_year[year][code]['value'] += item.get('value', 0)
    
    paid_data = {}
    
    # Process by year (שנת הבקשה)
    for year in range(2015, 2025):
        year_df = df[df['שנת הבקשה'] == year].copy()
        
        if len(year_df) == 0:
            continue
        
        # Get valid budget codes for this year
        valid_codes = budget_codes_by_year.get(year, set())
        
        # Mark matched vs orphan records
        year_df['is_matched'] = year_df['קוד_תקנה'].isin(valid_codes)
        
        matched_df = year_df[year_df['is_matched']]
        orphan_df = year_df[~year_df['is_matched']]
        
        # Aggregate by code
        # NOTE: Budget data is in thousands of ILS (אלפי ש"ח)
        # Paid supports data is in single ILS (ש"ח)
        # We convert paid amounts to thousands to match budget units
        by_code = {}
        for code, group in matched_df.groupby('קוד_תקנה'):
            code_int = int(code) if pd.notna(code) else 0
            paid_sum = group['סכום ששולם'].sum()
            # Convert from ILS to thousands of ILS to match budget units
            paid_in_thousands = (float(paid_sum) / 1000.0) if pd.notna(paid_sum) else 0.0
            by_code[str(code_int)] = {  # Use string key for JSON compatibility
                'paid': paid_in_thousands,
                'count': int(len(group)),
                'name': str(group['שם_תקנה'].iloc[0]) if len(group) > 0 and pd.notna(group['שם_תקנה'].iloc[0]) else '',
                'recipients': int(group['שם מגיש'].nunique())
            }
        
        # Build recipients grouped by code for drill-down (top 100 per code)
        recipients_by_code = {}
        for code, group in matched_df.groupby('קוד_תקנה'):
            code_int = int(code) if pd.notna(code) else 0
            code_str = str(code_int)
            top_recipients = group.nlargest(100, 'סכום ששולם')
            recipients_by_code[code_str] = []
            for _, row in top_recipients.iterrows():
                paid_val = row['סכום ששולם']
                hp_val = row.get('ח"פ מגיש', '')
                recipients_by_code[code_str].append({
                    'name': str(row['שם מגיש']) if pd.notna(row['שם מגיש']) else '',
                    'hp': str(hp_val) if pd.notna(hp_val) else '',
                    'paid': float(paid_val) if pd.notna(paid_val) else 0.0,  # Keep in ILS
                })
        
        # Build recipients list (limit to top 5000 by amount for performance)
        recipients_df = matched_df.nlargest(5000, 'סכום ששולם')
        recipients = []
        for _, row in recipients_df.iterrows():
            code_int = int(row['קוד_תקנה']) if pd.notna(row['קוד_תקנה']) else 0
            budget_info = budget_info_by_year.get(year, {}).get(code_int, {})
            
            paid_val = row['סכום ששולם']
            hp_val = row.get('ח"פ מגיש', '')
            # Keep individual recipient amounts in original ILS for display
            recipients.append({
                'name': str(row['שם מגיש']) if pd.notna(row['שם מגיש']) else '',
                'code': code_int,
                'hp': str(hp_val) if pd.notna(hp_val) else '',
                'takanName': str(row['שם_תקנה']) if pd.notna(row['שם_תקנה']) else '',
                'ministry': str(budget_info.get('path', [''])[0]) if budget_info.get('path') else '',
                'paid': float(paid_val) if pd.notna(paid_val) else 0.0,  # Keep in ILS for display
                'requestYear': int(row['שנת הבקשה']) if pd.notna(row['שנת הבקשה']) else year
            })
        
        # Build hierarchical flow data for Sankey diagram
        # Structure: רמה 1 → רמה 2 → סעיף → תקנה
        flow_data = build_flow_data(budget_info_by_year.get(year, {}), by_code)
        
        orphan_sum = orphan_df['סכום ששולם'].sum()
        matched_sum = matched_df['סכום ששולם'].sum() if len(matched_df) > 0 else 0
        
        # Convert totals to thousands of ILS to match budget units
        paid_data[year] = {
            'totalPaid': float(matched_sum) / 1000.0 if matched_sum else 0.0,  # In thousands
            'recipientCount': int(matched_df['שם מגיש'].nunique()) if len(matched_df) > 0 else 0,
            'byCode': by_code,
            'recipients': recipients,
            'recipientsByCode': recipients_by_code,
            'flowData': flow_data,
            'orphanRecords': int(len(orphan_df)),
            'orphanAmount': (float(orphan_sum) / 1000.0) if pd.notna(orphan_sum) else 0.0,  # In thousands
            'orphanCodes': int(orphan_df['קוד_תקנה'].nunique()) if len(orphan_df) > 0 else 0
        }
        
        print(f"    שנת {year}: {len(matched_df):,} רשומות מותאמות, {len(orphan_df):,} ללא התאמה")
    
    return paid_data


def build_flow_data(budget_info, paid_by_code):
    """
    Build hierarchical flow data for Sankey diagram.
    
    Returns:
    {
        'nodes': [
            { 'id': 'r1_שירותים חברתיים', 'name': 'שירותים חברתיים', 'level': 1, 'budget': X, 'paid': Y },
            ...
        ],
        'links': [
            { 'source': 'r1_שירותים חברתיים', 'target': 'r2_חינוך', 'budget': X, 'paid': Y },
            ...
        ]
    }
    """
    # Aggregate budget and paid amounts by hierarchy level
    # Level 1: רמה 1, Level 2: רמה 2, Level 3: סעיף, Level 4: תקנה (code)
    
    hierarchy = {}  # path_key -> { budget, paid, children }
    
    for code, info in budget_info.items():
        path = info.get('path', [])
        if len(path) < 3:
            continue
            
        rama1 = path[0] if len(path) > 0 and path[0] else 'לא מוגדר'
        rama2 = path[1] if len(path) > 1 and path[1] else 'לא מוגדר'
        seif = path[2] if len(path) > 2 and path[2] else 'לא מוגדר'
        takana_name = info.get('name', str(code))
        
        budget = info.get('value', 0)
        paid_info = paid_by_code.get(str(code), {})
        paid = paid_info.get('paid', 0)
        
        # Build hierarchy keys
        r1_key = f"r1_{rama1}"
        r2_key = f"r2_{rama1}_{rama2}"
        r3_key = f"r3_{rama1}_{rama2}_{seif}"
        r4_key = f"r4_{code}"
        
        # Initialize nodes if not exist
        if r1_key not in hierarchy:
            hierarchy[r1_key] = {'name': rama1, 'level': 1, 'budget': 0, 'paid': 0, 'children': set()}
        if r2_key not in hierarchy:
            hierarchy[r2_key] = {'name': rama2, 'level': 2, 'budget': 0, 'paid': 0, 'children': set(), 'parent': r1_key}
        if r3_key not in hierarchy:
            hierarchy[r3_key] = {'name': seif, 'level': 3, 'budget': 0, 'paid': 0, 'children': set(), 'parent': r2_key}
        if r4_key not in hierarchy:
            hierarchy[r4_key] = {'name': takana_name, 'level': 4, 'budget': 0, 'paid': 0, 'code': code, 'parent': r3_key}
        
        # Aggregate values
        hierarchy[r1_key]['budget'] += budget
        hierarchy[r1_key]['paid'] += paid
        hierarchy[r1_key]['children'].add(r2_key)
        
        hierarchy[r2_key]['budget'] += budget
        hierarchy[r2_key]['paid'] += paid
        hierarchy[r2_key]['children'].add(r3_key)
        
        hierarchy[r3_key]['budget'] += budget
        hierarchy[r3_key]['paid'] += paid
        hierarchy[r3_key]['children'].add(r4_key)
        
        hierarchy[r4_key]['budget'] += budget
        hierarchy[r4_key]['paid'] += paid
    
    # Convert to nodes and links format
    nodes = []
    links = []
    
    for key, data in hierarchy.items():
        node = {
            'id': key,
            'name': data['name'],
            'level': data['level'],
            'budget': data['budget'],
            'paid': data['paid']
        }
        if 'code' in data:
            node['code'] = data['code']
        nodes.append(node)
        
        # Create link to parent
        if 'parent' in data:
            links.append({
                'source': data['parent'],
                'target': key,
                'budget': data['budget'],
                'paid': data['paid']
            })
    
    # Sort nodes by level and budget
    nodes.sort(key=lambda x: (x['level'], -x['budget']))
    
    # Convert children sets to lists (for JSON serialization)
    for key in hierarchy:
        if 'children' in hierarchy[key]:
            hierarchy[key]['children'] = list(hierarchy[key]['children'])
    
    return {
        'nodes': nodes,
        'links': links
    }


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


def create_paid_supports_file(budget_data, paid_supports_data):
    """יצירת קובץ HTML לניתוח תמיכות ותקציב"""

    template_path = os.path.join(os.path.dirname(__file__), 'paid_supports_template.html')

    if not os.path.exists(template_path):
        print("  תבנית paid_supports_template.html לא נמצאה, מדלג...")
        return None

    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()

    # Convert budget data to include code for matching
    json_budget = json.dumps(budget_data, ensure_ascii=False)
    json_paid = json.dumps(paid_supports_data, ensure_ascii=False)
    
    final_html = html_template.replace('__BUDGET_DATA_PLACEHOLDER__', json_budget)
    final_html = final_html.replace('__PAID_SUPPORTS_PLACEHOLDER__', json_paid)

    output_path = os.path.join(os.path.dirname(__file__), 'paid_supports.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_html)

    return output_path


def main():
    print("=" * 60)
    print("יצירת ויזואליזציה אינטראקטיבית של תקציב המדינה")
    print("=" * 60)

    # טעינת נתונים
    print("\n[1/9] טוען נתונים מכל השנים...")
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

    # טעינת נתוני תמיכות
    print("\n[2/9] טוען נתוני תמיכות...")
    paid_supports_data = load_paid_supports_data(budget_data)
    if paid_supports_data:
        print(f"  נטענו נתוני תמיכות ל-{len(paid_supports_data)} שנים")

    # יצירת HTML ראשי
    print("\n[3/9] יוצר קובץ HTML ראשי...")
    output_path = create_html_file(budget_data, income_data, commitment_data)
    print(f"נשמר: {output_path}")

    # יצירת HTML להשוואה שנתית
    print("\n[4/9] יוצר קובץ השוואה שנתית...")
    time_series_path = create_time_series_file(budget_data)
    if time_series_path:
        print(f"נשמר: {time_series_path}")

    # יצירת HTML לאחוזי שכר
    print("\n[5/9] יוצר קובץ אחוזי שכר...")
    salary_path = create_salary_percentage_file(budget_data)
    if salary_path:
        print(f"נשמר: {salary_path}")

    # יצירת HTML לסקירת משרדים
    print("\n[6/9] יוצר קובץ סקירת משרדים...")
    ministry_path = create_ministry_overview_file(budget_data)
    if ministry_path:
        print(f"נשמר: {ministry_path}")

    # יצירת HTML לתרשים Sunburst
    print("\n[7/9] יוצר קובץ Sunburst...")
    sunburst_path = create_sunburst_file(budget_data)
    if sunburst_path:
        print(f"נשמר: {sunburst_path}")

    # יצירת HTML ל-5 עמודי התקציב
    print("\n[8/9] יוצר קובץ 5 עמודי התקציב...")
    five_pillars_path = create_five_pillars_file(budget_data)
    if five_pillars_path:
        print(f"נשמר: {five_pillars_path}")

    # יצירת HTML לתמיכות ותקציב
    print("\n[9/9] יוצר קובץ תמיכות ותקציב...")
    if paid_supports_data:
        paid_supports_path = create_paid_supports_file(budget_data, paid_supports_data)
        if paid_supports_path:
            print(f"נשמר: {paid_supports_path}")
    else:
        print("  דילוג - אין נתוני תמיכות")

    # פתיחה בדפדפן
    print("\nפותח בדפדפן...")
    webbrowser.open('file://' + os.path.abspath(output_path))

    print("\n" + "=" * 60)
    print("הסתיים בהצלחה!")
    print("")
    print("שימוש:")
    print("  • גרור את הסליידר לבחירת שנה")
    print("  • לחץ על כל קופסא לכניסה פנימה")
    print("  • לחץ 'חזרה' או על הנתיב למעלה לחזור")
    print("  • לחץ על כפתור הניווט לויזואליזציות נוספות")
    print("  • לחץ על 'תמיכות ותקציב' לניתוח ביצוע תקציבי")
    print("=" * 60)


if __name__ == "__main__":
    main()
