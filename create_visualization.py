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

    for filename in all_files:
        try:
            print(f"טוען {filename}...")
            df = pd.read_excel(filename)

            # חילוץ שנה
            year_str = ''.join(filter(str.isdigit, filename))[:4]
            year = int(year_str) if year_str else 0

            if year < 2015 or year > 2024:
                continue

            # סינון רק הוצאות וביצוע
            if 'הוצאה/הכנסה' in df.columns:
                df = df[df['הוצאה/הכנסה'] == 'הוצאה']

            if 'סוג תקציב' in df.columns:
                df = df[df['סוג תקציב'] == 'ביצוע']

            # המרת הוצאה נטו למספר
            df['הוצאה נטו'] = pd.to_numeric(df['הוצאה נטו'], errors='coerce').fillna(0)
            df = df[df['הוצאה נטו'] > 0]

            # יצירת מבנה נתונים להיררכיה
            # הוספת שם מיון רמה 1 כרמה האחרונה (שכר/קניות/העברות/השקעה)
            hierarchy_cols = ['שם רמה 1', 'שם רמה 2', 'שם סעיף', 'שם תחום', 'שם תקנה', 'שם מיון רמה 1']

            data_items = []
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

                if path[0] and value > 0:  # רק אם יש רמה 1 וערך חיובי
                    data_items.append({
                        'name': path[-1] if path[-1] else path[-2] if path[-2] else path[0],
                        'path': path,
                        'value': float(value),
                        'isSalary': is_salary,
                        'miunRama1': miun_rama1
                    })

            all_data[year] = data_items
            print(f"  נטענו {len(data_items)} רשומות לשנת {year}")

        except Exception as e:
            print(f"שגיאה בטעינת {filename}: {e}")

    return all_data


def create_html_file(budget_data):
    """יצירת קובץ HTML עם הנתונים"""

    # קריאת התבנית
    template_path = os.path.join(os.path.dirname(__file__), 'budget_visualization.html')

    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()

    # המרת הנתונים ל-JSON
    json_data = json.dumps(budget_data, ensure_ascii=False)

    # החלפת placeholder
    final_html = html_template.replace('__BUDGET_DATA_PLACEHOLDER__', json_data)

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


def main():
    print("=" * 60)
    print("יצירת ויזואליזציה אינטראקטיבית של תקציב המדינה")
    print("=" * 60)

    # טעינת נתונים
    print("\n[1/7] טוען נתונים מכל השנים...")
    budget_data = load_all_budget_data()

    print(f"\nנטענו נתונים ל-{len(budget_data)} שנים: {sorted(budget_data.keys())}")

    # יצירת HTML ראשי
    print("\n[2/7] יוצר קובץ HTML ראשי...")
    output_path = create_html_file(budget_data)
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
    print("\n[6/7] יוצר קובץ Sunburst...")
    sunburst_path = create_sunburst_file(budget_data)
    if sunburst_path:
        print(f"נשמר: {sunburst_path}")

    # פתיחה בדפדפן
    print("\n[7/7] פותח בדפדפן...")
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
