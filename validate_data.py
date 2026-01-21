
import json
import re
import os

def validate_json_in_html():
    file_path = 'budget_interactive.html'
    if not os.path.exists(file_path):
        print(f"File {file_path} not found.")
        return

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Extract JSON object
    # const BUDGET_DATA = { ... };
    match = re.search(r'const BUDGET_DATA = ({.*?});', content, re.DOTALL)
    if not match:
        print("Could not find BUDGET_DATA in HTML.")
        return

    json_str = match.group(1)
    
    try:
        data = json.loads(json_str)
        print("JSON is valid.")
        
        # Check structure of 2024 data
        if '2024' in data:
            items = data['2024']
            print(f"Found {len(items)} items for 2024.")
            if len(items) > 0:
                item = items[0]
                print("Sample item keys:", item.keys())
                if 'program' in item:
                    print(f"Sample program: '{item['program']}'")
                else:
                    print("WARNING: 'program' key missing in item.")
                
                if 'classification' in item:
                    print(f"Sample classification: '{item['classification']}'")
                else:
                    print("WARNING: 'classification' key missing in item.")
                    
                # Check for items with empty program
                empty_programs = sum(1 for i in items if not i.get('program'))
                print(f"Items with empty program: {empty_programs} out of {len(items)}")
                
    except json.JSONDecodeError as e:
        print(f"JSON Decode Error: {e}")
        # Print context around error
        print(json_str[max(0, e.pos-50):min(len(json_str), e.pos+50)])

if __name__ == "__main__":
    validate_json_in_html()
