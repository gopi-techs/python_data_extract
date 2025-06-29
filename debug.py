import os
import pandas as pd


def debug_excel_structure(file_path):
    """Debug the structure of Excel files to understand layout"""
    try:
        print(f"\n{'=' * 60}")
        print(f"DEBUGGING: {os.path.basename(file_path)}")
        print(f"{'=' * 60}")

        # Check all sheets
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        print(f"📋 Sheets found: {xls.sheet_names}")

        # Focus on the second sheet (index 1)
        if len(xls.sheet_names) > 1:
            sheet_name = xls.sheet_names[1]
            print(f"🎯 Analyzing sheet: '{sheet_name}'")

            # Read without header
            df = pd.read_excel(file_path, sheet_name=1, header=None, engine='openpyxl')
            df = df.astype(str).fillna('')

            print(f"📊 Sheet dimensions: {df.shape[0]} rows x {df.shape[1]} columns")

            # Show first 20 rows with content
            print(f"\n🔍 FIRST 20 ROWS WITH CONTENT:")
            print("-" * 80)

            for i in range(min(20, len(df))):
                row = df.iloc[i]
                # Only show rows that have meaningful content
                non_empty = [str(cell) for cell in row.values if str(cell).strip() not in ['', 'nan']]
                if non_empty:
                    print(f"Row {i:2d}: {non_empty}")

            # Look for specific patterns
            print(f"\n🎯 SEARCHING FOR KEY PATTERNS:")
            print("-" * 40)

            patterns_found = {
                'SSH Invoice Numbers': [],
                'Date Patterns': [],
                'Name Fields': [],
                'GST Invoice References': [],
                'Company Names': []
            }

            for i, row in df.iterrows():
                for j, cell in enumerate(row.values):
                    cell_str = str(cell).strip()

                    # SSH invoice pattern
                    if 'SSH-' in cell_str and len(cell_str) > 8:
                        patterns_found['SSH Invoice Numbers'].append(f"Row {i}, Col {j}: {cell_str}")

                    # Date patterns
                    if '-' in cell_str and any(char.isdigit() for char in cell_str):
                        # Simple date check
                        parts = cell_str.split('-')
                        if len(parts) == 3 and all(part.isdigit() for part in parts):
                            patterns_found['Date Patterns'].append(f"Row {i}, Col {j}: {cell_str}")

                    # Name field
                    if 'Name' in cell_str and ':' in cell_str:
                        patterns_found['Name Fields'].append(f"Row {i}, Col {j}: {cell_str}")

                    # GST invoice
                    if 'GST' in cell_str.upper() and 'INVOICE' in cell_str.upper():
                        patterns_found['GST Invoice References'].append(f"Row {i}, Col {j}: {cell_str}")

                    # Potential company names (alphabetic, longer than 5 chars, not common words)
                    if (len(cell_str) > 5 and
                            cell_str.replace(' ', '').replace('.', '').replace('&', '').isalpha() and
                            cell_str.lower() not in ['tamilnadu', 'karnataka', 'invoice', 'number', 'date', 'total']):
                        patterns_found['Company Names'].append(f"Row {i}, Col {j}: {cell_str}")

            # Display findings
            for pattern_type, findings in patterns_found.items():
                print(f"\n{pattern_type}:")
                if findings:
                    for finding in findings[:5]:  # Show first 5 matches
                        print(f"  ✅ {finding}")
                    if len(findings) > 5:
                        print(f"  ... and {len(findings) - 5} more")
                else:
                    print("  ❌ None found")

            # Show a specific area around row 6-10 where invoice details usually are
            print(f"\n🎯 DETAILED VIEW - ROWS 5-12 (Invoice Header Area):")
            print("-" * 80)
            for i in range(5, min(13, len(df))):
                row = df.iloc[i]
                print(f"Row {i:2d}: ", end="")
                for j, cell in enumerate(row.values):
                    cell_str = str(cell).strip()
                    if cell_str and cell_str != 'nan':
                        print(f"[Col{j}:{cell_str}] ", end="")
                print()  # New line

        else:
            print("❌ No second sheet found!")

    except Exception as e:
        print(f"❌ Error analyzing {file_path}: {e}")


def main():
    # 👉 CHANGE THIS TO YOUR FOLDER PATH
    folder_path = r"F:\Office\SRI SAI HEATERS APRIL 2024 - MAR 2025"

    if not os.path.exists(folder_path):
        print(f"❌ Folder not found: {folder_path}")
        return

    excel_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx', '.xls'))]

    if not excel_files:
        print("❌ No Excel files found!")
        return

    print(f"🔍 Found {len(excel_files)} Excel files")
    print("📝 Analyzing first 3 files for structure...")

    # Analyze first 3 files to understand the pattern
    for i, file in enumerate(excel_files[:3]):
        full_path = os.path.join(folder_path, file)
        debug_excel_structure(full_path)

        if i < 2:  # Don't ask after the last file
            input("\nPress Enter to continue to next file...")


if __name__ == "__main__":
    main()