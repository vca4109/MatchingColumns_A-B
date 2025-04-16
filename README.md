🔍 Match Columns in Excel
This Python script reads an Excel file containing two columns and compares their values to determine which items are:

✅ Matched — exist in both columns

❌ Unmatched — exist in only one column

The results are exported into a new Excel file with two sheets: Matched and Unmatched.

📂 Input
File: Match_A&B_Excel.xlsx

Structure: No header row (header=None), with two columns (A and B) to compare.

🛠️ How It Works
Reads the Excel file using pandas.

Checks each unique item in both columns:

If the item appears in both, it's marked as a match.

If it appears in only one, it's recorded as unmatched with its source column.

Outputs the results into a new Excel file: matched_unmatched_items.xlsx.

📤 Output
The script generates:

matched_unmatched_items.xlsx

Sheet 1: Matched – A list of all matched values.

Sheet 2: Unmatched – A list of values found only in one column, along with which column they came from.

💡 Example Use Case
Perfect for cross-referencing two lists of IDs, names, or any other values where you need to quickly spot overlaps and differences.
