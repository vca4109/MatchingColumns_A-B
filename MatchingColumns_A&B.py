import pandas as pd

def match_columns(df):
    matched_items = []
    unmatched_items = []

    
    for item_a in df.iloc[:, 0].unique():
        if item_a in df.iloc[:, 1].unique():
            matched_items.append(item_a)
        else:
            unmatched_items.append((item_a, 'Column A'))

    
    for item_b in df.iloc[:, 1].unique():
        if item_b not in matched_items and item_b not in unmatched_items:
            if item_b in df.iloc[:, 0].unique():
                matched_items.append(item_b)
            else:
                unmatched_items.append((item_b, 'Column B'))

    return matched_items, unmatched_items


file_path = "Match_A&B_Excel.xlsx"
try:
    df = pd.read_excel(file_path, header=None)
except FileNotFoundError:
    print(f"Error: File '{file_path}' not found.")
    exit()


matched_items, unmatched_items = match_columns(df)


matched_df = pd.DataFrame(matched_items, columns=["Matched Items"])
unmatched_df = pd.DataFrame(unmatched_items, columns=["Unmatched Items", "Source Column"])


output_file_path = "matched_unmatched_items.xlsx"
with pd.ExcelWriter(output_file_path) as writer:
    matched_df.to_excel(writer, sheet_name="Matched", index=False)
    unmatched_df.to_excel(writer, sheet_name="Unmatched", index=False)

print(f"Matched and unmatched items exported to '{output_file_path}'")
