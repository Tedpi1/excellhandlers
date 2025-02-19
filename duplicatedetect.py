import openpyxl
import pandas as pd
from thefuzz import fuzz
import tkinter as tk
from tkinter import filedialog


def merge_similar_items(file_path, similarity_threshold=80):
    # Load Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Ensure required columns exist
    required_columns = {"ItemId", "ItemCode", "ItemName"}
    if not required_columns.issubset(df.columns):
        print("Error: Required columns (ItemId, ItemCode, ItemName) not found!")
        return

    merged_items = []
    matched = set()

    for i, row in df.iterrows():
        if i in matched:
            continue  # Skip already merged items

        current_name = row["ItemName"]
        similar_rows = [row]

        for j, other_row in df.iterrows():
            if i != j and j not in matched:
                other_name = other_row["ItemName"]
                similarity_score = fuzz.ratio(
                    str(current_name).lower(), str(other_name).lower()
                )

                if similarity_score >= similarity_threshold:
                    similar_rows.append(other_row)
                    matched.add(j)

        # Merge similar items
        merged_item = similar_rows[0].copy()
        merged_item["ItemName"] = (
            f"{merged_item['ItemName']} (Merged {len(similar_rows)} items)"
        )

        # If there's a quantity column, sum up the values
        if "Quantity" in df.columns:
            merged_item["Quantity"] = sum(
                item["Quantity"]
                for item in similar_rows
                if pd.notnull(item["Quantity"])
            )

        merged_items.append(merged_item)

    # Create a new DataFrame with merged results
    merged_df = pd.DataFrame(merged_items)

    # Save the merged data
    output_file = file_path.replace(".xlsx", "_merged.xlsx")
    merged_df.to_excel(output_file, index=False)
    print(f"✅ Merged similar items. Saved as: {output_file}")


def upload_file():
    # Open file dialog to let the user choose an Excel file
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    if file_path:
        print(f"📂 File selected: {file_path}")
        merge_similar_items(file_path)


# Run the file upload function
upload_file()
