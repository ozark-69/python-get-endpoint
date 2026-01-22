import pandas as pd
import sys
from openpyxl.utils import get_column_letter


def process_interface_changes(input_file, output_file):
    print(f"Reading file: {input_file}...")

    try:
        # Load all sheets
        all_sheets_dict = pd.read_excel(input_file, sheet_name=None)
    except FileNotFoundError:
        print(f"Error: The file '{input_file}' was not found.")
        return
    except Exception as e:
        print(f"An error occurred: {e}")
        return

    sheet_names = list(all_sheets_dict.keys())

    if len(sheet_names) < 2:
        print("Error: The file has less than 2 sheets. Cannot skip the first sheet.")
        return

    # 1. Skip the first sheet (Summary)
    summary_sheet = sheet_names[0]
    potential_data_sheets = sheet_names[1:]

    print(f"Skipping summary sheet: '{summary_sheet}'")

    # 2. Filter remaining sheets and Add Device Name
    valid_dataframes = []
    required_comparison_cols = ["Link Status New", "Protocol Status New"]

    for sheet in potential_data_sheets:
        # Create a copy to avoid modifying the original read data unexpectedly
        df = all_sheets_dict[sheet].copy()

        # Check if the required 'New' columns exist in this sheet
        missing_cols = [
            col for col in required_comparison_cols if col not in df.columns
        ]

        if missing_cols:
            print(f"Skipping sheet '{sheet}': Missing columns {missing_cols}")
        else:
            # --- NEW STEP: Add Device Name column from Sheet Name ---
            df["Device"] = sheet
            valid_dataframes.append(df)

    if not valid_dataframes:
        print("No valid data sheets found.")
        return

    print(f"Processing {len(valid_dataframes)} valid sheet(s)...")

    # 3. Combine valid sheets
    df_combined = pd.concat(valid_dataframes, ignore_index=True)

    col_interface = "Interface"

    # 4. Filter rows where Status has changed
    changes = df_combined[
        (df_combined["Link Status"] != df_combined["Link Status New"])
        | (df_combined["Protocol Status"] != df_combined["Protocol Status New"])
    ].copy()

    if changes.empty:
        print("No status changes found in the valid sheets.")
        return

    # 5. Rename and Format
    changes.rename(
        columns={
            "Link Status": "Link Status (Before)",
            "Link Status New": "Link Status (After)",
            "Protocol Status": "Protocol Status (Before)",
            "Protocol Status New": "Protocol Status (After)",
        },
        inplace=True,
    )

    # Updated Column Order: Device Name is first
    final_cols = [
        "Device",
        col_interface,
        "Description",
        "IP Address",
        "Link Status (Before)",
        "Link Status (After)",
        "Protocol Status (Before)",
        "Protocol Status (After)",
    ]

    # Ensure we only select columns that exist (in case Description/IP are missing in source)
    # But strictly based on your request, we enforce this order:
    changes = changes[final_cols]

    # 6. Save with Auto-sized columns
    print("Saving and adjusting column widths...")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        changes.to_excel(writer, index=False, sheet_name="Status Changes")

        worksheet = writer.sheets["Status Changes"]

        for i, column in enumerate(changes.columns):
            # Calculate max length of data or header
            column_len = (
                max(changes[column].astype(str).map(len).max(), len(column)) + 2
            )

            col_letter = get_column_letter(i + 1)
            worksheet.column_dimensions[col_letter].width = column_len

    print(f"Success! Created '{output_file}' with {len(changes)} rows.")


# --- Usage ---
input_filename = "res-dci.xlsx"
output_filename = "status_changes_report_dci_2.xlsx"

if __name__ == "__main__":
    process_interface_changes(input_filename, output_filename)
