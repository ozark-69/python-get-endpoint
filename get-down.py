import pandas as pd
import sys
from openpyxl.utils import get_column_letter


def process_down_ports(input_file, output_file):
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
    # We need these columns to check the status
    required_cols = ["Link Status", "Link Status New"]

    for sheet in potential_data_sheets:
        df = all_sheets_dict[sheet].copy()

        # Check if the required columns exist
        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            print(f"Skipping sheet '{sheet}': Missing columns {missing_cols}")
        else:
            # Add Device Name column
            df["Device Name"] = sheet
            valid_dataframes.append(df)

    if not valid_dataframes:
        print("No valid data sheets found.")
        return

    print(f"Processing {len(valid_dataframes)} valid sheet(s)...")

    # 3. Combine valid sheets
    df_combined = pd.concat(valid_dataframes, ignore_index=True)

    # 4. Filter for ports that are DOWN Before AND After
    # We strip whitespace just in case ' down' or 'down ' exists
    down_ports = df_combined[
        (df_combined["Link Status"].astype(str).str.strip() == "down")
        & (df_combined["Link Status New"].astype(str).str.strip() == "down")
    ].copy()

    if down_ports.empty:
        print("No ports found that are down in both states.")
        return

    # 5. Rename and Format
    down_ports.rename(
        columns={
            "Link Status": "Link Status (Before)",
            "Link Status New": "Link Status (After)",
            "Protocol Status": "Protocol Status (Before)",
            "Protocol Status New": "Protocol Status (After)",
        },
        inplace=True,
    )

    # Define columns to keep
    col_interface = "Interface"
    final_cols = [
        "Device Name",
        col_interface,
        "Description",
        "IP Address",
        "Link Status (Before)",
        "Link Status (After)",
        "Protocol Status (Before)",
        "Protocol Status (After)",
    ]

    # Select columns (ensure they exist in df)
    # This handles cases where optional cols like Description might be missing
    available_cols = [c for c in final_cols if c in down_ports.columns]
    down_ports = down_ports[available_cols]

    # 6. Save with Auto-sized columns
    print("Saving and adjusting column widths...")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        down_ports.to_excel(writer, index=False, sheet_name="Down Ports")

        worksheet = writer.sheets["Down Ports"]

        for i, column in enumerate(down_ports.columns):
            # Calculate max length
            column_len = (
                max(down_ports[column].astype(str).map(len).max(), len(column)) + 2
            )

            col_letter = get_column_letter(i + 1)
            worksheet.column_dimensions[col_letter].width = column_len

    print(f"Success! Created '{output_file}' with {len(down_ports)} rows.")


# --- Usage ---
input_filename = "res-dc.xlsx"
output_filename = "ports-down-dc.xlsx"

if __name__ == "__main__":
    process_down_ports(input_filename, output_filename)
