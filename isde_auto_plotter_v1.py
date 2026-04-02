import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook


REQUIRED_COLUMNS = {
    "vdd", "input", "ion", "freq", "actual_freq",
    "sr_num", "cs", "upper", "lower"
}


def norm(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def extract_let_from_ion(ion_value):
    if pd.isna(ion_value):
        return None
    ion_str = str(ion_value).strip()
    if "-" not in ion_str:
        return None
    try:
        return float(ion_str.split("-")[-1])
    except ValueError:
        return None


def prompt_choice(prompt_text, valid_choices):
    valid = [v.lower() for v in valid_choices]
    while True:
        value = input(prompt_text).strip().lower()
        if value in valid:
            return value
        print(f"Please enter one of: {', '.join(valid_choices)}")


def prompt_float(prompt_text):
    while True:
        raw = input(prompt_text).strip()
        try:
            return float(raw)
        except ValueError:
            print("Please enter a valid number.")


def prompt_int_list(prompt_text):
    while True:
        raw = input(prompt_text).strip()
        try:
            values = [int(x.strip()) for x in raw.split(",") if x.strip()]
            if values:
                return values
            print("Please enter at least one SR number.")
        except ValueError:
            print("Please enter comma-separated integers like 5,10,12")


def prompt_sheet_choice(sheet_names):
    print("\nAvailable sheets:")
    for i, name in enumerate(sheet_names):
        print(f"{i}: {name}")

    while True:
        raw = input("Choose sheet number: ").strip()
        try:
            idx = int(raw)
            if 0 <= idx < len(sheet_names):
                return sheet_names[idx]
            print("Invalid sheet number.")
        except ValueError:
            print("Please enter a valid integer.")


def row_is_blank(values):
    return all(v is None or str(v).strip() == "" for v in values)


def find_tables_in_sheet(ws):
    tables = []
    max_row = ws.max_row
    max_col = ws.max_column
    scanned_header_rows = set()

    for r in range(1, max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        normalized = [norm(v) for v in row_vals]

        present = set(v for v in normalized if v)
        if not REQUIRED_COLUMNS.issubset(present):
            continue

        if r in scanned_header_rows:
            continue

        header_positions = {}
        for c_idx, cell_val in enumerate(normalized, start=1):
            if cell_val in REQUIRED_COLUMNS and cell_val not in header_positions:
                header_positions[cell_val] = c_idx

        if not REQUIRED_COLUMNS.issubset(set(header_positions.keys())):
            continue

        scanned_header_rows.add(r)

        data_rows = []
        blank_streak = 0

        for rr in range(r + 1, max_row + 1):
            row_dict = {}
            for col_name, col_idx in header_positions.items():
                row_dict[col_name] = ws.cell(rr, col_idx).value

            if row_is_blank(list(row_dict.values())):
                blank_streak += 1
                if blank_streak >= 2:
                    break
                continue
            else:
                blank_streak = 0

            data_rows.append(row_dict)

        if data_rows:
            df = pd.DataFrame(data_rows)
            df["source_sheet"] = ws.title
            df["header_row"] = r
            tables.append(df)

    return tables


def clean_table(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {}
    for c in df.columns:
        lc = c.strip().lower()
        if lc in REQUIRED_COLUMNS:
            rename_map[c] = lc
    df = df.rename(columns=rename_map)

    if "input" in df.columns:
        df["input"] = df["input"].astype(str).str.strip()

    for col in ["vdd", "freq", "actual_freq", "sr_num", "cs", "upper", "lower"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "ion" in df.columns:
        df["LET"] = df["ion"].apply(extract_let_from_ion)
    else:
        df["LET"] = None

    return df


def filter_table(df, x_axis, input_val, sr_nums, vdd_val=None, let_val=None, freq_val=None):
    filtered = df.copy()

    filtered = filtered[filtered["input"] == input_val]
    filtered = filtered[filtered["sr_num"].isin(sr_nums)]

    if x_axis != "vdd":
        filtered = filtered[filtered["vdd"] == vdd_val]

    if x_axis != "let":
        filtered = filtered[filtered["LET"] == let_val]

    if x_axis != "frq":
        filtered = filtered[filtered["freq"] == freq_val]

    return filtered


def main():
    print("=== ISDE Automatic Plotter ===")

    file_path = input("Excel file path: ").strip()

    try:
        wb = load_workbook(file_path, data_only=True)
    except Exception as e:
        print(f"Could not open workbook: {e}")
        return

    selected_sheet_name = prompt_sheet_choice(wb.sheetnames)
    ws = wb[selected_sheet_name]

    print("\nEnter all plot settings.\n")

    x_axis = prompt_choice("Choose x-axis (vdd, let, frq): ", ["vdd", "let", "frq"])
    input_val = input("Input value: ").strip()
    sr_nums = prompt_int_list("SR_NUMs (comma separated, e.g. 5,10,12): ")
    scale = prompt_choice("Scale (linear/log): ", ["linear", "log"])

    vdd_val = None
    let_val = None
    freq_val = None

    if x_axis != "vdd":
        vdd_val = prompt_float("Specify VDD: ")

    if x_axis != "let":
        let_val = prompt_float("Specify LET: ")

    if x_axis != "frq":
        freq_val = prompt_float("Specify freq: ")

    print(f"\nScanning sheet '{selected_sheet_name}' for matching embedded tables...")

    all_tables = find_tables_in_sheet(ws)

    if not all_tables:
        print("No tables with the required columns were found in that sheet.")
        return

    matched_frames = []

    for raw_df in all_tables:
        try:
            df = clean_table(raw_df)

            needed = {
                "vdd", "input", "ion", "freq", "actual_freq",
                "sr_num", "cs", "upper", "lower",
                "source_sheet", "header_row", "LET"
            }
            if not needed.issubset(set(df.columns)):
                continue

            filtered = filter_table(
                df,
                x_axis=x_axis,
                input_val=input_val,
                sr_nums=sr_nums,
                vdd_val=vdd_val,
                let_val=let_val,
                freq_val=freq_val
            )

            if not filtered.empty:
                matched_frames.append(filtered)

        except Exception:
            continue

    if not matched_frames:
        print("No matching data found in the selected sheet.")
        return

    combined = pd.concat(matched_frames, ignore_index=True)

    if x_axis == "vdd":
        combined["x"] = combined["vdd"]
        xlabel = "VDD"
    elif x_axis == "let":
        combined["x"] = combined["LET"]
        xlabel = "LET"
    else:
        combined["x"] = combined["actual_freq"]
        xlabel = "Frequency"

    combined["y"] = combined["cs"]
    combined = combined.dropna(subset=["x", "y", "lower", "upper", "sr_num"])

    if combined.empty:
        print("Matching data was found, but plotting columns were invalid after cleaning.")
        return

    combined = combined.sort_values(["sr_num", "x"])

    title_parts = [f"Input={input_val}", f"SR={','.join(map(str, sr_nums))}"]
    if x_axis != "vdd":
        title_parts.append(f"VDD={vdd_val}")
    if x_axis != "let":
        title_parts.append(f"LET={let_val}")
    if x_axis != "frq":
        title_parts.append(f"FRQ={freq_val}")

    plt.figure(figsize=(9, 6))

    plotted_any = False
    for sr in sr_nums:
        sr_df = combined[combined["sr_num"] == sr].sort_values("x")
        if sr_df.empty:
            continue

        plt.errorbar(
            sr_df["x"],
            sr_df["y"],
            yerr=[sr_df["lower"], sr_df["upper"]],
            fmt="o-",
            capsize=4,
            label=f"SR{sr}"
        )
        plotted_any = True

    if not plotted_any:
        print("No plottable data found.")
        return

    if scale == "log":
        plt.yscale("log")
        if x_axis == "frq":
            plt.xscale("log")

    plt.xlabel(xlabel)
    plt.ylabel("Cross Section")
    plt.title("Cross Section Plot | " + ", ".join(title_parts))
    plt.grid(True, which="both", linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()

    used_locations = combined[["source_sheet", "header_row"]].drop_duplicates()
    print("\nMatched table locations:")
    for _, row in used_locations.iterrows():
        print(f"- Sheet: {row['source_sheet']}, header row: {int(row['header_row'])}")

    plt.show()


if __name__ == "__main__":
    main()