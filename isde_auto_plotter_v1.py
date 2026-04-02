import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

REQUIRED_COLUMNS = {"vdd", "input", "ion", "freq", "actual_freq", "SR_NUM", "cs", "upper", "lower"}

def find_valid_sheets(file_path):
    xls = pd.ExcelFile(file_path)
    valid_sheets = []

    for sheet in xls.sheet_names:
        for header_row in range(5):  # try first 5 rows
            try:
                df = pd.read_excel(file_path, sheet_name=sheet, header=header_row)
                cols = set(df.columns.astype(str).str.strip())
                if REQUIRED_COLUMNS.issubset(cols):
                    valid_sheets.append((sheet, header_row))
                    break
            except:
                continue

    return valid_sheets

def extract_let(ion_series):
    return ion_series.str.extract(r'-(\d+)').astype(float)

def main():
    file_path = input("Excel file path: ").strip()

    valid_sheets = find_valid_sheets(file_path)

    if not valid_sheets:
        print("No valid data sheets found.")
        return

    print("\nAvailable data sheets:")
    for i, (sheet, header) in enumerate(valid_sheets):
        print(f"{i}: {sheet} (header row {header})")

    choice = int(input("Select sheet number: "))
    sheet_name, header_row = valid_sheets[choice]

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

    df.columns = df.columns.str.strip()

    df["LET"] = extract_let(df["ion"])

    x_axis = input("Choose x-axis (vdd, let, frq): ").strip().lower()
    input_val = input("Input value: ").strip()
    sr_nums = list(map(int, input("SR_NUMs (comma separated): ").split(",")))

    scale = input("Scale (linear/log): ").strip().lower()

    filtered = df[df["input"] == input_val]
    filtered = filtered[filtered["SR_NUM"].isin(sr_nums)]

    if x_axis != "vdd":
        vdd_val = float(input("Specify VDD: "))
        filtered = filtered[filtered["vdd"] == vdd_val]

    if x_axis != "let":
        let_val = float(input("Specify LET: "))
        filtered = filtered[filtered["LET"] == let_val]

    if x_axis != "frq":
        freq_val = float(input("Specify freq: "))
        filtered = filtered[filtered["freq"] == freq_val]

    if filtered.empty:
        print("No data after filtering.")
        return

    if x_axis == "vdd":
        x = filtered["vdd"]
    elif x_axis == "let":
        x = filtered["LET"]
    else:
        x = filtered["actual_freq"]

    y = filtered["cs"]
    yerr = [filtered["lower"], filtered["upper"]]

    plt.errorbar(x, y, yerr=yerr, fmt='o', capsize=5)

    if scale == "log":
        plt.xscale("log")
        plt.yscale("log")

    title_parts = []
    if x_axis != "vdd":
        title_parts.append(f"VDD={vdd_val}")
    if x_axis != "let":
        title_parts.append(f"LET={let_val}")
    if x_axis != "frq":
        title_parts.append(f"FRQ={freq_val}")

    plt.title("Cross Section Plot | " + ", ".join(title_parts))
    plt.xlabel(x_axis.upper())
    plt.ylabel("Cross Section")

    plt.grid(True)
    plt.show()

if __name__ == "__main__":
    main()