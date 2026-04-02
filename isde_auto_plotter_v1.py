import pandas as pd
import matplotlib.pyplot as plt


REQUIRED_COLUMNS = {"vdd", "input", "ion", "freq", "actual_freq", "sr_num", "cs", "upper", "lower"}


def normalize_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df


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
            return [int(x.strip()) for x in raw.split(",") if x.strip()]
        except ValueError:
            print("Please enter comma-separated integers like 5,10,12")


def prompt_choice(prompt_text, valid_choices):
    valid_choices = [x.lower() for x in valid_choices]
    while True:
        value = input(prompt_text).strip().lower()
        if value in valid_choices:
            return value
        print(f"Please enter one of: {', '.join(valid_choices)}")


def sheet_has_required_columns(file_path, sheet_name, header_row):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, nrows=5)
        df = normalize_columns(df)
        cols_lower = {str(c).strip().lower() for c in df.columns}
        return REQUIRED_COLUMNS.issubset(cols_lower)
    except Exception:
        return False


def find_candidate_sheets(file_path):
    xls = pd.ExcelFile(file_path)
    candidates = []

    for sheet in xls.sheet_names:
        for header_row in range(0, 20):
            if sheet_has_required_columns(file_path, sheet, header_row):
                candidates.append((sheet, header_row))
                break

    return candidates


def load_and_filter_sheet(file_path, sheet_name, header_row, x_axis, input_val, sr_nums,
                          vdd_val=None, let_val=None, freq_val=None):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
    except Exception:
        return pd.DataFrame()

    df = normalize_columns(df)
    colmap = {c.lower(): c for c in df.columns}

    df = df.rename(columns={
        colmap["vdd"]: "vdd",
        colmap["input"]: "input",
        colmap["ion"]: "ion",
        colmap["freq"]: "freq",
        colmap["actual_freq"]: "actual_freq",
        colmap["sr_num"]: "SR_NUM",
        colmap["cs"]: "cs",
        colmap["upper"]: "upper",
        colmap["lower"]: "lower",
    })

    df["LET"] = df["ion"].apply(extract_let_from_ion)
    df["input"] = df["input"].astype(str).str.strip()

    filtered = df.copy()
    filtered = filtered[filtered["input"] == input_val]
    filtered = filtered[filtered["SR_NUM"].isin(sr_nums)]

    if x_axis != "vdd" and vdd_val is not None:
        filtered = filtered[pd.to_numeric(filtered["vdd"], errors="coerce") == vdd_val]

    if x_axis != "let" and let_val is not None:
        filtered = filtered[pd.to_numeric(filtered["LET"], errors="coerce") == let_val]

    if x_axis != "frq" and freq_val is not None:
        filtered = filtered[pd.to_numeric(filtered["freq"], errors="coerce") == freq_val]

    if filtered.empty:
        return pd.DataFrame()

    filtered["source_sheet"] = sheet_name
    return filtered


def main():
    print("=== ISDE Automatic Plotter ===")
    file_path = input("Excel file path: ").strip()

    print("\nEnter plot settings first.\n")

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

    print("\nSearching across all sheets...")

    candidates = find_candidate_sheets(file_path)

    if not candidates:
        print("No sheets with the required columns were found.")
        return

    all_matches = []

    for sheet_name, header_row in candidates:
        filtered = load_and_filter_sheet(
            file_path=file_path,
            sheet_name=sheet_name,
            header_row=header_row,
            x_axis=x_axis,
            input_val=input_val,
            sr_nums=sr_nums,
            vdd_val=vdd_val,
            let_val=let_val,
            freq_val=freq_val,
        )
        if not filtered.empty:
            all_matches.append(filtered)

    if not all_matches:
        print("No matching data found across any sheets.")
        return

    combined = pd.concat(all_matches, ignore_index=True)

    if x_axis == "vdd":
        combined["x"] = pd.to_numeric(combined["vdd"], errors="coerce")
        xlabel = "VDD"
    elif x_axis == "let":
        combined["x"] = pd.to_numeric(combined["LET"], errors="coerce")
        xlabel = "LET"
    else:
        combined["x"] = pd.to_numeric(combined["actual_freq"], errors="coerce")
        xlabel = "Frequency"

    combined["y"] = pd.to_numeric(combined["cs"], errors="coerce")
    combined["lower"] = pd.to_numeric(combined["lower"], errors="coerce")
    combined["upper"] = pd.to_numeric(combined["upper"], errors="coerce")

    combined = combined.dropna(subset=["x", "y", "lower", "upper", "SR_NUM"])

    if combined.empty:
        print("Matching rows were found, but numeric plotting data was invalid after cleaning.")
        return

    combined = combined.sort_values(["SR_NUM", "x"])

    title_parts = [f"Input={input_val}", f"SR_NUM={','.join(map(str, sr_nums))}"]

    if x_axis != "vdd":
        title_parts.append(f"VDD={vdd_val}")
    if x_axis != "let":
        title_parts.append(f"LET={let_val}")
    if x_axis != "frq":
        title_parts.append(f"FRQ={freq_val}")

    plt.figure(figsize=(9, 6))

    plotted_any = False
    used_sheets = sorted(combined["source_sheet"].dropna().unique())

    for sr in sr_nums:
        sr_df = combined[combined["SR_NUM"] == sr].sort_values("x")
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

    print("\nUsed sheets:")
    for sheet in used_sheets:
        print(f"- {sheet}")

    plt.show()


if __name__ == "__main__":
    main()