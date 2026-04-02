
import re
from pathlib import Path
from typing import List, Optional

import matplotlib.pyplot as plt
import pandas as pd


DEFAULT_FILE = "heavy_ion_analysis_v2.xlsx"


def prompt_nonempty(message: str) -> str:
    while True:
        value = input(message).strip()
        if value:
            return value
        print("Please enter a value.")


def prompt_choice(message: str, choices: List[str]) -> str:
    normalized = {c.lower(): c for c in choices}
    while True:
        value = input(message).strip().lower()
        if value in normalized:
            return normalized[value]
        print(f"Please choose one of: {', '.join(choices)}")


def prompt_yes_no(message: str) -> bool:
    while True:
        value = input(message).strip().lower()
        if value in {"y", "yes"}:
            return True
        if value in {"n", "no"}:
            return False
        print("Please enter yes or no.")


def parse_sr_numbers(raw: str) -> List[int]:
    values = []
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            values.append(int(part))
        except ValueError:
            raise ValueError(f"Invalid SR number: {part}")
    if not values:
        raise ValueError("You must enter at least one SR number.")
    return values


def extract_let_from_ion(value) -> Optional[float]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    match = re.search(r"-\s*([0-9]*\.?[0-9]+)", text)
    if match:
        return float(match.group(1))
    try:
        return float(text)
    except ValueError:
        return None


def normalize_text(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.lower()


def load_data(file_path: str) -> pd.DataFrame:
    xls = pd.ExcelFile(file_path)

    preferred_sheet = None
    for name in xls.sheet_names:
        lowered = name.strip().lower()
        if "raw" in lowered or "data" in lowered:
            preferred_sheet = name
            break

    if preferred_sheet is None:
        preferred_sheet = xls.sheet_names[0]

    df = pd.read_excel(file_path, sheet_name=preferred_sheet)
    df.columns = [str(col).strip() for col in df.columns]

    required = {"vdd", "input", "ion", "freq", "actual_freq", "SR_NUM", "cs", "upper", "lower"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Missing required columns: {', '.join(sorted(missing))}\n"
            f"Columns found: {', '.join(df.columns)}"
        )

    df = df.copy()
    df["LET"] = df["ion"].apply(extract_let_from_ion)
    return df


def filter_equals(df: pd.DataFrame, column: str, target: str) -> pd.DataFrame:
    if column not in df.columns:
        raise ValueError(f"Column '{column}' is not in the spreadsheet.")

    original = df.copy()

    numeric_series = pd.to_numeric(df[column], errors="coerce")
    try:
        numeric_target = float(target)
        numeric_match = numeric_series.notna() & (numeric_series == numeric_target)
    except ValueError:
        numeric_match = pd.Series(False, index=df.index)

    text_match = normalize_text(df[column]) == target.strip().lower()
    filtered = df[numeric_match | text_match]

    if filtered.empty:
        sample_values = original[column].dropna().astype(str).unique()[:10]
        preview = ", ".join(sample_values) if len(sample_values) else "No non-empty values found"
        raise ValueError(
            f"No rows matched {column} = {target}.\n"
            f"Sample available values: {preview}"
        )
    return filtered


def build_title(x_axis_label: str, filters: dict, sr_numbers: List[int]) -> str:
    parts = [f"Cross Section vs {x_axis_label}"]
    parts.append(f"SR_NUM={','.join(map(str, sr_numbers))}")
    for key, value in filters.items():
        parts.append(f"{key}={value}")
    return " | ".join(parts)


def choose_output_name() -> str:
    raw = input("Output image file name [plot.png]: ").strip()
    return raw if raw else "plot.png"


def main() -> None:
    print("\nISDE Automatic Plotter - First Iteration")
    print("----------------------------------------")

    default_path = Path.cwd() / DEFAULT_FILE
    file_prompt = f"Excel file path [{default_path.name}]: "
    file_path = input(file_prompt).strip()
    if not file_path:
        file_path = str(default_path)

    df = load_data(file_path)

    print("\nAvailable x-axis choices: vdd, let, frq")
    x_axis = prompt_choice("Choose x-axis variable: ", ["vdd", "let", "frq"]).lower()

    raw_sr = prompt_nonempty("Enter SR_NUM values separated by commas (example: 5,10,12): ")
    sr_numbers = parse_sr_numbers(raw_sr)

    input_value = prompt_nonempty("Enter input value to filter by: ")

    scale_choice = prompt_choice("Choose y-axis scale (linear/log): ", ["linear", "log"]).lower()
    same_x_scale = prompt_yes_no("Use the same scale for the x-axis? (yes/no): ")
    x_scale_choice = scale_choice if same_x_scale else prompt_choice(
        "Choose x-axis scale (linear/log): ", ["linear", "log"]
    ).lower()

    filtered = df[df["SR_NUM"].isin(sr_numbers)].copy()
    if filtered.empty:
        raise ValueError("No rows matched the requested SR_NUM values.")

    filtered = filter_equals(filtered, "input", input_value)

    filters_for_title = {"input": input_value}

    if x_axis != "vdd":
        vdd_value = prompt_nonempty("Enter fixed VDD value: ")
        filtered = filter_equals(filtered, "vdd", vdd_value)
        filters_for_title["vdd"] = vdd_value

    if x_axis != "let":
        let_value = prompt_nonempty("Enter fixed LET value (taken from ion like NE-2 -> 2): ")
        numeric_let = pd.to_numeric(filtered["LET"], errors="coerce")
        try:
            let_target = float(let_value)
        except ValueError:
            raise ValueError("LET must be numeric.")
        filtered = filtered[numeric_let == let_target]
        if filtered.empty:
            raise ValueError(f"No rows matched LET = {let_target}.")
        filters_for_title["LET"] = let_target

    if x_axis != "frq":
        freq_value = prompt_nonempty("Enter fixed frequency using the freq column: ")
        filtered = filter_equals(filtered, "freq", freq_value)
        filters_for_title["freq"] = freq_value

    if filtered.empty:
        raise ValueError("No data remained after filtering.")

    x_column_map = {
        "vdd": "vdd",
        "let": "LET",
        "frq": "actual_freq",
    }
    x_label_map = {
        "vdd": "VDD",
        "let": "LET",
        "frq": "Frequency",
    }

    x_col = x_column_map[x_axis]
    x_label = x_label_map[x_axis]

    filtered = filtered.dropna(subset=[x_col, "cs", "upper", "lower"]).copy()
    if filtered.empty:
        raise ValueError("No plottable rows remained after removing missing values.")

    filtered[x_col] = pd.to_numeric(filtered[x_col], errors="coerce")
    filtered["cs"] = pd.to_numeric(filtered["cs"], errors="coerce")
    filtered["upper"] = pd.to_numeric(filtered["upper"], errors="coerce")
    filtered["lower"] = pd.to_numeric(filtered["lower"], errors="coerce")
    filtered = filtered.dropna(subset=[x_col, "cs", "upper", "lower"])

    if filtered.empty:
        raise ValueError("No numeric rows remained for plotting.")

    plt.figure(figsize=(10, 6))

    for sr_num in sr_numbers:
        sr_df = filtered[filtered["SR_NUM"] == sr_num].copy()
        if sr_df.empty:
            print(f"Warning: SR_NUM {sr_num} had no matching rows after filtering.")
            continue

        sr_df = sr_df.sort_values(by=x_col)
        yerr = [sr_df["lower"].to_numpy(), sr_df["upper"].to_numpy()]
        plt.errorbar(
            sr_df[x_col],
            sr_df["cs"],
            yerr=yerr,
            marker="o",
            linestyle="-",
            capsize=4,
            label=f"SR {sr_num}",
        )

    if not plt.gca().lines:
        raise ValueError("Nothing was plotted. Check your filters.")

    plt.xlabel(x_label)
    plt.ylabel("Cross Section")
    plt.title(build_title(x_label, filters_for_title, sr_numbers))
    plt.xscale(x_scale_choice)
    plt.yscale(scale_choice)
    plt.grid(True, which="both", alpha=0.3)
    plt.legend()
    plt.tight_layout()

    output_name = choose_output_name()
    plt.savefig(output_name, dpi=300)
    plt.show()

    print(f"\nPlot saved to: {Path(output_name).resolve()}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nError: {exc}")
