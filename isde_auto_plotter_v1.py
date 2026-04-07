import os
import re
import pandas as pd
import xlwings as xw


REQUIRED_COLUMNS = {
    "vdd", "input", "ion", "freq", "actual_freq",
    "sr_num", "cs", "upper", "lower"
}


def norm(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def row_is_blank(values):
    return all(v is None or str(v).strip() == "" for v in values)


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
    valid_lower = [v.lower() for v in valid_choices]
    while True:
        value = input(prompt_text).strip().lower()
        if value in valid_lower:
            return value
        print(f"Please enter one of: {', '.join(valid_choices)}")


def prompt_float_or_all(prompt_text):
    while True:
        raw = input(prompt_text).strip().lower()
        if raw == "all":
            return "all"
        try:
            return float(raw)
        except ValueError:
            print("Please enter a valid number or 'all'.")


def prompt_input_value(prompt_text):
    raw = input(prompt_text).strip()
    if raw.lower() == "all":
        return "all"
    return raw


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


def parse_shift_register_token(token):
    token = token.strip().upper()
    if not token:
        return None

    m = re.fullmatch(r"([AZ])\-S(\d+)", token)
    if m:
        return {"prefix": m.group(1), "number": int(m.group(2)), "raw": token}

    m = re.fullmatch(r"S(\d+)", token)
    if m:
        return {"prefix": None, "number": int(m.group(1)), "raw": token}

    m = re.fullmatch(r"(\d+)", token)
    if m:
        return {"prefix": None, "number": int(m.group(1)), "raw": token}

    return None


def prompt_shift_registers(prompt_text):
    while True:
        raw = input(prompt_text).strip()
        if raw.lower() == "all":
            return "all"

        parts = [x.strip() for x in raw.split(",") if x.strip()]
        parsed = [parse_shift_register_token(x) for x in parts]

        if parts and all(p is not None for p in parsed):
            return parsed

        print("Enter comma-separated shift registers like: A-S5, Z-S10, S3, or type 'all'")


def find_tables_in_sheet_xlwings(ws):
    used = ws.used_range
    values = used.value

    if not values:
        return []

    if not isinstance(values, list):
        return []

    if values and not isinstance(values[0], list):
        values = [values]

    max_row = len(values)
    max_col = max(len(r) if isinstance(r, list) else 1 for r in values)

    def get_cell(r, c):
        try:
            row = values[r - 1]
            if not isinstance(row, list):
                row = [row]
            return row[c - 1] if c - 1 < len(row) else None
        except Exception:
            return None

    tables = []
    scanned_header_rows = set()

    for r in range(1, max_row + 1):
        row_vals = [get_cell(r, c) for c in range(1, max_col + 1)]
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
                row_dict[col_name] = get_cell(rr, col_idx)

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
            df["source_sheet"] = ws.name
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

    numeric_cols = ["vdd", "freq", "actual_freq", "cs", "upper", "lower"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df["sr_num_raw"] = df["sr_num"].astype(str).str.strip()
    extracted_sr_num = df["sr_num_raw"].str.extract(r"(\d+)", expand=False)
    df["sr_num_numeric"] = pd.to_numeric(extracted_sr_num, errors="coerce")

    extracted_prefix = df["sr_num_raw"].str.extract(r"\b([AZ])\-S\d+\b", expand=False)
    df["sr_prefix"] = extracted_prefix.str.upper()

    if "ion" in df.columns:
        df["LET"] = df["ion"].apply(extract_let_from_ion)
    else:
        df["LET"] = None

    return df


def float_matches(series, value, tol=1e-9):
    return series.notna() & ((series - value).abs() < tol)


def unique_sorted_non_null(values):
    cleaned = []
    for v in values:
        if pd.isna(v):
            continue
        if v not in cleaned:
            cleaned.append(v)
    try:
        return sorted(cleaned)
    except TypeError:
        return cleaned


def filter_table(df, x_axis, input_val, selected_srs, vdd_val=None, let_val=None, freq_val=None):
    filtered = df.copy()

    if str(input_val).strip().lower() != "all":
        filtered = filtered[filtered["input"] == str(input_val).strip()]

    if selected_srs != "all":
        selected_numbers = sorted(set(sr["number"] for sr in selected_srs))
        filtered = filtered[filtered["sr_num_numeric"].isin(selected_numbers)]

        has_prefix_data = filtered["sr_prefix"].notna().any()
        if has_prefix_data:
            allowed_pairs = {(sr["prefix"], sr["number"]) for sr in selected_srs if sr["prefix"] is not None}
            no_prefix_requests = {sr["number"] for sr in selected_srs if sr["prefix"] is None}

            keep_mask = []
            for _, row in filtered.iterrows():
                sr_num = int(row["sr_num_numeric"]) if pd.notna(row["sr_num_numeric"]) else None
                pair = (row["sr_prefix"], sr_num)

                keep = False
                if sr_num in no_prefix_requests:
                    keep = True
                if pair in allowed_pairs:
                    keep = True

                keep_mask.append(keep)

            filtered = filtered[pd.Series(keep_mask, index=filtered.index)]

    if x_axis != "vdd" and vdd_val != "all":
        filtered = filtered[float_matches(filtered["vdd"], vdd_val)]

    if x_axis != "let" and let_val != "all":
        filtered = filtered[float_matches(filtered["LET"], let_val)]

    if x_axis != "frq" and freq_val != "all":
        filtered = filtered[float_matches(filtered["freq"], freq_val)]

    return filtered


def format_vdd_title(vdd_val):
    if vdd_val is None or vdd_val == "all" or pd.isna(vdd_val):
        return None
    if vdd_val < 10:
        return f"{int(round(vdd_val * 1000))} mV"
    return f"{int(round(vdd_val))} mV"


def format_freq_title(freq_val):
    if freq_val is None or freq_val == "all" or pd.isna(freq_val):
        return None
    if float(freq_val).is_integer():
        return f"{int(freq_val)} MHz"
    return f"{freq_val} MHz"


def format_let_title(let_val):
    if let_val is None or let_val == "all" or pd.isna(let_val):
        return None
    if float(let_val).is_integer():
        return f"LET {int(let_val)}"
    return f"LET {let_val}"


def build_plot_title(x_axis, input_val, vdd_val=None, let_val=None, freq_val=None):
    title_parts = []

    if x_axis != "vdd" and vdd_val not in [None, "all"]:
        title_parts.append(format_vdd_title(vdd_val))
    if x_axis != "frq" and freq_val not in [None, "all"]:
        title_parts.append(format_freq_title(freq_val))
    if x_axis != "let" and let_val not in [None, "all"]:
        title_parts.append(format_let_title(let_val))

    if str(input_val).strip().lower() != "all":
        title_parts.append(f"Input {input_val}")
    else:
        title_parts.append("All Inputs")

    return " ".join([p for p in title_parts if p])


def build_series_label_from_group(sr_df, requested_srs, sr_num, x_axis):
    prefixes = sorted(set(x for x in sr_df["sr_prefix"].dropna().unique()))
    if len(prefixes) == 1:
        base = f"{prefixes[0]}-S{int(sr_num)}"
    elif len(prefixes) > 1:
        base = f"S{int(sr_num)}"
    else:
        if requested_srs != "all":
            matching_requests = [sr for sr in requested_srs if sr["number"] == int(sr_num)]
            if len(matching_requests) == 1 and matching_requests[0]["prefix"] is not None:
                base = matching_requests[0]["raw"]
            else:
                base = f"S{int(sr_num)}"
        else:
            base = f"S{int(sr_num)}"

    extras = []

    if x_axis != "vdd" and sr_df["vdd"].notna().any():
        extras.append(format_vdd_title(sr_df["vdd"].iloc[0]))
    if x_axis != "let" and sr_df["LET"].notna().any():
        extras.append(format_let_title(sr_df["LET"].iloc[0]))
    if x_axis != "frq" and sr_df["freq"].notna().any():
        extras.append(format_freq_title(sr_df["freq"].iloc[0]))
    if sr_df["input"].notna().any():
        extras.append(f"Input {str(sr_df['input'].iloc[0]).strip()}")

    extras = [x for x in extras if x]
    return f"{base} ({', '.join(extras)})" if extras else base


def choose_split_dimension(x_axis, input_val, vdd_val, let_val, freq_val):
    if str(input_val).strip().lower() == "all":
        return "input"
    if x_axis != "vdd" and vdd_val == "all":
        return "vdd"
    if x_axis != "frq" and freq_val == "all":
        return "freq"
    if x_axis != "let" and let_val == "all":
        return "LET"
    return None


def chart_title_for_subset(x_axis, subset_df, input_val, vdd_val, let_val, freq_val):
    title_parts = []

    if x_axis != "vdd":
        if vdd_val == "all":
            if subset_df["vdd"].notna().any():
                title_parts.append(format_vdd_title(subset_df["vdd"].iloc[0]))
        elif vdd_val is not None:
            title_parts.append(format_vdd_title(vdd_val))

    if x_axis != "frq":
        if freq_val == "all":
            if subset_df["freq"].notna().any():
                title_parts.append(format_freq_title(subset_df["freq"].iloc[0]))
        elif freq_val is not None:
            title_parts.append(format_freq_title(freq_val))

    if x_axis != "let":
        if let_val == "all":
            if subset_df["LET"].notna().any():
                title_parts.append(format_let_title(subset_df["LET"].iloc[0]))
        elif let_val is not None:
            title_parts.append(format_let_title(let_val))

    if str(input_val).strip().lower() == "all":
        if subset_df["input"].notna().any():
            title_parts.append(f"Input {subset_df['input'].iloc[0]}")
    else:
        title_parts.append(f"Input {input_val}")

    return " ".join([p for p in title_parts if p])


def last_used_col(ws):
    return ws.used_range.last_cell.column


def clear_old_autoplotter_objects(ws):
    try:
        for ch in list(ws.charts):
            try:
                if str(ch.name).startswith("AutoPlotter_"):
                    ch.delete()
            except Exception:
                pass
    except Exception:
        pass


def write_helper_block(ws, start_col, start_row, series_dicts):
    """
    Write helper data as repeated X/Y column pairs:

    col A: series1_x
    col B: series1_y
    col C: series2_x
    col D: series2_y
    ...

    This layout works much better with Excel scatter charts when using
    chart.set_source_data(...).
    """
    current_col = start_col
    max_end_row = start_row

    nonempty = [s for s in series_dicts if not s["df"].empty]
    if not nonempty:
        return None

    for s in nonempty:
        df = s["df"].sort_values("x").copy()
        label = s["label"]

        x_col = current_col
        y_col = current_col + 1

        ws.range((start_row, x_col)).value = f"{label}_x"
        ws.range((start_row, y_col)).value = label

        row_ptr = start_row + 1
        for _, row in df.iterrows():
            if pd.notna(row["x"]) and pd.notna(row["y"]):
                ws.range((row_ptr, x_col)).value = float(row["x"])
                ws.range((row_ptr, y_col)).value = float(row["y"])
                row_ptr += 1

        max_end_row = max(max_end_row, row_ptr - 1)
        current_col += 2

    return {
        "start_row": start_row,
        "end_row": max_end_row,
        "start_col": start_col,
        "end_col": current_col - 1,
    }

def build_series_dicts(combined, selected_srs, x_axis, split_dim=None):
    group_cols = ["sr_num_numeric", "sr_prefix"]

    if x_axis != "vdd":
        group_cols.append("vdd")
    if x_axis != "let":
        group_cols.append("LET")
    if x_axis != "frq":
        group_cols.append("freq")

    group_cols.append("input")

    if split_dim is not None and split_dim in group_cols:
        group_cols.remove(split_dim)

    series_dicts = []
    grouped = combined.groupby(group_cols, dropna=False)

    for _, sr_df in grouped:
        sr_df = sr_df.sort_values("x").copy()
        if sr_df.empty:
            continue

        sr_num = sr_df["sr_num_numeric"].iloc[0]
        label = build_series_label_from_group(sr_df, selected_srs, sr_num, x_axis)

        series_dicts.append({
            "label": label,
            "df": sr_df
        })

    return series_dicts


def add_scatter_chart_mac(ws, helper_meta, chart_title, x_axis, scale, anchor_left=500, anchor_top=20):
    chart = ws.charts.add(left=anchor_left, top=anchor_top, width=900, height=450)
    try:
        chart.name = f"AutoPlotter_{abs(hash((ws.name, chart_title, anchor_left, anchor_top))) % 10**8}"
    except Exception:
        pass

    src = ws.range(
        (helper_meta["start_row"], helper_meta["start_col"]),
        (helper_meta["end_row"], helper_meta["end_col"])
    )

    chart.set_source_data(src)
    chart.chart_type = "xy_scatter_lines"

    api_raw = chart.api
    api_chart = api_raw[1] if isinstance(api_raw, tuple) else api_raw

    try:
        api_chart.has_title.set(True)
        api_chart.chart_title.characters().text.set(chart_title)
    except Exception:
        pass

    try:
        cat_axis = api_chart.axes(1)
        val_axis = api_chart.axes(2)

        cat_axis.has_title.set(True)
        cat_axis.axis_title.characters().text.set(
            "VDD" if x_axis == "vdd" else ("LET (MeV-cm²/mg)" if x_axis == "let" else "Frequency (MHz)")
        )

        val_axis.has_title.set(True)
        val_axis.axis_title.characters().text.set("SEU cross-section (cm²)")
    except Exception:
        pass

    if scale == "log":
        try:
            api_chart.axes(2).scale_type.set(-4133)
        except Exception:
            pass
        if x_axis == "frq":
            try:
                api_chart.axes(1).scale_type.set(-4133)
            except Exception:
                pass

    return chart

def split_and_plot_on_same_sheet(ws, combined, selected_srs, x_axis, scale, input_val, vdd_val, let_val, freq_val):
    split_dim = choose_split_dimension(x_axis, input_val, vdd_val, let_val, freq_val)

    clear_old_autoplotter_objects(ws)

    helper_start_col = last_used_col(ws) + 3
    helper_row_ptr = 1

    if split_dim is None:
        plot_title = build_plot_title(x_axis, input_val, vdd_val, let_val, freq_val)

        series_dicts = build_series_dicts(combined, selected_srs, x_axis, None)
        helper_meta = write_helper_block(ws, helper_start_col, helper_row_ptr, series_dicts)

        if not helper_meta:
            return 0

        anchor_left = max(500, ws.range((1, helper_start_col)).left + 50)
        add_scatter_chart_mac(ws, helper_meta, plot_title, x_axis, scale, anchor_left, 20)
        return 1

    values = unique_sorted_non_null(combined[split_dim].unique())
    if not values:
        return 0

    chart_count = 0
    chart_top = 20

    for value in values:
        subset_df = combined[combined[split_dim] == value].copy()
        if subset_df.empty:
            continue

        subset_df = subset_df.sort_values(["sr_num_numeric", "input", "x"])

        chart_title = chart_title_for_subset(x_axis, subset_df, input_val, vdd_val, let_val, freq_val)
        series_dicts = build_series_dicts(subset_df, selected_srs, x_axis, split_dim)
        helper_meta = write_helper_block(ws, helper_start_col, helper_row_ptr, series_dicts)

        if not helper_meta:
            continue

        anchor_left = max(500, ws.range((1, helper_start_col)).left + 50)
        add_scatter_chart_mac(ws, helper_meta, chart_title, x_axis, scale, anchor_left, chart_top)

        chart_count += 1
        chart_top += 470
        helper_row_ptr = helper_meta["end_row"] + 6

    return chart_count

def main():
    print("=== ISDE Automatic Plotter (Mac xlwings) ===")

    file_path = input("Excel file path: ").strip()
    if not os.path.exists(file_path):
        print("File not found.")
        return

    app = None
    wb = None

    try:
        app = xw.App(visible=True)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(file_path)

        selected_sheet_name = prompt_sheet_choice([s.name for s in wb.sheets])
        ws = wb.sheets[selected_sheet_name]

        print("\nEnter plot settings:\n")
        x_axis = prompt_choice("Choose x-axis (vdd, let, frq): ", ["vdd", "let", "frq"])
        input_val = prompt_input_value("Input value (or type 'all'): ")
        selected_srs = prompt_shift_registers("Shift registers (comma separated, e.g. A-S5, Z-S10, or type 'all'): ")
        scale = prompt_choice("Scale (linear/log): ", ["linear", "log"])

        vdd_val = None
        let_val = None
        freq_val = None

        if x_axis != "vdd":
            vdd_val = prompt_float_or_all("Specify VDD (or type 'all'): ")
        if x_axis != "let":
            let_val = prompt_float_or_all("Specify LET (or type 'all'): ")
        if x_axis != "frq":
            freq_val = prompt_float_or_all("Specify freq (or type 'all'): ")

        print(f"\nScanning sheet '{selected_sheet_name}' for matching embedded tables...")

        all_tables = find_tables_in_sheet_xlwings(ws)
        if not all_tables:
            print("No tables with the required columns were found in that sheet.")
            return

        matched_frames = []
        for raw_df in all_tables:
            try:
                df = clean_table(raw_df)
                filtered = filter_table(df, x_axis, input_val, selected_srs, vdd_val, let_val, freq_val)
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
        elif x_axis == "let":
            combined["x"] = combined["LET"]
        else:
            combined["x"] = combined["actual_freq"]

        combined["y"] = combined["cs"]
        combined = combined.dropna(subset=["x", "y", "lower", "upper", "sr_num_numeric", "input"])

        if combined.empty:
            print("Matching data was found, but plotting columns were invalid after cleaning.")
            return

        combined = combined.sort_values(["sr_num_numeric", "input", "x"])

        used_locations = combined[["source_sheet", "header_row"]].drop_duplicates()
        print("\nMatched table locations:")
        for _, row in used_locations.iterrows():
            print(f"- Sheet: {row['source_sheet']}, header row: {int(row['header_row'])}")

        chart_count = split_and_plot_on_same_sheet(
            ws, combined, selected_srs, x_axis, scale, input_val, vdd_val, let_val, freq_val
        )

        if chart_count == 0:
            print("No plottable chart was created.")
            return

        app.screen_updating = True
        wb.save()

        print(f"\nDone. Inserted {chart_count} chart(s) into the original workbook.")
        print(file_path)

    finally:
        try:
            if wb is not None:
                wb.save()
                wb.close()
        except Exception:
            pass
        try:
            if app is not None:
                app.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
