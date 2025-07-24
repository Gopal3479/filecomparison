def create_side_by_side_sheet(self, df1, df2, output_wb):
    """Optimized side-by-side comparison with total row."""
    ws = output_wb.create_sheet("Side by Side Comparison")
    common_cols = list(set(df1.columns) & set(df2.columns))
    
    num_cols1 = [col for col in df1.columns if self.is_numeric(df1[col])]
    num_cols2 = [col for col in df2.columns if self.is_numeric(df2[col])]
    common_num_cols = list(set(num_cols1) & set(num_cols2))
    
    header_row = list(df1.columns) + [" | "] + list(df2.columns) + ["Match Status"]
    ws.append(header_row)
    
    # Create dataframes for summation, excluding total rows
    df1_for_sum, df2_for_sum = df1, df2
    if self.total_row_identifier:
        if not df1.empty:
            df1_for_sum = df1[~df1.iloc[:, 0].astype(str).str.contains(self.total_row_identifier, case=False, na=False)]
        if not df2.empty:
            df2_for_sum = df2[~df2.iloc[:, 0].astype(str).str.contains(self.total_row_identifier, case=False, na=False)]
    
    # Create total row using filtered data
    total_row, total_match = [], True
    for col in df1.columns:
        total_row.append(df1_for_sum[col].sum() if col in num_cols1 else "")
    total_row.append(" | ")
    for col in df2.columns:
        total_row.append(df2_for_sum[col].sum() if col in num_cols2 else "")
    
    for col in common_num_cols:
        if not self.are_equal(df1_for_sum[col].sum(), df2_for_sum[col].sum()):
            total_match = False
            break
    total_row.append("Matched" if total_match else "Not Matched")
    ws.append(total_row)
    
    # Apply formatting to header and total rows
    for cell in ws[1]: cell.fill, cell.font, cell.border = HEADER_FILL, Font(bold=True), THIN_BORDER
    for cell in ws[2]: cell.fill, cell.font, cell.border = TOTAL_ROW_FILL, Font(bold=True), THIN_BORDER
    
    # Precompute indices and comparison results
    separator_col = len(df1.columns) + 1
    match_status_col = len(header_row)
    file2_start_col = len(df1.columns) + 2
    
    match_status, diff_positions, is_missing_list = [], [], []
    for i in range(max(len(df1), len(df2))):
        row_match, diffs_in_row = True, []
        if i < len(df1) and i < len(df2):
            for col in common_cols:
                val1, val2 = df1.at[i, col], df2.at[i, col]
                if not self.are_equal(val1, val2):
                    row_match = False
                    col_idx1 = df1.columns.get_loc(col) + 1
                    col_idx2 = df2.columns.get_loc(col) + file2_start_col
                    diffs_in_row.append((col_idx1, col_idx2))
            is_missing = False
        else:
            row_match = False
            is_missing = True
        
        match_status.append("Matched" if row_match else "Not Matched")
        diff_positions.append(diffs_in_row)
        is_missing_list.append(is_missing)
    
    # Write data rows
    for i in range(max(len(df1), len(df2))):
        row_data = []
        if i < len(df1): 
            row_data.extend(df1.iloc[i].values)
        else: 
            row_data.extend([""] * len(df1.columns))
        row_data.append(" | ")
        if i < len(df2): 
            row_data.extend(df2.iloc[i].values)
        else: 
            row_data.extend([""] * len(df2.columns))
        row_data.append(match_status[i])
        ws.append(row_data)

        row_idx = i + 3
        
        # Apply borders to all cells
        for col_idx in range(1, len(header_row) + 1):
            ws.cell(row=row_idx, column=col_idx).border = THIN_BORDER

        # Apply row-level formatting
        if match_status[i] == "Matched":
            # Green background for fully matched rows
            for col_idx in range(1, len(header_row) + 1):
                ws.cell(row=row_idx, column=col_idx).fill = ROW_MATCH_FILL
        elif is_missing_list[i]:
            # Gray background for missing rows
            for col_idx in range(1, len(header_row) + 1):
                ws.cell(row=row_idx, column=col_idx).fill = ROW_MISSING_FILL
        
        # Format separator column
        ws.cell(row=row_idx, column=separator_col).fill = PatternFill(
            start_color="000000", end_color="000000", fill_type="solid"
        )
        
        # Highlight differing cells in red for non-missing rows
        if match_status[i] == "Not Matched" and not is_missing_list[i]:
            for col1, col2 in diff_positions[i]:
                ws.cell(row=row_idx, column=col1).fill = CELL_DIFF_FILL
                ws.cell(row=row_idx, column=col2).fill = CELL_DIFF_FILL

    # Highlight differences in total row
    if not total_match:
        for col in common_num_cols:
            if not self.are_equal(df1_for_sum[col].sum(), df2_for_sum[col].sum()):
                col_idx1 = df1.columns.get_loc(col) + 1
                col_idx2 = df2.columns.get_loc(col) + file2_start_col
                ws.cell(row=2, column=col_idx1).fill = CELL_DIFF_FILL
                ws.cell(row=2, column=col_idx2).fill = CELL_DIFF_FILL
    
    # Set column widths
    for col_idx, column in enumerate(ws.columns, 1):
        max_length = max(len(str(cell.value or "")) for cell in column[:100])
        ws.column_dimensions[get_column_letter(col_idx)].width = (max_length + 2) * 1.2
    
    return len([s for s in match_status if s == "Matched"])
