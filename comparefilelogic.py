def create_side_by_side_sheet(self, df1, df2, output_wb):
    """Create side-by-side comparison sheet with single Match Status column at end"""
    # Create sheet
    ws = output_wb.create_sheet("Side by Side Comparison")
    
    # Get common columns
    common_cols = list(set(df1.columns) & set(df2.columns))
    
    # Get all string columns (excluding dates)
    str_cols1 = self.get_string_columns(df1)
    str_cols2 = self.get_string_columns(df2)
    all_str_cols = list(set(str_cols1) | set(str_cols2))
    
    # Create concatenation keys using all string columns
    concat_keys1 = {}
    concat_keys2 = {}
    
    # Create keys for df1
    for idx, row in df1.iterrows():
        key_parts = []
        for col in all_str_cols:
            if col in df1.columns:
                val = row[col]
                if pd.notna(val) and not self.is_date(val):
                    key_parts.append(str(val))
        concat_keys1[idx] = "_".join(key_parts) if key_parts else None
    
    # Create keys for df2
    for idx, row in df2.iterrows():
        key_parts = []
        for col in all_str_cols:
            if col in df2.columns:
                val = row[col]
                if pd.notna(val) and not self.is_date(val):
                    key_parts.append(str(val))
        concat_keys2[idx] = "_".join(key_parts) if key_parts else None
    
    # Create sets of keys for matching
    keys1_set = set(concat_keys1.values())
    keys2_set = set(concat_keys2.values())
    
    # Create match status for File1
    file1_match_status = []
    for idx in df1.index:
        key = concat_keys1[idx]
        if key is not None and key in keys2_set:
            file1_match_status.append("Matched")
        else:
            file1_match_status.append("Not Matched")
    
    # Write headers - File1 columns + separator + File2 columns + Match Status
    header_row = list(df1.columns) + ["|"] + list(df2.columns) + ["Match Status"]
    ws.append(header_row)
    
    # Apply header styling
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True)
        cell.border = THIN_BORDER
    
    # Special styling for separator column
    separator_col = len(df1.columns) + 1
    ws.cell(row=1, column=separator_col).fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    ws.cell(row=1, column=separator_col).font = Font(color="FFFFFF", bold=True)
    ws.cell(row=1, column=separator_col).alignment = Alignment(horizontal="center")
    
    # Style Match Status header
    match_status_col = len(header_row)
    ws.cell(row=1, column=match_status_col).fill = HEADER_FILL
    ws.cell(row=1, column=match_status_col).font = Font(bold=True)
    
    # Write data row by row
    max_rows = max(len(df1), len(df2))
    for i in range(max_rows):
        row_data = []
        
        # Add File1 data if exists
        if i < len(df1):
            row_data.extend(df1.iloc[i].tolist())
        else:
            row_data.extend([""] * len(df1.columns))
        
        # Add separator
        row_data.append("|")
        
        # Add File2 data if exists
        if i < len(df2):
            row_data.extend(df2.iloc[i].tolist())
        else:
            row_data.extend([""] * len(df2.columns))
        
        # Add Match Status (only for File1 rows)
        if i < len(file1_match_status):
            row_data.append(file1_match_status[i])
        else:
            row_data.append("")
        
        ws.append(row_data)
        
        # Apply row highlighting based on File1's match status
        if self.highlight_row_matches and i < len(file1_match_status):
            file1_match = file1_match_status[i]
            
            # Calculate column ranges
            file1_start_col = 1
            file1_end_col = len(df1.columns)
            file2_start_col = file1_end_col + 2  # +1 for separator, +1 for next column
            file2_end_col = file2_start_col + len(df2.columns) - 1
            
            # Apply styling to File1 columns
            if file1_match == "Matched":
                fill = ROW_MATCH_FILL
            else:
                fill = ROW_MISSING_FILL
                
            for col_idx in range(file1_start_col, file1_end_col + 1):
                ws.cell(row=i+2, column=col_idx).fill = fill
            
            # Apply same styling to File2 columns
            for col_idx in range(file2_start_col, file2_end_col + 1):
                ws.cell(row=i+2, column=col_idx).fill = fill
            
            # Style Match Status column
            ws.cell(row=i+2, column=match_status_col).fill = fill
        
        # Apply cell difference highlighting
        if self.highlight_cell_diffs:
            if i < len(df1) and i < len(df2):
                for col in common_cols:
                    val1 = df1.at[i, col]
                    val2 = df2.at[i, col]
                    
                    if not self.are_equal(val1, val2):
                        col_idx1 = list(df1.columns).index(col) + 1
                        col_idx2 = list(df2.columns).index(col) + len(df1.columns) + 2  # +1 for separator, +1 for next column
                        
                        ws.cell(row=i+2, column=col_idx1).fill = CELL_DIFF_FILL
                        ws.cell(row=i+2, column=col_idx2).fill = CELL_DIFF_FILL
        
        # Style the separator column
        ws.cell(row=i+2, column=separator_col).fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        ws.cell(row=i+2, column=separator_col).font = Font(color="FFFFFF", bold=True)
        ws.cell(row=i+2, column=separator_col).alignment = Alignment(horizontal="center")
        
        # Apply borders
        for col_idx in range(1, len(header_row) + 1):
            ws.cell(row=i+2, column=col_idx).border = THIN_BORDER
    
    # Auto-size columns
    for col_idx in range(1, len(header_row) + 1):
        max_length = 0
        col_letter = get_column_letter(col_idx)
        
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
            for cell in row:
                try:
                    if cell.value is not None:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
        
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[col_letter].width = adjusted_width
    
    # Set a fixed width for the separator column
    ws.column_dimensions[get_column_letter(separator_col)].width = 3
    
    # Freeze panes
    ws.freeze_panes = "A2"
    
    matched_count = file1_match_status.count("Matched")
    unmatched_count = file1_match_status.count("Not Matched")
    
    return matched_count, unmatched_count, len(df2) - matched_count
