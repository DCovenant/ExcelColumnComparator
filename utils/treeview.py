def calculate_column_width(col_index, all_values, heading_text):
    max_len = len(heading_text)
    for row_values in all_values:
        if col_index < len(row_values):
            max_len = max(max_len, len(row_values[col_index]))
    char_width = 8
    padding = 20
    return min(max(max_len * char_width + padding, 60), 400)


def auto_size_columns(tree, column_ids, all_values, headings):
    for i, col_id in enumerate(column_ids):
        width = calculate_column_width(i, all_values, headings[i])
        tree.column(col_id, width=width, minwidth=60, stretch=False)
