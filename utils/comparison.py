from utils.text import normalize


def collect_col_data(df, col_name, header_row):
    if col_name not in df.columns:
        return {}
    result = {}
    for idx, val in df[col_name].dropna().items():
        v = normalize(val)
        if v and v.lower() != "nan":
            result[header_row + 2 + idx] = v
    return result


def get_rows_with_unique_values(col_data, unique_values):
    rows = {}
    for col_name, data in col_data.items():
        for excel_row, val in data.items():
            if val in unique_values:
                rows.setdefault(excel_row, {})[col_name] = val
    return rows


def matches_search_pattern(text, pattern):
    text = text.lower()
    pattern = pattern.lower()
    if pattern.startswith("*") and pattern.endswith("*"):
        return pattern[1:-1] in text
    if pattern.startswith("*"):
        return text.endswith(pattern[1:])
    if pattern.endswith("*"):
        return text.startswith(pattern[:-1])
    return text == pattern


def row_matches_search(values, pattern):
    for val in values:
        if matches_search_pattern(str(val), pattern):
            return True
    return False
