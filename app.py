@st.cache_data
def load_defaults(project_name: str):
    wb = openpyxl.load_workbook(io.BytesIO(_download_xlsx()), data_only=True)
    ws = wb[project_name]

    header = next(ws.iter_rows(min_row=2, max_row=2, values_only=True))
    years = []
    for v in header[2:]:
        try:
            y = int(v)
            if 1900 < y < 2200:
                years.append(y)
        except (TypeError, ValueError):
            pass
    n = len(years)

    KNOWN = {"INFLOWS", "OUTFLOWS", "FINANCING"}
    sections = {"INFLOWS": [], "OUTFLOWS": [], "FINANCING": []}
    current = None

    def _f(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    for row in ws.iter_rows(min_row=3, values_only=True):
        sec     = str(row[0]).strip() if row[0] else ""
        concept = str(row[1]).strip() if row[1] else ""

        if sec in KNOWN:
            current = sec
            if concept:
                vals = [_f(v) for v in row[2:2 + n]]
                sections[current].append((concept, vals))
            continue

        if current and concept:
            vals = [_f(v) for v in row[2:2 + n]]
            sections[current].append((concept, vals))

    return sections, years
