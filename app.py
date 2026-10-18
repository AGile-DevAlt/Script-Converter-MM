import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Excel Sheet Converter", page_icon="📊", layout="wide")
st.title("Excel Sheet Converter")
st.write("Upload a registration spreadsheet, map its columns, and download a clean Excel file.")

TARGET_COLUMNS = [
    "School Name",
    "Head Teacher name",
    "Name of Participant",
    "No of participant",
    "Event",
    "type of partici",
    "Payment status",
]

ALIASES = {
    "School Name": ["school", "school name", "institution", "institution name", "organisation", "organization"],
    "Head Teacher name": ["head teacher", "head teacher name", "principal", "principal name", "school head", "teacher in charge", "coordinator"],
    "Name of Participant": ["participant", "participant name", "student name", "name of participant", "delegate name", "attendee", "name"],
    "No of participant": ["no of participant", "number of participants", "no. of participants", "participants count", "participant count", "quantity", "number of participant", "no of participants"],
    "Event": ["event", "event name", "competition", "activity", "programme", "program"],
    "type of partici": ["type of participation", "participation type", "type of participant", "category", "participant type", "type of partici"],
    "Payment status": ["payment status", "payment", "paid status", "fee status", "payment received", "status"],
}

def norm(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()

def suggest_column(columns, target):
    normalized = {norm(c): c for c in columns}
    for alias in ALIASES[target]:
        if norm(alias) in normalized:
            return normalized[norm(alias)]
    for c in columns:
        nc = norm(c)
        if any(norm(alias) in nc or nc in norm(alias) for alias in ALIASES[target] if len(norm(alias)) > 3):
            return c
    return "(Not available)"

def read_excel(uploaded, sheet_name):
    uploaded.seek(0)
    return pd.read_excel(uploaded, sheet_name=sheet_name, dtype=object)

uploaded = st.file_uploader("Upload source Excel file", type=["xlsx", "xls"])
if uploaded:
    try:
        xls = pd.ExcelFile(uploaded)
        sheet = st.selectbox("Choose worksheet", xls.sheet_names)
        df = read_excel(uploaded, sheet)
        df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed column {i+1}" for i, c in enumerate(df.columns)]
        st.subheader("Preview of source data")
        st.dataframe(df.head(10), use_container_width=True)

        st.subheader("Match source columns to output columns")
        st.caption("We suggest matches automatically. Check each one before converting. Choose “Not available” if the source has no matching column.")
        source_options = ["(Not available)"] + list(df.columns)
        mapping = {}
        left, right = st.columns(2)
        for i, target in enumerate(TARGET_COLUMNS):
            container = left if i % 2 == 0 else right
            with container:
                suggestion = suggest_column(df.columns, target)
                default_index = source_options.index(suggestion) if suggestion in source_options else 0
                mapping[target] = st.selectbox(
                    f"{target}",
                    source_options,
                    index=default_index,
                    key=f"map_{target}",
                )

        st.subheader("Conversion options")
        keep_empty = st.checkbox("Keep rows even if some mapped fields are blank", value=True)
        one_row_per_participant = st.checkbox(
            "Split comma-separated participant names into separate rows",
            value=False,
            help="For example, 'Asha, Rohan' becomes two rows. Other mapped fields are copied to both rows."
        )
        extra_columns = st.checkbox("Append any unmapped source columns after the required columns", value=False)

        if st.button("Convert Excel", type="primary", use_container_width=True):
            out = pd.DataFrame(index=df.index)
            for target in TARGET_COLUMNS:
                src = mapping[target]
                out[target] = df[src] if src != "(Not available)" else ""

            if not keep_empty:
                mapped_targets = [t for t in TARGET_COLUMNS if mapping[t] != "(Not available)"]
                if mapped_targets:
                    out = out.loc[out[mapped_targets].apply(
                        lambda row: any(pd.notna(v) and str(v).strip() for v in row), axis=1
                    )].copy()

            if one_row_per_participant and mapping["Name of Participant"] != "(Not available)":
                out["Name of Participant"] = out["Name of Participant"].fillna("").astype(str).str.split(r"\s*[,;\n]\s*")
                out = out.explode("Name of Participant", ignore_index=True)
                out["Name of Participant"] = out["Name of Participant"].astype(str).str.strip()
                out = out[out["Name of Participant"] != ""]

            if extra_columns:
                used = {v for v in mapping.values() if v != "(Not available)"}
                for col in df.columns:
                    if col not in used:
                        out[col] = df.loc[out.index, col] if len(out.index) and out.index.max() < len(df) else ""

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                out.to_excel(writer, index=False, sheet_name="Converted Data")
                ws = writer.sheets["Converted Data"]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for column_cells in ws.columns:
                    max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in column_cells)
                    ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 14), 36)
            output.seek(0)

            st.success(f"Converted {len(out)} row(s).")
            st.dataframe(out.head(20), use_container_width=True)
            st.download_button(
                "Download converted Excel file",
                data=output.getvalue(),
                file_name="converted_registration_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    except Exception as exc:
        st.error(f"Could not read or convert this workbook: {exc}")
        st.info("If the file is an older .xls workbook, try saving it as .xlsx in Excel first.")
else:
    st.info("Upload an Excel file to begin.")
    st.markdown("**Output columns**")
    st.dataframe(pd.DataFrame(columns=TARGET_COLUMNS), use_container_width=True)
