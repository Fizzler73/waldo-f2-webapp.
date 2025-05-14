import streamlit as st
import pandas as pd
import io
import re
import os

st.set_page_config(page_title="Waldo F2 Job Creator", layout="wide")
st.title("\U0001F4C4 Waldo F2 Web-Based Job Creator")

st.markdown("""
<style>
    .stTextInput>div>div>input {
        border-radius: 0.375rem;
        padding: 8px;
    }
    .stSelectbox>div>div {
        border-radius: 0.375rem;
    }
</style>
""", unsafe_allow_html=True)

iolm_options = ["ATT F2 PON SHORT LINK", "ATT F2 PON", "ATT F2 PON MEDIUM LINK"]
opm_options = ["ATT F2 PFP -21dBm", "ATT F2 Terminal -24dBm"]

st.header("📁 Step 1: Upload PON TEST SHEET")
uploaded_file = st.file_uploader("Upload PON TEST SHEET (.xlsx)", type=["xlsx"], key="uploaded_file")

if uploaded_file:
    st.success(f"✅ File '{uploaded_file.name}' loaded successfully.")

if "opm_count" in st.session_state and "iolm_count" in st.session_state:
    st.markdown(f"### 🧾 Test Summary: {st.session_state['opm_count']} OPM test points and {st.session_state['iolm_count']} iOLM test points generated.")

st.header("✍️ Step 2: Enter Job Metadata")
with st.form("job_config"):
    col1, col2, col3 = st.columns(3)
    with col1:
        x2_default = ""
        if uploaded_file:
            try:
                x2_default = pd.read_excel(uploaded_file, sheet_name="PON TEST SHEET", header=None).iat[1, 23]
            except:
                pass
        if "clli" not in st.session_state:
            st.session_state["clli"] = ""
        if "co" not in st.session_state:
            st.session_state["co"] = ""
        cfas = st.text_input("CFAS # (Required)", value=str(x2_default).strip(), max_chars=20, key="cfas")
        clli = st.text_input("Wire Center CLLI", value=st.session_state["clli"], key="clli")
        co = st.text_input("Central Office", value=st.session_state["co"], key="co")
    with col2:
        if "tech_id" not in st.session_state:
            st.session_state["tech_id"] = ""
        if "pfp" not in st.session_state:
            st.session_state["pfp"] = ""
        tech_id = st.text_input("Technician UID", value=st.session_state["tech_id"], key="tech_id") + "@att.com"
        pfp = st.text_input("PFP", value=st.session_state["pfp"], key="pfp")
    with col3:
        olm_config = st.selectbox("iOLM Config", options=iolm_options)
        opm_config = st.selectbox("OPM Config", options=opm_options)

    col_a, col_b = st.columns([3, 1])
    with col_a:
        submitted = st.form_submit_button("📊 Analyze Test Sheet")
    with col_b:
        clear_form = st.form_submit_button("🧹 Clear Form")

if clear_form:
    for key in ["cfas", "clli", "co", "tech_id", "pfp", "opm_count", "iolm_count", "uploaded_file"]:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

if uploaded_file and submitted:
    if not cfas:
        st.error("CFAS # is required.")
        st.stop()

    xl = pd.ExcelFile(uploaded_file)

    try:
        pon_test_df = xl.parse("PON TEST SHEET", header=None)
    except Exception as e:
        st.error(f"Error loading PON TEST SHEET: {e}")
        st.stop()

    st.success("File loaded and form submitted. Parsing test sheet and preparing CSV output.")

    expected_columns = {
        "TERMINAL": None,
        "CABLE ID": None,
        "POWER TEST STRAND": None,
        "OTDR TEST STRAND": None
    }

    header_row_index = None
    found_headers = []

    def normalize(text):
        return re.sub(r"[^A-Z0-9]", "", text.upper())

    for idx, row in pon_test_df.iterrows():
        for col_idx, cell in row.items():
            if pd.isna(cell):
                continue
            cell_str = normalize(str(cell))
            found_headers.append(cell_str)
            for key in expected_columns:
                if normalize(key) in cell_str and expected_columns[key] is None:
                    expected_columns[key] = col_idx
        if all(v is not None for v in expected_columns.values()):
            header_row_index = idx
            break

    st.info(f"Detected normalized headers in sheet: {set(found_headers)}")

    if all(v is not None for v in expected_columns.values()):
        parsed_df = xl.parse("PON TEST SHEET", header=header_row_index)

        parsed_df.columns = [normalize(str(col)) for col in parsed_df.columns]

        column_map = {
            "TERMINAL": next((col for col in parsed_df.columns if "TERMINAL" in col), None),
            "CABLE ID": next((col for col in parsed_df.columns if "CABLEID" in col), None),
            "POWER TEST STRAND": next((col for col in parsed_df.columns if "POWERTESTSTRAND" in col), None),
            "OTDR TEST STRAND": next((col for col in parsed_df.columns if "OTDRTESTSTRAND" in col), None)
        }

        if None in column_map.values():
            st.error(f"Column mapping failed. Found: {column_map}")
            st.stop()

        extracted_df = parsed_df[[column_map[k] for k in column_map]]
        extracted_df.columns = list(column_map.keys())

        st.subheader("\U0001F4C2 Extracted Test Data")
        st.dataframe(extracted_df.head(20))

        test_rows = []
        has_olm = olm_config and olm_config.lower() != "no configuration"
        has_opm = opm_config and opm_config.lower() != "no configuration"
        test_config = ""
        if has_olm:
            test_config = olm_config + ".iolmcfg"
        if has_opm:
            test_config = (test_config + "|" if test_config else "") + opm_config + ".opmcfg"

        for idx, row in extracted_df.iterrows():
            terminal = str(row["TERMINAL"]).strip() if pd.notna(row["TERMINAL"]) else ""
            caid = str(row["CABLE ID"]).strip() if pd.notna(row["CABLE ID"]) else ""
            power = row["POWER TEST STRAND"]
            otdr = row["OTDR TEST STRAND"]

            try:
                power_port = int(power)
                test_name = f"{power_port} - {terminal}_1_{caid}"
                line = [cfas, tech_id, "AT&T", "", "", test_name, caid, power_port, clli, co, pfp, "OPM", "", test_config]
                test_rows.append(line)
            except:
                pass

            if isinstance(otdr, str):
                ports = []
                parts = otdr.split("/")
                for part in parts:
                    if "-" in part:
                        start, end = map(int, part.split("-"))
                        ports.extend(range(start, end+1))
                    else:
                        ports.append(int(part))
                for i, port in enumerate(ports):
                    test_name = f"{port} - {terminal}_{i+2}_{caid}"
                    line = ["", "", "", "", "", test_name, caid, port, clli, co, pfp, "", "iOLM", ""]
                    test_rows.append(line)

        if test_rows:
            st.session_state["opm_count"] = sum(1 for row in test_rows if row[11] == "OPM")
            st.session_state["iolm_count"] = sum(1 for row in test_rows if row[12] == "iOLM")

            st.subheader("✅ Select Test Types to Include in Export")
            include_opm = st.checkbox("Include OPM Test Points", value=True)
            include_iolm = st.checkbox("Include iOLM Test Points", value=True)

            filtered_rows = [row for row in test_rows if (row[11] == "OPM" and include_opm) or (row[12] == "iOLM" and include_iolm)]
            csv_output = io.StringIO()
            df_out = pd.DataFrame(filtered_rows, columns=[
                "name", "assignees", "company", "customer", "dueDate",
                "testPointName", "identifier_Cable ID", "identifier_Fiber ID",
                "identifier_ALoc", "identifier_ZLoc", "identifier_WireCenterClli",
                "testType_01", "testType_02", "testConfigurations"
            ])

            if not df_out.empty:
                df_out.iloc[1:, 0] = ""
                df_out.iloc[1:, 1] = ""
                df_out.iloc[1:, 2] = ""
                df_out.iloc[1:, 13] = ""

            df_out.to_csv(csv_output, index=False)

            st.subheader("\U0001F4C3 Preview of CSV Output")
            st.dataframe(df_out, use_container_width=True)

            st.download_button(
                label="\U0001F4E5 Download Job CSV",
                data=csv_output.getvalue(),
                file_name=f"{cfas}_job.csv",
                mime="text/csv"
            )
        else:
            st.warning("No valid test points were generated from the data.")
    else:
        st.error("Could not find all required columns in PON TEST SHEET. Check that headers like 'TERMINAL', 'CABLE ID', 'POWER TEST STRAND', and 'OTDR TEST STRAND(S)' exist.")
