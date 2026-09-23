import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="ASME Sec II Part D Property Viewer", layout="wide"
)

st.title("ASME Section II Part D - Material Properties Explorer")


@st.cache_data
def load_data():
  excel_path = "ASME SecII Part D.xlsx"
  xls = pd.ExcelFile(excel_path)

  db_as = pd.read_excel(excel_path, sheet_name="DB_AS")
  db_y1 = pd.read_excel(excel_path, sheet_name="DB_Y1")
  db_u = pd.read_excel(excel_path, sheet_name="DB_U")
  db_te = pd.read_excel(excel_path, sheet_name="DB_TE")
  db_tcd = pd.read_excel(excel_path, sheet_name="DB_TCD")
  db_tm = pd.read_excel(excel_path, sheet_name="DB_TM")
  db_map = pd.read_excel(excel_path, sheet_name="DB_MAP")

  return db_as, db_y1, db_u, db_te, db_tcd, db_tm, db_map


try:
  db_as, db_y1, db_u, db_te, db_tcd, db_tm, db_map = load_data()

  # --- INDEPENDENT MASTER MATERIAL LIST POOLING ---
  def extract_material_specs(df, spec_col):
    if df is None or df.empty:
      return pd.DataFrame()

    # Normalize column name variations if needed
    cols = {spec_col: "Spec No.", "Type/Grade": "Type/Grade"}
    # Locate alloy / UNS column dynamically
    uns_cols = [c for c in df.columns if "UNS" in c or "Alloy" in c]
    uns_col_name = uns_cols[0] if uns_cols else None

    sub_cols = ["Spec No.", "Type/Grade"]
    if uns_col_name:
      cols[uns_col_name] = "UNS No."
      sub_cols.append("UNS No.")

    temp_df = df.rename(columns=cols)
    available_cols = [c for c in sub_cols if c in temp_df.columns]
    sub = temp_df[available_cols].copy()

    for c in available_cols:
      sub[c] = sub[c].fillna("").astype(str).str.strip()
    return sub


  # Extract keys from all three stress databases independently
  master_as = extract_material_specs(
      db_as, "Spec No." if "Spec No." in db_as.columns else "SpecNo."
  )
  master_y1 = extract_material_specs(
      db_y1, "Spec No." if "Spec No." in db_y1.columns else "SpecNo."
  )
  master_u = extract_material_specs(
      db_u, "Spec No." if "Spec No." in db_u.columns else "SpecNo."
  )

  # Combine and drop duplicates to make sure extra materials in DB_Y1/DB_U are not lost
  df_master = (
      pd.concat([master_as, master_y1, master_u])
      .drop_duplicates()
      .reset_index(drop=True)
  )

  # Sidebar Filters (Preserving previous UI structure)
  st.sidebar.header("Material Filters")

  spec_list = ["All"] + sorted(df_master["Spec No."].unique().tolist())
  selected_spec = st.sidebar.selectbox("Specification (Spec No.)", spec_list)

  if selected_spec != "All":
    filtered_df = df_master[df_master["Spec No."] == selected_spec]
  else:
    filtered_df = df_master

  grade_list = sorted(filtered_df["Type/Grade"].unique().tolist())
  selected_grade = st.sidebar.selectbox("Type / Grade", grade_list)

  # Main Dashboard Layout
  st.subheader(
      f"Selected Specification: **{selected_spec}** | Grade: **{selected_grade}**"
  )

  # Filter underlying tables based on user choice
  match_as = pd.DataFrame()
  match_y1 = pd.DataFrame()
  match_u = pd.DataFrame()

  if not db_as.empty:
    spec_col_as = "Spec No." if "Spec No." in db_as.columns else "SpecNo."
    match_as = db_as[
        (db_as[spec_col_as].astype(str).str.strip() == str(selected_spec))
        & (db_as["Type/Grade"].astype(str).str.strip() == str(selected_grade))
    ]

  if not db_y1.empty:
    spec_col_y1 = "Spec No." if "Spec No." in db_y1.columns else "SpecNo."
    match_y1 = db_y1[
        (db_y1[spec_col_y1].astype(str).str.strip() == str(selected_spec))
        & (db_y1["Type/Grade"].astype(str).str.strip() == str(selected_grade))
    ]

  if not db_u.empty:
    spec_col_u = "Spec No." if "Spec No." in db_u.columns else "SpecNo."
    match_u = db_u[
        (db_u[spec_col_u].astype(str).str.strip() == str(selected_spec))
        & (db_u["Type/Grade"].astype(str).str.strip() == str(selected_grade))
    ]

  # Display data across tabs keeping original structure
  tab1, tab2, tab3 = st.tabs(
      ["Allowable Stress (DB_AS)", "Yield Strength (DB_Y1)", "Tensile Strength (DB_U)"]
  )

  with tab1:
    st.markdown("### Allowable Stress Data")
    if not match_as.empty:
      st.dataframe(match_as, use_container_width=True)
    else:
      st.info("No allowable stress entry found for this specific material.")

  with tab2:
    st.markdown("### Yield Strength Data")
    if not match_y1.empty:
      st.dataframe(match_y1, use_container_width=True)
    else:
      st.info("No yield strength entry found in DB_Y1 for this selection.")

  with tab3:
    st.markdown("### Ultimate Tensile Strength Data")
    if not match_u.empty:
      st.dataframe(match_u, use_container_width=True)
    else:
      st.info("No ultimate tensile strength entry found in DB_U for this selection.")

except Exception as e:
  st.error(f"Error loading or processing application: {e}")
    
