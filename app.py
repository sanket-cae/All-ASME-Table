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
  def extract_material_specs(df, spec_col_candidates):
    if df is None or df.empty:
      return pd.DataFrame()

    # Find the actual spec column name
    df_cols_str = [str(c) for c in df.columns]
    actual_spec_col = None
    for cand in spec_col_candidates:
      matches = [c for c in df.columns if cand.lower() in str(c).lower()]
      if matches:
        actual_spec_col = matches[0]
        break

    if not actual_spec_col:
      return pd.DataFrame()

    # Find alloy / UNS column safely by casting column names to string
    uns_cols = [
        c
        for c in df.columns
        if "uns" in str(c).lower() or "alloy" in str(c).lower()
    ]
    uns_col_name = uns_cols[0] if uns_cols else None

    cols_mapping = {actual_spec_col: "Spec No.", "Type/Grade": "Type/Grade"}
    sub_cols = ["Spec No.", "Type/Grade"]

    if uns_col_name:
      cols_mapping[uns_col_name] = "UNS No."
      sub_cols.append("UNS No.")

    temp_df = df.rename(columns=cols_mapping)
    available_cols = [c for c in sub_cols if c in temp_df.columns]
    sub = temp_df[available_cols].copy()

    for c in available_cols:
      sub[c] = sub[c].fillna("").astype(str).str.strip()

    # Filter out empty or NaN spec rows
    sub = sub[
        ~sub["Spec No."].isin(["", "nan", "None", "NaN", "NAT"])
        & sub["Type/Grade"].ne("")
        & sub["Type/Grade"].ne("nan")
    ]
    return sub


  # Extract records independently from all 3 stress sheets
  master_as = extract_material_specs(db_as, ["Spec No.", "SpecNo."])
  master_y1 = extract_material_specs(db_y1, ["Spec No.", "SpecNo."])
  master_u = extract_material_specs(db_u, ["Spec No.", "SpecNo."])

  # Combine and drop duplicates to account for extra materials in DB_Y1 / DB_U
  df_master = (
      pd.concat([master_as, master_y1, master_u])
      .drop_duplicates(subset=["Spec No.", "Type/Grade"])
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

  # Filter underlying tables based on user choice safely
  def filter_table(df, spec_val, grade_val):
    if df is None or df.empty:
      return pd.DataFrame()
    # Find matching spec column name
    spec_col = next(
        (c for c in df.columns if "spec" in str(c).lower()), df.columns[3]
    )
    if "Type/Grade" not in df.columns:
      return pd.DataFrame()

    return df[
        (df[spec_col].fillna("").astype(str).str.strip() == str(spec_val))
        & (
            df["Type/Grade"].fillna("").astype(str).str.strip()
            == str(grade_val)
        )
    ]


  match_as = filter_table(db_as, selected_spec, selected_grade)
  match_y1 = filter_table(db_y1, selected_spec, selected_grade)
  match_u = filter_table(db_u, selected_spec, selected_grade)

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
    
