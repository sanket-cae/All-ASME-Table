import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="ASME Sec II Part D Property Viewer", layout="wide"
)

st.title("ASME Section II Part D - Material Properties Explorer (Unified)")


@st.cache_data
def load_all_data():
  excel_path = "ASME SecII Part D.xlsx"

  # Load sheets
  db_as = pd.read_excel(excel_path, sheet_name="DB_AS")
  db_y1 = pd.read_excel(excel_path, sheet_name="DB_Y1")
  db_u = pd.read_excel(excel_path, sheet_name="DB_U")

  # Standardize column naming variations across sheets
  if "SpecNo." in db_y1.columns:
    db_y1 = db_y1.rename(columns={"SpecNo.": "Spec No."})
  if "SpecNo." in db_u.columns:
    db_u = db_u.rename(columns={"SpecNo.": "Spec No."})

  # Find standard column names for Alloy / UNS
  for df in [db_as, db_y1, db_u]:
    uns_col = [c for c in df.columns if "UNS" in c or "Alloy" in c]
    if uns_col and uns_col[0] != "UNS No.":
      df.rename(columns={uns_col[0]: "UNS No."}, inplace=True)

  return db_as, db_y1, db_u


try:
  db_as, db_y1, db_u = load_all_data()

  # Create a unified master list of materials from all three databases
  def extract_materials(df):
    if df is None or df.empty:
      return pd.DataFrame()
    cols = ["Spec No.", "Type/Grade", "UNS No.", "Nominal Composition"]
    # Keep only columns that exist
    available_cols = [c for c in cols if c in df.columns]
    sub = df[available_cols].copy()
    for c in available_cols:
      sub[c] = sub[c].fillna("").astype(str).str.strip()
    return sub


  df_master = pd.concat(
      [
          extract_materials(db_as),
          extract_materials(db_y1),
          extract_materials(db_u),
      ]
  ).drop_duplicates(subset=["Spec No.", "Type/Grade", "UNS No."])
  df_master = df_master.sort_values(by=["Spec No.", "Type/Grade"])

  st.sidebar.header("Material Search & Filter")

  # Filter by Specification
  spec_options = ["All"] + list(df_master["Spec No."].unique())
  selected_spec = st.sidebar.selectbox("Filter by Spec No.", spec_options)

  if selected_spec != "All":
    filtered_master = df_master[df_master["Spec No."] == selected_spec]
  else:
    filtered_master = df_master

  selected_grade = st.sidebar.selectbox(
      "Select Material (Type/Grade - UNS)",
      filtered_master.apply(
          lambda row: f"{row['Spec No.']} | {row['Type/Grade']} | UNS: {row['UNS No']}",
          axis=1,
      ),
  )

  if selected_grade:
    # Parse back the selection
    parts = selected_grade.split(" | ")
    s_spec, s_type, s_uns = parts[0], parts[1], parts[2].replace("UNS: ", "")

    st.subheader(f"Properties for: {s_spec} / {s_type} / {s_uns}")

    # Fetch data from respective sheets
    match_as = db_as[
        (db_as["Spec No."].astype(str).str.strip() == s_spec)
        & (db_as["Type/Grade"].astype(str).str.strip() == s_type)
    ]
    match_y1 = db_y1[
        (db_y1["Spec No."].astype(str).str.strip() == s_spec)
        & (db_y1["Type/Grade"].astype(str).str.strip() == s_type)
    ]
    match_u = db_u[
        (db_u["Spec No."].astype(str).str.strip() == s_spec)
        & (db_u["Type/Grade"].astype(str).str.strip() == s_type)
    ]

    col1, col2, col3 = st.columns(3)

    with col1:
      st.markdown("### Allowable Stress (`DB_AS`)")
      if not match_as.empty:
        st.dataframe(match_as)
      else:
        st.info("Not available in Allowable Stress table.")

    with col2:
      st.markdown("### Yield Strength (`DB_Y1`)")
      if not match_y1.empty:
        st.dataframe(match_y1)
      else:
        st.info("Not available in Yield Strength table.")

    with col3:
      st.markdown("### Ultimate Tensile (`DB_U`)")
      if not match_u.empty:
        st.dataframe(match_u)
      else:
        st.info("Not available in Ultimate Tensile table.")

except Exception as e:
  st.error(f"An error occurred: {e}")
          
