import datetime
import numpy as np
import pandas as pd
import streamlit as st

# Page configuration
st.set_page_config(
    page_title="ASME Sec II Part D Property Viewer",
    page_icon="⚙️",
    layout="wide",
)

# Custom CSS styling for professional engineering UI
st.markdown(
    """
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        border-radius: 6px;
        font-weight: 600;
    }
    .meta-box {
        background-color: #ffffff;
        padding: 12px 15px;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        margin-bottom: 8px;
    }
    .meta-title {
        font-size: 11px;
        color: #666666;
        text-transform: uppercase;
        font-weight: 700;
        margin-bottom: 2px;
    }
    .meta-value {
        font-size: 14px;
        color: #111111;
        font-weight: 600;
    }
    .sidebar-constant-box {
        background-color: #2b3035;
        color: #ffffff !important;
        padding: 12px;
        border-radius: 6px;
        border: 1px solid #495057;
        margin-top: 10px;
        margin-bottom: 10px;
    }
    .sidebar-constant-box b, .sidebar-constant-box {
        color: #ffffff !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("⚙️ ASME Section II, Part D Material Property Database")
st.markdown(
    "Interactive engineering workspace for material property lookups, historical"
    " stacking, and dynamic temperature evaluations."
)


# Load Excel Database (Cleaning dates and ensuring string formats)
@st.cache_data
def load_data(file_path):
  xls = pd.ExcelFile(file_path)
  db_as = pd.read_excel(file_path, sheet_name="DB_AS", dtype=str)
  db_y1 = pd.read_excel(file_path, sheet_name="DB_Y1", dtype=str)
  db_u = pd.read_excel(file_path, sheet_name="DB_U", dtype=str)
  db_te = pd.read_excel(file_path, sheet_name="DB_TE", dtype=str)
  db_tcd = pd.read_excel(file_path, sheet_name="DB_TCD", dtype=str)
  db_tm = pd.read_excel(file_path, sheet_name="DB_TM", dtype=str)

  db_map_raw = pd.read_excel(file_path, sheet_name="DB_MAP", dtype=str)
  db_map_raw.columns = db_map_raw.iloc[0]
  db_map = db_map_raw[1:].reset_index(drop=True)

  # Normalize SpecNo. naming in DB_Y1 & DB_U if needed
  for df in [db_y1, db_u]:
    if "SpecNo." in df.columns and "Spec No." not in df.columns:
      df.rename(columns={"SpecNo.": "Spec No."}, inplace=True)

  for df in [db_as, db_y1, db_u, db_map]:
    if "Type/Grade" in df.columns:
      df["Type/Grade"] = df["Type/Grade"].apply(
          lambda x: (
              x.strftime("%b-%d").upper()
              if isinstance(
                  x, (datetime.datetime, datetime.date, pd.Timestamp)
              )
              else str(x).strip()
          )
      )
    if "Type/\nGrade" in df.columns:
      df["Type/\nGrade"] = df["Type/\nGrade"].apply(
          lambda x: (
              x.strftime("%b-%d").upper()
              if isinstance(
                  x, (datetime.datetime, datetime.date, pd.Timestamp)
              )
              else str(x).strip()
          )
      )

  return db_as, db_y1, db_u, db_te, db_tcd, db_tm, db_map


excel_file = "ASME SecII Part D.xlsx"

try:
  db_as, db_y1, db_u, db_te, db_tcd, db_tm, db_map = load_data(excel_file)
except Exception as e:
  st.error(
      f"Error loading Excel file: {e}. Please ensure 'ASME SecII Part D.xlsx'"
      " is uploaded in your repository."
  )
  st.stop()

if "history" not in st.session_state:
  st.session_state.history = []

# --- SIDEBAR INPUTS (BIDIRECTIONAL SEARCH ACROSS ALL SHEETS) ---
st.sidebar.markdown("### 🎛️ Configuration Panel")

with st.sidebar.expander("1. Code & Material Selection", expanded=True):
  code_section_map = {
      "I": "I",
      "III": "III",
      "VIII-1": "VIII-1",
      "VIII-2 Cl. 2": "VIII-2\nCl. 2",
      "XII": "XII",
  }
  selected_section_label = st.selectbox(
      "Applicability (Code Section)", list(code_section_map.keys())
  )
  db_column_name = code_section_map[selected_section_label]

  # Independent Pooling across DB_AS, DB_Y1, and DB_U
  all_specs = sorted(
      list(
          set(
              [
                  str(s).strip()
                  for s in db_as["Spec No."].dropna().unique()
                  if str(s).strip() not in ["nan", ""]
              ]
              + [
                  str(s).strip()
                  for s in db_y1["Spec No."].dropna().unique()
                  if str(s).strip() not in ["nan", ""]
              ]
              + [
                  str(s).strip()
                  for s in db_u["Spec No."].dropna().unique()
                  if str(s).strip() not in ["nan", ""]
              ]
          )
      )
  )

  all_grades = sorted(
      list(
          set(
              [
                  str(g).strip()
                  for g in db_as["Type/Grade"].dropna().unique()
                  if str(g).strip() not in ["nan", ""]
              ]
              + [
                  str(g).strip()
                  for g in db_y1["Type/Grade"].dropna().unique()
                  if str(g).strip() not in ["nan", ""]
              ]
              + [
                  str(g).strip()
                  for g in db_u["Type/Grade"].dropna().unique()
                  if str(g).strip() not in ["nan", ""]
              ]
          )
      )
  )

  search_mode = st.radio("Search Direction", ["By Spec No. first", "By Grade first"])

  if search_mode == "By Spec No. first":
    selected_spec = st.selectbox("Specification Number (Spec No.)", all_specs)
    spec_grades = sorted(
        list(
            set(
                db_as[db_as["Spec No."].str.strip() == selected_spec][
                    "Type/Grade"
                ]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
                + db_y1[db_y1["Spec No."].str.strip() == selected_spec][
                    "Type/Grade"
                ]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
                + db_u[db_u["Spec No."].str.strip() == selected_spec][
                    "Type/Grade"
                ]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
            )
        )
    )
    selected_grade = st.selectbox(
        "Type / Grade", spec_grades if spec_grades else all_grades
    )
  else:
    selected_grade = st.selectbox("Type / Grade", all_grades)
    grade_specs = sorted(
        list(
            set(
                db_as[db_as["Type/Grade"].str.strip() == selected_grade][
                    "Spec No."
                ]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
                + db_y1[db_y1["Type/Grade"].str.strip() == selected_grade][
                    "Spec No."
                ]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
                + db_u[db_u["Type/Grade"].str.strip() == selected_grade][
                    "Spec No."
                ]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
            )
        )
    )
    selected_spec = st.selectbox(
        "Specification Number (Spec No.)", grade_specs if grade_specs else all_specs
    )

  combined_db = pd.concat([db_as, db_y1, db_u], ignore_index=True)
  filtered_uns_df = combined_db[
      (combined_db["Spec No."].str.strip() == selected_spec)
      & (combined_db["Type/Grade"].str.strip() == selected_grade)
  ]
  uns_col = [
      c
      for c in filtered_uns_df.columns
      if "Alloy" in str(c) or "UNS" in str(c)
  ]

  filtered_uns = []
  if uns_col:
    filtered_uns = sorted(
        [
            str(u).strip()
            for u in filtered_uns_df[uns_col[0]].dropna().unique()
            if str(u).strip() not in ["nan", "…", ""]
        ]
    )
  selected_uns = st.selectbox("Alloy Designation / UNS No.", ["All"] + filtered_uns)

# Pre-lookup fixed material constants
sidebar_poisson, sidebar_density = 0.3, 7850.0
if not db_map.empty:
  map_spec_col = (
      [c for c in db_map.columns if "Spec" in str(c)][0]
      if [c for c in db_map.columns if "Spec" in str(c)]
      else db_map.columns[2]
  )
  map_grade_col = (
      [c for c in db_map.columns if "Grade" in str(c) or "Type" in str(c)][0]
      if [c for c in db_map.columns if "Grade" in str(c) or "Type" in str(c)]
      else db_map.columns[3]
  )

  matched_map_sb = pd.DataFrame()
  for _, m_row in db_map.iterrows():
    m_spec = str(m_row.get(map_spec_col, "")).strip()
    m_grade = str(m_row.get(map_grade_col, "")).strip()
    if (
        m_spec == str(selected_spec).strip()
        and m_grade == str(selected_grade).strip()
    ):
      matched_map_sb = pd.DataFrame([m_row.to_dict()])
      break
  if matched_map_sb.empty:
    for _, m_row in db_map.iterrows():
      m_spec = str(m_row.get(map_spec_col, "")).strip()
      if m_spec == str(selected_spec).strip():
        matched_map_sb = pd.DataFrame([m_row.to_dict()])
        break

  if not matched_map_sb.empty:
    sidebar_poisson = (
        float(matched_map_sb.iloc[0].get("Poisson's\nRatio", 0.3))
        if pd.notnull(matched_map_sb.iloc[0].get("Poisson's\nRatio"))
        else 0.3
    )
    sidebar_density = (
        float(matched_map_sb.iloc[0].get("Density\nkg/m3", 7850))
        if pd.notnull(matched_map_sb.iloc[0].get("Density\nkg/m3"))
        else 7850
    )

# Variant Source Strategy across sheets
variant_df = db_as[
    (db_as["Spec No."].str.strip() == selected_spec)
    & (db_as["Type/Grade"].str.strip() == selected_grade)
].copy()
source_sheet_used = "DB_AS"

if variant_df.empty:
  variant_df = db_y1[
      (db_y1["Spec No."].str.strip() == selected_spec)
      & (db_y1["Type/Grade"].str.strip() == selected_grade)
  ].copy()
  source_sheet_used = "DB_Y1"

if variant_df.empty:
  variant_df = db_u[
      (db_u["Spec No."].str.strip() == selected_spec)
      & (db_u["Type/Grade"].str.strip() == selected_grade)
  ].copy()
  source_sheet_used = "DB_U"

if selected_uns != "All" and uns_col:
  variant_df = variant_df[variant_df[uns_col[0]].str.strip() == selected_uns]

if source_sheet_used == "DB_AS" and db_column_name in variant_df.columns:
  variant_df = variant_df[
      variant_df[db_column_name].str.strip().str.upper() != "NP"
  ]

variant_df = variant_df.reset_index(drop=True)
variant_df["Variant_No"] = variant_df.index + 1

tensile_col_candidates = [c for c in variant_df.columns if "Tensile" in str(c)]
tensile_col = tensile_col_candidates[0] if tensile_col_candidates else None

with st.sidebar.expander("2. Variant & Initial Temp", expanded=True):
  if len(variant_df) > 1 and tensile_col:
    selected_variant = st.selectbox(
        "Max. Allow. Stress Variant",
        variant_df["Variant_No"].tolist(),
        format_func=lambda x: f"Variant {x} (Min Tensile: {variant_df[variant_df['Variant_No']==x][tensile_col].values[0]} MPa)",
    )
  elif len(variant_df) >= 1:
    selected_variant = 1
    st.info(f"Variant 1 (Loaded from {source_sheet_used})")
  else:
    st.warning("No permitted records found for this section.")
    st.stop()

  initial_temp = st.number_input("Initial Target Temp (°C)", value=20.0, step=25.0)

st.sidebar.markdown("---")
if st.sidebar.button(
    "🔍 Find Properties & Add to History", type="primary", use_container_width=True
):
  record = variant_df[variant_df["Variant_No"] == selected_variant].iloc[0]
  history_item = {
      "spec": selected_spec,
      "grade": selected_grade,
      "uns": selected_uns,
      "section": selected_section_label,
      "variant": selected_variant,
      "record": record,
      "source": source_sheet_used,
      "eval_temps": [initial_temp],
  }
  st.session_state.history.insert(0, history_item)

st.sidebar.markdown("### 📌 Material Constants")
st.sidebar.markdown(
    f"""<div class="sidebar-constant-box">
        <b>Poisson's Ratio (μ):</b> {sidebar_poisson:.3f} (–)<br>
        <b>Density (ρ):</b> {sidebar_density:.3f} kg/m³
    </div>""",
    unsafe_allow_html=True,
)

st.sidebar.markdown("### 📥 Database Export")
try:
  with open(excel_file, "rb") as f:
    excel_bytes = f.read()
  st.sidebar.download_button(
      label="Download Master Excel File",
      data=excel_bytes,
      file_name="ASME SecII Part D.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
      use_container_width=True,
  )
except Exception as e:
  pass


# --- ROBUST INTERPOLATION HELPERS ---
def interpolate_prop(
    df, temp_col_target, val_col_target, target, group_col_target=None, group_val=None
):
  if df.empty:
    return "-"

  cols_lower = {str(c).strip().lower(): c for c in df.columns}
  temp_col = cols_lower.get(str(temp_col_target).strip().lower())
  val_col = cols_lower.get(str(val_col_target).strip().lower())

  if not temp_col or not val_col:
    return "-"

  sub_df = df.copy()
  if group_col_target and group_val:
    group_cols_lower = {str(c).strip().lower(): c for c in df.columns}
    g_col = group_cols_lower.get(str(group_col_target).strip().lower())
    if g_col:
      sub_df = sub_df[
          sub_df[g_col].astype(str).str.strip() == str(group_val).strip()
      ]

  if sub_df.empty:
    return "-"

  sub_df[temp_col] = pd.to_numeric(sub_df[temp_col], errors="coerce")
  sub_df[val_col] = pd.to_numeric(sub_df[val_col], errors="coerce")
  sub_df = sub_df.dropna(subset=[temp_col, val_col]).sort_values(by=temp_col)

  if sub_df.empty:
    return "-"

  temps = sub_df[temp_col].values
  vals = sub_df[val_col].values

  if target in temps:
    return vals[np.where(temps == target)[0][0]]
  elif target < temps[0] or target > temps[-1]:
    return "-"
  else:
    return np.interp(target, temps, vals)


def get_row_stress_at_temp(row_data, t_cols, target):
  t_vals, s_vals = [], []
  for t in t_cols:
    val = row_data.get(t, "NP")
    if pd.notnull(val):
      try:
        f_val = float(val)
        t_vals.append(float(t))
        s_vals.append(f_val)
      except ValueError:
        continue
  if not t_vals or not s_vals:
    return "NP"
  t_vals = np.array(t_vals)
  s_vals = np.array(s_vals)
  sorted_indices = np.argsort(t_vals)
  t_vals = t_vals[sorted_indices]
  s_vals = s_vals[sorted_indices]
  if target in t_vals:
    return s_vals[np.where(t_vals == target)[0][0]]
  elif target < t_vals[0] or target > t_vals[-1]:
    return "NP"
  else:
    return np.interp(target, t_vals, s_vals)


temp_cols = [
    c for c in db_as.columns if str(c).strip().replace(".", "", 1).isdigit()
]


def render_meta_item(title, value):
  st.markdown(
      f"""
        <div class="meta-box">
            <div class="meta-title">{title}</div>
            <div class="meta-value">{value}</div>
        </div>
    """,
      unsafe_allow_html=True,
  )


# --- DISPLAY MATERIAL HISTORY STACK ---
if not st.session_state.history:
  st.info(
      "👈 Please configure your material selection in the sidebar and click"
      " **'Find Properties & Add to History'**."
  )
else:
  st.markdown(
      "### 📚 Evaluated Material History (Click any material card to"
      " expand/collapse)"
  )

  if st.button("🗑️ Clear All History"):
    st.session_state.history = []
    st.rerun()

  for idx, item in enumerate(st.session_state.history):
    spec = item["spec"]
    grade = item["grade"]
    uns = item["uns"]
    section = item["section"]
    variant = item["variant"]
    record = item["record"]
    source = item["source"]

    db_col = code_section_map[section]
    t_col_name = [c for c in record.index if "Tensile" in str(c)]
    tensile = record.get(t_col_name[0], "-") if t_col_name else "-"

    card_label = (
        f"📦 [{idx+1}] **Spec:** {spec} | **Grade:** {grade} | **Variant:**"
        f" {variant} | **Source:** {source} | **Section:** {section} (Min."
        f" Tensile: {tensile} MPa)"
    )

    with st.expander(card_label, expanded=(idx == 0)):
      nom_comp = (
          record.get([c for c in record.index if "Nominal" in str(c)][0], "-")
          if [c for c in record.index if "Nominal" in str(c)]
          else "-"
      )
      prod_form = (
          record.get([c for c in record.index if "Product" in str(c)][0], "-")
          if [c for c in record.index if "Product" in str(c)]
          else "-"
      )
      class_cond = (
          record.get([c for c in record.index if "Class" in str(c)][0], "-")
          if [c for c in record.index if "Class" in str(c)]
          else "-"
      )
      size_thick = (
          record.get([c for c in record.index if "Size" in str(c)][0], "-")
          if [c for c in record.index if "Size" in str(c)]
          else "-"
      )
      y_col_name = [c for c in record.index if "Yield" in str(c)]
      min_yield = record.get(y_col_name[0], "-") if y_col_name else "-"

      raw_temp_limit = record.get(db_col, "NP") if source == "DB_AS" else "-"
      max_temp_limit = (
          f"{raw_temp_limit} °C"
          if pd.notnull(raw_temp_limit)
          and str(raw_temp_limit).strip() not in ["NP", "-", ""]
          else str(raw_temp_limit)
      )

      ext_chart = (
          record.get([c for c in record.index if "Ext" in str(c)][0], "-")
          if [c for c in record.index if "Ext" in str(c)]
          else "-"
      )
      notes = (
          record.get([c for c in record.index if "Note" in str(c)][0], "-")
          if [c for c in record.index if "Note" in str(c)]
          else "-"
      )

      te_group, tcd_group, tm_group, poisson, density = (
          "Group 1",
          "Group A",
          "C<=0.3%",
          0.3,
          7850,
      )
      if not db_map.empty:
        map_spec_col = (
            [c for c in db_map.columns if "Spec" in str(c)][0]
            if [c for c in db_map.columns if "Spec" in str(c)]
            else db_map.columns[2]
        )
        map_grade_col = (
            [
                c
                for c in db_map.columns
                if "Grade" in str(c) or "Type" in str(c)
            ][0]
            if [
                c
                for c in db_map.columns
                if "Grade" in str(c) or "Type" in str(c)
            ]
            else db_map.columns[3]
        )

        matched_map = pd.DataFrame()
        for _, m_row in db_map.iterrows():
          m_spec = str(m_row.get(map_spec_col, "")).strip()
          m_grade = str(m_row.get(map_grade_col, "")).strip()
          if m_spec == str(spec).strip() and m_grade == str(grade).strip():
            matched_map = pd.DataFrame([m_row.to_dict()])
            break
        if matched_map.empty:
          for _, m_row in db_map.iterrows():
            m_spec = str(m_row.get(map_spec_col, "")).strip()
            if m_spec == str(spec).strip():
              matched_map = pd.DataFrame([m_row.to_dict()])
              break

        if not matched_map.empty:
          te_group = matched_map.iloc[0].get("Table TE Group", "Group 1")
          tcd_group = matched_map.iloc[0].get("Table TCD Group", "Group A")
          tm_group = matched_map.iloc[0].get("Table TM Group", "C<=0.3%")
          poisson = (
              float(matched_map.iloc[0].get("Poisson's\nRatio", 0.3))
              if pd.notnull(matched_map.iloc[0].get("Poisson's\nRatio"))
              else 0.3
          )
          density = (
              float(matched_map.iloc[0].get("Density\nkg/m3", 7850))
              if pd.notnull(matched_map.iloc[0].get("Density\nkg/m3"))
              else 7850
          )

      r1_c1, r1_c2, r1_c3, r1_c4, r1_c5 = st.columns(5)
      with r1_c1:
        render_meta_item("Nominal Comp.", str(nom_comp))
      with r1_c2:
        render_meta_item("Product Form", str(prod_form))
      with r1_c3:
        render_meta_item("Class / Cond.", str(class_cond))
      with r1_c4:
        render_meta_item("Size / Thick.", str(size_thick))
      with r1_c5:
        render_meta_item(
 
