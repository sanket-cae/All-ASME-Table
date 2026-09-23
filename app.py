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


# Load Excel Database with duplicate column handling
@st.cache_data
def load_data(file_path):
  xls = pd.ExcelFile(file_path)

  def read_and_dedup(sheet):
    df = pd.read_excel(file_path, sheet_name=sheet, dtype=str)
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
      cols[cols == dup] = [
          f"{dup}_{i}" if i != 0 else dup
          for i in range(sum(cols == dup))
      ]
    df.columns = cols
    return df

  db_as = read_and_dedup("DB_AS")
  db_y1 = read_and_dedup("DB_Y1")
  db_u = read_and_dedup("DB_U")
  db_te = read_and_dedup("DB_TE")
  db_tcd = read_and_dedup("DB_TCD")
  db_tm = read_and_dedup("DB_TM")

  db_map_raw = pd.read_excel(file_path, sheet_name="DB_MAP", dtype=str)
  db_map_raw.columns = db_map_raw.iloc[0]
  db_map = db_map_raw[1:].reset_index(drop=True)

  # Normalize spec column name
  for df in [db_y1, db_u]:
    for col in df.columns:
      if "spec" in str(col).lower() and "no" in str(col).lower():
        if "Spec No." not in df.columns:
          df.rename(columns={col: "Spec No."}, inplace=True)
          break

  for df in [db_as, db_y1, db_u]:
    if "SpecNo." in df.columns and "Spec No." not in df.columns:
      df.rename(columns={"SpecNo.": "Spec No."}, inplace=True)

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

# --- SAFE COLUMN FINDERS & EXTRACTORS ---
def get_spec_col(df):
  for c in df.columns:
    if "spec" in str(c).lower() and "no" in str(c).lower():
      return c
  return None


def get_grade_col(df):
  for c in df.columns:
    if "grade" in str(c).lower() or "type" in str(c).lower():
      return c
  return None


spec_col_as = get_spec_col(db_as)
grade_col_as = get_grade_col(db_as)
spec_col_y1 = get_spec_col(db_y1)
grade_col_y1 = get_grade_col(db_y1)
spec_col_u = get_spec_col(db_u)
grade_col_u = get_grade_col(db_u)


def get_unique_list(df, col):
  if df is None or not col or col not in df.columns:
    return []
  s = df[col]
  if isinstance(s, pd.DataFrame):
    s = s.iloc[:, 0]
  return sorted(
      [
          str(x).strip()
          for x in s.dropna().unique()
          if str(x).strip().lower() not in ["nan", "none", ""]
      ]
  )


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

  all_specs = sorted(
      list(
          set(
              get_unique_list(db_as, spec_col_as)
              + get_unique_list(db_y1, spec_col_y1)
              + get_unique_list(db_u, spec_col_u)
          )
      )
  )

  all_grades = sorted(
      list(
          set(
              get_unique_list(db_as, grade_col_as)
              + get_unique_list(db_y1, grade_col_y1)
              + get_unique_list(db_u, grade_col_u)
          )
      )
  )

  search_mode = st.radio("Search Direction", ["By Spec No. first", "By Grade first"])

  if search_mode == "By Spec No. first":
    selected_spec = st.selectbox(
        "Specification Number (Spec No.)",
        all_specs if all_specs else ["None"],
    )

    # Get grades matching this spec safely across sheets
    def filter_grades_for_spec(df, s_c, g_c, spec):
      if df is None or not s_c or not g_c:
        return []
      sub = df[df[s_c].astype(str).str.strip() == spec]
      return get_unique_list(sub, g_c)

    spec_grades = sorted(
        list(
            set(
                filter_grades_for_spec(db_as, spec_col_as, grade_col_as, selected_spec)
                + filter_grades_for_spec(db_y1, spec_col_y1, grade_col_y1, selected_spec)
                + filter_grades_for_spec(db_u, spec_col_u, grade_col_u, selected_spec)
            )
        )
    )
    selected_grade = st.selectbox(
        "Type / Grade", spec_grades if spec_grades else all_grades
    )
  else:
    selected_grade = st.selectbox(
        "Type / Grade", all_grades if all_grades else ["None"]
    )

    def filter_specs_for_grade(df, s_c, g_c, grade):
      if df is None or not s_c or not g_c:
        return []
      sub = df[df[g_c].astype(str).str.strip() == grade]
      return get_unique_list(sub, s_c)

    grade_specs = sorted(
        list(
            set(
                filter_specs_for_grade(db_as, spec_col_as, grade_col_as, selected_grade)
                + filter_specs_for_grade(db_y1, spec_col_y1, grade_col_y1, selected_grade)
                + filter_specs_for_grade(db_u, spec_col_u, grade_col_u, selected_grade)
            )
        )
    )
    selected_spec = st.selectbox(
        "Specification Number (Spec No.)",
        grade_specs if grade_specs else all_specs,
    )

  # Find UNS / Alloy designation across combined DB safely
  combined_db = pd.concat([db_as, db_y1, db_u], ignore_index=True)
  c_s_col = get_spec_col(combined_db)
  c_g_col = get_grade_col(combined_db)

  filtered_uns = []
  if c_s_col and c_g_col:
    f_df = combined_db[
        (combined_db[c_s_col].astype(str).str.strip() == str(selected_spec))
        & (combined_db[c_g_col].astype(str).str.strip() == str(selected_grade))
    ]
    uns_col = [
        c
        for c in f_df.columns
        if "alloy" in str(c).lower() or "uns" in str(c).lower()
    ]
    if uns_col:
      filtered_uns = get_unique_list(f_df, uns_col[0])
      filtered_uns = [
          u
          for u in filtered_uns
          if u and u.lower() not in ["nan", "…", "", "none"]
      ]

  selected_uns = st.selectbox("Alloy Designation / UNS No.", ["All"] + filtered_uns)

# Pre-lookup fixed material constants
sidebar_poisson, sidebar_density = 0.3, 7850.0
if not db_map.empty:
  map_spec_col = get_spec_col(db_map) or db_map.columns[2]
  map_grade_col = get_grade_col(db_map) or db_map.columns[3]

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
    p_col = next(
        (c for c in matched_map_sb.columns if "poisson" in str(c).lower()), None
    )
    d_col = next(
        (c for c in matched_map_sb.columns if "density" in str(c).lower()), None
    )
    if p_col and pd.notnull(matched_map_sb.iloc[0][p_col]):
      sidebar_poisson = float(matched_map_sb.iloc[0][p_col])
    if d_col and pd.notnull(matched_map_sb.iloc[0][d_col]):
      sidebar_density = float(matched_map_sb.iloc[0][d_col])


# Safe variant extractor function
def get_variant_subset(df, s_c, g_c):
  if df is None or df.empty or not s_c or not g_c:
    return pd.DataFrame()
  mask = (df[s_c].astype(str).str.strip() == str(selected_spec)) & (
      df[g_c].astype(str).str.strip() == str(selected_grade)
  )
  return df[mask].copy()


variant_df = get_variant_subset(db_as, spec_col_as, grade_col_as)
source_sheet_used = "DB_AS"

if variant_df.empty:
  variant_df = get_variant_subset(db_y1, spec_col_y1, grade_col_y1)
  source_sheet_used = "DB_Y1"

if variant_df.empty:
  variant_df = get_variant_subset(db_u, spec_col_u, grade_col_u)
  source_sheet_used = "DB_U"

if selected_uns != "All" and uns_col:
  variant_df = variant_df[variant_df[uns_col[0]].str.strip() == selected_uns]

if source_sheet_used == "DB_AS" and db_column_name in variant_df.columns:
  variant_df = variant_df[
      variant_df[db_column_name].astype(str).str.strip().str.upper() != "NP"
  ]

variant_df = variant_df.reset_index(drop=True)
variant_df["Variant_No"] = variant_df.index + 1

tensile_col_candidates = [
    c for c in variant_df.columns if "tensile" in str(c).lower()
]
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

    # Explicit Ultimate Tensile Strength lookup from DB_U
    u_match = get_variant_subset(db_u, spec_col_u, grade_col_u)
    if not u_match.empty:
      u_t_candidates = [
          c for c in u_match.columns if "tensile" in str(c).lower()
      ]
      tensile = (
          u_match.iloc[0].get(u_t_candidates[0], "-")
          if u_t_candidates
          else "-"
      )
    else:
      t_col_name = [c for c in record.index if "tensile" in str(c).lower()]
      tensile = record.get(t_col_name[0], "-") if t_col_name else "-"

    card_label = (
        f"📦 [{idx+1}] **Spec:** {spec} | **Grade:** {grade} | **Variant:**"
        f" {variant} | **Source:** {source} | **Section:** {section} (Min."
        f" Tensile: {tensile} MPa)"
    )

    with st.expander(card_label, expanded=(idx == 0)):
      nom_comp = record.get(
          next((c for c in record.index if "nominal" in str(c).lower()), "-"), "-"
      )
      prod_form = record.get(
          next((c for c in record.index if "product" in str(c).lower()), "-"), "-"
      )
      class_cond = record.get(
          next((c for c in record.index if "class" in str(c).lower()), "-"), "-"
      )
      size_thick = record.get(
          next((c for c in record.index if "size" in str(c).lower()), "-"), "-"
      )
      y_col_name = [c for c in record.index if "yield" in str(c).lower()]
      min_yield = record.get(y_col_name[0], "-") if y_col_name else "-"

      raw_temp_limit = record.get(db_col, "NP") if source == "DB_AS" else "-"
      max_temp_limit = (
          f"{raw_temp_limit} °C"
          if pd.notnull(raw_temp_limit)
          and str(raw_temp_limit).strip() not in ["NP", "-", ""]
          else str(raw_temp_limit)
      )

      ext_chart = record.get(
          next((c for c in record.index if "ext" in str(c).lower()), "-"), "-"
      )
      notes = record.get(
          next((c for c in record.index if "note" in str(c).lower()), "-"), "-"
      )

      te_group, tcd_group, tm_group, poisson, density = (
          "Group 1",
          "Group A",
          "C<=0.3%",
          sidebar_poisson,
          sidebar_density,
      )
      if not db_map.empty:
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
            "Min. Tensile (Ultimate)",
            (
                f"{float(tensile):.3f} MPa"
                if pd.notnull(tensile)
                and str(tensile).replace(".", "", 1).isdigit()
                else str(tensile)
            ),
        )

      r2_c1, r2_c2, r2_c3, r2_c4, r2_c5 = st.columns(5)
      with r2_c1:
        render_meta_item(
            "Min. Yield",
            (
                f"{float(min_yield):.3f} MPa"
                if pd.notnull(min_yield)
                and str(min_yield).replace(".", "", 1).isdigit()
                else "-"
            ),
        )
      with r2_c2:
        render_meta_item("Max Temp Limit", str(max_temp_limit))
      with r2_c3:
        render_meta_item("Ext. Chart No.", str(ext_chart))
      with r2_c4:
        render_meta_item(
            "Property Groups",
            f"TE: {te_group} | TCD: {tcd_group} | TM: {tm_group}",
        )
      with r2_c5:
        render_meta_item("Notes", str(notes) if pd.notnull(notes) else "-")

      st.markdown("---")

      # --- DYNAMIC TEMPERATURE EVALUATION INPUT ---
      st.markdown("#### 🎯 Evaluate at Additional Temperature(s)")
      col_t1, col_t2 = st.columns([3, 1])
      with col_t1:
        new_temp_input = st.text_input(
            "Enter Temperature(s) in °C (comma-separated)",
            key=f"temp_input_{idx}",
            placeholder="e.g. 150, 250, 350",
        )
      with col_t2:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if st.button("Evaluate Temp", key=f"eval_btn_{idx}"):
          try:
            parsed_temps = [
                float(t.strip())
                for t in new_temp_input.split(",")
                if t.strip().replace(".", "", 1).isdigit()
            ]
            for pt in parsed_temps:
              if pt not in item["eval_temps"]:
                item["eval_temps"].append(pt)
            st.rerun()
          except Exception:
            st.error("Invalid temperature format.")

      if not item.get("eval_temps"):
        item["eval_temps"] = [20.0]

      y1_row = get_variant_subset(db_y1, spec_col_y1, grade_col_y1)
      as_row = get_variant_subset(db_as, spec_col_as, grade_col_as)
      u_row = get_variant_subset(db_u, spec_col_u, grade_col_u)

      multi_temp_records = []
      for t in sorted(item["eval_temps"]):
        if source == "DB_AS":
          s_val = get_row_stress_at_temp(record, temp_cols, t)
        else:
          if not as_row.empty:
            s_val = get_row_stress_at_temp(as_row.iloc[0], temp_cols, t)
          else:
            s_val = "N/A"

        y_val = (
            get_row_stress_at_temp(y1_row.iloc[0], temp_cols, t)
            if not y1_row.empty
            else "-"
        )
        u_val = (
            get_row_stress_at_temp(u_row.iloc[0], temp_cols, t)
            if not u_row.empty
            else "-"
        )

        e_i = interpolate_prop(db_tm, "T (˚C)", "E", t, "TM GR.", tm_group)
        tc_i = interpolate_prop(db_tcd, "T (˚C)", "TC", t, "TCD GR.", tcd_group)
        td_i = interpolate_prop(db_tcd, "T (˚C)", "TD", t, "TCD GR.", tcd_group)
        a_i = interpolate_prop(db_te, "T (˚C)", "A", t, "TE GROUP", te_group)
        b_i = interpolate_prop(db_te, "T (˚C)", "B", t, "TE GROUP", te_group)

        cp_i = "-"
        try:
          if (
              isinstance(tc_i, (int, float))
              and isinstance(td_i, (int, float))
              and td_i > 0
              and density
          ):
            cp_i = (tc_i * (10**6)) / (float(density) * td_i)
        except Exception:
          cp_i = "-"

        multi_temp_records.append({
            "Temp (°C)": t,
            "Allowable Stress (MPa)": (
                round(s_val, 3) if isinstance(s_val, (int, float)) else s_val
            ),
            "Yield Strength (MPa)": (
                round(y_val, 3) if isinstance(y_val, (int, float)) else y_val
            ),
            "Ultimate Tensile Strength (MPa)": (
                round(u_val, 3) if isinstance(u_val, (int, float)) else u_val
            ),
            "Modulus E (GPa)": (
                round(e_i, 3) if isinstance(e_i, (int, float)) else e_i
            ),
            "Poisson Ratio (–)": round(poisson, 3),
            "Density (kg/m³)": round(density, 3),
            "Thermal Cond. TC (W/m·°C)": (
                round(tc_i, 3) if isinstance(tc_i, (int, float)) else tc_i
            ),
            "Thermal Diff. TD (10⁻⁶ m²/s)": (
                round(td_i, 3) if isinstance(td_i, (int, float)) else td_i
            ),
            "Specific Heat Cp (J/kg·°C)": (
                round(cp_i, 3) if isinstance(cp_i, (int, float)) else cp_i
            ),
            "Thermal Exp. A (mm/mm/°C)": (
                f"{a_i:.3e}" if isinstance(a_i, (int, float)) else a_i
            ),
            "Thermal Exp. B (mm/mm/°C)": (
                f"{b_i:.3e}" if isinstance(b_i, (int, float)) else b_i
            ),
        })

      st.dataframe(pd.DataFrame(multi_temp_records), use_container_width=True)

      sub_tab1, sub_tab2 = st.tabs(
          ["📑 Full Temperature Breakdown", "📈 Trend Graph"]
      )

      with sub_tab1:
        if temp_cols:
          stress_values = (
              [record.get(t, "NP") for t in temp_cols]
              if source == "DB_AS"
              else ["-"] * len(temp_cols)
          )
          yield_values = [
              y1_row.iloc[0].get(t, "-") if not y1_row.empty else "-"
              for t in temp_cols
          ]
          tensile_values = [
              u_row.iloc[0].get(t, "-") if not u_row.empty else "-"
              for t in temp_cols
          ]
          te_a_v, te_b_v, tc_v, td_v, cp_v, e_v = [], [], [], [], [], []

          for t_str in temp_cols:
            try:
              t_num = float(t_str)
            except ValueError:
              continue
            a_v = interpolate_prop(
                db_te, "T (˚C)", "A", t_num, "TE GROUP", te_group
            )
            b_v = interpolate_prop(
                db_te, "T (˚C)", "B", t_num, "TE GROUP", te_group
            )
            te_a_v.append(
                f"{a_v:.3e}" if isinstance(a_v, (int, float)) else a_v
            )
            te_b_v.append(
                f"{b_v:.3e}" if isinstance(b_v, (int, float)) else b_v
            )

            tc_item = interpolate_prop(
                db_tcd, "T (˚C)", "TC", t_num, "TCD GR.", tcd_group
            )
            td_item = interpolate_prop(
                db_tcd, "T (˚C)", "TD", t_num, "TCD GR.", tcd_group
            )
            tc_v.append(
                round(tc_item, 3) if isinstance(tc_item, (int, float)) else tc_item
            )
            td_v.append(
                round(td_item, 3) if isinstance(td_item, (int, float)) else td_item
            )

            try:
              if (
                  isinstance(tc_item, (int, float))
                  and isinstance(td_item, (int, float))
                  and td_item > 0
                  and density
              ):
                cp_v.append(
                    round(
                        (tc_item * (10**6)) / (float(density) * td_item), 3
                    )
                )
              else:
                cp_v.append("-")
            except Exception:
              cp_v.append("-")

            e_item = interpolate_prop(
                db_tm, "T (˚C)", "E", t_num, "TM GR.", tm_group
            )
            e_v.append(
                round(e_item, 3) if isinstance(e_item, (int, float)) else e_item
            )

          valid_temp_cols = [
              t
              for t in temp_cols
              if str(t).strip().replace(".", "", 1).isdigit()
          ]
          full_df = pd.DataFrame({
              "Temp (°C)": valid_temp_cols[: len(e_v)],
              "Allowable Stress (MPa)": stress_values[: len(e_v)],
              "Yield Strength (MPa)": yield_values[: len(e_v)],
              "Ultimate Tensile Strength (MPa)": tensile_values[: len(e_v)],
              "Modulus E (GPa)": e_v,
              "Poisson Ratio (–)": round(poisson, 3),
              "Density (kg/m³)": round(density, 3),
              "Thermal Cond. TC (W/m·°C)": tc_v,
              "Thermal Diff. TD (10⁻⁶ m²/s)": td_v,
              "Specific Heat Cp (J/kg·°C)": cp_v,
              "Thermal Exp. A (mm/mm/°C)": te_a_v,
              "Thermal Exp. B (mm/mm/°C)": te_b_v,
          })
          st.dataframe(full_df, use_container_width=True, height=300)

      with sub_tab2:
        if temp_cols and "full_df" in locals() and not full_df.empty:
          chart_metric = st.selectbox(
              "Select Property to Plot:",
              [
                  "Allowable Stress (MPa)",
                  "Yield Strength (MPa)",
                  "Ultimate Tensile Strength (MPa)",
                  "Modulus E (GPa)",
                  "Thermal Cond. TC (W/m·°C)",
                  "Specific Heat Cp (J/kg·°C)",
              ],
              key=f"chart_{idx}",
          )
          plot_df = full_df[["Temp (°C)", chart_metric]].copy()
          plot_df["Temp (°C)"] = pd.to_numeric(
              plot_df["Temp (°C)"], errors="coerce"
          )
          plot_df[chart_metric] = pd.to_numeric(
              plot_df[chart_metric], errors="coerce"
          )
          plot_df = plot_df.dropna().sort_values("Temp (°C)")

          if not plot_df.empty:
            st.line_chart(
                plot_df.set_index("Temp (°C)")[chart_metric],
                use_container_width=True,
            )
          else:
            st.info("No numeric data available to plot.")
