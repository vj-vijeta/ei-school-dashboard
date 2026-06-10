import streamlit as st
import pandas as pd
import io
import os

# 1. Page Setup & Theme Styling
st.set_page_config(page_title="Schools Master Dashboard", layout="wide", page_icon="🏫")

st.markdown("""
    <style>
    div[data-testid="metric-container"] {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #e1e4e8;
    }
    .main-header {
        font-size: 20px;
        font-weight: 700;
        color: #1A365D;
        margin-top: 15px;
        margin-bottom: 10px;
    }
    .sub-header {
        font-size: 16px;
        font-weight: 600;
        color: #4A5568;
        background-color: #EDF2F7;
        padding: 6px 12px;
        border-radius: 4px;
        margin-top: 20px;
        margin-bottom: 10px;
    }
    /* Custom colored status pills */
    .status-badge {
        display: inline-block;
        padding: 4px 10px;
        font-size: 14px;
        font-weight: bold;
        border-radius: 12px;
        text-align: center;
        min-width: 100px;
    }
    .status-completed { background-color: #C6F6D5; color: #22543D; }
    .status-pending { background-color: #FEEBC8; color: #7B341E; }
    .status-progress { background-color: #BEE3F8; color: #2A4365; }
    .status-none { background-color: #EDF2F7; color: #4A5568; }
    </style>
""", unsafe_allow_html=True)

# Projects to exclude
EXCLUSIONS = ['Thrive', 'Kalinga']

# --- HELPER: AUTO-CLEANER ---
def auto_clean_data(df):
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip().str.title()
            if 'status' in col.lower() or col == 'Confirmation':
                replacements = {
                    'Done': 'Completed', 'Yes': 'Completed', 'Uploaded': 'Completed', 'Active': 'Completed',
                    'No': 'Pending', 'Not Started': 'Pending', 'Nan': 'Pending', 'None': 'Pending',
                    'Ongoing': 'In Progress'
                }
                df[col] = df[col].replace(replacements)
            df[col] = df[col].replace({'Nan': ''})
    return df

# --- HELPER: FORM DISPLAY ---
def display_as_readonly_form(row_df, exclude_cols):
    """Takes a single row dataframe and displays it beautifully as a read-only form."""
    if row_df.empty: return
    
    valid_cols = [c for c in row_df.columns if c not in exclude_cols and pd.notna(row_df[c].values[0]) and str(row_df[c].values[0]).strip() != '']
    
    for i in range(0, len(valid_cols), 2):
        col1, col2 = st.columns(2)
        with col1:
            val1 = str(row_df[valid_cols[i]].values[0])
            st.text_input(valid_cols[i], value=val1, disabled=True, key=f"{valid_cols[i]}_{i}")
        with col2:
            if i + 1 < len(valid_cols):
                val2 = str(row_df[valid_cols[i+1]].values[0])
                st.text_input(valid_cols[i+1], value=val2, disabled=True, key=f"{valid_cols[i+1]}_{i+1}")


# --- 2. DATA LOADER & COMPILER ---
@st.cache_data(ttl=5)
def load_all_local_excel_files():
    combined_data = {}
    master_df = None
    
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$')]
    
    for file in excel_files:
        try:
            xls = pd.ExcelFile(file)
            
            if "For Vijeta" in file or len(xls.sheet_names) == 1:
                df = pd.read_excel(file, header=1) 
                
                if 'School Name' not in df.columns and 'School Type' not in df.columns:
                    df = pd.read_excel(file)
                    df.columns = df.iloc[0]
                    df = df.drop(0).reset_index(drop=True)
                
                master_df = df
                continue
                
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name)
                if sheet_name in ['CARES', 'Mindspark', 'Check List  For CARES', 'Check List for Mindspark'] and 'School Name' in df.columns:
                    df = df[~df['School Name'].isin(EXCLUSIONS)]
                    
                df = auto_clean_data(df)
                
                if sheet_name not in combined_data:
                    combined_data[sheet_name] = []
                combined_data[sheet_name].append(df)
                
        except Exception as e:
            st.error(f"Error reading {file}: {e}")
            
    final_sheets = {}
    for sheet_name, df_list in combined_data.items():
        if df_list:
            final_sheets[sheet_name] = pd.concat(df_list, ignore_index=True)
            
    return final_sheets, master_df

# Load Data into Session State
if 'master_sheets' not in st.session_state or 'master_db' not in st.session_state:
    sheets, master_db = load_all_local_excel_files()
    st.session_state.master_sheets = sheets
    st.session_state.master_db = master_db
    
    if st.session_state.master_sheets and 'ASSET' not in st.session_state.master_sheets:
        st.session_state.master_sheets['ASSET'] = pd.DataFrame(columns=['School Name', 'ASSET Onboarding Status', 'Notes'])

if not st.session_state.master_sheets and st.session_state.master_db is None:
    st.warning("⚠️ No Excel files found. Please make sure your .xlsx files (including the Master) are in the same folder.")
    st.stop()

# Assign Variables
df_master = st.session_state.master_db if st.session_state.master_db is not None else pd.DataFrame()
df_cares = st.session_state.master_sheets.get('CARES', pd.DataFrame())
df_ms = st.session_state.master_sheets.get('Mindspark', pd.DataFrame())
df_checklist_cares = st.session_state.master_sheets.get('Check List  For CARES', pd.DataFrame())
df_asset = st.session_state.master_sheets.get('ASSET', pd.DataFrame())

# Build the master school list
all_school_names = set()
if not df_master.empty and 'School Name' in df_master.columns:
    all_school_names.update(df_master['School Name'].replace('', pd.NA).dropna().unique())
else:
    for df in [df_cares, df_ms, df_asset]:
        if 'School Name' in df.columns:
            all_school_names.update([n for n in df['School Name'].dropna().unique() if n != ''])
            
all_schools_list = sorted(list(all_school_names))

# Helper Functions
def get_excel_download(sheets_dict):
    output_stream = io.BytesIO()
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        for s_name, data_df in sheets_dict.items():
            data_df.to_excel(writer, sheet_name=s_name, index=False)
    return output_stream.getvalue()

def render_status_pill(status_text):
    text_str = str(status_text).strip().lower()
    if text_str in ['completed', 'done', 'active', 'yes', 'uploaded']:
        return f'<span class="status-badge status-completed">{status_text.title()}</span>'
    elif text_str in ['pending', 'no', 'not started', 'nan', '']:
        return f'<span class="status-badge status-pending">Pending</span>'
    elif text_str in ['in progress', 'ongoing']:
        return f'<span class="status-badge status-progress">In Progress</span>'
    else:
        return f'<span class="status-badge status-none">{status_text}</span>'

def has_product_in_master(school_name, product_keyword):
    if df_master.empty or 'School Name' not in df_master.columns or 'Offering' not in df_master.columns:
        return False
    row = df_master[df_master['School Name'] == school_name]
    if not row.empty:
        offerings = str(row['Offering'].values[0]).lower()
        return product_keyword.lower() in offerings
    return False


# 3. Sidebar Menu
st.sidebar.title("Main Menu")
app_view = st.sidebar.radio(
    "Choose a Section", 
    ["🏠 Home: Region Overview", "🎯 View School Details", "✏️ Edit Tracking Data"]
)

st.sidebar.divider()
st.sidebar.subheader("Download Tracking Data")
st.sidebar.download_button(
    label="📥 Download Updated Tracker",
    data=get_excel_download(st.session_state.master_sheets),
    file_name="Combined_Master_Tracker.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

# =====================================================================================
# VIEW 1: HOME PAGE (REGION OVERVIEW)
# =====================================================================================
if app_view == "🏠 Home: Region Overview":
    st.title("🏠 Region Overview")
    st.write("Hover your mouse over any metric card or category to see the complete list of matching schools.")
    st.divider()
    
    # 1. Gather lists for metric hover text
    cares_unique = sorted(list(df_cares['School Name'].replace('', pd.NA).dropna().unique())) if not df_cares.empty else []
    ms_unique = sorted(list(df_ms['School Name'].replace('', pd.NA).dropna().unique())) if not df_ms.empty else []
    asset_unique = sorted(list(df_asset[df_asset['ASSET Onboarding Status'] != 'Not Started']['School Name'].unique())) if not df_asset.empty else []
    
    cares_hover = "🏫 Schools with CARES:\n" + "\n".join([f"- {s}" for s in cares_unique]) if cares_unique else "No schools found."
    ms_hover = "🏫 Schools with Mindspark:\n" + "\n".join([f"- {s}" for s in ms_unique]) if ms_unique else "No schools found."
    asset_hover = "🏫 Schools with ASSET:\n" + "\n".join([f"- {s}" for s in asset_unique]) if asset_unique else "No schools found."
    total_hover = "🏫 All Schools List:\n" + "\n".join([f"- {s}" for s in all_schools_list])
    
    # Core Aggregation Metrics Row with Hover Ability
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Master Schools", len(all_schools_list), help=total_hover)
    col2.metric("Mapped in CARES Tracker", len(cares_unique), help=cares_hover)
    col3.metric("Mapped in Mindspark Tracker", len(ms_unique), help=ms_hover)
    col4.metric("Manually Tracking ASSET", len(asset_unique), help=asset_hover)
    
    st.divider()
    
    # 2. Regional Institutional Classifications with Hover Lists
    st.markdown('<p class="main-header">Regional Institutional Classifications</p>', unsafe_allow_html=True)
    st.write("Hover over the cards below to see which schools belong to each category.")
    
    # Combine classifications across sheets to get accurate mappings
    school_type_map = {}
    for df in [df_cares, df_ms]:
        if 'School Name' in df.columns and 'School Type' in df.columns:
            for _, row in df.iterrows():
                name, stype = str(row['School Name']), str(row['School Type'])
                if name and stype and name != 'Nan' and stype != 'Nan' and name != '':
                    school_type_map[name] = stype
                    
    # Group school names by their type category
    type_groups = {}
    for name, stype in school_type_map.items():
        if stype not in type_groups:
            type_groups[stype] = []
        type_groups[stype].append(name)
        
    if type_groups:
        type_cols = st.columns(len(type_groups))
        for idx, (stype, schools) in enumerate(type_groups.items()):
            with type_cols[idx]:
                category_hover = f"🏫 Schools in {stype}:\n" + "\n".join([f"- {s}" for s in sorted(schools)])
                st.metric(label=stype, value=len(schools), help=category_hover)
    else:
        st.info("No school category classifications found in tracking files.")
        
    st.divider()
    
    # 3. Interactive Summary Table List
    st.markdown('<p class="main-header">Complete School Summary & Status List</p>', unsafe_allow_html=True)
    st.write("Click any column header header below (such as **Division**) to sort the rows instantly.")
    
    summary_data = []
    for school in all_schools_list:
        m_match = df_master[df_master['School Name'] == school] if not df_master.empty else pd.DataFrame()
        division = m_match['Division'].values[0] if not m_match.empty and 'Division' in m_match.columns else "N/A"
        
        a_match = df_asset[df_asset['School Name'] == school] if not df_asset.empty else pd.DataFrame()
        a_status = a_match['ASSET Onboarding Status'].values[0] if not a_match.empty else "Not Started"
        
        c_match = df_cares[df_cares['School Name'] == school] if not df_cares.empty else pd.DataFrame()
        c_status = c_match['Introduction with School Status'].values[0] if not c_match.empty and 'Introduction with School Status' in c_match.columns else "No Program"
        
        m_match_ms = df_ms[df_ms['School Name'] == school] if not df_ms.empty else pd.DataFrame()
        m_status = m_match_ms['Introduction with School Status'].values[0] if not m_match_ms.empty and 'Introduction with School Status' in m_match_ms.columns else "No Program"
        
        summary_data.append({
            "School Name": school,
            "Division": division,
            "CARES Status": c_status,
            "Mindspark Status": m_status,
            "ASSET Status": a_status
        })
        
    summary_df = pd.DataFrame(summary_data)
    
    # Local table division filter
    home_divs = ["All Divisions"] + sorted([str(d) for d in summary_df['Division'].unique() if str(d) != 'nan' and str(d).strip() != ''])
    selected_home_div = st.selectbox("Filter Summary List by Division:", home_divs)
    
    if selected_home_div != "All Divisions":
        summary_df = summary_df[summary_df['Division'] == selected_home_div]
        
    st.dataframe(summary_df, use_container_width=True, hide_index=True)


# =====================================================================================
# VIEW 2: VIEW SCHOOL DETAILS
# =====================================================================================
elif app_view == "🎯 View School Details":
    st.title("🎯 School Explorer & Tracker")
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filters")
    
    # Filter by Zone
    available_zones = ["All Zones"]
    if not df_master.empty and 'Zone' in df_master.columns:
        valid_zones = [z for z in df_master['Zone'].dropna().unique() if str(z).strip() != '']
        available_zones.extend(sorted(valid_zones))
    selected_zone = st.sidebar.selectbox("Filter by Zone", available_zones)
    
    # Filter by Division
    available_divisions = ["All Divisions"]
    if not df_master.empty and 'Division' in df_master.columns:
        valid_divs = [d for d in df_master['Division'].dropna().unique() if str(d).strip() != '']
        available_divisions.extend(sorted(valid_divs))
    selected_division = st.sidebar.selectbox("Filter by Division", available_divisions)
    
    search_query = st.sidebar.text_input("🔍 Search School Name or No.", "").strip()
    
    # Filter school pool matching constraints
    filtered_schools = all_schools_list
    
    if selected_zone != "All Zones" and not df_master.empty:
        zone_schools = df_master[df_master['Zone'] == selected_zone]['School Name'].tolist()
        filtered_schools = [s for s in filtered_schools if s in zone_schools]
        
    if selected_division != "All Divisions" and not df_master.empty:
        div_schools = df_master[df_master['Division'] == selected_division]['School Name'].tolist()
        filtered_schools = [s for s in filtered_schools if s in div_schools]
        
    if search_query:
        if not df_master.empty and 'School No' in df_master.columns and search_query.isdigit():
            number_matches = df_master[df_master['School No'].astype(str).str.contains(search_query)]['School Name'].tolist()
            filtered_schools = [s for s in filtered_schools if s in number_matches or search_query.lower() in s.lower()]
        else:
            filtered_schools = [s for s in filtered_schools if search_query.lower() in s.lower()]
            
    filtered_schools = sorted(filtered_schools)
    
    if not filtered_schools:
        st.warning("No schools match your search or filters.")
    else:
        chosen_school = st.selectbox("Select a School to View:", filtered_schools)
        st.divider()
        
        master_row = df_master[df_master['School Name'] == chosen_school] if not df_master.empty else pd.DataFrame()
        c_row = df_cares[df_cares['School Name'] == chosen_school] if not df_cares.empty and 'School Name' in df_cares.columns else pd.DataFrame()
        m_row = df_ms[df_ms['School Name'] == chosen_school] if not df_ms.empty and 'School Name' in df_ms.columns else pd.DataFrame()
        a_row = df_asset[df_asset['School Name'] == chosen_school] if not df_asset.empty and 'School Name' in df_asset.columns else pd.DataFrame()
        
        tab_main, tab_cares, tab_ms, tab_asset, tab_raw = st.tabs(["📋 Master Details & Status", "📊 CARES Tracking Form", "⚡ Mindspark Tracking Form", "📘 ASSET Tracker", "📂 View Spreadsheets"])
        
        # --- SUB-TAB 1: MASTER PROFILE INFO ---
        with tab_main:
            st.markdown('<p class="main-header">Master Database Profile</p>', unsafe_allow_html=True)
            if not master_row.empty:
                m1, m2, m3 = st.columns(3)
                m1.metric("School No.", str(master_row['School No'].values[0]))
                m2.metric("Location", f"{master_row['City'].values[0]}, {master_row['State'].values[0]}")
                m3.metric("Division & Zone", f"{master_row['Division'].values[0]} | {master_row['Zone'].values[0]}")
                st.info(f"**Purchased Master Offerings:** {master_row['Offering'].values[0]}")
            else:
                st.warning("No details found for this school in the Master Database file.")

            st.markdown('<p class="main-header">Implementation Alignment Status</p>', unsafe_allow_html=True)
            k1, k2, k3 = st.columns(3)
            with k1:
                st.write("**CARES Implementation**")
                if has_product_in_master(chosen_school, "CARES"):
                    status = str(c_row['Introduction with School Status'].values[0]) if not c_row.empty else "Pending Configuration"
                    st.markdown(render_status_pill(status), unsafe_allow_html=True)
                else:
                    st.markdown(render_status_pill("No Program"), unsafe_allow_html=True)
                    
            with k2:
                st.write("**Mindspark Implementation**")
                if has_product_in_master(chosen_school, "Mindspark"):
                    status = str(m_row['Introduction with School Status'].values[0]) if not m_row.empty else "Pending Configuration"
                    st.markdown(render_status_pill(status), unsafe_allow_html=True)
                else:
                    st.markdown(render_status_pill("No Program"), unsafe_allow_html=True)
                    
            with k3:
                st.write("**ASSET Tracking**")
                if has_product_in_master(chosen_school, "ASSET"):
                    current_a_status = a_row['ASSET Onboarding Status'].values[0] if not a_row.empty else "Not Started"
                    st.markdown(render_status_pill(current_a_status), unsafe_allow_html=True)
                else:
                    st.markdown(render_status_pill("No Program"), unsafe_allow_html=True)

        # --- SUB-TAB 2: CARES (FORM LAYOUT) ---
        with tab_cares:
            if c_row.empty:
                st.warning("No CARES tracking row found for this school in the operational spreadsheet.")
            else:
                st.markdown('<p class="main-header">CARES Tracking Record</p>', unsafe_allow_html=True)
                status_c1, status_c2 = st.columns(2)
                with status_c1:
                    st.write("**Book Mapping Status:**")
                    st.markdown(render_status_pill(c_row.get('Book Mapping status', pd.Series(['N/A'])).values[0]), unsafe_allow_html=True)
                with status_c2:
                    st.write("**Student Onboarding Status:**")
                    st.markdown(render_status_pill(c_row.get('On Boardning Status For students', pd.Series(['N/A'])).values[0]), unsafe_allow_html=True)
                
                st.markdown('<p class="sub-header">All Tracker Details (Form View)</p>', unsafe_allow_html=True)
                # Exclude Sr No out of the form
                display_as_readonly_form(c_row, exclude_cols=['Sr No'])

        # --- SUB-TAB 3: MINDSPARK (FORM LAYOUT) ---
        with tab_ms:
            if m_row.empty:
                st.warning("No Mindspark tracking row found for this school in the operational spreadsheet.")
            else:
                st.markdown('<p class="main-header">Mindspark Tracking Record</p>', unsafe_allow_html=True)
                st.write("**Infra Check Status:**")
                st.markdown(render_status_pill(m_row.get('Infra check status', pd.Series(['Pending'])).values[0]), unsafe_allow_html=True)
                    
                st.markdown('<p class="sub-header">All Tracker Details (Form View)</p>', unsafe_allow_html=True)
                # Exclude Sr No and PDF copy status as requested
                display_as_readonly_form(m_row, exclude_cols=['Sr No', 'PDF copy status'])

        # --- SUB-TAB 4: ASSET TRACKER ---
        with tab_asset:
            st.markdown('<p class="main-header">Manual ASSET Onboarding</p>', unsafe_allow_html=True)
            if not has_product_in_master(chosen_school, "ASSET"):
                st.info("The master database does not list ASSET in the offerings for this school. You can still track it manually below if needed.")
                
            with st.form("asset_update_form"):
                current_status = a_row['ASSET Onboarding Status'].values[0] if not a_row.empty and 'ASSET Onboarding Status' in a_row.columns else "Not Started"
                current_notes = a_row['Notes'].values[0] if not a_row.empty and 'Notes' in a_row.columns else ""
                
                status_options = ["Not Started", "In Progress", "Completed"]
                default_index = status_options.index(current_status) if current_status in status_options else 0
                
                new_status = st.radio("Is ASSET Onboarding Done?", status_options, index=default_index)
                new_notes = st.text_area("Any Additional Notes?", value=str(current_notes) if pd.notna(current_notes) else "")
                
                if st.form_submit_button("Save ASSET Update", type="primary"):
                    if not a_row.empty:
                        idx = a_row.index[0]
                        st.session_state.master_sheets['ASSET'].at[idx, 'ASSET Onboarding Status'] = new_status
                        st.session_state.master_sheets['ASSET'].at[idx, 'Notes'] = new_notes
                    else:
                        new_row = pd.DataFrame([{'School Name': chosen_school, 'ASSET Onboarding Status': new_status, 'Notes': new_notes}])
                        st.session_state.master_sheets['ASSET'] = pd.concat([st.session_state.master_sheets['ASSET'], new_row], ignore_index=True)
                    
                    st.success("ASSET details saved! Download the updated file from the sidebar.")
                    st.rerun()

        # --- SUB-TAB 5: RAW SPREADSHEETS VIEW ---
        with tab_raw:
            st.markdown('<p class="main-header">View Operational Spreadsheets</p>', unsafe_allow_html=True)
            target_sheet_view = st.selectbox("Choose a tracker sheet to view:", list(st.session_state.master_sheets.keys()), key="viewer_sheet_select")
            st.dataframe(st.session_state.master_sheets[target_sheet_view], use_container_width=True)


# =====================================================================================
# VIEW 3: BULK DATA EDITOR ROOM
# =====================================================================================
elif app_view == "✏️ Edit Tracking Data":
    st.title("✏️ Edit School Tracking Data")
    st.write("Changes made here will update the tracking tables (not the Master DB).")
    st.divider()
    
    target_write_sheet = st.selectbox("Which tracker sheet do you want to edit?", list(st.session_state.master_sheets.keys()))
    sheet_data_buffer = st.session_state.master_sheets[target_write_sheet].copy()
    
    entry_interface_style = st.radio("How do you want to edit?", ["Edit in a Table", "Fill out a Form"])
    st.divider()
    
    if entry_interface_style == "Edit in a Table":
        st.markdown('<p class="main-header">Table Editor</p>', unsafe_allow_html=True)
        updated_grid_df = st.data_editor(sheet_data_buffer, use_container_width=True, num_rows="dynamic", key=f"grid_writer_{target_write_sheet}")
        
        if st.button("Save Table Changes", type="primary"):
            st.session_state.master_sheets[target_write_sheet] = updated_grid_df
            st.success("Changes saved! Download the updated file from the side menu.")
            st.cache_data.clear()
            
    else:
        st.markdown('<p class="main-header">Form Data Entry</p>', unsafe_allow_html=True)
        action_intent_profile = st.radio("What are you trying to do?", ["Update an Existing Tracker Row", "Add a Completely New Tracker Row"])
        
        with st.form(key=f"form_writer_channel_{target_write_sheet}"):
            form_payload_map = {}
            target_index_loc = None
            
            if action_intent_profile == "Update an Existing Tracker Row" and 'School Name' in sheet_data_buffer.columns:
                valid_school_names = [s for s in sheet_data_buffer['School Name'].dropna().unique() if s != '']
                target_form_school = st.selectbox("Select the school to update:", sorted(valid_school_names))
                
                if target_form_school:
                    loc_matches = sheet_data_buffer[sheet_data_buffer['School Name'] == target_form_school].index
                    if len(loc_matches) > 0:
                        target_index_loc = loc_matches[0]
                        extant_row_metrics = sheet_data_buffer.loc[target_index_loc]
            
            st.markdown("---")
            for col_field in sheet_data_buffer.columns:
                if col_field == "Sr No":
                    continue
                    
                current_field_fallback = ""
                if action_intent_profile == "Update an Existing Tracker Row" and target_index_loc not in [None]:
                    current_field_fallback = extant_row_metrics[col_field] if pd.notna(extant_row_metrics[col_field]) else ""
                    
                if sheet_data_buffer[col_field].dtype in ['int64', 'float64']:
                    try:
                        form_payload_map[col_field] = st.number_input(f"{col_field}", value=int(current_field_fallback) if current_field_fallback != "" else 0)
                    except:
                        form_payload_map[col_field] = st.number_input(f"{col_field}", value=float(current_field_fallback) if current_field_fallback != "" else 0.0)
                elif isinstance(current_field_fallback, pd.Timestamp):
                    form_payload_map[col_field] = st.date_input(f"{col_field}", value=current_field_fallback.date())
                else:
                    form_payload_map[col_field] = st.text_input(f"{col_field}", value=str(current_field_fallback))
                    
            submit_form_execution = st.form_submit_button("Save Form Data", type="primary")
            
            if submit_form_execution:
                payload_series_object = pd.Series(form_payload_map)
                
                if action_intent_profile == "Update an Existing Tracker Row" and target_index_loc is not None:
                    if "Sr No" in sheet_data_buffer.columns:
                        payload_series_object["Sr No"] = sheet_data_buffer.loc[target_index_loc, "Sr No"]
                    sheet_data_buffer.loc[target_index_loc] = payload_series_object
                    st.session_state.master_sheets[target_write_sheet] = sheet_data_buffer
                    st.success("School tracker data updated successfully!")
                else:
                    if "Sr No" in sheet_data_buffer.columns:
                        payload_series_object["Sr No"] = len(sheet_data_buffer) + 1
                    sheet_data_buffer = pd.concat([sheet_data_buffer, pd.DataFrame([payload_series_object])], ignore_index=True)
                    st.session_state.master_sheets[target_write_sheet] = sheet_data_buffer
                    st.success("New school row added successfully!")
                st.cache_data.clear()