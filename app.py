import streamlit as st
import pandas as pd
import io
import os

# 1. Platform Design Architecture & Theme Styling
st.set_page_config(page_title="East Zone Enterprise Hub", layout="wide", page_icon="🏢")

# Custom UI status tags and styling adjustments
st.markdown("""
    <style>
    div[data-testid="metric-container"] {
        background-color: #fafbfc;
        padding: 20px;
        border-radius: 8px;
        border: 1px solid #e1e4e8;
    }
    .macro-header {
        font-size: 20px;
        font-weight: 700;
        color: #1A365D;
        border-left: 5px solid #3182CE;
        padding-left: 10px;
        margin-top: 15px;
        margin-bottom: 15px;
    }
    .micro-header {
        font-size: 16px;
        font-weight: 600;
        color: #4A5568;
        background-color: #EDF2F7;
        padding: 6px 12px;
        border-radius: 4px;
        margin-top: 25px;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

FILE_NAME = "PI Team East Zone.xlsx"
EXCLUSIONS = ['Thrive', 'Kalinga']

# 2. Resilient Data Compilation Engine
@st.cache_data(ttl=5)
def load_and_compile_sheets():
    if not os.path.exists(FILE_NAME):
        return {}
    try:
        xls = pd.ExcelFile(FILE_NAME)
        compiled_data = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(FILE_NAME, sheet_name=sheet)
            # Filter specific high-risk exclusion entries from key data spaces
            if sheet in ['CARES', 'Mindspark', 'Check List  For CARES', 'Check List for Mindspark'] and 'School Name' in df.columns:
                df = df[~df['School Name'].isin(EXCLUSIONS)]
            compiled_data[sheet] = df
        return compiled_data
    except Exception as e:
        st.error(f"Execution Error loading workbook layers: {e}")
        return {}

sheets = load_and_compile_sheets()

if not sheets:
    st.error(f"⚠️ Critical System Fault: File source '{FILE_NAME}' is missing from the working directory.")
    st.stop()

# Base Datasets Isolate
df_cares = sheets.get('CARES', pd.DataFrame())
df_ms = sheets.get('Mindspark', pd.DataFrame())
df_checklist_cares = sheets.get('Check List  For CARES', pd.DataFrame())

# Extract master key indices
cares_schools_set = set(df_cares['School Name'].dropna().unique()) if 'School Name' in df_cares.columns else set()
ms_schools_set = set(df_ms['School Name'].dropna().unique()) if 'School Name' in df_ms.columns else set()
unified_schools_list = sorted(list(cares_schools_set.union(ms_schools_set)))

def generate_binary_workbook(sheets_dict):
    output_stream = io.BytesIO()
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        for s_name, data_df in sheets_dict.items():
            data_df.to_excel(writer, sheet_name=s_name, index=False)
    return output_stream.getvalue()

def render_color_status(status_text):
    text_str = str(status_text).strip().lower()
    if text_str in ['completed', 'done', 'active', 'uploaded']:
        st.success(f"🟢 **{status_text}**")
    elif text_str in ['pending', 'not started', 'onboarding', 'in progress']:
        st.warning(f"🟡 **{status_text}**")
    else:
        st.info(f"🔵 **{status_text}**")


# 3. Main Enterprise Navigation Split
st.sidebar.title("🌐 Enterprise Hub Control")
app_workspace = st.sidebar.radio(
    "Select System Workspace", 
    ["🏠 Regional Home Summary", "🎯 Granular School Data Explorer", "⚙️ CRM Database Editor Room"]
)

# Shared Global Download Component across views
st.sidebar.divider()
st.sidebar.subheader("Export Center")
st.sidebar.download_button(
    label="📥 Download Master Workbook (.xlsx)",
    data=generate_binary_workbook(sheets),
    file_name="PI_Team_East_Zone_Master.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)


# =====================================================================================
# INTERFACE VIEW 1: HOME PAGE GENERAL MACRO ANALYTICS
# =====================================================================================
if app_workspace == "🏠 Regional Home Summary":
    st.title("🏠 Regional Home Summary Overview")
    st.write("Macro-level footprint and implementation benchmarks across the East Zone region.")
    st.divider()
    
    # Core Aggregations Row
    kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
    kpi_col1.metric("Total Unique Target Accounts", len(unified_schools_list))
    kpi_col2.metric("Total Active CARES Platforms", len(cares_schools_set))
    kpi_col3.metric("Total Active Mindspark Environments", len(ms_schools_set))
    
    st.divider()
    
    sec_col1, sec_col2 = st.columns(2)
    with sec_col1:
        st.markdown('<p class="macro-header">🏷️ Introduction with School Status Count</p>', unsafe_allow_html=True)
        all_intro_statuses = []
        if 'Introduction with School Status' in df_cares.columns:
            all_intro_statuses.extend(df_cares['Introduction with School Status'].dropna().astype(str).tolist())
        if 'Introduction with School Status' in df_ms.columns:
            all_intro_statuses.extend(df_ms['Introduction with School Status'].dropna().astype(str).tolist())
        
        intro_series = pd.Series(all_intro_statuses)
        if not intro_series.empty:
            st.dataframe(intro_series.value_counts().to_frame("Total Accounts"), use_container_width=True)
        else:
            st.info("No recorded configuration alignment items found.")
            
        st.markdown('<p class="macro-header">🎓 CARES Student Onboarding Tracker</p>', unsafe_allow_html=True)
        if 'On Boardning Status For students' in df_cares.columns:
            st.dataframe(df_cares['On Boardning Status For students'].fillna('Unassigned/Pending').value_counts().to_frame("Total Schools"), use_container_width=True)
        else:
            st.info("Onboarding metrics missing inside source file tables.")

    with sec_col2:
        st.markdown('<p class="macro-header">📋 Student Orientation Status Verification</p>', unsafe_allow_html=True)
        if not df_checklist_cares.empty and 'Student Orientation status' in df_checklist_cares.columns:
            st.dataframe(df_checklist_cares['Student Orientation status'].fillna('Not Scheduled').value_counts().to_frame("Accounts Count"), use_container_width=True)
        else:
            st.info("No records mapped inside the 'Check List For CARES' workspace layer.")
            
        st.markdown('<p class="macro-header">🗺️ Regional Institutional Classifications</p>', unsafe_allow_html=True)
        combined_types = []
        if 'School Type' in df_cares.columns: combined_types.extend(df_cares['School Type'].dropna().tolist())
        if 'School Type' in df_ms.columns: combined_types.extend(df_ms['School Type'].dropna().tolist())
        type_series = pd.Series(combined_types)
        if not type_series.empty:
            st.dataframe(type_series.value_counts().to_frame("Active Environments"), use_container_width=True)


# =====================================================================================
# INTERFACE VIEW 2: GRANULAR SCHOOL DATA EXPLORER
# =====================================================================================
elif app_workspace == "🎯 Granular School Data Explorer":
    st.title("🎯 Granular School Data Workspace")
    
    # Scope Modifiers local to this discovery area
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎯 Explorer Search Filters")
    query_search = st.sidebar.text_input("Search Account Portfolio By Name", "").strip()
    
    deployment_filter = st.sidebar.radio(
        "Program Configuration Scope",
        ["Show All Accounts", "CARES Accounts Only", "Mindspark Accounts Only", "Dual Systems Running (Both)"]
    )
    
    # Narrow down scope array
    if deployment_filter == "CARES Accounts Only":
        working_pool = list(cares_schools_set)
    elif deployment_filter == "Mindspark Accounts Only":
        working_pool = list(ms_schools_set)
    elif deployment_filter == "Dual Systems Running (Both)":
        working_pool = list(cares_schools_set.intersection(ms_schools_set))
    else:
        working_pool = unified_schools_list
        
    all_types = set()
    for d_set in [df_cares, df_ms]:
        if 'School Type' in d_set.columns:
            all_types.update(d_set['School Type'].dropna().unique())
            
    selected_class_type = st.sidebar.selectbox("School Classification Type", ["All Types"] + list(all_types))
    
    # Process final arrays matching conditions
    final_active_pool = []
    for s_entry in working_pool:
        is_valid_type = True
        if selected_class_type != "All Types":
            c_check = df_cares[df_cares['School Name'] == s_entry]['School Type'].tolist()
            m_check = df_ms[df_ms['School Name'] == s_entry]['School Type'].tolist()
            if selected_class_type not in (c_check + m_check):
                is_valid_type = False
                
        is_valid_search = True
        if query_search and query_search.lower() not in s_entry.lower():
            is_valid_search = False
            
        if is_valid_type and is_valid_search:
            final_active_pool.append(s_entry)
            
    final_active_pool = sorted(final_active_pool)
    
    if not final_active_pool:
        st.warning("No institutional records match your current discovery settings.")
    else:
        chosen_school = st.selectbox("Select Target Institution Portfolio", final_active_pool)
        st.divider()
        
        # Slices Isolation
        c_row = df_cares[df_cares['School Name'] == chosen_school]
        m_row = df_ms[df_ms['School Name'] == chosen_school]
        ch_row = df_checklist_cares[df_checklist_cares['School Name'] == chosen_school]
        
        # Tabs Construction (Read-Only Layout)
        tab_key, tab_cares, tab_ms, tab_immutable = st.tabs([
            "📋 Combined Summary Key Details", 
            "📊 Comprehensive CARES Details", 
            "⚡ Comprehensive Mindspark Details",
            "📂 Immutable Database Sheet Viewer"
        ])
        
        # --- SUB-TAB 1: INITIAL COMBINED PROFILE KEY DETAILS ---
        with tab_key:
            st.markdown('<p class="macro-header">Macroscopic Performance Baseline Parameters</p>', unsafe_allow_html=True)
            k_col1, k_col2, k_col3 = st.columns(3)
            
            c_intro_val = str(c_row['Introduction with School Status'].values[0]) if not c_row.empty else "Not Deployed"
            m_intro_val = str(m_row['Introduction with School Status'].values[0]) if not m_row.empty else "Not Deployed"
            
            with k_col1:
                st.write("**CARES Account Alignment Status**")
                render_color_status(c_intro_val)
            with k_col2:
                st.write("**Mindspark Engagement Status**")
                render_color_status(m_intro_val)
            with k_col3:
                total_enrolled = 0
                if not c_row.empty: total_enrolled += int(c_row['Total Students'].fillna(0).values[0])
                if not m_row.empty: total_enrolled += int(m_row['Total Students'].fillna(0).values[0])
                st.metric("Unified Registration Base", f"{total_enrolled} Students")
                
            st.markdown('<p class="micro-header">Microscopic Execution Records & Data Trails</p>', unsafe_allow_html=True)
            st.write("Granular operational tracking rows associated with the active asset layers:")
            if not c_row.empty:
                st.caption("Combined CARES Core Footprint Record Row:")
                st.dataframe(c_row.dropna(axis=1, how='all'), use_container_width=True)
            if not m_row.empty:
                st.caption("Combined Mindspark Core Footprint Record Row:")
                st.dataframe(m_row.dropna(axis=1, how='all'), use_container_width=True)

        # --- SUB-TAB 2: COMPREHENSIVE CARES DETAILS ---
        with tab_cares:
            if c_row.empty:
                st.warning("This enterprise entity is not mapped to use CARES program matrices.")
            else:
                st.markdown('<p class="macro-header">Macroscopic Metrics Summary Profile</p>', unsafe_allow_html=True)
                cm1, cm2, cm3 = st.columns(3)
                cm1.metric("Current Classification", str(c_row['School Type'].values[0]))
                cm2.metric("Subject Tracking Vol", str(c_row['Total Subjects'].values[0]))
                cm3.metric("Assigned Cycle Allocation", str(c_row['Total Cycles'].values[0]))
                
                status_c1, status_c2 = st.columns(2)
                with status_c1:
                    st.write("**Book Mapping Milestone Execution:**")
                    render_color_status(c_row.get('Book Mapping status', pd.Series(['N/A'])).values[0])
                with status_c2:
                    st.write("**Onboarding Student Framework Milestone:**")
                    render_color_status(c_row.get('On Boardning Status For students', pd.Series(['N/A'])).values[0])
                    
                st.markdown('<p class="micro-header">Microscopic Fields Matrix Audit Ledger</p>', unsafe_allow_html=True)
                st.dataframe(c_row.dropna(axis=1, how='all').T, use_container_width=True)

        # --- SUB-TAB 3: COMPREHENSIVE MINDSPARK DETAILS ---
        with tab_ms:
            if m_row.empty:
                st.warning("This enterprise entity is not mapped to use Mindspark program matrices.")
            else:
                st.markdown('<p class="macro-header">Macroscopic Metrics Summary Profile</p>', unsafe_allow_html=True)
                mm1, mm2, mm3 = st.columns(3)
                mm1.metric("Platform Start Date", str(m_row['Programme Start Date'].values[0]))
                mm2.metric("Account Value Volume", f"INR {m_row['Revenue'].values[0]:,}" if 'Revenue' in m_row.columns and pd.notna(m_row['Revenue'].values[0]) else "N/A")
                mm3.metric("Total Active Enrollees", f"{int(m_row['Total Students'].values[0])} Students")
                
                status_m1, status_m2 = st.columns(2)
                with status_m1:
                    st.write("**Infrastructure Baseline Architecture Clearance Status:**")
                    render_color_status(m_row.get('Infra check status', pd.Series(['Pending'])).values[0])
                with status_m2:
                    st.write("**PDF Operations Documentation Upload Status:**")
                    render_color_status(m_row.get('PDF copy status', pd.Series(['No File'])).values[0])
                    
                st.markdown('<p class="micro-header">Microscopic Fields Matrix Audit Ledger</p>', unsafe_allow_html=True)
                st.dataframe(m_row.dropna(axis=1, how='all').T, use_container_width=True)

        # --- SUB-TAB 4: IMMUTABLE REFERENCE VAULT OVERVIEW ---
        with tab_immutable:
            st.markdown('<p class="macro-header">Global Read-Only File Ledger Vault</p>', unsafe_allow_html=True)
            st.write("Uneditable reference lookups covering all underlying structural spreadsheet worksheets:")
            target_sheet_view = st.selectbox("Choose Sheet Model to Inspect", list(sheets.keys()), key="viewer_sheet_select")
            st.dataframe(sheets[target_sheet_view], use_container_width=True)


# =====================================================================================
# INTERFACE VIEW 3: CRM WRITER & UPDATE SUITE
# =====================================================================================
elif app_workspace == "⚙️ CRM Database Editor Room":
    st.title("⚙️ CRM Production Operations Suite")
    st.write("Secure workspace environment reserved strictly for configuration alterations and data processing.")
    st.divider()
    
    target_write_sheet = st.selectbox("Target Update Sheet Module Layer", list(sheets.keys()))
    sheet_data_buffer = sheets[target_write_sheet].copy()
    
    entry_interface_style = st.radio("Modification Methodology Form-Factor", ["Edit as Sheet (Tabular Interaction)", "Edit as Form (Direct Form Inputs Fields)"])
    st.divider()
    
    # 1. METHODOLOGY OPTION A: EDIT AS SPREADSHEET
    if entry_interface_style == "Edit as Sheet (Tabular Interaction)":
        st.markdown('<p class="macro-header">Tabular Update Grid Workspace Ledger</p>', unsafe_allow_html=True)
        st.info("💡 Double-click any variable or record value inside the grid array below to make updates. Hit 'Commit Workspace Alterations' to save modifications.")
        
        updated_grid_df = st.data_editor(sheet_data_buffer, use_container_width=True, num_rows="dynamic", key=f"grid_writer_{target_write_sheet}")
        
        if st.button("Commit Workspace Alterations", type="primary"):
            sheets[target_write_sheet] = updated_grid_df
            st.success(f"Updates successfully written to working memory for `{target_write_sheet}` sheet module layer! Download the compiled sheet using the sidebar tool.")
            st.cache_data.clear()
            
    # 2. METHODOLOGY OPTION B: EDIT AS DATA INPUT FORM
    else:
        st.markdown('<p class="macro-header">Form Field Record Writer Channel</p>', unsafe_allow_html=True)
        action_intent_profile = st.radio("Form Operation Intention Type", ["Modify Extant Row Profile Record", "Append Completely New Row Data"])
        
        with st.form(key=f"form_writer_channel_{target_write_sheet}"):
            form_payload_map = {}
            target_index_loc = None
            
            if action_intent_profile == "Modify Extant Row Profile Record" and 'School Name' in sheet_data_buffer.columns:
                target_form_school = st.selectbox("Select Target Destination Institution to Modify", sorted(sheet_data_buffer['School Name'].dropna().unique().tolist()))
                
                if target_form_school:
                    loc_matches = sheet_data_buffer[sheet_data_buffer['School Name'] == target_form_school].index
                    if len(loc_matches) > 0:
                        target_index_loc = loc_matches[0]
                        extant_row_metrics = sheet_data_buffer.loc[target_index_loc]
            
            st.markdown("---")
            # Generate input structures reflecting dataframe schemas dynamically
            for col_field in sheet_data_buffer.columns:
                if col_field == "Sr No":
                    continue
                    
                current_field_fallback = ""
                if action_intent_profile == "Modify Extant Row Profile Record" and target_index_loc not in [None]:
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
                    
            submit_form_execution = st.form_submit_button("Commit Form Parameters to Memory", type="primary")
            
            if submit_form_execution:
                payload_series_object = pd.Series(form_payload_map)
                
                if action_intent_profile == "Modify Extant Row Profile Record" and target_index_loc is not None:
                    if "Sr No" in sheet_data_buffer.columns:
                        payload_series_object["Sr No"] = sheet_data_buffer.loc[target_index_loc, "Sr No"]
                    sheet_data_buffer.loc[target_index_loc] = payload_series_object
                    sheets[target_write_sheet] = sheet_data_buffer
                    st.success("Target institutional record updated inside current session workspace data arrays.")
                else:
                    if "Sr No" in sheet_data_buffer.columns:
                        payload_series_object["Sr No"] = len(sheet_data_buffer) + 1
                    sheet_data_buffer = pd.concat([sheet_data_buffer, pd.DataFrame([payload_series_object])], ignore_index=True)
                    sheets[target_write_sheet] = sheet_data_buffer
                    st.success("Successfully structuralized new entity parameters within master catalog data layers.")
                st.cache_data.clear()