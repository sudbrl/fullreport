import streamlit as st
import pandas as pd
from openpyxl.styles import Font
from datetime import datetime
import io, os, tempfile, traceback

# --- Enhanced UI Styling ---
st.markdown("""
    <style>
    /* Hide default Streamlit components */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {visibility: hidden;}
    
    /* Global styling */
    .stApp {
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }
    
    /* Main container styling */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        background: rgba(255, 255, 255, 0.8);
        border-radius: 12px;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.06);
        border: 1px solid rgba(226, 232, 240, 0.8);
        margin: 2rem auto;
        max-width: 800px;
    }
    
    /* Title styling */
    .main-title {
        text-align: center;
        color: #1e293b;
        font-size: 2.2rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .sub-title {
        text-align: center;
        color: #64748b;
        font-size: 1rem;
        margin-bottom: 2rem;
        font-weight: 400;
    }
    
    /* File uploader styling */
    .stFileUploader > div > div {
        background: #ffffff;
        border: 2px solid #e2e8f0;
        border-radius: 8px;
        padding: 1.5rem;
        transition: all 0.2s ease;
    }
    
    .stFileUploader > div > div:hover {
        border-color: #3b82f6;
        background: #fefefe;
    }
    
    /* Button styling */
    .stButton > button {
        width: 100%;
        background: #1e293b;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.75rem 1.5rem;
        font-size: 0.95rem;
        font-weight: 500;
        transition: all 0.2s ease;
        letter-spacing: 0.025em;
    }
    
    .stButton > button:hover {
        background: #0f172a;
        transform: translateY(-1px);
    }
    
    .stButton > button:active {
        transform: translateY(0px);
    }
    
    /* Download button styling */
    .stDownloadButton > button {
        background: #059669;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.75rem 1.5rem;
        font-size: 0.95rem;
        font-weight: 500;
        transition: all 0.2s ease;
        width: 100%;
    }
    
    .stDownloadButton > button:hover {
        background: #047857;
        transform: translateY(-1px);
    }
    
    /* Progress bar styling */
    .stProgress > div > div > div > div {
        background: #3b82f6;
        border-radius: 4px;
    }
    
    /* Success/Error messages */
    .stSuccess {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 6px;
        padding: 0.75rem;
    }
    
    .stError {
        background: #fef2f2;
        border: 1px solid #fecaca;
        border-radius: 6px;
        padding: 0.75rem;
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background: linear-gradient(180deg, #f8fafc 0%, #e2e8f0 100%);
    }
    
    .css-1d391kg .stButton > button {
        background: #374151;
        border: 1px solid #4b5563;
        color: white;
    }
    
    .css-1d391kg .stButton > button:hover {
        background: #1f2937;
    }
    
    /* Spinner styling */
    .stSpinner > div {
        border-top-color: #3b82f6 !important;
    }
    </style>
""", unsafe_allow_html=True)

###############################################################################
# -------------------------  LOGIN PAGE  --------------------------------------
def login_page():
    st.markdown("""
        <style>
        .login-container {
            max-width: 380px;
            margin: 10vh auto;
            padding: 2.5rem;
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 4px 16px rgba(0,0,0,0.08);
            border: 1px solid #e2e8f0;
        }
        .login-header {
            font-size: 1.8rem;
            font-weight: 600;
            color: #1e293b;
            margin-bottom: 0.5rem;
            text-align: center;
        }
        .login-subtitle {
            font-size: 0.9rem;
            color: #64748b;
            text-align: center;
            margin-bottom: 2rem;
        }
        .stTextInput > div > div > input {
            border-radius: 6px;
            border: 2px solid #e2e8f0;
            padding: 0.75rem;
            font-size: 0.95rem;
            transition: all 0.2s ease;
        }
        .stTextInput > div > div > input:focus {
            border-color: #3b82f6;
        }
        </style>
        <div class="login-container">
            <div class="login-header">🔐 Secure Login</div>
            <div class="login-subtitle">Access your unified report generator</div>
        </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        username = st.text_input("👤 Username", placeholder="Enter your username")
        password = st.text_input("🔑 Password", type="password", placeholder="Enter your password")
        submitted = st.form_submit_button("🚀 Login")

    if submitted:
        if username in st.secrets.get("auth", {}) and password == st.secrets["auth"][username]:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ Invalid username or password. Please try again.")

###############################################################################
# -------------------------  APP ENTRY POINT  ---------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    login_page()
    st.stop()

###############################################################################
# -------------------------  SIDEBAR LOGOUT  ----------------------------------
with st.sidebar:
    st.markdown("### 🎛️ Controls")
    st.markdown("---")
    if st.button("🚪 Logout"):
        st.session_state["authenticated"] = False
        st.rerun()
    st.markdown("---")
    st.markdown("### ℹ️ Info")
    st.markdown("**Version:** 2.0")
    st.markdown("**Last Updated:** " + datetime.now().strftime("%Y-%m-%d"))

###############################################################################
# ----------------------------- CONSTANTS -------------------------------------
KEEP_SLIPPAGE = ['Branch Name', 'Main Code', 'Ac Type Desc', 'Name',
                 'Limit', 'Balance', 'Provision']
PROV_MAP = {'G': 1, 'W': 2, 'S': 3, 'D': 4, 'B': 5}
CAT_NAMES = {'G': 'Good', 'W': 'Watchlist', 'S': 'Substandard',
             'D': 'Doubtful', 'B': 'Bad'}
CAT_ORDER = ['Good', 'Watchlist', 'Substandard', 'Doubtful', 'Bad']

STAFF_LOANS = {
    'STAFF SOCIAL LOAN', 'STAFF VEHICLE LOAN', 'STAFF HOME LOAN',
    'STAFF FLEXIBLE LOAN', 'STAFF HOME LOAN(COF)',
    'STAFF VEHICLE FACILITY LOAN (EVF)'
}

###############################################################################
# ----------------------------- UTILITIES -------------------------------------
def autofit_excel(writer):
    for ws in writer.sheets.values():
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2

###############################################################################
# --------------------------- PRE-PROCESS -------------------------------------
def preprocess_slippage(df):
    df.columns = df.columns.str.strip()
    miss = [c for c in KEEP_SLIPPAGE if c not in df.columns]
    if miss:
        raise ValueError(f"Slippage – missing columns: {miss}")
    df = df[KEEP_SLIPPAGE].copy()

    df['Limit'] = pd.to_numeric(df['Limit'], errors='coerce')
    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
    df = df.dropna(subset=['Limit', 'Balance'])
    df = df[df['Limit'] != 0]

    df['Prov_init'] = df['Provision'].astype(str).str.upper().str[0]
    bad = df[~df['Prov_init'].isin(PROV_MAP)]
    if not bad.empty:
        raise ValueError(f"Invalid provision codes: {bad['Prov_init'].unique()}")

    df['Prov_rank'] = df['Prov_init'].map(PROV_MAP)
    df['Prov_cat'] = df['Prov_init'].map(CAT_NAMES)
    return df

def preprocess_comp(df):
    df = df.copy()
    df['Ac Type Desc'] = df['Ac Type Desc'].str.strip().str.upper()
    df = df[~df['Ac Type Desc'].isin({s.upper() for s in STAFF_LOANS})]
    df = df[df['Limit'] != 0]
    df = df[~df['Main Code'].isin({'AcType Total', 'Grand Total'})]
    return df

###############################################################################
# --------------------------- SLIPPAGE ----------------------------------------
def detect_slippage(df_prev, df_curr):
    prev = df_prev[['Main Code', 'Prov_rank', 'Prov_cat']].rename(
        columns={'Prov_rank': 'rank_prev', 'Prov_cat': 'cat_prev'})
    curr = df_curr[['Main Code', 'Prov_rank', 'Prov_cat']].rename(
        columns={'Prov_rank': 'rank_curr', 'Prov_cat': 'cat_curr'})

    merged = pd.merge(prev, curr, on='Main Code', how='inner')
    full = (df_curr[df_curr['Main Code'].isin(merged['Main Code'])]
            .merge(merged[['Main Code', 'rank_prev', 'cat_prev']], on='Main Code'))

    full['Movement'] = full.apply(
        lambda r: "Slippage" if r['Prov_rank'] > r['rank_prev'] else
                  "Upgrade" if r['Prov_rank'] < r['rank_prev'] else
                  "Stable", axis=1)

    cols = ['Branch Name', 'Main Code', 'Ac Type Desc', 'Name',
            'Limit', 'Balance', 'cat_prev', 'Prov_cat', 'Movement']
    return full[cols].rename(columns={'Prov_cat': 'cat_curr'})

def category_matrix(df, group_col=None):
    index = group_col if group_col else pd.Series(0, index=df.index, name='dummy')
    
    # Create pivot table with previous category breakdown
    mat = (df
           .pivot_table(index=index,
                        columns='cat_prev',
                        values='cat_curr',
                        aggfunc='size',
                        fill_value=0)
           .reindex(columns=CAT_ORDER, fill_value=0)
           .astype(int))
    
    # Add previous category column for summary sheets
    if group_col:
        # Get the previous category for each group
        prev_cat_mapping = df.groupby(group_col)['cat_prev'].first()
        result = mat.reset_index()
        result.insert(1, 'Previous_Category', result[group_col].map(prev_cat_mapping))
        return result
    else:
        return mat.reset_index(drop=True)

###############################################################################
# --------------------------- BALANCE COMPARE ---------------------------------
def balance_comparison(df_prev, df_curr, writer):
    req = {'Main Code', 'Balance'}
    for col in req:
        for d, name in ((df_prev, 'Previous'), (df_curr, 'Current')):
            if col not in d.columns:
                raise ValueError(f"{name} file – missing column '{col}'")

    prev_codes = set(df_prev['Main Code'])
    curr_codes = set(df_curr['Main Code'])

    only_prev = df_prev[df_prev['Main Code'].isin(prev_codes - curr_codes)]
    only_curr = df_curr[df_curr['Main Code'].isin(curr_codes - prev_codes)]
    both = pd.merge(
        df_prev[['Main Code', 'Balance']].rename(columns={'Balance': 'Prev_Bal'}),
        df_curr[['Main Code', 'Branch Name', 'Name', 'Ac Type Desc', 'Balance']].rename(columns={'Balance': 'Curr_Bal'}),
        on='Main Code')
    both['Change'] = both['Curr_Bal'] - both['Prev_Bal']

    reco = pd.DataFrame({
        'Description': ['Opening', 'Settled', 'New', 'Inc/Dec', 'Adjusted', 'Closing'],
        'Amount': [
            df_prev['Balance'].sum(),
            -only_prev['Balance'].sum(),
            only_curr['Balance'].sum(),
            both['Change'].sum(),
            0,
            df_curr['Balance'].sum()],
        'No of Acs': [
            len(prev_codes),
            -len(prev_codes - curr_codes),
            len(curr_codes - prev_codes),
            "", "", len(curr_codes)]
    })
    reco.at[4, 'Amount'] = reco.loc[0:4, 'Amount'].sum()

    only_prev.to_excel(writer, sheet_name='Settled', index=False)
    only_curr.to_excel(writer, sheet_name='New', index=False)
    both[['Main Code', 'Ac Type Desc', 'Branch Name', 'Name',
          'Curr_Bal', 'Prev_Bal', 'Change']].to_excel(writer, sheet_name='Movement', index=False)
    reco.to_excel(writer, sheet_name='Reco', index=False)

###############################################################################
# --------------------------- PIVOT COMPARE  (FIXED) --------------------------
def pivot_compare(df_prev, df_curr, by, writer, sheet_name):
    # summaries
    g1 = (df_prev.groupby(by)
          .agg(Prev_Sum=('Balance', 'sum'), Prev_Cnt=(by, 'count'))
          .rename(columns={by: 'tmp'}))
    g2 = (df_curr.groupby(by)
          .agg(New_Sum=('Balance', 'sum'), New_Cnt=(by, 'count'))
          .rename(columns={by: 'tmp'}))
    merged = g1.join(g2, how='outer').fillna(0)
    merged['Change'] = merged['New_Sum'] - merged['Prev_Sum']
    merged['Pct'] = (merged['Change'] / merged['Prev_Sum'].replace(0, pd.NA) * 100).fillna(0)
    merged['Pct'] = merged['Pct'].map('{:.2f}%'.format)

    # grand-total row
    total = merged.sum(numeric_only=True)
    total.name = 'Total'
    total['Pct'] = '{:.2f}%'.format(
        (total['New_Sum'] - total['Prev_Sum']) / total['Prev_Sum'] * 100
        if total['Prev_Sum'] else 0)
    out = pd.concat([merged, total.to_frame().T]).reset_index().rename(columns={'index': by})

    # write to Excel
    out.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]

    # bold ONLY the last (grand-total) row
    last_row = len(out) + 1          # +1 for header offset
    for col in range(1, len(out.columns) + 1):
        ws.cell(row=last_row, column=col).font = Font(bold=True)

###############################################################################
# ------------------------------ MAIN APP -------------------------------------
def main():
    # Page config
    st.set_page_config(
        page_title="📊 Unified Report Generator",
        page_icon="📊",
        layout="centered",
        initial_sidebar_state="expanded"
    )
    
    # Main title and subtitle
    st.markdown('<h1 class="main-title">📊 Unified Report Generator</h1>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Transform your Excel data into comprehensive analytical reports</p>', unsafe_allow_html=True)
    
    # Create two columns for file upload
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📅 Previous Period")
        prev_upl = st.file_uploader("Upload previous period data", type=['xlsx'], key="prev")
        
    with col2:
        st.markdown("### 📅 Current Period")
        curr_upl = st.file_uploader("Upload current period data", type=['xlsx'], key="curr")

    # Show upload status
    if prev_upl:
        st.success(f"✅ Previous file loaded: {prev_upl.name}")
    if curr_upl:
        st.success(f"✅ Current file loaded: {curr_upl.name}")

    if prev_upl and curr_upl:
        st.markdown("---")
        
        # Generate report button
        if st.button("🚀 Generate Unified Report"):
            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Step 1: Reading files
                status_text.text("📖 Reading Excel files...")
                progress_bar.progress(10)
                
                with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_prev, \
                     tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_curr:
                    tmp_prev_path = tmp_prev.name
                    tmp_curr_path = tmp_curr.name
                    tmp_prev.write(prev_upl.getbuffer())
                    tmp_curr.write(curr_upl.getbuffer())

                progress_bar.progress(25)
                
                # Step 2: Loading data
                status_text.text("🔄 Loading and preprocessing data...")
                df_prev_raw = pd.read_excel(tmp_prev_path)
                df_curr_raw = pd.read_excel(tmp_curr_path)
                progress_bar.progress(40)

                # Step 3: Processing slippage
                status_text.text("📊 Analyzing slippage patterns...")
                df_prev_sl = preprocess_slippage(df_prev_raw)
                df_curr_sl = preprocess_slippage(df_curr_raw)
                slip = detect_slippage(df_prev_sl, df_curr_sl)
                branch_sum = category_matrix(slip, 'Branch Name')
                actype_sum = category_matrix(slip, 'Ac Type Desc')
                progress_bar.progress(60)

                # Step 4: Balance comparison
                status_text.text("💰 Processing balance comparisons...")
                df_prev_cp = preprocess_comp(df_prev_raw)
                df_curr_cp = preprocess_comp(df_curr_raw)
                progress_bar.progress(80)

                # Step 5: Generating report
                status_text.text("📝 Generating final report...")
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as w:
                    slip.to_excel(w, sheet_name='Slippage', index=False)
                    branch_sum.to_excel(w, sheet_name='Summary_Branch', index=False)
                    actype_sum.to_excel(w, sheet_name='Summary_AcType', index=False)

                    balance_comparison(df_prev_cp, df_curr_cp, w)
                    pivot_compare(df_prev_cp, df_curr_cp,
                                  by='Ac Type Desc', writer=w, sheet_name='Compare')
                    pivot_compare(df_prev_cp, df_curr_cp,
                                  by='Branch Name', writer=w, sheet_name='Branch')
                    autofit_excel(w)

                progress_bar.progress(100)
                status_text.text("✅ Report generated successfully!")
                
                out.seek(0)
                
                # Success message with download button
                st.success("🎉 Your unified report is ready for download!")
                st.download_button(
                    label="📥 Download Unified Report",
                    data=out,
                    file_name=f"unified_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                # Clean temp files
                os.unlink(tmp_prev_path)
                os.unlink(tmp_curr_path)

            except Exception as ex:
                st.error("❌ An error occurred while processing your files")
                with st.expander("🔍 View Error Details"):
                    st.code(str(ex))
                    st.text("Full traceback:")
                    st.code(traceback.format_exc())
                    
                # Clean temp files if they exist
                try:
                    os.unlink(tmp_prev_path)
                    os.unlink(tmp_curr_path)
                except:
                    pass
    else:
        st.info("👆 Please upload both previous and current period Excel files to proceed")

if __name__ == "__main__":
    main()
