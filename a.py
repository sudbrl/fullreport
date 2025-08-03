import streamlit as st
import pandas as pd
from openpyxl.styles import Font
from datetime import datetime
import io, os, tempfile, traceback

# --- Hide Streamlit UI components ---
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

###############################################################################
# -------------------------  LOGIN PAGE  --------------------------------------
def login_page():
    st.markdown("""
        <style>
        .login-container {
            max-width: 280px;
            margin: 60px auto;
            padding: 15px 20px;
            background: #f0f2f6;
            border-radius: 6px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        }
        .login-header {
            font-size: 20px;
            font-weight: 600;
            color: #333;
            margin-bottom: 15px;
            text-align: center;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        </style>
        <div class="login-container">
            <div class="login-header">Please Log In</div>
        </div>
    """, unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
    if submitted:
        if username in st.secrets.get("auth", {}) and password == st.secrets["auth"][username]:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Invalid username or password.")
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
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()
###############################################################################
# ---------------------------  ORIGINAL APP  ----------------------------------
st.set_page_config(page_title="📊 Unified Report Generator", layout="centered")
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)
###############################################################################
# ----------------------------- CONSTANTS -------------------------------------
KEEP_SLIPPAGE = ['Branch Name', 'Main Code', 'Ac Type Desc', 'Name',
                 'Limit', 'Balance', 'Provision']
PROV_MAP = {'G': 1, 'W': 2, 'S': 3, 'D': 4, 'B': 5}
CAT_NAMES = {'G': 'Good', 'W': 'Watchlist', 'S': 'Substandard',
             'D': 'Doubtful', 'B': 'Bad'}
CAT_ORDER = ['Good', 'Substandard', 'Doubtful', 'Bad']
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
            'Limit', 'Balance', 'cat_prev', 'cat_curr', 'Movement']
    return full[cols]
def category_matrix(df, group_col=None):
    index = group_col if group_col else pd.Series(0, index=df.index, name='dummy')
    mat = (df
           .pivot_table(index=index,
                        columns='cat_prev',
                        values='cat_curr',
                        aggfunc='size',
                        fill_value=0)
           .reindex(columns=CAT_ORDER, fill_value=0)
           .astype(int))
    if group_col:
        return mat.reset_index()
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
    st.title("📊 Unified Report Generator")
    st.write("Upload **Previous** and **Current** Excel files to generate one consolidated report.")
    prev_upl = st.file_uploader("📅 Previous period", type=['xlsx'])
    curr_upl = st.file_uploader("📅 Current period",  type=['xlsx'])
    if prev_upl and curr_upl:
        if st.button("Generate Report"):
            with st.spinner("Processing…"):
                try:
                    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_prev, \
                         tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_curr:
                        tmp_prev_path = tmp_prev.name
                        tmp_curr_path = tmp_curr.name
                        tmp_prev.write(prev_upl.getbuffer())
                        tmp_curr.write(curr_upl.getbuffer())
                    df_prev_raw = pd.read_excel(tmp_prev_path)
                    df_curr_raw = pd.read_excel(tmp_curr_path)
                    # Slippage
                    df_prev_sl = preprocess_slippage(df_prev_raw)
                    df_curr_sl = preprocess_slippage(df_curr_raw)
                    slip = detect_slippage(df_prev_sl, df_curr_sl)
                    branch_sum = category_matrix(slip, 'Branch Name')
                    actype_sum = category_matrix(slip, 'Ac Type Desc')
                    
                    # Add cat_prev column as second column in summary sheets
                    if not branch_sum.empty:
                        branch_sum.insert(1, 'cat_prev', slip['cat_prev'].values[:len(branch_sum)])
                    if not actype_sum.empty:
                        actype_sum.insert(1, 'cat_prev', slip['cat_prev'].values[:len(actype_sum)])
                    
                    # Balance
                    df_prev_cp = preprocess_comp(df_prev_raw)
                    df_curr_cp = preprocess_comp(df_curr_raw)
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
                    out.seek(0)
                    st.success("✅ Report ready!")
                    st.download_button(
                        label="📥 Download unified report",
                        data=out,
                        file_name=f"unified_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    # Clean temp files
                    os.unlink(tmp_prev_path)
                    os.unlink(tmp_curr_path)
                except Exception as ex:
                    st.error("❌ Processing failed")
                    with st.expander("Show error"):
                        st.code(traceback.format_exc())
if __name__ == "__main__":
    main()
