import streamlit as st
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import io
import os
import tempfile
import traceback
from typing import Dict, List, Tuple, Optional

# Configure Streamlit page
st.set_page_config(
    page_title="📊 Unified Report Generator",
    page_icon="📊",
    layout="centered",  # Changed from "wide" to "centered"
    initial_sidebar_state="expanded"
)

# Hide Streamlit UI components
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Custom styling */
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
    }
    
    .metric-container {
        background: #f8f9fa;
        padding: 0.5rem;
        border-radius: 8px;
        border-left: 4px solid #007acc;
        margin: 1rem 0;
    }
    
    .error-container {
        background: #fff5f5;
        border: 1px solid #fed7d7;
        color: #c53030;
        padding: 0.5rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .success-container {
        background: #f0fff4;
        border: 1px solid #9ae6b4;
        color: #276749;
        padding: 0.5rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

###############################################################################

# Rest of the code remains the same...

def create_summary_statistics(slippage_df: pd.DataFrame) -> Dict:
    """Create summary statistics for display"""
    total_accounts = len(slippage_df)
    
    movement_counts = slippage_df['Movement'].value_counts()
    
    return {
        'total_accounts': total_accounts,
        'slippage_count': movement_counts.get('Slippage', 0),
        'upgrade_count': movement_counts.get('Upgrade', 0),
        'stable_count': movement_counts.get('Stable', 0),
        'slippage_rate': (movement_counts.get('Slippage', 0) / total_accounts * 100) if total_accounts > 0 else 0
    }

def main_app():
    """Main application interface"""
    
    # Show logout option
    AuthManager.show_logout_sidebar()
    
    # Main header
    st.markdown("""
        <div class="main-header">
            <h1>📊 Unified Report Generator</h1>
            <p>Upload Previous and Current Excel files to generate comprehensive analysis reports</p>
        </div>
    """, unsafe_allow_html=True)
    
    # File upload section
    st.subheader("📁 File Upload")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**📅 Previous Period File**")
        prev_file = st.file_uploader(
            "Choose previous period Excel file",
            type=['xlsx', 'xls'],
            key="prev_file"
        )
    
    with col2:
        st.markdown("**📅 Current Period File**")
        curr_file = st.file_uploader(
            "Choose current period Excel file", 
            type=['xlsx', 'xls'],
            key="curr_file"
        )
    
    # Process files if both uploaded
    if prev_file and curr_file:
        
        # Analysis options
        st.subheader("⚙️ Analysis Options")
        
        analysis_options = st.multiselect(
            "Select analysis types to include:",
            options=[
                "Slippage Analysis",
                "Balance Comparison", 
                "Branch Comparison",
                "Account Type Comparison"
            ],
            default=[
                "Slippage Analysis",
                "Balance Comparison",
                "Branch Comparison", 
                "Account Type Comparison"
            ]
        )
        
        # Generate report button
        if st.button("🚀 Generate Comprehensive Report", type="primary", use_container_width=True):
            
            with st.spinner("🔄 Processing files and generating report..."):
                try:
                    # Save uploaded files temporarily
                    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_prev, \
                         tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_curr:
                        
                        tmp_prev.write(prev_file.getbuffer())
                        tmp_curr.write(curr_file.getbuffer())
                        tmp_prev_path = tmp_prev.name
                        tmp_curr_path = tmp_curr.name
                    
                    # Read Excel files
                    df_prev_raw = pd.read_excel(tmp_prev_path)
                    df_curr_raw = pd.read_excel(tmp_curr_path)
                    
                    # Create output buffer
                    output_buffer = io.BytesIO()
                    
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        
                        # Slippage Analysis
                        if "Slippage Analysis" in analysis_options:
                            df_prev_slippage = DataProcessor.preprocess_slippage(df_prev_raw)
                            df_curr_slippage = DataProcessor.preprocess_slippage(df_curr_raw)
                            
                            slippage_data = DataProcessor.detect_slippage(df_prev_slippage, df_curr_slippage)
                            
                            # Write slippage data
                            slippage_data.to_excel(writer, sheet_name='Slippage_Detail', index=False)
                            
                            # Create summaries - FIXED to show previous categories properly
                            branch_matrix = ReportGenerator.create_category_matrix(slippage_data, 'Branch Name')
                            actype_matrix = ReportGenerator.create_category_matrix(slippage_data, 'Ac Type Desc')
                            movement_summary = ReportGenerator.create_movement_summary(slippage_data)
                            
                            # Add previous category column to summaries
                            branch_matrix.insert(1, 'Previous_Category', slippage_data['cat_prev'].unique())
                            actype_matrix.insert(1, 'Previous_Category', slippage_data['cat_prev'].unique())
                            
                            # Write summaries
                            branch_matrix.to_excel(writer, sheet_name='Summary_Branch', index=False)
                            actype_matrix.to_excel(writer, sheet_name='Summary_AcType', index=False)
                            movement_summary.to_excel(writer, sheet_name='Movement_Summary', index=False)
                        
                        # Balance and other comparisons
                        if any(opt in analysis_options for opt in ["Balance Comparison", "Branch Comparison", "Account Type Comparison"]):
                            df_prev_comp = DataProcessor.preprocess_comparison(df_prev_raw)
                            df_curr_comp = DataProcessor.preprocess_comparison(df_curr_raw)
                            
                            if "Balance Comparison" in analysis_options:
                                balance_stats = ReportGenerator.create_balance_comparison(df_prev_comp, df_curr_comp, writer)
                            
                            if "Account Type Comparison" in analysis_options:
                                ReportGenerator.create_pivot_comparison(
                                    df_prev_comp, df_curr_comp, 'Ac Type Desc', writer, 'AcType_Comparison'
                                )
                            
                            if "Branch Comparison" in analysis_options:
                                ReportGenerator.create_pivot_comparison(
                                    df_prev_comp, df_curr_comp, 'Branch Name', writer, 'Branch_Comparison'
                                )
                        
                        # Apply formatting
                        ExcelFormatter.autofit_columns(writer)
                        ExcelFormatter.format_headers(writer)
                    
                    output_buffer.seek(0)
                    
                    # Show success message and summary stats
                    st.success("✅ Report generated successfully!")
                    
                    # Display summary statistics if slippage analysis was performed
                    if "Slippage Analysis" in analysis_options:
                        stats = create_summary_statistics(slippage_data)
                        
                        st.subheader("📈 Summary Statistics")
                        
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Total Accounts", f"{stats['total_accounts']:,}")
                        with col2:
                            st.metric("Slipped Accounts", f"{stats['slippage_count']:,}")
                        with col3:
                            st.metric("Upgraded Accounts", f"{stats['upgrade_count']:,}")
                        with col4:
                            st.metric("Slippage Rate", f"{stats['slippage_rate']:.1f}%")
                    
                    # Download button
                    st.download_button(
                        label="📥 Download Comprehensive Report",
                        data=output_buffer,
                        file_name=f"unified_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    # Clean up temporary files
                    try:
                        os.unlink(tmp_prev_path)
                        os.unlink(tmp_curr_path)
                    except:
                        pass
                            
                except Exception as e:
                    st.error("❌ An error occurred while processing the files")
                    
                    with st.expander("🔍 View Error Details"):
                        st.code(f"Error: {str(e)}")
                        st.code(traceback.format_exc())

def main():
    """Main entry point"""
    # Initialize authentication state
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    
    # Show login or main app
    if not st.session_state["authenticated"]:
        AuthManager.show_login()
    else:
        main_app()

if __name__ == "__main__":
    main()
