"""
Financial Crime Risk Monitor - Professional Dashboard
Wells Fargo Style: Red & White Theme
"""

import streamlit as st
import pandas as pd
import subprocess
import sys
from pathlib import Path
from datetime import datetime
import time
import os
import re

# Page Configuration
st.set_page_config(
    page_title="Financial Crime Risk Monitor",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS - Wells Fargo Theme (Red & White)
st.markdown("""
    <style>
    /* Main theme colors */
    :root {
        --wells-red: #D71E28;
        --wells-dark-red: #A41E22;
        --wells-light-gray: #F5F5F5;
        --wells-dark-gray: #333333;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Main container */
    .main {
        background-color: white;
    }
    
    /* Header styling */
    .header-container {
        background: linear-gradient(135deg, #D71E28 0%, #A41E22 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    .header-title {
        color: white;
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
        text-align: center;
        letter-spacing: 1px;
    }
    
    .header-subtitle {
        color: rgba(255,255,255,0.9);
        font-size: 1.1rem;
        text-align: center;
        margin-top: 0.5rem;
    }
    
    /* Control panel */
    .control-panel {
        background: white;
        padding: 2rem;
        border-radius: 10px;
        border: 2px solid #D71E28;
        margin-bottom: 2rem;
    }
    
    /* Section titles */
    h3 {
        color: #333333 !important;
        font-weight: 700 !important;
        font-size: 1.2rem !important;
        margin-bottom: 0.5rem !important;
    }
    
    /* All text visibility */
    p, span, div, label {
        color: #333333 !important;
    }
    
    /* Selectbox - Make options visible */
    div[data-baseweb="select"] {
        background-color: white !important;
    }
    
    div[data-baseweb="select"] > div {
        background-color: white !important;
        color: #333333 !important;
    }
    
    div[data-baseweb="select"] [role="option"] {
        background-color: white !important;
        color: #333333 !important;
        padding: 0.5rem !important;
    }
    
    div[data-baseweb="select"] [role="option"]:hover {
        background-color: #FFF5F5 !important;
        color: #D71E28 !important;
    }
    
    div[data-baseweb="select"] [role="listbox"] {
        background-color: white !important;
        max-height: 300px !important;
    }
    
    /* Force all select-related elements to white background */
    [data-baseweb="popover"] {
        background-color: white !important;
    }
    
    ul[role="listbox"] {
        background-color: white !important;
    }
    
    li[role="option"] {
        background-color: white !important;
        color: #333333 !important;
    }
    
    li[role="option"]:hover {
        background-color: #FFF5F5 !important;
        color: #D71E28 !important;
    }
    
    /* Status box */
    .status-box {
        padding: 1.5rem;
        border-radius: 8px;
        margin: 1rem 0;
        border-left: 5px solid #D71E28;
        background: white;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    .status-title {
        font-weight: 700;
        font-size: 1.2rem;
        color: #333333;
        margin-bottom: 0.5rem;
    }
    
    .status-text {
        color: #666666;
        font-size: 1rem;
    }
    
    /* Selectbox styling */
    .stSelectbox > div > div {
        background-color: white;
        border: 2px solid #D71E28;
        border-radius: 8px;
        font-weight: 600;
        color: #333333;
    }
    
    .stSelectbox label {
        font-weight: 700;
        color: #333333;
        font-size: 1.1rem;
    }
    
    /* Progress indicator */
    .progress-step {
        background: white;
        padding: 1rem;
        margin: 0.5rem 0;
        border-radius: 5px;
        border-left: 4px solid #D71E28;
    }
    
    .progress-step.active {
        background: #FFF5F5;
        border-left-color: #D71E28;
    }
    
    .progress-step.completed {
        background: #F0FFF0;
        border-left-color: #28a745;
    }
    
    /* Results table */
    .dataframe {
        border: 2px solid #D71E28 !important;
        border-radius: 8px;
        overflow: hidden;
    }
    
    .dataframe thead tr th {
        background-color: #D71E28 !important;
        color: white !important;
        font-weight: 700 !important;
        padding: 12px !important;
        border: none !important;
    }
    
    .dataframe tbody tr:nth-child(even) {
        background-color: #FFF5F5;
    }
    
    .dataframe tbody tr:hover {
        background-color: #FFE5E5;
    }
    
    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #D71E28 0%, #A41E22 100%);
        color: white;
        font-weight: 700;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        border: none;
        font-size: 1.1rem;
        cursor: pointer;
        transition: all 0.3s;
        width: 100%;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #A41E22 0%, #8B1A1E 100%);
        box-shadow: 0 4px 8px rgba(215,30,40,0.3);
        transform: translateY(-2px);
    }
    
    .stButton > button:active {
        transform: translateY(0);
    }
    
    /* Download button */
    .stDownloadButton > button {
        background: #333333;
        color: white;
        font-weight: 600;
        padding: 0.6rem 1.5rem;
        border-radius: 6px;
        border: 2px solid #D71E28;
        transition: all 0.3s;
    }
    
    .stDownloadButton > button:hover {
        background: #D71E28;
        border-color: #A41E22;
    }
    
    /* Metrics */
    .metric-container {
        background: white;
        padding: 1.5rem;
        border-radius: 8px;
        border: 2px solid #D71E28;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    .metric-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: #D71E28;
    }
    
    .metric-label {
        font-size: 1rem;
        color: #666666;
        margin-top: 0.5rem;
    }
    
    /* Info boxes */
    .info-box {
        background: white;
        border: 2px solid #D71E28;
        border-radius: 8px;
        padding: 1.5rem;
        margin: 1rem 0;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    .info-box strong {
        color: #333333 !important;
    }
    
    .info-box small {
        color: #666666 !important;
    }
    
    /* Spinner */
    .stSpinner > div {
        border-top-color: #D71E28 !important;
    }
    
    /* Success message */
    .success-banner {
        background: linear-gradient(135deg, #28a745 0%, #20893a 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 8px;
        text-align: center;
        font-size: 1.2rem;
        font-weight: 700;
        margin: 1rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    /* Error message */
    .error-banner {
        background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 8px;
        text-align: center;
        font-size: 1.2rem;
        font-weight: 700;
        margin: 1rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    /* Log container */
    .log-container {
        background: #1e1e1e !important;
        padding: 1rem;
        border-radius: 5px;
        font-family: 'Courier New', monospace;
        max-height: 400px;
        overflow-y: auto;
        margin: 1rem 0;
        border: 1px solid #333;
    }
    
    .log-container pre {
        color: #ffffff !important;
        margin: 0;
        white-space: pre-wrap;
        word-wrap: break-word;
        font-size: 0.9rem;
        background: #1e1e1e !important;
    }
    
    .log-container * {
        color: #ffffff !important;
        background: #1e1e1e !important;
    }
    </style>
""", unsafe_allow_html=True)

# Base directory
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / 'output'

def get_month_year_options():
    """Generate month/year options from current date backwards"""
    current_date = datetime.now()
    options = []
    
    # Generate 24 months backwards from current month
    for i in range(24):
        month = current_date.month - i
        year = current_date.year
        
        while month <= 0:
            month += 12
            year -= 1
        
        month_name = datetime(year, month, 1).strftime('%B')
        options.append({
            'display': f"{month_name} {year}",
            'month': month,
            'year': year
        })
    
    return options

def run_web_scraper_realtime(year, month, log_placeholder):
    """Run web scraper with real-time log streaming"""
    try:
        # Set environment to handle encoding properly
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        env['PYTHONUNBUFFERED'] = '1'
        
        process = subprocess.Popen(
            [sys.executable, 'web_scraper.py', '--year', str(year), '--month', str(month)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            cwd=BASE_DIR,
            bufsize=1,
            universal_newlines=True,
            encoding='utf-8',
            errors='replace',  # Replace bad Unicode characters
            env=env
        )
        
        log_lines = []
        
        while True:
            line = process.stdout.readline()
            if line == '' and process.poll() is not None:
                break
            if line:
                # Clean up the line - remove ANSI codes and weird Unicode
                clean_line = re.sub(r'\x1b\[[0-9;]*m', '', line.rstrip())  # Remove ANSI colors
                clean_line = re.sub(r'[^\x20-\x7E\n]', '', clean_line)  # Keep only ASCII printable
                log_lines.append(clean_line)
                # Show last 50 lines in real-time
                display_logs = '\n'.join(log_lines[-50:])
                log_placeholder.markdown(f"""
                    <div style="background: #1e1e1e; padding: 1rem; border-radius: 5px; font-family: 'Courier New', monospace; max-height: 400px; overflow-y: auto; border: 1px solid #333;">
                        <pre style="color: #ffffff; margin: 0; white-space: pre-wrap; word-wrap: break-word; font-size: 0.9rem;">{display_logs}</pre>
                    </div>
                """, unsafe_allow_html=True)
        
        return process.returncode == 0, '\n'.join(log_lines), ""
    except Exception as e:
        return False, "", str(e)

def run_extraction_realtime(year, month, log_placeholder):
    """Run extraction with real-time log streaming"""
    try:
        # Set environment to handle encoding properly
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        env['PYTHONUNBUFFERED'] = '1'
        
        process = subprocess.Popen(
            [sys.executable, 'extract_red_flags.py', '--year', str(year), '--month', str(month)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            cwd=BASE_DIR,
            bufsize=1,
            universal_newlines=True,
            encoding='utf-8',
            errors='replace',  # Replace bad Unicode characters
            env=env
        )
        
        log_lines = []
        
        while True:
            line = process.stdout.readline()
            if line == '' and process.poll() is not None:
                break
            if line:
                # Clean up the line - remove ANSI codes and weird Unicode
                clean_line = re.sub(r'\x1b\[[0-9;]*m', '', line.rstrip())  # Remove ANSI colors
                clean_line = re.sub(r'[^\x20-\x7E\n]', '', clean_line)  # Keep only ASCII printable
                log_lines.append(clean_line)
                # Show last 50 lines in real-time
                display_logs = '\n'.join(log_lines[-50:])
                log_placeholder.markdown(f"""
                    <div style="background: #1e1e1e; padding: 1rem; border-radius: 5px; font-family: 'Courier New', monospace; max-height: 400px; overflow-y: auto; border: 1px solid #333;">
                        <pre style="color: #ffffff; margin: 0; white-space: pre-wrap; word-wrap: break-word; font-size: 0.9rem;">{display_logs}</pre>
                    </div>
                """, unsafe_allow_html=True)
        
        return process.returncode == 0, '\n'.join(log_lines), ""
    except Exception as e:
        return False, "", str(e)

def load_results(year, month):
    """Load extraction results from Excel file"""
    try:
        filename = f"red_flags_analysis_enhanced_{year}_{month:02d}.xlsx"
        filepath = OUTPUT_DIR / filename
        
        if filepath.exists():
            df = pd.read_excel(filepath)
            return df, filepath
        else:
            return None, None
    except Exception as e:
        st.error(f"Error loading results: {str(e)}")
        return None, None

def main():
    # Header
    st.markdown("""
        <div class="header-container">
            <h1 class="header-title">FINANCIAL CRIME RISK MONITOR</h1>
            <p class="header-subtitle">AI-Powered Red Flag Detection & Analysis System</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Control Panel
    st.markdown('<div class="control-panel">', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([2, 2, 1])
    
    with col1:
        st.markdown('<h3>Select Time Period</h3>', unsafe_allow_html=True)
        month_year_options = get_month_year_options()
        
        # Create display list
        display_options = [opt['display'] for opt in month_year_options]
        
        selected_display = st.selectbox(
            "Month & Year",
            options=display_options,
            index=0,
            label_visibility="collapsed"
        )
        
        # Find the selected option
        selected_idx = display_options.index(selected_display)
        selected_month = month_year_options[selected_idx]['month']
        selected_year = month_year_options[selected_idx]['year']
    
    with col2:
        st.markdown('<div style="margin-top: 2.5rem;"></div>', unsafe_allow_html=True)
        start_button = st.button("START ANALYSIS", use_container_width=True)
    
    with col3:
        st.markdown('<h3>Status</h3>', unsafe_allow_html=True)
        if 'processing' in st.session_state and st.session_state.processing:
            st.markdown('<div style="background: #FFF5F5; padding: 0.5rem; border-radius: 5px; text-align: center; border: 2px solid #D71E28;"><span style="color: #D71E28; font-weight: 700; font-size: 1.1rem;">RUNNING</span></div>', unsafe_allow_html=True)
        elif 'completed' in st.session_state and st.session_state.completed:
            st.markdown('<div style="background: #F0FFF0; padding: 0.5rem; border-radius: 5px; text-align: center; border: 2px solid #28a745;"><span style="color: #28a745; font-weight: 700; font-size: 1.1rem;">COMPLETED</span></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div style="background: white; padding: 0.5rem; border-radius: 5px; text-align: center; border: 2px solid #666666;"><span style="color: #666666; font-weight: 700; font-size: 1.1rem;">READY</span></div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Initialize session state
    if 'processing' not in st.session_state:
        st.session_state.processing = False
    if 'completed' not in st.session_state:
        st.session_state.completed = False
    if 'results_df' not in st.session_state:
        st.session_state.results_df = None
    if 'results_file' not in st.session_state:
        st.session_state.results_file = None
    
    # Start Analysis Process
    if start_button:
        st.session_state.processing = True
        st.session_state.completed = False
        st.session_state.results_df = None
        
        progress_container = st.container()
        
        with progress_container:
            st.markdown("---")
            st.markdown('<h2 style="color: #D71E28;">Processing Pipeline</h2>', unsafe_allow_html=True)
            
            # Step 1: Web Scraping
            st.markdown("""
                <div class="progress-step active">
                    <strong style="color: #D71E28; font-size: 1.1rem;">STEP 1/2: Web Scraping</strong><br>
                    <span style="color: #333333;">Collecting data from 9 regulatory sources</span>
                </div>
            """, unsafe_allow_html=True)
            
            # Real-time log display
            scraper_log = st.empty()
            success, stdout, stderr = run_web_scraper_realtime(selected_year, selected_month, scraper_log)
            
            if success:
                st.markdown("""
                    <div class="progress-step completed">
                        <strong style="color: #28a745; font-size: 1.1rem;">✓ STEP 1/2: Web Scraping - COMPLETED</strong>
                    </div>
                """, unsafe_allow_html=True)
                
                # Collapsible complete logs
                with st.expander("📋 View Complete Web Scraping Logs", expanded=False):
                    st.markdown(f"""
                        <div class="log-container">
                            <pre>{stdout if stdout else "No output"}</pre>
                        </div>
                    """, unsafe_allow_html=True)
            else:
                st.error("❌ Web scraping failed")
                with st.expander("🔍 View Error Logs", expanded=True):
                    st.markdown(f"""
                        <div class="log-container">
                            <pre style="color: #ff6b6b;">{stderr if stderr else stdout}</pre>
                        </div>
                    """, unsafe_allow_html=True)
                st.session_state.processing = False
                return
            
            # Step 2: Red Flag Extraction
            st.markdown("""
                <div class="progress-step active">
                    <strong style="color: #D71E28; font-size: 1.1rem;">STEP 2/2: AI Analysis</strong><br>
                    <span style="color: #333333;">Extracting red flags with multi-model ensemble</span>
                </div>
            """, unsafe_allow_html=True)
            
            # Real-time log display
            extraction_log = st.empty()
            success, stdout, stderr = run_extraction_realtime(selected_year, selected_month, extraction_log)
            
            if success:
                st.markdown("""
                    <div class="progress-step completed">
                        <strong style="color: #28a745; font-size: 1.1rem;">✓ STEP 2/2: AI Analysis - COMPLETED</strong>
                    </div>
                """, unsafe_allow_html=True)
                
                # Collapsible complete logs
                with st.expander("📋 View Complete Extraction Logs", expanded=False):
                    st.markdown(f"""
                        <div class="log-container">
                            <pre>{stdout if stdout else "No output"}</pre>
                        </div>
                    """, unsafe_allow_html=True)
                
                # Load results
                df, filepath = load_results(selected_year, selected_month)
                if df is not None:
                    st.session_state.results_df = df
                    st.session_state.results_file = filepath
                    st.session_state.completed = True
                    
                    st.markdown("""
                        <div class="success-banner">
                            ✓ ANALYSIS COMPLETED SUCCESSFULLY
                        </div>
                    """, unsafe_allow_html=True)
                else:
                    st.error("Results file not found. Please check output directory.")
            else:
                st.error("❌ Red flag extraction failed")
                with st.expander("🔍 View Error Logs", expanded=True):
                    st.markdown(f"""
                        <div class="log-container">
                            <pre style="color: #ff6b6b;">{stderr if stderr else stdout}</pre>
                        </div>
                    """, unsafe_allow_html=True)
                st.session_state.processing = False
                return
        
        st.session_state.processing = False
    
    # Display Results
    if st.session_state.completed and st.session_state.results_df is not None:
        st.markdown("---")
        st.markdown('<h2 style="color: #D71E28; font-weight: 700;">Analysis Results</h2>', unsafe_allow_html=True)
        
        df = st.session_state.results_df
        
        # Summary Metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-value">{len(df)}</div>
                    <div class="metric-label">Total Red Flags</div>
                </div>
            """, unsafe_allow_html=True)
        
        with col2:
            avg_confidence = df['Confidence Score'].mean() if 'Confidence Score' in df.columns else 0
            st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-value">{avg_confidence:.1f}</div>
                    <div class="metric-label">Avg Confidence</div>
                </div>
            """, unsafe_allow_html=True)
        
        with col3:
            high_conf = len(df[df['Confidence Score'] >= 70]) if 'Confidence Score' in df.columns else 0
            st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-value">{high_conf}</div>
                    <div class="metric-label">High Confidence</div>
                </div>
            """, unsafe_allow_html=True)
        
        with col4:
            not_covered = len(df[df['Coverage Status'] == 'Not Covered']) if 'Coverage Status' in df.columns else 0
            st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-value">{not_covered}</div>
                    <div class="metric-label">New Patterns</div>
                </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Display 7 columns as requested
        display_columns = [
            'Source Link',
            'Date',
            'Extracted Red Flag',
            'Associated Paragraph',
            'Category',
            'Coverage Status',
            'Confidence Score'
        ]
        
        # Filter to only show available columns
        available_columns = [col for col in display_columns if col in df.columns]
        display_df = df[available_columns].copy()
        
        # Format for display
        if 'Confidence Score' in display_df.columns:
            display_df['Confidence Score'] = display_df['Confidence Score'].round(1)
        
        if 'Date' in display_df.columns:
            display_df['Date'] = pd.to_datetime(display_df['Date']).dt.strftime('%Y-%m-%d')
        
        # Truncate long text for display
        for col in ['Extracted Red Flag', 'Associated Paragraph']:
            if col in display_df.columns:
                display_df[col] = display_df[col].apply(
                    lambda x: (x[:100] + '...') if isinstance(x, str) and len(x) > 100 else x
                )
        
        # Download Section
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.session_state.results_file:
                with open(st.session_state.results_file, 'rb') as f:
                    st.download_button(
                        label="📥 DOWNLOAD FULL REPORT (Excel)",
                        data=f,
                        file_name=st.session_state.results_file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Data Table
        st.markdown('<h3 style="color: #333333; font-weight: 700;">Extracted Red Flags</h3>', unsafe_allow_html=True)
        st.dataframe(
            display_df,
            use_container_width=True,
            height=600,
            hide_index=True
        )
        
        # Coverage Breakdown
        st.markdown('<h3 style="color: #333333; font-weight: 700;">Coverage Analysis</h3>', unsafe_allow_html=True)
        if 'Coverage Status' in df.columns:
            coverage_counts = df['Coverage Status'].value_counts()
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                fully = coverage_counts.get('Fully Covered', 0)
                st.markdown(f"""
                    <div class="info-box">
                        <strong>Fully Covered:</strong> 
                        <span style="color: #D71E28; font-size: 1.3rem; font-weight: 700;">{fully}</span><br>
                        <small>Matches existing red flag patterns</small>
                    </div>
                """, unsafe_allow_html=True)
            
            with col2:
                partially = coverage_counts.get('Partially Covered', 0)
                st.markdown(f"""
                    <div class="info-box">
                        <strong>Partially Covered:</strong> 
                        <span style="color: #D71E28; font-size: 1.3rem; font-weight: 700;">{partially}</span><br>
                        <small>Similar to existing patterns</small>
                    </div>
                """, unsafe_allow_html=True)
            
            with col3:
                not_cov = coverage_counts.get('Not Covered', 0)
                st.markdown(f"""
                    <div class="info-box">
                        <strong>Not Covered:</strong> 
                        <span style="color: #D71E28; font-size: 1.3rem; font-weight: 700;">{not_cov}</span><br>
                        <small>New red flag patterns detected</small>
                    </div>
                """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
