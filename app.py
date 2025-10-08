"""
DSS TEMPLATE AUTO-FILL - Streamlit Web App
Telecommunications Expert System for Dynamic Spectrum Sharing
"""

import streamlit as st
import pandas as pd
import re
from io import StringIO, BytesIO
import sys

# Page config
st.set_page_config(
    page_title="QUADGEN DSS Template Auto-Fill",
    page_icon="📡",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .log-box {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 5px;
        font-family: monospace;
        font-size: 0.9rem;
        max-height: 500px;
        overflow-y: auto;
        color: #000000;
    }
    </style>
""", unsafe_allow_html=True)

# Import helper functions
from utils import find_column, safe_get_value, safe_load_sheet, process_template

# Header
st.markdown('<div class="main-header">📡 QUADGEN dss Template Auto-Fill System</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Telecommunications Expert System for Dynamic Spectrum Sharing Activation</div>', unsafe_allow_html=True)

st.markdown("---")

# File upload section
st.markdown("### 📁 Upload Files")
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("Upload Excel File (.xlsx)", type=['xlsx', 'xls'], key="excel")
    if excel_file:
        st.success(f"✓ {excel_file.name}")

with col2:
    template_file = st.file_uploader("Upload Template File (.txt)", type=['txt'], key="template")
    if template_file:
        st.success(f"✓ {template_file.name}")

st.markdown("---")

# Process button
if excel_file and template_file:
    if st.button("🚀 Process Template", type="primary", use_container_width=True):
        # Create log container
        log_container = st.empty()
        logs = []
        
        def log(message):
            """Add message to logs"""
            logs.append(message)
            log_container.markdown(
                f'<div class="log-box">{"<br>".join(logs)}</div>',
                unsafe_allow_html=True
            )
        
        try:
            # Start processing
            log("="*80)
            log("🔄 STARTING PROCESSING...")
            log("="*80)
            
            # Save uploaded files temporarily
            excel_path = f"temp_{excel_file.name}"
            template_path = f"temp_{template_file.name}"
            
            with open(excel_path, "wb") as f:
                f.write(excel_file.getbuffer())
            
            with open(template_path, "wb") as f:
                f.write(template_file.getbuffer())
            
            log("\n✓ Files uploaded successfully")
            
            # Process the template
            filled_content, replacements, warnings = process_template(
                excel_path, 
                template_path, 
                log_callback=log
            )
            
            # Summary
            log("\n" + "="*80)
            log("✅ PROCESSING COMPLETED!")
            log("="*80)
            log(f"\n📊 Summary:")
            log(f"  • Placeholders replaced: {len(replacements)}")
            log(f"  • Total replacements: {sum(replacements.values())}")
            log(f"  • Warnings: {len(warnings)}")
            
            if warnings:
                log(f"\n⚠️  Warnings:")
                for warning in warnings[:5]:
                    log(f"  • {warning}")
                if len(warnings) > 5:
                    log(f"  • ... and {len(warnings)-5} more warnings")
            
            # Store in session state
            st.session_state['filled_content'] = filled_content
            st.session_state['output_filename'] = template_file.name.replace('.txt', '_FILLED.txt')
            st.session_state['processed'] = True
            
            # Clean up temp files
            import os
            try:
                os.remove(excel_path)
                os.remove(template_path)
            except:
                pass
                
        except Exception as e:
            log(f"\n❌ ERROR: {str(e)}")
            log("\nProcessing failed. Please check your files and try again.")
            st.session_state['processed'] = False

else:
    st.info("👆 Please upload both Excel and Template files to proceed")

# Download button (only show if processed)
if st.session_state.get('processed', False):
    st.markdown("---")
    st.markdown("### 💾 Download Result")
    
    filled_content = st.session_state.get('filled_content', '')
    output_filename = st.session_state.get('output_filename', 'output_FILLED.txt')
    
    st.download_button(
        label="📥 Download Filled Template",
        data=filled_content,
        file_name=output_filename,
        mime="text/plain",
        type="primary",
        use_container_width=True
    )

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; font-size: 0.9rem;">
    <p>DSS Template Auto-Fill System v1.0 | Telecom Automation Expert</p>
    <p>Handles missing worksheets, columns, and data gracefully</p>
</div>
""", unsafe_allow_html=True)
