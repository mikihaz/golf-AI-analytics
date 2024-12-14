import streamlit as st
import pandas as pd
from document_processor import process_document
from ppt_generator import create_presentation
import tempfile
import os
from config import OPENAI_API_KEY

def main():
    try:
        # Basic page config
        st.set_page_config(page_title="AI Document Analyzer", layout="wide")
        
        # Remove API key from session state
        if 'processed_file' not in st.session_state:
            st.session_state.processed_file = None
        if 'analysis_result' not in st.session_state:
            st.session_state.analysis_result = None

        # Remove API Key input section from sidebar
        st.sidebar.title("⚙️ Settings")

        # Add reference PPT upload section
        st.sidebar.markdown("---")
        st.sidebar.markdown("### Reference Template")
        reference_ppt = st.sidebar.file_uploader(
            "Upload a reference PPTX file",
            type=['pptx'],
            key='reference_ppt'
        )

        if reference_ppt:
            st.session_state.reference_template = reference_ppt
            st.sidebar.success("✅ Reference template loaded")
        
        # Main interface
        st.title("📄 AI Document Analyzer")
        st.markdown("### Upload your CV or Excel file for AI analysis")

        # File uploader with clear instructions
        uploaded_file = st.file_uploader(
            "Supported formats: XLSX, CSV, DOCX, PDF",
            type=['xlsx', 'csv', 'docx', 'pdf'],
            key='file_uploader'
        )
        
        if uploaded_file is not None:
            # Show file info
            st.info(f"File uploaded: {uploaded_file.name}")
            
            if uploaded_file != st.session_state.processed_file:
                st.session_state.processed_file = uploaded_file
                
                # Process file with reference template if available
                with st.spinner('📊 Analyzing document...'):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        file_path = tmp_file.name
                        
                        try:
                            # Save reference template if available
                            reference_path = None
                            if 'reference_template' in st.session_state:
                                with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as ref_file:
                                    ref_file.write(st.session_state.reference_template.getvalue())
                                    reference_path = ref_file.name

                            # Call process_document without API key
                            analysis_result = process_document(file_path, reference_path)
                            st.session_state.analysis_result = analysis_result
                        except Exception as e:
                            st.error(f"Error processing document: {str(e)}")
                        finally:
                            os.unlink(file_path)
                            if reference_path:
                                os.unlink(reference_path)

            # Display results
            if st.session_state.analysis_result:
                st.subheader("🔍 Analysis Results")
                st.write(st.session_state.analysis_result)

                if ("Error" not in st.session_state.analysis_result and 
                    "Unsupported" not in st.session_state.analysis_result):
                    try:
                        # Generate PPT
                        with st.spinner('Generating presentation...'):
                            ppt_path = create_presentation(st.session_state.analysis_result)
                            
                            # Provide download button
                            with open(ppt_path, "rb") as ppt:
                                st.download_button(
                                    label="📥 Download Analysis PPT",
                                    data=ppt,
                                    file_name="analysis_report.pptx",
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                                )
                    except Exception as e:
                        st.error(f"Error generating presentation: {str(e)}")

        # Add footer
        st.markdown("---")
        st.markdown("### Instructions")
        st.markdown("""
        1. Upload your document using the file uploader above
        2. Wait for the AI analysis to complete
        3. Review the analysis results
        4. Download the PowerPoint presentation
        """)

    except Exception as e:
        st.error(f"Application error: {str(e)}")
        st.error("Please refresh the page and try again")

if __name__ == "__main__":
    main()
