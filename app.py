import streamlit as st
from backend.vector_engine import create_vector_store
from backend.llm_engine import extract_keywords, generate_tailored_resume_json
from backend.docx_engine import extract_text_from_docx, update_docx_and_export

st.set_page_config(page_title="AI Resume Tailor", layout="wide")
st.markdown("<style>.stApp {background-color: #FDFBF7;}</style>", unsafe_allow_html=True)

st.title("✨ AI Resume Tailor (Pro Word Edition)")

col1, col2 = st.columns(2)
with col1:
    resume_file = st.file_uploader("Upload Resume (Word .docx ONLY)", type=["docx"])
with col2:
    jd_text = st.text_area("Paste Job Description", height=200)

if st.button("🚀 Process & Generate ATS Resume", use_container_width=True):
    if resume_file and jd_text:
        
        # 1. INGESTION 
        with st.spinner("Reading Word Document..."):
            file_bytes = resume_file.read()
            original_text, doc_object = extract_text_from_docx(file_bytes)
            
        # 2. KEYWORDS 
        with st.spinner("Targeting JD Keywords..."):
            top_keywords = extract_keywords(jd_text)
            st.info(f"🎯 **Target ATS Keywords Found:** {top_keywords}")
            
        # 3. RETRIEVAL 
        with st.spinner("Finding matching experience..."):
            vector_store = create_vector_store(original_text)
            relevant_docs = vector_store.similarity_search(jd_text, k=5) 
            
        # 4. GENERATION 
        with st.spinner("AI is rewriting your bullets (preserving exact formatting)..."):
            changes_dict = generate_tailored_resume_json(jd_text, relevant_docs, original_text)
            
        if not changes_dict:
            st.error("⚠️ AI could not generate valid replacements. Please try again.")
        else:
            st.success("✅ Resume Successfully Tailored!")
            
            st.subheader("🔍 What Changed?")
            for old, new in changes_dict.items():
                with st.expander(f"Modified: {old[:50]}..."):
                    st.write(f"❌ **Old:** {old}")
                    st.write(f"✅ **New:** {new}")
            
            # 5. EXPORT & DOWNLOAD 
            with st.spinner("Packaging your new Word file..."):
                final_docx_bytes = update_docx_and_export(doc_object, changes_dict, verbose=False)
            
            st.divider()
            st.download_button(
                label="📝 Download Tailored Resume (DOCX)",
                data=final_docx_bytes,
                file_name="ATS_Optimized_Resume.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
    else:
        st.warning("Please upload a resume and paste a JD first!")