import streamlit as st
import io

st.set_page_config(page_title="Minimalist PPT Redesigner", layout="wide")

# Custom CSS for modern look
st.markdown("""
    <style>
    .main {
        background-color: #ffffff;
    }
    .stButton>button {
        background-color: #FF6B00;
        color: white;
        border-radius: 5px;
        border: none;
        padding: 10px 24px;
    }
    .stButton>button:hover {
        background-color: #e56000;
        color: white;
    }
    h1 {
        color: #FF6B00;
        font-family: 'Arial', sans-serif;
    }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.title("PPT Redesign Tool")
    st.subheader("Transform your presentation into a clean, minimalist version.")

    with st.sidebar:
        st.header("Settings")
        use_ai = st.checkbox("Enable AI Summarization (Optional)", value=False)
        api_key = ""
        if use_ai:
            model_choice = st.selectbox("AI Model", ["Gemini", "OpenAI"])
            api_key = st.text_input(f"{model_choice} API Key", type="password")
            st.info("AI will be used to summarize text into 3-5 concise bullet points.")

    uploaded_file = st.file_uploader("Choose a PowerPoint file", type=["pptx"])

    if uploaded_file is not None:
        st.success("File uploaded successfully!")
        
        if st.button("Redesign Presentation"):
            import importlib
            import pptx_processor
            importlib.reload(pptx_processor)
            from pptx_processor import PPTXParser, InitialPlacement, PPTXRenderer, validate_layout
            
            # --- 1. PARSE ---
            with st.spinner("Layer 1: Parsing original content..."):
                try:
                    parser = PPTXParser()
                    raw_model = parser.extract(uploaded_file)
                except Exception as e:
                    st.error(f"Parsing Error: {e}")
                    return

            # --- 2. INITIAL PLACEMENT & LAYOUT ---
            with st.spinner("Layer 2: Setting initial placement & resolving overlaps..."):
                # AI Summarization Logic
                ai_summarizer = None
                if use_ai and api_key:
                    from ai_service import get_ai_summarizer
                    ai_summarizer = get_ai_summarizer(model_choice, api_key)
                
                try:
                    engine = InitialPlacement()
                    processed_model = engine.apply(raw_model, ai_summarizer=ai_summarizer)
                except Exception as e:
                    st.error(f"Placement Error: {e}")
                    return

            # --- 3. VALIDATE ---
            with st.spinner("Safety Check: Validating layout..."):
                is_valid, msg = validate_layout(processed_model)
                if not is_valid:
                    st.error(f"Validation Failed: {msg}")
                    return
                st.info("Layout validated successfully (No overlaps or overflows).")

            # --- 4. RENDER ---
            with st.spinner("Layer 3: Rendering final PowerPoint..."):
                try:
                    renderer = PPTXRenderer()
                    output_pptx = renderer.render(processed_model)
                    
                    st.balloons()
                    st.success("Redesign complete!")
                    
                    st.download_button(
                        label="Download Redesigned PPTX",
                        data=output_pptx,
                        file_name="redesigned_presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                except Exception as e:
                    st.error(f"Rendering Error: {e}")

if __name__ == "__main__":
    main()
