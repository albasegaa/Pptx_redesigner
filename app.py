import streamlit as st
import io
import importlib

from merge_pptx import merge_presentations

st.set_page_config(
    page_title="Minimalist PPT Redesigner",
    layout="wide"
)

# ---------------- CSS ----------------
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


# ---------------- APP ----------------
def main():

    st.title("PPT Redesign Tool")
    st.subheader("Merge multiple presentations and transform them into a clean, minimalist version.")

    # -------- SIDEBAR --------
    with st.sidebar:

        st.header("Settings")

        use_ai = st.checkbox("Enable AI Summarization (Optional)", value=False)

        api_key = ""
        model_choice = None

        if use_ai:
            model_choice = st.selectbox(
                "AI Model",
                ["Gemini", "OpenAI"]
            )

            api_key = st.text_input(
                f"{model_choice} API Key",
                type="password"
            )

            st.info("AI will summarize text into concise bullet points.")

    # -------- FILE UPLOAD --------
    uploaded_files = st.file_uploader(
        "Upload PowerPoint files",
        type=["pptx"],
        accept_multiple_files=True
    )

    # -------- MAIN FLOW --------
    if uploaded_files:

        st.success(f"{len(uploaded_files)} file(s) uploaded successfully!")

        if st.button("Merge & Redesign Presentation"):

            import pptx_processor
            importlib.reload(pptx_processor)

            from pptx_processor import (
                PPTXParser,
                InitialPlacement,
                PPTXRenderer,
                validate_layout
            )

            # ---------------- 1. MERGE ----------------
            with st.spinner("Step 1/4: Merging PowerPoints..."):

                try:
                    merged_path = merge_presentations(uploaded_files)

                    with open(merged_path, "rb") as f:
                        merged_bytes = io.BytesIO(f.read())

                except Exception as e:
                    st.error(f"Merge Error: {e}")
                    return

            # ---------------- 2. PARSE ----------------
            with st.spinner("Step 2/4: Parsing content..."):

                try:
                    parser = PPTXParser()
                    raw_model = parser.extract(merged_bytes)

                except Exception as e:
                    st.error(f"Parsing Error: {e}")
                    return

            # ---------------- 3. LAYOUT ----------------
            with st.spinner("Step 3/4: Optimizing layout..."):

                ai_summarizer = None

                if use_ai and api_key:
                    from ai_service import get_ai_summarizer

                    ai_summarizer = get_ai_summarizer(
                        model_choice,
                        api_key
                    )

                try:
                    engine = InitialPlacement()

                    processed_model = engine.apply(
                        raw_model,
                        ai_summarizer=ai_summarizer
                    )

                except Exception as e:
                    st.error(f"Placement Error: {e}")
                    return

            # ---------------- 4. VALIDATION ----------------
            with st.spinner("Validating layout..."):

                is_valid, msg = validate_layout(processed_model)

                if not is_valid:
                    st.error(f"Validation Failed: {msg}")
                    return

                st.info("Layout validated successfully.")

            # ---------------- 5. RENDER ----------------
            with st.spinner("Step 4/4: Rendering PowerPoint..."):

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