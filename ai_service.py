import google.generativeai as genai
from openai import OpenAI

def get_ai_summarizer(model_type, api_key):
    """Returns a function that summarizes text using the selected AI model."""
    
    if model_type == "Gemini":
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-pro')
        
        def summarize(text):
            if not text.strip():
                return []
            prompt = f"""
            You are a presentation expert. 
            Summarize the following text into exactly 3 to 5 concise bullet points.
            Strictly follow these rules:
            - DO NOT add new information.
            - Preserve the original meaning.
            - Format each bullet point as a single string in a list.
            - Use a professional, minimalist tone.
            
            Text:
            {text}
            
            Bullet points:
            """
            try:
                response = model.generate_content(prompt)
                lines = response.text.strip().split('\n')
                # Clean up bullets (remove leading dashes, numbers)
                bullets = [line.lstrip('-*•1234567890. ').strip() for line in lines if line.strip()]
                return bullets[:5]
            except Exception:
                # Fallback to basic splitting if AI fails
                return [line.strip() for line in text.split('\n') if line.strip()][:5]
        
        return summarize

    elif model_type == "OpenAI":
        client = OpenAI(api_key=api_key)
        
        def summarize(text):
            if not text.strip():
                return []
            prompt = f"""
            Summarize the following text into exactly 3 to 5 concise bullet points for a professional presentation.
            DO NOT add new information.
            
            Text:
            {text}
            """
            try:
                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[{"role": "user", "content": prompt}]
                )
                content = response.choices[0].message.content
                lines = content.strip().split('\n')
                bullets = [line.lstrip('-*•1234567890. ').strip() for line in lines if line.strip()]
                return bullets[:5]
            except Exception:
                return [line.strip() for line in text.split('\n') if line.strip()][:5]
                
        return summarize

    return None
