@echo off

cd /d %~dp0

echo Installing requirements...
pip install -r requirements.txt

echo Starting Streamlit app...
streamlit run app.py

pause