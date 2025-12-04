cd /d "%~dp0"
call repgen\Scripts\activate.bat
streamlit cache clear
streamlit run ui.py
pause