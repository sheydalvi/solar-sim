*The core scripts were initially developed by Kyle Graham*

================ ONE CLICK LAUNCH (LOCAL) ==================
For running locally (recommended)

One time setup:
	1. Create a local copy of Report_Generation folder on your PC
	2. Click on SETUP.bat to install
Everytime after the setup:
	3. Click on RUNME.bat to launch 

====================== USING ANACONDA ======================
No setup needed

1. Open Anaconda Navigator
2. Launch VS Code
3. Change directory to Report Generation
4. Open a bash terminal
5. Run the following command
	streamlit run ui.py
6. Should open up a tab in your default browser


If needed, install modules with the following commands
	conda install conda-forge::python-docx
	conda install conda-forge::streamlit
	etc.
