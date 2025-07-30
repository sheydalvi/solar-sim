import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from core.SciImports import sudatImport, ivdatImports, ssdatImport
from docx import Document
from docx.shared import Inches
from docx.shared import RGBColor
from core.SolarSimScripts import *
import io

doc = Document("input/template.docx")

# access a specific table by index
projrect = doc.tables[0]  
model = doc.tables[1]  
performer = doc.tables[2]  
extras = doc.tables[3]  
equip = doc.tables[4]  
specs = doc.tables[5]
nonuni = doc.tables[6]      # NU
spectral = doc.tables[7]    # SM  
tempinst = doc.tables[8]    # TI  

st.title("Data Analysis Dashboard")

# --- NU ---
st.header("Upload Your Data Files")

uploaded_NU = st.file_uploader("Upload .sudat file", type=['sudat'])
# print(uploaded_NU)

if uploaded_NU is not None:
    result_NU = sudatImport(uploaded_NU)
    st.write(result_NU)
else:
    st.write("upload the file")


NU_report = NUScript(result_NU)
# TABLE
nonuni.cell(1, 1).text = NU_report['Date of Measurement']           # [YYYY-MM-DD]
nonuni.cell(2, 1).text = '5 × 5' # NU_report['Measurement Grid Size [cm]']            # Detector Area:
nonuni.cell(3, 1).text = str(NU_report['Number of Measurement Points'])            # Number of Measurement Points:
nonuni.cell(4, 1).text = f"{float(NU_report['Ideal Detector Area [cm^2]']):.3f}"             # Measurement Point Area:
nonuni.cell(5, 1).text = f"{float(NU_report['Maximum Irradiance [Suns]']):.3f}"            # Maximum Irradiance:
nonuni.cell(6, 1).text = f"{float(NU_report['Minimum Irradiance [Suns]']):.3f}"            # Minimum Irradiance:
nonuni.cell(7, 1).text = f"{float(NU_report['Standard Deviation [Suns]']):.3f}"            # Sample Standard Deviation of Spatial Non-Uniformity:
nonuni.cell(8, 1).text = f"{float(NU_report['Spatial Non-Uniformity of Irradiance [%]']):.3f}"           # Spatial Non-Uniformity of Irradiance:
# Classification:
if NU_report['Spatial Non-Uniformity of Irradiance [%]'] <= 2.0:
    nonuni.cell(9, 1).text = "A"         # Classification:
elif NU_report['Spatial Non-Uniformity of Irradiance [%]'] <= 5.0:
    nonuni.cell(9, 1).text = "B"         
else:
    nonuni.cell(9, 1).text = "C"         


# PLOT
# search for the placeholder and replace with the image
for i, para in enumerate(doc.paragraphs):
    if "{{INSERT_PLOT_NU}}" in para.text:
        para.text = para.text.replace("{{INSERT_PLOT_NU}}", "")
        
        # insert image right after the current paragraph
        run = para.insert_paragraph_before().add_run()
        run.add_picture("output/NU.png", width=Inches(10))
        break


# --- TI ---
st.header("Upload Your Data Files")



uploaded_TI = st.file_uploader("Upload .ivdat file", type=['ivdat'], accept_multiple_files=True)

if uploaded_TI:
    result_TI = ivdatImports(uploaded_TI) 
    st.success("Files successfully parsed.")

    # Display parsed info (optional)
    for key, value in result_TI.items():
        st.subheader(value['filename'])
        st.write(value)

    # Optional: Call your analysis
    TI_report = TIScript(result_TI)
    st.write("Simulation Results:", TI_report)


# TABLE
tempinst.cell(1, 1).text = TI_report['Date of Measurement']                                   # [YYYY-MM-DD]
tempinst.cell(2, 1).text = str(TI_report['Detector Area'])                    # Detector Area:
tempinst.cell(3, 1).text = str(TI_report['Time Between Data Points [s]'])                     # Number of Measurement Points:
tempinst.cell(4, 1).text = str(TI_report['Number of Power Line Cycles'])             # Measurement Point Area:
tempinst.cell(5, 1).text = str(TI_report['Total Measurement Points'])          # Maximum Irradiance:
tempinst.cell(6, 1).text = f"{float(TI_report['Minimum Irradiance [Suns]']):.3f}"            # Maximum Irradiance:
tempinst.cell(7, 1).text = f"{float(TI_report['Maximum Irradiance [Suns]']):.3f}"            # Minimum Irradiance:
tempinst.cell(8, 1).text = f"{float(TI_report['Temporal Instability [%]']):.3f}"           # Spatial Non-Uniformity of Irradiance:
# Classification:
if TI_report['Temporal Instability [%]'] <= 4:
    tempinst.cell(9, 1).text = "A"         # Classification:
elif TI_report['Temporal Instability [%]'] <= 10:
    tempinst.cell(9, 1).text = "B"         
else:
    tempinst.cell(9, 1).text = "C"         

# PLOT
# search for the placeholder and replace with the image
for i, para in enumerate(doc.paragraphs):
    if "{{INSERT_PLOT_TI}}" in para.text:
        para.text = para.text.replace("{{INSERT_PLOT_TI}}", "")
        
        # insert image right after the current paragraph
        run = para.insert_paragraph_before().add_run()
        run.add_picture("output/TI.png", width=Inches(19))
        break


# --- SM ---

# uploaded_SM = st.file_uploader("Upload .ssdat file", type=['ssdat'])

# if uploaded_SM is not None:
#     result_SM = sudatImport(uploaded_SM)
#     st.write(result_SM)
# else:
#     st.write("upload the file")

# SM_report = SMScript(result_SM)


# --- Choice Inputs ---
silicons = ["high gain", "low gain", "no silicon"]
igas = ["high gain", "low gain", "no in gas"]
standards = ["AM1.5G", "AM1.5D", "AM0"]

st.header("Spectrometry Analysis")

silicon = st.selectbox("First Choice", silicons)
iga = st.selectbox("First Choice", igas)
standard = st.selectbox("First Choice", standards)


# --- Graphs ---

# --- Report Gen ---

# doc.save("output/report_saved.docx")


if st.button("Generate Report"):

    # save to in-memory file
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)

    st.download_button(
        label="Download Report as .docx",
        data=doc_io,
        file_name="report_downloaded.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
