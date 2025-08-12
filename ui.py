import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from core.SciImports import sudatImport, ivdatImports, ssdatImport
from docx import Document
from docx.shared import Inches
from docx.shared import RGBColor
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
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

def cell_formatting(table):
    # Loop through each row in the table
    for row in table.rows:
        cell = row.cells[1]

        # Add paragraph and apply alignment
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_after = Pt(0)

            # Apply formatting to all runs
            for run in paragraph.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(12)
    
                # Fix font name for Word compatibility
                r = run._element
                r.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')



st.title("SciSun Test Analysis Dashbord")

col1, col2 = st.columns(2)
with col1:
    PN = st.text_input(
    "Project Number",
    "1234",
    key="pn",
)

with col2:
    SN = st.text_input(
    "SciSun Serial Number",
    "0001234",
    key="sn",
)


(projrect.cell(0, 1).paragraphs[0]).paragraph_format.space_after = Pt(0)
run = (projrect.cell(0, 1).paragraphs[0]).add_run(PN)
run.font.name = 'Arial'
run.font.size = Pt(12)

(model.cell(1, 1).paragraphs[0]).paragraph_format.space_after = Pt(0)
run = (model.cell(1, 1).paragraphs[0]).add_run(SN)
run.font.name = 'Arial'
run.font.size = Pt(12)



# --- NU ---
st.header("Non-Uniformity")

uploaded_NU = st.file_uploader("Upload .sudat file", type=['sudat'])
# print(uploaded_NU)

if uploaded_NU is not None:
    result_NU = sudatImport(uploaded_NU)
    st.success("File successfully parsed.")

    # Optional: Call your analysis
    NU_report = NUScript(result_NU)
    st.write("Results:", NU_report)


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
    cell_formatting(nonuni)

    # PLOT
    # search for the placeholder and replace with the image
    for i, para in enumerate(doc.paragraphs):
        if "{{INSERT_PLOT_NU}}" in para.text:
            para.text = para.text.replace("{{INSERT_PLOT_NU}}", "")
            
            # insert image right after the current paragraph
            run = para.insert_paragraph_before().add_run()
            run.add_picture("output/NU.png", width=Inches(8))
            break
else:
    st.error("NU file not uploaded")


# --- TI ---
st.header("Temporal Instability")

uploaded_TI = st.file_uploader("Upload .ivdat file", type=['ivdat'], accept_multiple_files=True)

if uploaded_TI:
    result_TI = ivdatImports(uploaded_TI) 
    st.success("Files successfully parsed.")

    # Display parsed info (optional)
    # for key, value in result_TI.items():
    #     st.subheader(value['filename'])
    #     st.write(value)

    # Optional: Call your analysis
    TI_report = TIScript(result_TI)
    st.write("Results:", TI_report)

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
    cell_formatting(tempinst)

    # PLOT
    # search for the placeholder and replace with the image
    for i, para in enumerate(doc.paragraphs):
        if "{{INSERT_PLOT_TI}}" in para.text:
            para.text = para.text.replace("{{INSERT_PLOT_TI}}", "")
            
            # insert image right after the current paragraph
            run = para.insert_paragraph_before().add_run()
            run.add_picture("output/TI.png", width=Inches(8))
            break
else:
    st.error("TI file not uploaded")


# --- SM ---
st.header("Spectrometry Analysis")


silicons = {
    "Si High Gain": '1',
    "Si Low Gain": '2',
    "No Si": '3'
}

igas = {
    "IGA High Gain": '1',
    "IGA Low Gain": '2',
    "No IGA": '3'
}

standards = {
    "AM1.5D Direct Normal ASTM E927-19": '1',
    "AM1.5G Hemispherical ASTM E927-19": '2',
    "AM0 Extra-Terrestrial ASTM E927-19": '3',
    "AM1.5G IEC 60904-9 Ed.3 Table 1": '4',
    "AM1.5G IEC 60904-9 Ed. 3 Table 2": '5',
    "AM1.5G Hemispherical ASTM E927-19, Limited Range [700 - 1100 nm]": '6',
}


col11, col22 = st.columns(2)
with col11:
    user_si = st.selectbox("Silicon:", list(silicons.keys()))
    status_Si = silicons[user_si]
with col11:
    user_iga = st.selectbox("IGA:", list(igas.keys()))
    status_IGA = igas[user_iga]    
with col22:
    user_am = st.selectbox("The Standard:", list(standards.keys()))
    AMType = standards[user_am]  

with col22:
    crosspoint = st.text_input(
        "Crosspoint:",
        "1100",
        key="cross",
    )
# --- Choice Inputs ---
# silicons = ["high gain", "low gain", "no silicon"]
# igas = ["high gain", "low gain", "no in gas"]
# standards = ["AM1.5G", "AM1.5D", "AM0"]


label = st.text_input(
    "Label the data for the plot:",
    "SciSun300 SN0001234",
    key="label",
)
rawdata = st.checkbox("Save the raw .csv data file", key="disabled")



uploaded_Si = st.file_uploader("Upload .ssdat file Si", type=['ssdat'], key='Si')
uploaded_IGA = st.file_uploader("Upload .ssdat file IGA", type=['ssdat'], key='IGA')

if uploaded_Si or uploaded_IGA:
    result_Si = ssdatImport(uploaded_Si)
    result_IGA = ssdatImport(uploaded_IGA)
    st.success("File successfully parsed.")


    SM_report = SMScript(status_Si, status_IGA, AMType, result_Si, result_IGA, label, rawdata, crosspoint)
    st.write("Results:", SM_report)

    # TABLE
    spectral.cell(1, 1).text = SM_report['Date of Si Measurement']                                   # [YYYY-MM-DD]
    spectral.cell(2, 1).text = '20' # str(SM_report['Scan Time'])                
    spectral.cell(3, 1).text = '1' # str(SM_report['number of spectra avg'])                 
    spectral.cell(4, 1).text = SM_report['Classification']                
    cell_formatting(spectral)

    # PLOT
    # search for the placeholder and replace with the image
    for i, para in enumerate(doc.paragraphs):
        if "{{INSERT_PLOT_SM}}" in para.text:
            para.text = para.text.replace("{{INSERT_PLOT_SM}}", "")
            
            # insert image right after the current paragraph
            run = para.insert_paragraph_before().add_run()
            run.add_picture("output/SM.png", width=Inches(8))
            break
else:
    st.error("SM file not uploaded")





# doc.save("output/report_saved.docx")


if st.button("Generate Report"):

    # save to in-memory file
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)

    st.download_button(
        label="Download Report as .docx",
        data=doc_io,
        file_name=f"TR-SCISUN-01 E927-19 [ProjectNo{PN} SN{SN}].docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


