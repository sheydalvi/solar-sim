# -*- coding: utf-8 -*-

import os
from io import StringIO
from tkinter import filedialog
from .Conversions import conv2Irrad
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
import io


def ivdatImports(uploaded_files):
    """
    Parses uploaded .ivdat files (Streamlit file-like objects) and returns a dictionary
    with structured IV data suitable for use in SolarSimScripts.
    """
    theDictionary = {}

    for fileIndex, uploaded_file in enumerate(uploaded_files):
        # Decode and read the uploaded file
        content = uploaded_file.read().decode('utf-8')
        lines = content.splitlines()

        # --- Find header start ---
        for hLineNo, hLine in enumerate(lines):
            if hLine.strip().startswith('Voltage'):
                headerRows = hLineNo - 1
                break

        # --- Find footer start and length ---
        for fLineNo, fLine in enumerate(lines):
            if fLine.strip().startswith('END DATA'):
                footerStartLine = fLineNo
            if fLine.strip().startswith('END FOOTER'):
                footerRows = fLineNo - footerStartLine + 1
                break

        # --- Parse main data block ---
        data_io = io.StringIO(content)
        fileData = pd.read_csv(data_io, sep='\t', header=headerRows, skipfooter=footerRows, engine='python')

        # Clean column names
        fileData.columns = [col.strip() for col in fileData.columns]

        # --- Extract key columns as Series (not DataFrames) ---
        try:
            voltage_col = [col for col in fileData.columns if 'voltage' in col.lower()][0]
            current_col = [col for col in fileData.columns if 'current' in col.lower()][0]
            power_col = [col for col in fileData.columns if 'power' in col.lower()][0]

            voltage = fileData[voltage_col]
            current = fileData[current_col]
            power = fileData[power_col]
        except IndexError:
            raise KeyError(f"Missing expected column(s) in file: {uploaded_file.name}. Found columns: {fileData.columns.tolist()}")

        numberOfData = len(voltage)

        # --- Parse header info ---
        header_io = io.StringIO(content)
        headerData = pd.read_csv(header_io, sep='\t', header=0,
                                 skipfooter=numberOfData + footerRows + 4,
                                 names=('Parameters', 'Values'), engine='python')

        # --- Parse footer info ---
        footer_io = io.StringIO(content)
        footerData = pd.read_csv(footer_io, sep='\t',
                                 header=headerRows + numberOfData + 2,
                                 skipfooter=1, names=('Parameters', 'Values'), engine='python')

        # --- Structure file-specific dictionary ---
        tempD = {
            'voltage': voltage,
            'current': current,
            'power': power,
            'filename': headerData['Values'][0],
            'date': headerData['Values'][1],
            'detArea': headerData['Values'][6],
            'numPoints': headerData['Values'][10],
            'spp': headerData['Values'][11],
            'wait': headerData['Values'][12],
            'dwell': headerData['Values'][13],
            'Voc': footerData['Values'][0],
            'Isc': footerData['Values'][1],
            'maxP': footerData['Values'][2],
            'FF': footerData['Values'][3]
        }

        theDictionary[f'File {fileIndex}'] = tempD

    return theDictionary

def ssdatImport(uploaded_file):
    if uploaded_file is None:
        print('No file was uploaded.')
        return None

    # Decode entire file and split into lines
    file_lines = uploaded_file.getvalue().decode('utf-8').splitlines()

    # Find header row for data
    for hLineNo, hLine in enumerate(file_lines):
        # print(hLine)
        if hLine.startswith('Wavelength'):
            headerRows = hLineNo - 1
            print(headerRows)
            break
    else:
        raise ValueError("Could not find 'Wavelength' header in file.")

    # Read data section
    file_str = '\n'.join(file_lines)
    file_io = StringIO(file_str)
    fileData = pd.read_csv(file_io, sep='\t', header=headerRows, skipfooter=1, engine='python')

    numberOfData = len(fileData['Wavelength(nm)'])

    # Rewind to read header section
    file_io.seek(0)
    headerData = pd.read_csv(file_io, sep='\t', header=0,
                             skipfooter=numberOfData + 4,
                             names=('Parameters', 'Values'),
                             engine='python')

    # Build the dictionary to return
    theDictionary = {
        'wavelengths': fileData['Wavelength(nm)'],
        'signal': fileData['Signal(V)'],
        'dIndices': fileData['Detector Index(#)'],
        'gIndices': fileData['Grating Index(#)'],
        'fIndices': fileData['Filter Index(#)'],
        'filepath': uploaded_file.name,   # Just the filename, no full path
        'filename': headerData['Values'][0],
        'date': headerData['Values'][1],
        'monoModel': headerData['Values'][5],
        'startWave': headerData['Values'][7],
        'stopWave': headerData['Values'][8],
        'stepSize': headerData['Values'][9],
        'folder': ''  # Not available in Streamlit context
    }

    print(theDictionary['filename'] + ' has been successfully imported.')
    return theDictionary



def sudatImport(uploaded_file):
    if uploaded_file is None:
        print('No file was uploaded.')
        return None

    # read all lines to memory for header/footer scanning
    file_lines = uploaded_file.getvalue().decode('utf-8').splitlines()

    # find header
    for hLineNo, hLine in enumerate(file_lines):
        if hLine[:10] == 'X Position':
            headerRows = hLineNo - 1
            break

    # find footer
    for fLineNo, fLine in enumerate(file_lines):
        if fLine[:8] == 'END DATA':
            footerStartLine = fLineNo
        if fLine[:10] == 'END FOOTER':
            footerRows = fLineNo - footerStartLine + 1
            break

    # recreate the file-like object for pandas since we've consumed it
    uploaded_file.seek(0)

    suData = pd.read_csv(uploaded_file, sep='\t', header=headerRows, skipfooter=footerRows, engine='python')

    numData = len(suData['X Position(cm)'])

    uploaded_file.seek(0)
    suHeader = pd.read_csv(uploaded_file, sep='\t', header=0,
                           skipfooter=numData + footerRows + 1,
                           names=('Parameters', 'Values'),
                           engine='python')

    uploaded_file.seek(0)
    suFooter = pd.read_csv(uploaded_file, sep='\t', header=headerRows + numData + 2,
                           skipfooter=1, names=('Parameters', 'Values'),
                           engine='python')

    measurementGeometry = suHeader['Values'][12]

    if measurementGeometry == 'Rectangular':
        theDictionary = {
            'Xs': suData['X Position(cm)'],
            'Ys': suData['Y Position(cm)'],
            'Zs': suData['Z Position(cm)'],
            'signal': suData['Front panel [Raw](A)'],
            'calSignal': suData['Front panel [Cal.](A)'],
            'filename': uploaded_file.name,
            'date': suHeader['Values'][11],
            'geometry': suHeader['Values'][12],
            'xSize': suHeader['Values'][14],
            'ySize': suHeader['Values'][15],
            'zSize': suHeader['Values'][16],
            'xSpacing': suHeader['Values'][17],
            'ySpacing': suHeader['Values'][18],
            'zSpacing': suHeader['Values'][19],
            'xNum': suHeader['Values'][20],
            'yNum': suHeader['Values'][21],
            'zNum': suHeader['Values'][22],
            'detArea': suHeader['Values'][23],
            'NU': suFooter['Values'][0],
            'minSignal': suFooter['Values'][1],
            'maxSignal': suFooter['Values'][2]
        }
    elif measurementGeometry == 'Circular':
        theDictionary = {
            'Xs': suData['X Position(cm)'],
            'Ys': suData['Y Position(cm)'],
            'Zs': suData['Z Position(cm)'],
            'signal': suData['Front panel [Raw](A)'],
            'calSignal': suData['Front panel [Cal.](A)'],
            'filename': uploaded_file.name,
            'date': suHeader['Values'][11],
            'geometry': suHeader['Values'][12],
            'diameter': suHeader['Values'][14],
            'zSize': suHeader['Values'][15],
            'ppp': suHeader['Values'][16],
            'zPlanes': suHeader['Values'][17],
            'detArea': suHeader['Values'][18],
            'NU': suFooter['Values'][0],
            'minSignal': suFooter['Values'][1],
            'maxSignal': suFooter['Values'][2]
        }
    else:
        print('Invalid geometry format in file.')
        return None

    print(theDictionary['filename'] + ' has been successfully imported.')
    return theDictionary


    # Ask the user to select the desired .tsmdat file and store its address.
    if guideString == None:
        samplePath = filedialog.askopenfilename(initialdir = r'\\10.0.0.11\Projects', filetypes = [('Transmittance Sample Scan Data', '*.tsmdat')], title = 'Select desired .tsmdat file.')
    elif isinstance(guideString, str):
        samplePath = filedialog.askopenfilename(initialdir = r'\\10.0.0.11\Projects', filetypes = [('Transmittance Sample Scan Data', '*.tsmdat')], title = guideString)
    else:
        print('Invalid optional argument. trnsImport accepts either 0 arguments, or a single, optional argument which must be a string.')
    
    # Check how many lines are in the sample scan's header so we can read in the header and data separately.
    with open(samplePath, 'r') as headerCheck:
        for hLineNo, hLine in enumerate(headerCheck):
            if hLine[:10] == 'Wavelength':
                sampleHeaderRows = hLineNo - 1
                break
    
    # Check how many lines are in the sample scan's footer so we can read in the data and footer separately.
    with open(samplePath, 'r') as footerCheck:
        for fLineNo, fLine in enumerate(footerCheck):
            if fLine[:8] == 'END DATA':
                footerStartLine = fLineNo
                sampleFooterRows = 2
            if fLine[:10] == 'END FOOTER':
                sampleFooterRows = fLineNo - footerStartLine + 1
                break
    
    # Read in the sample scan's header, footer, and data separately.
    sampleData = pd.read_csv(samplePath, sep = '\t', header = sampleHeaderRows, skipfooter = sampleFooterRows, engine = 'python')
    sampleNumData = len(sampleData['Wavelength(nm)'])
    sampleHeader = pd.read_csv(samplePath, sep = '\t', header = 0, skipfooter = sampleNumData + sampleFooterRows + 4, names = ('Parameters', 'Values'), engine = 'python')
    sampleFooter = pd.read_csv(samplePath, sep = '\t', header = sampleHeaderRows + sampleNumData + 2, skipfooter = 1, names = ('Parameters', 'Values'), engine = 'python')
    
    # Extract the original address of the associated reference scan and check if it is still located there. If not, ask the user where it is.
    if os.path.exists(sampleHeader['Values'][6]):
        refPath = sampleHeader['Values'][6]
    else:
        refPath = filedialog.askopenfilename(filetypes = [('Transmittance Reference Scan Data', '*.tclrdat')], title = 'Select desired .tclrdat file.')
    
    # Check how many lines are in the reference scan's header so we can read in the header and data separately.
    with open(refPath, 'r') as headerCheck:
        for hLineNo, hLine in enumerate(headerCheck):
            if hLine[:10] == 'Wavelength':
                refHeaderRows = hLineNo - 1
                break
    
    # Read in the reference scan's header, and data separately. There is no footer.
    refData = pd.read_csv(refPath, sep = '\t', header = refHeaderRows, skipfooter = 1, engine = 'python')
    refNumData = len(refData['Wavelength(nm)'])
    refHeader = pd.read_csv(refPath, sep = '\t', header = 0, skipfooter = refNumData + 4, names = ('Parameters', 'Values'), engine = 'python')
    
    # Store all of the important data to be included in scan-specific dictionaries.
    sampleDictionary = {
        'wavelengths': sampleData['Wavelength(nm)'],
        'signal': sampleData['Raw_Signal(V)'],
        'lightSignal': sampleData['Light Signal(V)'],
        'refSignal': sampleData['Reference Signal(V)'],
        'transmittance': sampleData['Transmittance(%)'],
        'gIndices': sampleData['GratingIndex(#)'],
        'fIndices': sampleData['FilterIndex(#)'],
        'lIndices': sampleData['SourceIndex(#)'],
        'outPercent': sampleData['LampOutput(%)'],
        'outCurrent': sampleData['LampCurrent(A)'],
        'outVoltage': sampleData['LampVoltage(V)'],
        'outPower': sampleData['LampPower(W)'],
        'itIndices': sampleData['IntegrationTime_Index(#)'],
        'sIndices': sampleData['Sensitivity_Index(#)'],
        'filename': sampleHeader['Values'][0],
        'date': sampleHeader['Values'][1],
        'refFile': sampleHeader['Values'][6],
        'startWave': sampleHeader['Values'][8],
        'stopWave': sampleHeader['Values'][9],
        'stepSize': sampleHeader['Values'][10],
        'delay': sampleHeader['Values'][11],
        'range': sampleHeader['Values'][12],
        'intTime': sampleHeader['Values'][13],
        'chopFreq': sampleHeader['Values'][14],
        'maxTrns': sampleFooter['Values'][0],
        'maxTrnsWave': sampleFooter['Values'][1],
        'minTrns': min(sampleData['Transmittance(%)']),
        'minTrnsWave': sampleData['Wavelength(nm)'][np.where(sampleData['Transmittance(%)'] == min(sampleData['Transmittance(%)']))[0][0]]
        }
    referenceDictionary = {
        'wavelengths': refData['Wavelength(nm)'],
        'signal': refData['Wavelength(nm)'],
        'lightSignal': refData['Light Signal(V)'],
        'gIndices': refData['GratingIndex(#)'],
        'fIndices': refData['FilterIndex(#)'],
        'lIndices': refData['SourceIndex(#)'],
        'outPercent': refData['LampOutput(%)'],
        'outCurrent': refData['LampCurrent(A)'],
        'outVoltage': refData['LampVoltage(V)'],
        'outPower': refData['LampPower(W)'],
        'itIndices': refData['IntegrationTime_Index(#)'],
        'sIndices': refData['Sensitivity_Index(#)'],
        'filename': refHeader['Values'][0],
        'date': refHeader['Values'][1],
        'startWave': refHeader['Values'][7],
        'stopWave': refHeader['Values'][8],
        'stepSize': refHeader['Values'][9],
        'delay': refHeader['Values'][10],
        'range': refHeader['Values'][11],
        'intTime': refHeader['Values'][12],
        'chopFreq': refHeader['Values'][13]
        }
    
    # Store each of the file dictionaries into a master dictionary to be returned.
    theDictionary = {
        'referenceScan': referenceDictionary,
        'sampleScan': sampleDictionary
        }
    
    print(sampleDictionary['filename'] + ' has been successfully imported.')
    print(referenceDictionary['filename'] + ' has been successfully imported.')
    
    return theDictionary