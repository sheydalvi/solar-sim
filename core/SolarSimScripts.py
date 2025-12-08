# -*- coding: utf-8 -*-
# Import Statements
import os
import sys
import statistics
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.cm as cm
import matplotlib.colors as mcolors
from . import SciImports as ScImp
from scipy import interpolate
from tabulate import tabulate
from .Conversions import conv2Irrad


def NUScript(nuData):
    '''
    This function takes a single .sudat file as input, calculates its spatial non-uniformity, then outputs the results along with a plot for the test report.
    It can be used for rectangular or circular data and will output raw and normalized plots for a given dataset.
    Note that the normalization assumes that the target illumination level is 1 Sun.
    '''
    # Import the NU scan data
    # nuData = ScImp.sudatImport('Select the desired .sudat file.')
    
    # Determine the target area geometry.
    geo = nuData['geometry']
    
    if geo == 'Rectangular':
        # Order the signal data into an array
        blockData = []
        for y in range(int(nuData['yNum'])):
            dataRow = []
            for x in range(int(nuData['xNum'])):
                dataRow.append(nuData['signal'][y * int(nuData['xNum']) + x] * 1000)    # x1000 to convert to mA
            blockData.append(dataRow)
        
        # Generate a plot with normalized values
        blockData = np.flip(blockData, 0)   # reverse the data so it displays in the correct order
        normArray = blockData / max(blockData.min(), blockData.max(), key = abs)        # normalize against the maximum detected current
        normArray = normArray * ((normArray.max() - normArray.min()) / 2) + normArray   # shift the dataset up such that it is centered around 1 Sun
        normPlot = plt.imshow(normArray, interpolation = 'none', cmap = 'summer', vmin = normArray.min(), vmax = normArray.max(), extent = [0, abs(nuData['Xs'].min()) + abs(nuData['Xs'].max()) + float(nuData['xSpacing']), 0, abs(nuData['Ys'].min()) + abs(nuData['Ys'].max()) + float(nuData['ySpacing'])])
        cbar = plt.colorbar(normPlot)
        cbar.set_label('Irradiance [Suns]')
        plt.title('Non-Uniformity Plot - Normalized')
        plt.xlabel('X Position [cm]')
        plt.ylabel('Y Position [cm]')
        plt.savefig("output/NU.png")
        plt.clf() 

        
        report_d = {
            'Filename': nuData['filename'],
            'Date of Measurement': nuData['date'],
            'Measurement Grid Size [cm]': str(nuData['xSize']) + ' x ' + str(nuData['ySize']),
            'Number of Measurement Points': len(nuData['Xs']),
            'Ideal Detector Area [cm^2]': nuData['detArea'],
            'Maximum Irradiance [Suns]': normArray.max(),
            'Minimum Irradiance [Suns]': normArray.min(),
            'Standard Deviation [Suns]': np.std(normArray),
            'Spatial Non-Uniformity of Irradiance [%]': nuData['NU']
            }
        resultsFrame = pd.DataFrame.from_dict(report_d, orient = 'index')
        # print(tabulate(resultsFrame, colalign = ('right',)))
    
    elif geo == 'Circular':
        detDiam = 2 * np.sqrt(float(nuData['detArea']) / np.pi)
        
        # Normalize the current values
        normCurrents = nuData['signal'] / max(nuData['signal'].min(), nuData['signal'].max(), key = abs)    # normalize against the maximum detected current
        normCurrents = normCurrents * ((normCurrents.max() - normCurrents.min()) / 2) + normCurrents        # shift the dataset up such that it is centered around 1 Sun.
        

        
        # Associate the normalized current values with a colourmap.
        cmapNorm = cm.summer
        normColourNorm = mcolors.Normalize(vmin = normCurrents.min(), vmax = normCurrents.max())
        
        # Generate a plot with the normalized detector values.
        fig2, ax2 = plt.subplots()
        for xCoord, yCoord, ptColour in zip(nuData['Xs'], nuData['Ys'], normCurrents):
            colour = cmapNorm(normColourNorm(ptColour))
            circle = mpatches.Circle((xCoord, yCoord), radius = detDiam / 2, color = colour, alpha = 0.8)
            ax2.add_patch(circle)
        
        # Formatting the normalized plot.
        ax2.set_xlim(nuData['Xs'].min() - detDiam, nuData['Xs'].max() + detDiam)
        ax2.set_ylim(nuData['Ys'].min() - detDiam, nuData['Ys'].max() + detDiam)
        ax2.set_aspect('equal')
        ax2.set_title('Non-Uniformity Plot: Normalized')
        ax2.set_xlabel('X Position [cm]')
        ax2.set_ylabel('Y Position [cm]')
        
        # Generate a ScalarMappable and colourbar
        smNorm = cm.ScalarMappable(cmap = cmapNorm, norm = normColourNorm)
        smNorm.set_array([])
        cbarNorm = plt.colorbar(smNorm, ax = ax2)
        cbarNorm.set_label('Irradiance [Suns]')
        plt.savefig("output/NUbar.png")
        plt.clf() 


        
        report_d = {
            'Filename': nuData['filename'],
            'Date of Measurement': nuData['date'],
            'Target Area Diameter [cm]': statistics.mean([np.max(nuData['Xs']) - np.min(nuData['Xs']), np.max(nuData['Ys']) - np.min(nuData['Ys'])]),
            'Ideal Detector Area [cm^2]': nuData['detArea'],
            'Maximum Irradiance [Suns]': normCurrents.max(),
            'Minimum Irradiance [Suns]': normCurrents.min(),
            'Standard Deviation [Suns]': np.std(normCurrents),
            'Spatial Non-Uniformity of Irradiance [%]': nuData['NU']
            }
        resultsFrame = pd.DataFrame.from_dict(report_d, orient = 'index')
        # print(tabulate(resultsFrame, colalign = ('right', )))
    
    return report_d

def TIScript(tiData):
    '''
    TIScript: This function allows the user to import any number of .ivdat files, finds the dataset with the worst temporal instability, and outputs the results along with a plot for the test report.
    '''
    # Hardcoded values
    NPLC = 1                # number of power line cycles
    time_btw_points = 0.19  # time in seconds

    # Find the dataset with the worst temporal instability
    for setIndex, dataset in enumerate(tiData):
        # print(dataset)
        # We require a minimum of 20 data points. Reject any datasets with insufficient data.
        if len(tiData[dataset]['current']) < 20:
            print('The following datafile contains insufficient (<20) data points to properly calculate the temporal instability and will be ignored: ' + dataset['filepath'])
            continue
        # Calculate the TI
        TI = 100 * abs((tiData[dataset]['current'].max() - tiData[dataset]['current'].min()) / (tiData[dataset]['current'].max() + tiData[dataset]['current'].min()))
        # We only care about the dataset with the worst TI
        if setIndex == 0:
            worstTI = TI
            worstIndex = setIndex
        elif TI > worstTI:
            worstTI = TI
            worstIndex = setIndex
    
    worstData = tiData['File ' + str(worstIndex)]
    

    # Calculate the STI of the worst dataset
    for currentIndex, current in enumerate(worstData['current']):
        if currentIndex == 0:
            continue
        sti = 100 * abs((current - worstData['current'][currentIndex - 1]) / (current + worstData['current'][currentIndex - 1]))
        if currentIndex == 1:
            STI = sti
        elif sti > STI:
            STI = sti
    
    # Assume that the average current corresponds to 1 Sun (this should be very close to true if the measurement was performed correctly)
    sun = statistics.mean(worstData['current'])
    
    # Calculate the min/max irradiance in Suns based on this assumption.
    maxIrrad = worstData['current'].max() / sun
    minIrrad = worstData['current'].min() / sun
    
    # Generate a plot for the report.
    timeValues = np.linspace(0, (len(worstData['current']) - 1) * time_btw_points, len(worstData['current']))
    plt.yticks(np.arange(0.95, 1.05, step = 0.01))
    plt.ylim(0.95, 1.05)
    hLineHandle = plt.axhline(1.0, ls = '--', color = 'k')
    scatterHandle = plt.scatter(timeValues, worstData['current'] / sun, color = 'r')
    plt.legend([hLineHandle, scatterHandle], ['1 Sun Line', 'Temporal Instability Data Points'])
    plt.xlabel('Time [s]')
    plt.ylabel('Irradiance [Suns]')
    plt.savefig("output/TI.png")
    plt.clf() 



    # Display information for the test report formatted as a convenient table
    report_d = {
        'Filename': worstData['filename'],
        'Date of Measurement': worstData['date'],
        'Detector Area': worstData['detArea'],
        'Time Between Data Points [s]': time_btw_points,
        'Number of Power Line Cycles': NPLC,
        'Total Measurement Points': len(worstData['current']),
        'Maximum Irradiance [Suns]': maxIrrad,
        'Minimum Irradiance [Suns]': minIrrad,
        'Short Term Instability [%]': STI,
        'Temporal Instability [%]': worstTI
        }
    # resultsFrame = pd.DataFrame.from_dict(report_d, orient = 'index')
    # print(tabulate(resultsFrame, colalign = ('right',)))
    return report_d


def SMScript(status_Si, status_IGA, AMType, SiData, IGAData, label, rawdata, crosspoint):
    '''
    specScript: This function take a single .ssdat file from the spectroradiometer as input, calculates its degree of matching to a given solar spectrum, then outputs the results along with a plot for the test report.
    '''

    dir = os.path.abspath(os.path.dirname(__file__))
    parent_dir = os.path.dirname(dir)
    files = os.path.join(parent_dir, 'required_files')

    # Hardcoded values (here for easy access/updating)
    # Transfer function file locations. Currently given relative to the script's run location.
    SiHi = os.path.join(files, 'Transfer-Si-HI.csv')
    SiLo = os.path.join(files, 'Transfer-Si-LO.csv')
    IGAHi = os.path.join(files, 'Transfer-IGA-HI.csv')
    IGALo = os.path.join(files, 'Transfer-IGA-LO.csv')
    
    # Range over which each transfer function is defined.
    SiHiRange = '250 - 1100 nm'
    SiLoRange = '250 - 1100 nm'
    IGAHiRange = '900 - 1749 nm'
    IGALoRange = '1001 - 1749 nm'
    
    # Standard spectra file location. Currently given relative to the script's run location.
    stdSpecPath = os.path.join(files, 'Solar_Standards.csv')
    
    # The main function starts here.
    ssData = {}
    
    # Ask the user if there is any data taken with the Si detector. If so, import it.
    if status_Si == '1' or status_Si == '2':
        # Import the Si detector data.
        ssData['siData'] = SiData
        
        # Use the correct transfer function based on the gain selection.
        if status_Si == '1':
            tFunc_Si = pd.read_csv(SiHi, names = ('tWaves', 'tdata'))
            print('\nThe HI gain transfer function for this Si detector is defined across the range of ' + SiHiRange + '.')
        elif status_Si == '2':
            tFunc_Si = pd.read_csv(SiLo, names = ('tWaves', 'tdata'))
            print('\nThe LO gain transfer function for this Si detector is defined across the range of ' + SiLoRange + '.')
        
        # Use the transfer function to convert the Si voltage data to spectral irradiance.
        irrad_Si = conv2Irrad(ssData['siData']['wavelengths'], ssData['siData']['signal'], tFunc_Si)
    
    # Ask the user if there is any data taken with the InGaAs detector. If so, import it.
    if status_IGA == '1' or status_IGA == '2':
        # Import the InGaAs detector data.
        ssData['igaData'] = IGAData
        
        # Use the correct transfer function based on the gain selection.
        if status_IGA == '1':
            tFunc_IGA = pd.read_csv(IGAHi, names = ('tWaves', 'tdata'))
            print('\nThe HI gain transfer function for this InGaAs detector is defined across the range of ' + IGAHiRange + '.')
        elif status_IGA == '2':
            tFunc_IGA = pd.read_csv(IGALo, names = ('tWaves', 'tdata'))
            print('\nThe LO gain transfer function for this InGaAs detector is defined across the range of ' + IGALoRange + '.')
        
        # Use the transfer function to convert the InGaAs voltage data to spectral irradiance.
        irrad_IGA = conv2Irrad(ssData['igaData']['wavelengths'], ssData['igaData']['signal'], tFunc_IGA)
    
    # If there are both Si and InGaAs measurements, combine them into a single dataset.
    if 'siData' in ssData and 'igaData' in ssData:
        plt.xticks(np.arange(int(ssData['siData']['startWave']), int(ssData['igaData']['stopWave']), 100), fontsize = 6)
        plt.plot(irrad_Si['irradWaves'], irrad_Si['irrad'], label = 'Si Data')
        plt.plot(irrad_IGA['irradWaves'], irrad_IGA['irrad'], label = 'InGaAs Data')
        plt.xlabel('Wavlength [nm]')
        plt.ylabel('Spectral Irradiance [W/m^2/nm]')
                
        # trim the Si data
        trimmedIndices_Si = np.where(irrad_Si['irradWaves'] <= crosspoint)[0]
        trimmedWaves_Si = irrad_Si['irradWaves'][trimmedIndices_Si]
        trimmedIrrad_Si = irrad_Si['irrad'][trimmedIndices_Si]
        
        # trim the InGaAs data.
        trimmedIndices_IGA = np.where(irrad_IGA['irradWaves'] > crosspoint)[0]
        trimmedWaves_IGA = irrad_IGA['irradWaves'][trimmedIndices_IGA]
        trimmedIrrad_IGA = irrad_IGA['irrad'][trimmedIndices_IGA]
        
        # combined wavelength and irradiance data
        waves = np.concatenate((trimmedWaves_Si, trimmedWaves_IGA))
        irrad = np.concatenate((trimmedIrrad_Si, trimmedIrrad_IGA))
        
        # Location to save results to. Defaults to the location of the Si data file when both types are present.
        savePath = ssData['siData']['folder']
    
    elif 'siData' in ssData:
        # Just Si data
        waves = irrad_Si['irradWaves']
        irrad = irrad_Si['irrad']
        savePath = ssData['siData']['folder']
    elif 'igaData' in ssData:
        # Just InGaAs data
        waves = irrad_IGA['irradWaves']
        irrad = irrad_IGA['irradWaves']
        savePath = ssData['igaData']['folder']
    
    # Import the reference spectra.
    specRef = pd.read_csv(
        stdSpecPath,
        sep=',',
        header=0,
        engine='python',
        names=['Wavelengths', 'AM0 Irrad', 'AM1.5G Irrad', 'AM1.5D Irrad', 'AM0 Rad', 'AM1.5G Rad', 'AM1.5D Rad'],
    )

    
    if AMType == '1':
        refIrrad = specRef['AM1.5D Irrad'] / 10000
    elif AMType == '2':
        refIrrad = specRef['AM1.5G Irrad'] / 10000
    elif AMType == '3':
        refIrrad = specRef['AM0 Irrad'] / 10000
    elif AMType == '4':
        refIrrad = specRef['AM1.5G Irrad'] / 10000
    elif AMType == '5':
        refIrrad = specRef['AM1.5G Irrad'] / 10000
    elif AMType == '6':
        refIrrad = specRef['AM1.5G Irrad'] / 10000
    else:
        pass
    
    # Interpolate the collected data so it has the same resolution as the reference spectrum (1 nm steps)
    interpWaves = np.round(np.arange(round(min(waves), 1), round(max(waves), 1) + 0.1, 0.1), 3)
    
    interper = interpolate.interp1d(waves, irrad)
    interpIrrad = interper(interpWaves)
    
    # Load the ASTM wavelength bins and intensity percentage, depending on the reference spectrum.
    if AMType == '1':       # ASTM AM1.5D
        print('\nTesting for AM1.5D Direct Normal ASTM E927-19')
        lowerBounds = np.array((400, 500, 600, 700, 800, 900))
        upperBounds = np.array((500, 600, 700, 800, 900, 1100))
        ratios = np.array((16.75, 19.49, 18.36, 15.08, 12.82, 16.69))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['400-500', '500-600', '600-700', '700-800', '800-900', '900-1100']
        
        # Ensure that there is enough data to fulfill standard requirements
        if interpWaves[-1] < upperBounds[5]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5D classification under ASTM E927-19.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5D classification under ASTM E927-19.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[5])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    if AMType == '2':   # ASTM AM1.5G
        print('\nTesting for AM1.5G Hemispherical ASTM E927-19')
        lowerBounds = np.array((400, 500, 600, 700, 800, 900))
        upperBounds = np.array((500, 600, 700, 800, 900, 1100))
        ratios = np.array((18.21, 19.73, 18.20, 14.79, 12.39, 15.89))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['400-500', '500-600', '600-700', '700-800', '800-900', '900-1100']
        
        # Ensure that there is enough data to fulfill standard requirements
        if interpWaves[-1] < upperBounds[5]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5G classification under ASTM E927-19.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5G classification under ASTM E927-19.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[5])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    if AMType == '3':   # ASTM AM0
        print('\nTesting for AM0 Extra-Terrestrial ASTM E927-19')
        lowerBounds = np.array((350, 400, 500, 600, 700, 800, 900, 1100))
        upperBounds = np.array((400, 500, 600, 700, 800, 900, 1100, 1400))
        ratios = np.array((4.67, 16.80, 16.68, 14.28, 11.31, 8.98, 13.50, 12.56))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['350-400', '400-500', '500-600', '600-700', '700-800', '800-900', '900-1100', '1100-1400']
        
        # Ensure that there is enough data to fulfill standard requirements.
        if interpWaves[-1] < upperBounds[7]:
            print('Your data does not cover the complete wavelength range (350 - 1400 nm) required for AM0 classification under ASTM E927-19.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (350 - 1400 nm) required for AM0 classification under ASTM E927-19.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[7])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    if AMType == '4':   # AM1.5G: IEC Table 1
        print('\nTesting for AM1.5G IEC 60904-9 Ed.3 Table 1')
        lowerBounds = np.array((400, 500, 600, 700, 800, 900))
        upperBounds = np.array((500, 600, 700, 800, 900, 1100))
        ratios = np.array((18.4, 19.9, 18.4, 14.9, 12.5, 15.9))
        stairsLO = ratios * 0.75
        stairsHI = ratios * 1.25
        binLabels = ['400-500', '500-600', '600-700', '700-800', '800-900', '900-1100']
        
        # Ensure that there is enough data to fulfill standard requirements.
        if interpWaves[-1] < upperBounds[5]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5G classification under IEC 60904-9, Table 1.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5G classification under IEC 60904-9, Table 1.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[5])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    elif AMType == '5':     # AM1.5G: IEC Table 2
        print('\nTesting for AM1.5G IEC 60904-9 Ed. 3 Table 2')
        lowerBounds = np.array((300, 470, 561, 657, 772, 919))
        upperBounds = np.array((470, 561, 657, 772, 919, 1200))
        ratios = np.array((16.61, 16.74, 16.67, 16.63, 16.66, 16.69))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['300-470', '470-561', '561-657', '657-772', '772-919', '919-1200']
        
        # Ensure that there is enough data to fulfill standard requirements.
        if interpWaves[-1] < upperBounds[5]:
            print('Your data does not cover the complete wavelenth range (300 - 1200 nm) required for AM1.5G classification under IEC 60904-9, Table 2.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (300 - 1200 nm) required for AM1.5G classification under IEC 60904-9, Table 2.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[5])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    elif AMType == '6':     # AM1.5G, 700 - 1100 nm only
        print('\nTesting for AM1.5G Hemispherical ASTM E927-19, Limited Range [700 - 1100 nm]')
        lowerBounds = np.array((700, 800, 900))
        upperBounds = np.array((800, 900, 1100))
        ratios = np.array((34.4, 28.7, 36.9))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['700-800', '800-900', '900-1100']
        
        # Ensure that there is enough data to fulfill standard requirements.
        if interpWaves[-1] < upperBounds[2]:
            print('Your data does not cover the complete wavelength (700 - 1100) for this comparison.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength (700 - 1100) for this comparison.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[2])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    # Calculate the absolute error between the normalized measured spectrum and the reference spectrum (This was erroneously called the SPD in the original MATLAB script).
    errorCheckWaves = np.round(np.arange(max(300, min(waves)), min(1200.1, max(waves) + 0.1), 0.1), 3)
    errorCheckIrrad = interper(errorCheckWaves)
    errorCheckInterper = interpolate.interp1d(specRef['Wavelengths'], refIrrad)
    errorCheck = errorCheckInterper(errorCheckWaves)
    errorCheckIrrad = errorCheckIrrad / sum(errorCheckIrrad) * sum(errorCheck)
    absError = sum(abs(errorCheckIrrad - errorCheck)) / sum(errorCheck) * 100
    
    # Calculate the aggregate SPC
    SPC_loop = np.zeros(len(errorCheckWaves))
    for loop in np.arange(len(SPC_loop)):
        if errorCheckIrrad[loop] > errorCheck[loop] * 0.1:
            SPC_loop[loop] = errorCheck[loop]
        else:
            SPC_loop[loop] = 0
    SPC = sum(SPC_loop) / sum(errorCheck) * 100
    
    # Generating Plots
    # Instantiate some lists to fill in loops
    midBounds = []
    stepWaves = []
    stepsLO = []
    stepsHI = []
    
    # Build the bounds for the stair plot
    for bound in np.arange(len(lowerBounds)):
        midBounds.append(np.mean([lowerBounds[bound], upperBounds[bound]]))
        stepWaves.append(lowerBounds[bound])
        stepWaves.append(upperBounds[bound])
        stepsLO.append(stairsLO[bound])
        stepsLO.append(stairsLO[bound])
        stepsHI.append(stairsHI[bound])
        stepsHI.append(stairsHI[bound])
    
    # Set the wavelength bounds depending on the spectrum being matched to
    if AMType == '1':
        plt1x = np.arange(400, 1101, 1)
        xLims = [400, 1100]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5D'
        AM = 'AM1.5D, ASTM E927-19'
    elif AMType == '2':
        plt1x = np.arange(400, 1101, 1)
        xLims = [400, 1100]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5G'
        AM = 'AM1.5G, ASTM E927-19'
    elif AMType == '3':
        plt1x = np.arange(350, 1401, 1)
        xLims = [350, 1400]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM0'
        AM = 'AM0, ASTM E927-19'
    elif AMType == '4':
        plt1x = np.arange(300, 1101, 1)
        xLims = [400, 1100]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5G'
        AM = 'AM1.5G, IEC Table 1'
    elif AMType == '5':
        plt1x = np.arange(300, 1201, 1)
        xLims = [300, 1200]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5G'
        AM = 'AM1.5G, IEC Table 2'
    elif AMType == '6':
        plt1x = np.arange(700, 1101, 1)
        xLims = [700, 1100]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5G'
        AM = 'AM1.5G, ASTM E927-19'
    
    # Normalize the data to overlap with SMARTS standard data.
    normFactor = errorCheckInterper(plt1x)
    plotInterper = interpolate.interp1d(waves, irrad)
    plt1y = plotInterper(plt1x)
    plt1y = plt1y / sum(plt1y) * sum(normFactor)    # Normalize the data to overlap with SMARTS standard data
    dataLabel = label
    
    # Subplots
    ax1 = plt.subplot2grid((2, 2), (0, 0), colspan = 2)
    ax2 = plt.subplot2grid((2, 2), (1, 0))
    ax3 = plt.subplot2grid((2, 2), (1, 1))
    
    # Top plot
    ax1.plot(specRef['Wavelengths'], refIrrad, label = stdLeg, lw = 0.5)
    ax1.plot(plt1x, plt1y, c = 'k', label = dataLabel, lw = 0.5)
    ax1.set_xlim(xLims)
    ax1.set_xlabel('Wavelength [nm]', fontsize = 6)
    ax1.set_ylabel(r'Spectral Irradiance [W/$cm^2$/nm]', fontsize = 6)
    ax1.ticklabel_format(axis = 'y', style = 'sci', scilimits = (0, 0))
    ax1.tick_params(axis = 'both', which = 'major', labelsize = 6)
    ax1.yaxis.get_offset_text().set_fontsize(6)
    ax1.legend(fontsize = 5)
    ax1.set_title('Solar Simulator Spectral Comparison for ' + AM, weight = 'bold')
    
    # Bottom Left Plot
    ax2.scatter(midBounds, percents, s = 10, c = 'k', label = dataLabel)
    ax2.plot(stepWaves, stepsLO, 'r--', label = 'Lower Limit, Class A', lw = 0.5)
    ax2.plot(stepWaves, stepsHI, 'b--', label = 'Upper Limit, Class A', lw = 0.5)
    ax2.set_xlim(xLims)
    ax2.set_xlabel('Wavelength [nm]', fontsize = 6)
    ax2.set_ylabel('Ratio of Interval Irradiance [%]', fontsize = 6)
    ax2.tick_params(axis = 'both', which = 'major', labelsize = 6)
    ax2.legend(fontsize = 5)
    ax2.set_title('Spectral Irradiance Ratios, Compared to ' + AM, weight = 'bold', fontsize = 5)
    
    # Bottom Right Table
    percents[bucket] = np.around(percents[bucket], 2)   # round them to 2 decimal places
    plotTable = {
        'Wavelength Interval [nm]': binLabels,
        'Ratio of Interval Irradiance [%]': np.round(percents, 2),
        'Classification': classLetters
        }
    tableDF = pd.DataFrame(plotTable)
    table = ax3.table(cellText = tableDF.values, colLabels = tableDF.columns, loc = 'center', cellLoc = 'center')
    ax3.set_axis_off()
    table.auto_set_font_size(False)
    table.set_fontsize(5)
    table.auto_set_column_width(col = list(range(len(tableDF.columns))))
    plt.tight_layout()
    plt.savefig("output/SM.png", dpi=300)
    
    # classification
    grade_order = ['A+', 'A', 'B', 'C', 'D', 'U']
    classification = max(classLetters, key=lambda g: grade_order.index(g))

    saveCheck = rawdata
    savePath = 'output'
    print(saveCheck)
    if saveCheck == True:
        dataToSave = {
            'Wavelengths [nm]': waves,
            'Spectral Irradiance [W/cm^2/nm]': irrad
            }
        saveName = '\Spectral Irradiance Data.txt'
        saveFrame = pd.DataFrame(dataToSave)
        saveFrame.to_csv(savePath + saveName, sep = '\t', index = False)
        print('Your irradiance data has been saved to ' + savePath + saveName + '.')
    elif saveCheck == False:
        print('Your irradiance data has not been saved.')

    # Display a small inline summary table containing information for the test report.
    if 'siData' in ssData and 'igaData' in ssData:
        report_d = {
            'Filename of Si Measurement': ssData['siData']['filename'],
            'Date of Si Measurement': ssData['siData']['date'],
            'Filename of InGaAs Measurement': ssData['igaData']['filename'],
            'Date of InGaAs Measurement': ssData['igaData']['filename'],
            'SPD Absolute Error [%]': round(absError, 4),
            'Aggregate SPC [%]': round(SPC, 4),
            'Classification': classification,
            'df' : saveFrame
            }
    elif 'siData' in ssData:
        report_d = {
            'Filename of Si Measurement': ssData['siData']['filename'],
            'Date of Si Measurement': ssData['siData']['date'],
            'SPD Absolute Error [%]': round(absError, 4),
            'Aggregate SPC [%]': round(SPC, 4),
            'Classification': classification,
            'df' : saveFrame
            }
    elif 'igaData' in ssData:
        report_d = {
            'Filename of InGaAs Measurement': ssData['igaData']['filename'],
            'Date of InGaAs Measurement': ssData['igaData']['date'],
            'SPD Absolute Error [%]': round(absError, 4),
            'Aggregate SPC [%]': round(SPC, 4),
            'Classification': classification,
            'df' : saveFrame
            }
    resultsFrame = pd.DataFrame.from_dict(report_d, orient = 'index')
    print(tabulate(resultsFrame, colalign = ('right',)))

    return report_d



def SMScript2(AMType, siData, label):
    '''
    specScript2: This function works with sidat file instead
    '''

    dir = os.path.abspath(os.path.dirname(__file__))
    parent_dir = os.path.dirname(dir)
    files = os.path.join(parent_dir, 'required_files')

    
    # Standard spectra file location. Currently given relative to the script's run location.
    stdSpecPath = os.path.join(files, 'Solar_Standards.csv')
    
    # The main function starts here.
    ssData = {}
    print(siData)

    waves = [float(w) for w in siData['wavelengths'][:-5]]
    irrad = [float(i) for i in siData['irradiance'][:-5]]

    # savePath = ssData['siData']['folder']

    # Import the reference spectra.
    specRef = pd.read_csv(
        stdSpecPath,
        sep=',',
        header=0,
        engine='python',
        names=['Wavelengths', 'AM0 Irrad', 'AM1.5G Irrad', 'AM1.5D Irrad', 'AM0 Rad', 'AM1.5G Rad', 'AM1.5D Rad'],
    )

    
    if AMType == '1':
        refIrrad = specRef['AM1.5D Irrad'] / 10000
    elif AMType == '2':
        refIrrad = specRef['AM1.5G Irrad'] / 10000
    elif AMType == '3':
        refIrrad = specRef['AM0 Irrad'] / 10000
    elif AMType == '4':
        refIrrad = specRef['AM1.5G Irrad'] / 10000
    elif AMType == '5':
        refIrrad = specRef['AM1.5G Irrad'] / 10000
    elif AMType == '6':
        refIrrad = specRef['AM1.5G Irrad'] / 10000
    else:
        pass
    
    # Interpolate the collected data so it has the same resolution as the reference spectrum (1 nm steps)
    interpWaves = np.round(np.arange(round(min(waves), 1), round(max(waves), 1) + 0.1, 0.1), 3)
    interper = interpolate.interp1d(waves, irrad)
    interpIrrad = interper(interpWaves)
    
    # Load the ASTM wavelength bins and intensity percentage, depending on the reference spectrum.
    if AMType == '1':       # ASTM AM1.5D
        print('\nTesting for AM1.5D Direct Normal ASTM E927-19')
        lowerBounds = np.array((400, 500, 600, 700, 800, 900))
        upperBounds = np.array((500, 600, 700, 800, 900, 1100))
        ratios = np.array((16.75, 19.49, 18.36, 15.08, 12.82, 16.69))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['400-500', '500-600', '600-700', '700-800', '800-900', '900-1100']
        
        # Ensure that there is enough data to fulfill standard requirements
        if interpWaves[-1] < upperBounds[5]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5D classification under ASTM E927-19.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5D classification under ASTM E927-19.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[5])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    if AMType == '2':   # ASTM AM1.5G
        print('\nTesting for AM1.5G Hemispherical ASTM E927-19')
        lowerBounds = np.array((400, 500, 600, 700, 800, 900))
        upperBounds = np.array((500, 600, 700, 800, 900, 1100))
        ratios = np.array((18.21, 19.73, 18.20, 14.79, 12.39, 15.89))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['400-500', '500-600', '600-700', '700-800', '800-900', '900-1100']
        
        # Ensure that there is enough data to fulfill standard requirements
        if interpWaves[-1] < upperBounds[5]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5G classification under ASTM E927-19.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5G classification under ASTM E927-19.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[5])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    if AMType == '3':   # ASTM AM0
        print('\nTesting for AM0 Extra-Terrestrial ASTM E927-19')
        lowerBounds = np.array((350, 400, 500, 600, 700, 800, 900, 1100))
        upperBounds = np.array((400, 500, 600, 700, 800, 900, 1100, 1400))
        ratios = np.array((4.67, 16.80, 16.68, 14.28, 11.31, 8.98, 13.50, 12.56))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['350-400', '400-500', '500-600', '600-700', '700-800', '800-900', '900-1100', '1100-1400']
        
        # Ensure that there is enough data to fulfill standard requirements.
        if interpWaves[-1] < upperBounds[7]:
            print('Your data does not cover the complete wavelength range (350 - 1400 nm) required for AM0 classification under ASTM E927-19.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (350 - 1400 nm) required for AM0 classification under ASTM E927-19.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[7])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    if AMType == '4':   # AM1.5G: IEC Table 1
        print('\nTesting for AM1.5G IEC 60904-9 Ed.3 Table 1')
        lowerBounds = np.array((400, 500, 600, 700, 800, 900))
        upperBounds = np.array((500, 600, 700, 800, 900, 1100))
        ratios = np.array((18.4, 19.9, 18.4, 14.9, 12.5, 15.9))
        stairsLO = ratios * 0.75
        stairsHI = ratios * 1.25
        binLabels = ['400-500', '500-600', '600-700', '700-800', '800-900', '900-1100']
        
        # Ensure that there is enough data to fulfill standard requirements.
        if interpWaves[-1] < upperBounds[5]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5G classification under IEC 60904-9, Table 1.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (400 - 1100 nm) required for AM1.5G classification under IEC 60904-9, Table 1.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[5])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    elif AMType == '5':     # AM1.5G: IEC Table 2
        print('\nTesting for AM1.5G IEC 60904-9 Ed. 3 Table 2')
        lowerBounds = np.array((300, 470, 561, 657, 772, 919))
        upperBounds = np.array((470, 561, 657, 772, 919, 1200))
        ratios = np.array((16.61, 16.74, 16.67, 16.63, 16.66, 16.69))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['300-470', '470-561', '561-657', '657-772', '772-919', '919-1200']
        
        # Ensure that there is enough data to fulfill standard requirements.
        if interpWaves[-1] < upperBounds[5]:
            print('Your data does not cover the complete wavelenth range (300 - 1200 nm) required for AM1.5G classification under IEC 60904-9, Table 2.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength range (300 - 1200 nm) required for AM1.5G classification under IEC 60904-9, Table 2.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[5])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    elif AMType == '6':     # AM1.5G, 700 - 1100 nm only
        print('\nTesting for AM1.5G Hemispherical ASTM E927-19, Limited Range [700 - 1100 nm]')
        lowerBounds = np.array((700, 800, 900))
        upperBounds = np.array((800, 900, 1100))
        ratios = np.array((34.4, 28.7, 36.9))
        stairsLO = ratios * 0.75    # Class A lower bounds for stair plot
        stairsHI = ratios * 1.25    # Class A upper bounds for stair plot
        binLabels = ['700-800', '800-900', '900-1100']
        
        # Ensure that there is enough data to fulfill standard requirements.
        if interpWaves[-1] < upperBounds[2]:
            print('Your data does not cover the complete wavelength (700 - 1100) for this comparison.')
            return
        if interpWaves[0] > lowerBounds[0]:
            print('Your data does not cover the complete wavelength (700 - 1100) for this comparison.')
            return
        
        topLimit = np.where(interpWaves == upperBounds[2])[0][0]
        lowLimit = np.where(interpWaves == lowerBounds[0])[0][0]
        
        # Sum the irradiance between the limits. This will only be used for % comparisons against itself, so we don't need to account for the change in number of data points due to interpolation.
        irradSum = np.sum(interpIrrad[lowLimit:topLimit + 1])
        
        # Compare what % of the total irradiance falls into each wavelength bucket.
        classLetters = []
        percents = []
        for bucket in np.arange(len(binLabels)):
            topBucketLim = np.where(interpWaves == upperBounds[bucket])[0][0]
            lowBucketLim = np.where(interpWaves == lowerBounds[bucket])[0][0]
            buckSum = np.sum(interpIrrad[lowBucketLim:topBucketLim + 1])
            percents.append(buckSum / irradSum * 100)
            
            # Record the spectral match classification for each bucket.
            if percents[bucket] >= ratios[bucket] * 0.875 and percents[bucket] <= ratios[bucket] * 1.125:
                classLetters.append('A+')
            elif percents[bucket] >= ratios[bucket] * 0.75 and percents[bucket] <= ratios[bucket] * 1.25:
                classLetters.append('A')
            elif percents[bucket] >= ratios[bucket] * 0.6 and percents[bucket] <= ratios[bucket] * 1.4:
                classLetters.append('B')
            elif percents[bucket] >= ratios[bucket] * 0.4 and percents[bucket] <= ratios[bucket] * 2.0:
                classLetters.append('C')
            else:
                classLetters.append('U')
    
    # Calculate the absolute error between the normalized measured spectrum and the reference spectrum (This was erroneously called the SPD in the original MATLAB script).
    errorCheckWaves = np.round(np.arange(max(300, min(waves)), min(1200.1, max(waves) + 0.1), 0.1), 3)
    errorCheckIrrad = interper(errorCheckWaves)
    errorCheckInterper = interpolate.interp1d(specRef['Wavelengths'], refIrrad)
    errorCheck = errorCheckInterper(errorCheckWaves)
    errorCheckIrrad = errorCheckIrrad / sum(errorCheckIrrad) * sum(errorCheck)
    absError = sum(abs(errorCheckIrrad - errorCheck)) / sum(errorCheck) * 100
    
    # Calculate the aggregate SPC
    SPC_loop = np.zeros(len(errorCheckWaves))
    for loop in np.arange(len(SPC_loop)):
        if errorCheckIrrad[loop] > errorCheck[loop] * 0.1:
            SPC_loop[loop] = errorCheck[loop]
        else:
            SPC_loop[loop] = 0
    SPC = sum(SPC_loop) / sum(errorCheck) * 100
    
    # Generating Plots
    # Instantiate some lists to fill in loops
    midBounds = []
    stepWaves = []
    stepsLO = []
    stepsHI = []
    
    # Build the bounds for the stair plot
    for bound in np.arange(len(lowerBounds)):
        midBounds.append(np.mean([lowerBounds[bound], upperBounds[bound]]))
        stepWaves.append(lowerBounds[bound])
        stepWaves.append(upperBounds[bound])
        stepsLO.append(stairsLO[bound])
        stepsLO.append(stairsLO[bound])
        stepsHI.append(stairsHI[bound])
        stepsHI.append(stairsHI[bound])
    
    # Set the wavelength bounds depending on the spectrum being matched to
    if AMType == '1':
        plt1x = np.arange(400, 1101, 1)
        xLims = [400, 1100]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5D'
        AM = 'AM1.5D, ASTM E927-19'
    elif AMType == '2':
        plt1x = np.arange(400, 1101, 1)
        xLims = [400, 1100]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5G'
        AM = 'AM1.5G, ASTM E927-19'
    elif AMType == '3':
        plt1x = np.arange(350, 1401, 1)
        xLims = [350, 1400]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM0'
        AM = 'AM0, ASTM E927-19'
    elif AMType == '4':
        plt1x = np.arange(300, 1101, 1)
        xLims = [400, 1100]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5G'
        AM = 'AM1.5G, IEC Table 1'
    elif AMType == '5':
        plt1x = np.arange(300, 1201, 1)
        xLims = [300, 1200]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5G'
        AM = 'AM1.5G, IEC Table 2'
    elif AMType == '6':
        plt1x = np.arange(700, 1101, 1)
        xLims = [700, 1100]
        stdLeg = 'ASTM G173-03 Reference Spectrum: AM1.5G'
        AM = 'AM1.5G, ASTM E927-19'
    
    # Normalize the data to overlap with SMARTS standard data.
    normFactor = errorCheckInterper(plt1x)
    plotInterper = interpolate.interp1d(waves, irrad)
    plt1y = plotInterper(plt1x)
    plt1y = plt1y / sum(plt1y) * sum(normFactor)    # Normalize the data to overlap with SMARTS standard data
    dataLabel = label
    
    # Subplots
    ax1 = plt.subplot2grid((2, 2), (0, 0), colspan = 2)
    ax2 = plt.subplot2grid((2, 2), (1, 0))
    ax3 = plt.subplot2grid((2, 2), (1, 1))
    
    # Top plot
    ax1.plot(specRef['Wavelengths'], refIrrad, label = stdLeg, lw = 0.5)
    ax1.plot(plt1x, plt1y, c = 'k', label = dataLabel, lw = 0.5)
    ax1.set_xlim(xLims)
    ax1.set_xlabel('Wavelength [nm]', fontsize = 6)
    ax1.set_ylabel(r'Spectral Irradiance [W/$cm^2$/nm]', fontsize = 6)
    ax1.ticklabel_format(axis = 'y', style = 'sci', scilimits = (0, 0))
    ax1.tick_params(axis = 'both', which = 'major', labelsize = 6)
    ax1.yaxis.get_offset_text().set_fontsize(6)
    ax1.legend(fontsize = 5)
    ax1.set_title('Solar Simulator Spectral Comparison for ' + AM, weight = 'bold')
    
    # Bottom Left Plot
    ax2.scatter(midBounds, percents, s = 10, c = 'k', label = dataLabel)
    ax2.plot(stepWaves, stepsLO, 'r--', label = 'Lower Limit, Class A', lw = 0.5)
    ax2.plot(stepWaves, stepsHI, 'b--', label = 'Upper Limit, Class A', lw = 0.5)
    ax2.set_xlim(xLims)
    ax2.set_xlabel('Wavelength [nm]', fontsize = 6)
    ax2.set_ylabel('Ratio of Interval Irradiance [%]', fontsize = 6)
    ax2.tick_params(axis = 'both', which = 'major', labelsize = 6)
    ax2.legend(fontsize = 5)
    ax2.set_title('Spectral Irradiance Ratios, Compared to ' + AM, weight = 'bold', fontsize = 5)
    
    # Bottom Right Table
    percents[bucket] = np.around(percents[bucket], 2)   # round them to 2 decimal places
    plotTable = {
        'Wavelength Interval [nm]': binLabels,
        'Ratio of Interval Irradiance [%]': np.round(percents, 2),
        'Classification': classLetters
        }
    tableDF = pd.DataFrame(plotTable)
    table = ax3.table(cellText = tableDF.values, colLabels = tableDF.columns, loc = 'center', cellLoc = 'center')
    ax3.set_axis_off()
    table.auto_set_font_size(False)
    table.set_fontsize(5)
    table.auto_set_column_width(col = list(range(len(tableDF.columns))))
    plt.tight_layout()
    plt.savefig("output/SM.png", dpi=300)
    
    # classification
    grade_order = ['A+', 'A', 'B', 'C', 'D', 'U']
    classification = max(classLetters, key=lambda g: grade_order.index(g))

    dataToSave = {
        'Wavelengths [nm]': waves,
        'Spectral Irradiance [W/m^2/nm]': irrad
        }
    saveFrame = pd.DataFrame(dataToSave)

    # Display a small inline summary table containing information for the test report.
    report_d = {
        # 'Filename of Si Measurement': ssData['siData']['filename'],
        # 'Date of Si Measurement': ssData['siData']['date'],
        # 'Filename of InGaAs Measurement': ssData['igaData']['filename'],
        # 'Date of InGaAs Measurement': ssData['igaData']['filename'],
        'SPD Absolute Error [%]': round(absError, 4),
        'Aggregate SPC [%]': round(SPC, 4),
        'Classification': classification,
        'df': saveFrame
        }

    # resultsFrame = pd.DataFrame.from_dict(report_d, orient = 'index')
    # print(tabulate(resultsFrame, colalign = ('right',)))


    return report_d
