# -*- coding: utf-8 -*-
"""
Created on Wed Apr 16 14:42:56 2025

@author: KyleGraham
"""
import numpy as np
from scipy import interpolate

def conv2Irrad(scanX, scanY, tFunc):
    '''
    conv2Irrad will convert an array of raw spectral data to spectral irradiance
    scanX should be wavelength values from the user's scan
    scanY should be voltage values from the user's scan
    tFunc is a transfer function for the measurement instrument in units of W/cm^2/nm/V
    '''    
    minX = np.ceil(max(min(scanX), min(tFunc['tWaves'])))    # Finds the min wavelength of each input, chooses the largest one, then rounds to the nearest integer.
    maxX = np.floor(min(max(scanX), max(tFunc['tWaves'])))   # Finds the max wavelength of each input, chooses the smallest one, then rounds to the nearest integer.
    print('\nInterpolating over range: ' + str(minX) + ' to ' + str(maxX) + '.')
    
    Xs = np.arange(minX,maxX + 1,1);
    
    # interpolate the test data over the wave range
    interper = interpolate.interp1d(scanX, scanY)
    interpTFunc = interper(Xs)
    
    k_min = tFunc[tFunc['tWaves'] == minX].index.values[0]
    k_max = tFunc[tFunc['tWaves'] == maxX].index.values[0]
    transfer = tFunc['tdata'].iloc[k_min:k_max + 1]
    transfer = transfer.reset_index()
    print('\nUnits in W/cm^2')
    converted = interpTFunc * transfer['tdata']
    
    theDictionary = {
        'irradWaves': Xs,
        'irrad': converted
        }
    
    return theDictionary