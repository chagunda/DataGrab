####  MassLynx DataGrab 2.9   ####

## Use this program for quick automated copying and exporting of MassLynx single ion chromatogram (sic) data.
## SIC data is normalised to TIC, then plotted onto a single chromatogram plot, and extracted into an Excel file.
## Input your experiment name (exp_name), m/z of species to copy (species), and a file output directory (output_dir) below, then run script!

## ATTN: Rogue automated mouse movements can be cancelled by quickly moving cursor to any of the four corners of the primary monitor.

import pyautogui as pg
pg.FAILSAFE = True  
import time as tm
import matplotlib.pyplot as plt
import pandas as pd
import openpyxl as opx
from openpyxl.drawing.image import Image
import string

exp_name = 'IC_20220410_01-193'  # Write your Experiment name here, eg: exp_name = 'My Experiment'
output_dir = r"C:\Users\User\Desktop\IC\DataGrab DEMO" # Input the directory to save processed data, eg: output_dir = r"C:\Users\IanC\Documents\Experiments"
species = [479.40, 459.54, 127.20, 257.38, 181.35] # Enter m/z of ions of interest to be copied, eg: species = [100, 150, 1234]

time_delay = int(1.0) # Insert a time delay between operations if MassLynx is running slow in format: int(x.x). 1.0 should be enough.


#  1. COPY TIC DATA AND SET UP MAIN DATAFRAME (df)

def copy_tic():
    chrom_coord = pg.locateCenterOnScreen('Chromatogram.png', confidence=0.7) # find the Chromatogram.png image on the desktop. Screenshot of Chromatogram window header must be saved in Directory
    if chrom_coord is None:
        print('Chromatogram Window was not found.') 
            
    else:
        pg.click(chrom_coord)
        pg.click(chrom_coord[0]+70, chrom_coord[1]+50) # Using coordinates of the Chromatogram to click on the "Copy" button, which is about x+70 and y+50 pixels away.
        tm.sleep(time_delay)
        data = pd.read_clipboard() # Copy clipboard data and setup dataframe df as floats
        df = data.astype(float)
        df.columns = ['Time / min', 'TIC'] # Label the first 2 columns of the dataframe

    return df # After df is made, return takes it out of the function copy_tic and available to other functions

df = copy_tic()
x = len(df.columns)


#  2. COPY OTHER SPECIES DATA, INSERT Y-VALUES (INTENSITY) AND NORMALIZED INTENSITY INTO df

def copy_sic(species):
        chrom_coord = pg.locateCenterOnScreen('Chromatogram.png', confidence=0.7)
        pg.click(chrom_coord[0]+60, chrom_coord[1]+30)  # Click on the "Display" button.
        pg.click(chrom_coord[0]+60, chrom_coord[1]+45)  # Click on the "Mass" button.
        #tm.sleep(time_delay)
        pg.typewrite(str(species)) # type in the m/z value (species), but first converted to a string.
        pg.hotkey('enter') # Press enter to load species chromatogram
        tm.sleep(time_delay) # Activates time lag in case MassLynx is slow, eg due to windows animations.
        pg.click(chrom_coord[0]+70, chrom_coord[1]+50) # Click on the "Copy" button.
        #tm.sleep(time_delay)

def insert_sic(df, i, x):
        global species
        data2 = pd.read_clipboard(header=None) # Make a new temporary dataframe of new species data, separate from the TIC. header=None so 1st row is copied, and not excluded as a label
        df2 = data2.astype(float)
        df.insert(x, "m/z "+str(species[i])+" Raw", df2[1], True)  # Insert into original dataframe. df.insert(location index, column label, values or array [specific index], allow duplicates=True).

def norm_sic(df, i, x): # Normalize each species intensity to TIC
        global species
        data3 = pd.read_clipboard() # Make a new temporary dataframe of new species data.
        df3 = data3.astype(float)  
        df4 = df3.iloc[:, 1].div(df.iloc[:, 1].values) # Divide all values in all rows of column 1 in df3 (SIC intensity), by all of column 1 in df (TIC intensity).
        df.insert(x+i+1, "m/z "+str(species[i]), df4, True)


# 3 PLOT FIGURE(S)

def plot_raw(): # Plots the Raw dataset
    ax1 = df.plot(x=df.columns[0], y=df.columns[2:len(species)+2], figsize=(6.5, 5), linewidth=1, kind='line', legend=True, fontsize=9, cmap=plt.cm.tab10)
    ax1.set_title(exp_name+" - Raw", fontdict={'fontsize': 12, 'color': 'k'})
    ax1.legend(loc=0, frameon=False, fontsize=9)
    ax1.set_xlabel('Time / min', fontdict={'fontsize': 10})
    ax1.set_ylabel('Intensity', fontdict={'fontsize': 10})
    ax1.spines['right'].set_visible(False)  # Removing the spines top and right
    ax1.spines['top'].set_visible(False)
    #plt.show()
    plt.savefig(output_dir+"\PLOT_"+exp_name+" - Raw"+".png", dpi=300) # Save the figure, dpi specifies size.

def plot_norm(): # Plots the Normalized dataset
    ax2 = df.plot(x=df.columns[0], y=df.columns[len(species)+2:], figsize=(6.5, 5), linewidth=1, kind='line', legend=True, fontsize=9, cmap=plt.cm.tab10)
    ax2.set_title(exp_name, fontdict={'fontsize': 12, 'color': 'k'})
    ax2.legend(loc=0, frameon=False, fontsize=9)
    ax2.set_xlabel('Time / min', fontdict={'fontsize': 10}) # Fontdict can also include; 'family': 'Arial',  'color': 'r', 'weight': 'bold', 'fontsize':10, 'style': 'italic', etc
    ax2.set_ylabel('Relative Intensity', fontdict={'fontsize': 10})
    ax2.spines['right'].set_visible(False) # Removing the spines top and right
    ax2.spines['top'].set_visible(False)
    #plt.show()
    plt.savefig(output_dir+"\PLOT_"+exp_name+".png", dpi=300) # Save the figure, dpi specifies size.


# 4. EXPORT DATA INTO EXCEL FILE, INCLUDING FIGURE

def save_excel(): # Saves excel with the normalized plot included.
    df.to_excel(output_dir+"\DATA_"+exp_name+".xlsx", index=False) 
    tm.sleep(time_delay)
    img = Image(output_dir+"\PLOT_"+exp_name+".png") # Opens previously made norm plot, to be inserted in Excel.
    #img = Image(output_dir+"\PLOT_"+exp_name+"- Raw"+".png") # Activate to plot the raw image instead
    wb = opx.load_workbook(output_dir+"\DATA_"+exp_name+".xlsx", data_only=True)
    ws = wb.worksheets[0] # Worksheet number where figure should be added into
    ws.add_image(img, string.ascii_uppercase[(len(species))+10]+'1') # convert the length of species +5 into corresponding letter, and input image at that column letter and row 1.
    wb.save(output_dir+"\DATA_"+exp_name+".xlsx") # Save the workbook with image.


# 5. RUN IT ALL!!!

for i in range(len(species)): # specifying that for every instance in species, run copy_sic, insert_sic and norm_sic
    copy_sic(species[i])
    insert_sic(df, i, x)
    norm_sic(df, i, x)

plot_norm()
save_excel()

# DONE!  © chagunda@uvic.ca
