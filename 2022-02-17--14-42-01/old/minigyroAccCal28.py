################################################################################
# Company       : Devico AS
# Author        : Geir Bjarte
#
#
#	 ______   _______          _________ _______  _______
#	(  __  \ (  ____ \|\     /|\__   __/(  ____ \(  ___  )
#	| (  \  )| (    \/| )   ( |   ) (   | (    \/| (   ) |
#	| |   ) || (__    | |   | |   | |   | |      | |   | |
#	| |   | ||  __)   ( (   ) )   | |   | |      | |   | |
#	| |   ) || (       \ \_/ /    | |   | |      | |   | |
#	| (__/  )| (____/\  \   /  ___) (___| (____/\| (___) |
#	(______/ (_______/   \_/   \_______/(_______/(_______)
#
# All Rights Reserved
################################################################################
# R1 initial version
# R2 improving structure with shorter main program and writing and reading files as needed-, not completed gyro cal
# R3 rewrite file import functions to adapt to minigyro file formats 
# R4 include writing temperature data to calibration files
# R5 correcting g sensitivity for earth rotation. Format for acc gsens timing files changed, also labelled R5 
# R6 updated to change in file format from A101 to A108 (E lines, mag temp and following columns moved one column to the left)
# R7 change to one sequence with rotations, then acc and gsens. Only one timing file, with four sheets
# R8 automate file names
# R9 change read file format to match that written "split cal files.py" script 
# R10 include calibration of ST gyro, adapt to FW A110
# R11 adapt to FW A112, include ST acceleromter data. Only change is in file format, causing edit of read file function
# R12 adapt to calibarting in one cycle only. Only minor changes
# R13 remove little used plots. Introduce processteps in automated sequence
# R14 include support for control signal at 28C, write figures to file
# R15 change local gravity at Heimdal to 9.821m/s^2, used from 10.06.2020
# R16 for python 3.8.5
# R17 fixed MAA to MAB bug in write constant file script. No change in this script
# R18 introduced gyro result image script, fixed minor error in plot 117

# R99 Commented out plt.show() for automatic processing

from __future__ import division
import matplotlib.pyplot as plt
import numpy as np
import xlrd
import xlsxwriter
from scipy import signal
from pandas import DataFrame


sequencemode = 3            #  0: run old style one processtep at a time
                            #  1: run processtep 1 only (use for first tool in a batch, where timing files need to be tuned)
                            #  2: run processteps 2-6 (use for first tool in a batch)
                            #  3: run processteps 1-6 (when timing is known in advance)   

processtep = 1              # processtep = 1: read acc and gyro raw data files from calibration, adjust timing until happy (in file_path_timefile)
                            #                 Plot QC figures 101-107. Write selected data to file
                            # processtep = 2: read selected accelerometer data, calculate calibration constants (ortomatrix and 3rd order)
                            #                 write calibration constants to file and plot QC plots 111-113
                            # processtep = 3: read raw data from gyro bias temperature test. Plot QC plot 121, and write temperature coeffcients to file
                            # processtep = 4: temperature compensate gyro data, calculate g sensitivity. For now, this is done without 
                            #                 taking earth rotation into account
                            # processtep = 5: read three raw data files from gyro calibration, read timing file (external excel document), plot data. Calculate 
                            #                 gyro orto matrix. Option to select GX1, GX2 or average. But shall be 3x3 matrix! 
                            # processtep = 6: calculate earth rotation compensated g sensitivity
                        



# parameters processtep 1

abstemp = 28
toolSN = '3331'

g_range_ADXL = 4         # for cal file names only
g_range_ST = 16          # for cal file names only

tempfile_read =  toolSN + ' Tcal cut.xlsx' 
tempsheet_read =  toolSN + ' Tcal cut'

av_time = 10
av_timeAC = 5

autostarttime = False
standstill_time = 10    

if autostarttime == False: 
	starttime_Tcomp = 80
	stoptime_Tcomp = 4722
                                                                # seconds to average gyro bias and temperature standstill values



logfile_acc_gsens = toolSN + ' ' + str(abstemp) + 'C' + ' ' + 'rotgsens.xlsx' 
sheetname_acc_gsens = toolSN + ' ' + str(abstemp) + 'C' + ' ' + 'rotgsens'

file_path_timefile = toolSN + ' ' + str(abstemp) + 'C' + ' ' + 'rotgsens timing.xlsx'       # file containing timing and set angles/accelerometer values
sheetname_time = 'acc gsens'
sheetname_control = 'acc control'

pickfile_write = toolSN + ' ' + str(abstemp) + 'C' + 'acc gsens picked data.xlsx'
picksheet_write =  'Ark1'

pickfile_writeAC = toolSN + ' ' + str(abstemp) + 'C' + 'acc gsens picked dataAC.xlsx'
picksheet_writeAC =  'Ark1'

f_s_acc = 10    


# plot figures in gsens wo erc
plotfigs1 = False
   
# plot figures in gsens with erc
plotfigs2 = True

controlseq = False
                                                         
if abstemp == 28:
    controlseq = True
    

# parameters processtep 2

pickfile_read = pickfile_write 
picksheet_read =  picksheet_write

pickfile_readAC =  pickfile_writeAC 
picksheet_readAC =  picksheet_writeAC


acc_cal_file = toolSN + ' ' + str(abstemp) + 'C' + ' ' + str(g_range_ADXL) + 'g acc cal.xlsx'          # file for writing accelerometer calibration constants
acc_cal_sheet =  'Ark1'


# parameters processtep 3



f_s_temp = 10                                                            

gyroTcal_file = toolSN + ' gyro Tconst.xlsx'



gyroTcal_sheet = 'Ark1'


# parameters processtep 4
                                            


gyrogsenscal_file = toolSN + ' ' + str(abstemp) + 'C' + ' ' + 'gyro gsens ier.xlsx'

    
gyrogsenscal_sheet = 'Ark1'


# parameters processtep 5

Xrotfile = logfile_acc_gsens
Xrotsheet = sheetname_acc_gsens

Yrotfile = logfile_acc_gsens
Yrotsheet = sheetname_acc_gsens

Zrotfile = logfile_acc_gsens
Zrotsheet = sheetname_acc_gsens

rot_timefile = file_path_timefile
Xrot_timesheet = 'Xrot'
Yrot_timesheet = 'Yrot'
Zrot_timesheet = 'Zrot'

f_s_rot = 10   
av_time_standstill = 10         # seconds to average data in the standstill periods
av_time_rot = 180               # seconds to average rotation data, 180s corresponds to five full rounds 
margin_standstill = 5           # margin between end of averaging period and initiation of next sequence
margin_rotation = 10            # margin between end of rotation period and initiation of next sequence

bestGX = 0                      # recommended setting for GX counter values. All options are calibrated and stored, 
                                # but this parameter should set for use in later calibration of data 
                                # 0: Recommend use (GX1+GX2)/2 in data calibration. This should be the default setting
                                # 1: recommend use GX1 only in data calibration. Use if GX2 is clearly poorer than GX1
                                # 2: use GX2 only in data calibration

gyro_orto_file = toolSN + ' ' + str(abstemp) + 'C' + ' ' + 'gyro orto.xlsx'

gyro_orto_sheet = 'Ark1'


# parameters processtep 6

earthrate = 15.04       # deg/hour
latitude = 63.35        # deg at Heimdal

gyrogsenscal_erc_file = toolSN + ' ' + str(abstemp) + 'C' + ' ' + 'gyro gsens erc.xlsx'

gyrogsenscal_erc_sheet = 'Ark1'

# function definitions

          
def readrawdata(file_path, sheetname_data):
          
    # read data file (flex data)
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheetname_data)
       
    # extract data from X accelerometer
    offset = 1
    col = worksheet.col_values(1)
    accXcount = col[offset:]
    accXcount = np.array(accXcount)
    # extract data from Y accelerometer
    col = worksheet.col_values(2)
    accYcount = col[offset:]
    accYcount = np.array(accYcount)
    # extract data from Z accelerometer
    col = worksheet.col_values(3)
    accZcount = col[offset:]
    accZcount = np.array(accZcount)   
    
    # extract index
    col = worksheet.col_values(0)
    index = col[offset:]
    index = np.array(index)
    
    # extract gyro counter values
    col = worksheet.col_values(7)
    GX1count = col[offset:]
    GX1count = np.array(GX1count)
    
    col = worksheet.col_values(8)
    GY1count = col[offset:]
    GY1count = np.array(GY1count)
    
    col = worksheet.col_values(9)
    GX2count = col[offset:]
    GX2count = np.array(GX2count)    
    
    col = worksheet.col_values(10)
    GZ2count = col[offset:]
    GZ2count = np.array(GZ2count)           
    
    # extract ST gyro counter values
    col = worksheet.col_values(11)
    GX3count = col[offset:]
    GX3count = np.array(GX3count)
    
    col = worksheet.col_values(12)
    GY3count = col[offset:]
    GY3count = np.array(GY3count)
    
    col = worksheet.col_values(13)
    GZ3count = col[offset:]
    GZ3count = np.array(GZ3count)    
    
   
    # extract temperature data
    col = worksheet.col_values(14)
    TAcount = col[offset:]
    TAcount = np.array(TAcount)
    
    col = worksheet.col_values(15)   
    TG1count = col[offset:]
    TG1count = np.array(TG1count)
    
    col = worksheet.col_values(16)
    TG2count = col[offset:]
    TG2count = np.array(TG2count)
    
    col = worksheet.col_values(17)
    TG3count = col[offset:]
    TG3count = np.array(TG3count)
    
    index = index - index[0]
    
    return index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, \
            TAcount, TG1count, TG2count, TG3count 


# changed file format from R5
def readtimefile(file_path, sheetname_data):

    # read data file 
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheetname_data)
       
    # extract index data
    offset = 1
    col = worksheet.col_values(0)
    meas_point = col[offset:]
    meas_point = np.array(meas_point)
    
    # extract time gap data
    col = worksheet.col_values(1)
    timegap = col[offset:]
    timegap = np.array(timegap)
    
    # extract absolute timing data
    col = worksheet.col_values(2)
    timeabs = col[offset:]
    timeabs = np.array(timeabs)   
    
    # extract reference acceleration values
    col = worksheet.col_values(6)
    AXRef = col[offset:]
    AXRef = np.array(AXRef)
    
    col = worksheet.col_values(7)
    AYRef = col[offset:]
    AYRef = np.array(AYRef)
    
    col = worksheet.col_values(8)
    AZRef = col[offset:]
    AZRef = np.array(AZRef)         
    
    return meas_point, timegap, timeabs, AXRef, AYRef, AZRef 


def readanglefile(file_path, sheetname_data):

    # read data file 
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheetname_data)
    
    offset = 1
    
    # extract angle values
    col = worksheet.col_values(3)
    azimuth = col[offset:]
    azimuth = np.array(azimuth)
    
    col = worksheet.col_values(4)
    inclination = col[offset:]
    inclination = np.array(inclination)
    
    col = worksheet.col_values(5)
    toolface = col[offset:]
    toolface = np.array(toolface)         
    
    return azimuth, inclination, toolface 


def plot_figures_ps1(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                     timeabs, av_points, GX1count, GY1count, GX2count, GZ2count, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, \
                     GY1pick_av, GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, \
                     GX2pick_std, GZ2pick_std, TG1pick_av, TG2pick_av, TApick_av):
    
    #plot figures for processtep #1
    
    plt.close(101)
    plt.close(102)
    plt.close(103)
    plt.close(104)
    plt.close(105)
    plt.close(106)
    plt.close(107)
    
   
    plt.figure(101, figsize=(12,9))
    plt.plot(time, accXcount, label = 'accXcount')
 
    plt.plot(time, accYcount, label = 'accYcount')
    plt.plot(time, accZcount, label = 'accZcount')
    plt.plot(timeabs, AXpick, 'C0o', label = 'accXpick')
    plt.plot(timeabs, AYpick, 'C1o', label = 'accYpick')
    plt.plot(timeabs, AZpick, 'C2o', label = 'accZpick')
    plt.plot(timeabs, AXpick_av, 'C0s', label = 'accXpick_av')
    plt.plot(timeabs, AYpick_av, 'C1s', label = 'accYpick_av')
    plt.plot(timeabs, AZpick_av, 'C2s', label = 'accZpick_av') 
    
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))],accXcount[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C3')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))],accYcount[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C4')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))],accZcount[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C5')
    for n in range(len(meas_point)): 
        plt.annotate(n, xy=(timeabs[n],AXpick[n]+50))
        plt.annotate(n, xy=(timeabs[n],AYpick[n]+50))
        plt.annotate(n, xy=(timeabs[n],AZpick[n]+50))
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    
    
    plt.figure(102, figsize=(12,9))
    plt.plot(time, GX1count, label = 'GX1count')

    plt.plot(time, GY1count, label = 'GY1count')
    plt.plot(time, GX2count, label = 'GX2count')
    plt.plot(time, GZ2count, label = 'GZ2count')
    
    plt.plot(timeabs, GX1pick, 'C0o', label = 'GX1pick')
    plt.plot(timeabs, GY1pick, 'C1o', label = 'GY1pick')
    plt.plot(timeabs, GX2pick, 'C2o', label = 'GX2pick')
    plt.plot(timeabs, GZ2pick, 'C3o', label = 'GZ2pick')
 
    plt.plot(timeabs, GX1pick_av, 'C0s', label = 'GX1pick_av')
    plt.plot(timeabs, GY1pick_av, 'C1s', label = 'GY1pick_av')
    plt.plot(timeabs, GX2pick_av, 'C2s', label = 'GX2pick_av')
    plt.plot(timeabs, GZ2pick_av, 'C3s', label = 'GZ2pick_av')
    
    
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GX1count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C4')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GY1count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C5')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GX2count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C6')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GZ2count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C7')
    for n in range(len(meas_point)): 
        plt.annotate(n, xy=(timeabs[n],GX1pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GY1pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GX2pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GZ2pick[n]+50))
    
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    
    plt.savefig(toolSN + ' prim '+ str(abstemp)+ 'C plot 102 timing.png', bbox_inches='tight', dpi=200)
    
#    plt.show()
    
    
    plt.figure(103,figsize=(12,9))
    plt.plot(meas_point, AXpick_std, label = 'AXstd')

    plt.plot(meas_point, AYpick_std, label = 'AYstd')
    plt.plot(meas_point, AZpick_std, label = 'AZstd')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' prim '+ str(abstemp)+ 'C plot 103 acc stdev.png')
#    plt.show()
    
    plt.figure(104, figsize=(12,9))
    plt.plot(meas_point, GX1pick_std, label = 'GX1std')

    plt.plot(meas_point, GY1pick_std, label = 'GY1std')
    plt.plot(meas_point, GX2pick_std, label = 'GX2std')
    plt.plot(meas_point, GZ2pick_std, label = 'GZ2std')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' prim '+ str(abstemp)+ 'C plot 104 gyro stdev.png')
#    plt.show()

    plt.figure(105, figsize=(12,9))
    plt.plot(meas_point, AXpick_av, label = 'AX')

    plt.plot(meas_point, AYpick_av, label = 'AY')
    plt.plot(meas_point, AZpick_av, label = 'AZ')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    
    plt.figure(106, figsize=(12,9))
    plt.plot(meas_point, GX1pick_av, label = 'GX1av')

    plt.plot(meas_point, GY1pick_av, label = 'GY1av')
    plt.plot(meas_point, GX2pick_av, label = 'GX2av')
    plt.plot(meas_point, GZ2pick_av, label = 'GZ2av')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    
    plt.figure(107, figsize=(12,9))
    plt.plot(meas_point, TG1pick_av, label = 'TG1av')

    plt.plot(meas_point, TG2pick_av, label = 'TG2av')
    plt.plot(meas_point, TApick_av, label = 'TAav')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    return

def plot_figures_ps1AC(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                     timeabs, av_points, GX1count, GY1count, GX2count, GZ2count, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, \
                     GY1pick_av, GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, \
                     GX2pick_std, GZ2pick_std, TG1pick_av, TG2pick_av, TApick_av):
    
    #plot figures for processtep #1
    
    plt.close(108)
    plt.close(109)
    plt.close(110)
    
    plt.figure(108, figsize=(12,9))
    plt.plot(time, accXcount, label = 'accXcount')

    plt.plot(time, accYcount, label = 'accYcount')
    plt.plot(time, accZcount, label = 'accZcount')
    plt.plot(timeabs, AXpick, 'C0o', label = 'accXpick')
    plt.plot(timeabs, AYpick, 'C1o', label = 'accYpick')
    plt.plot(timeabs, AZpick, 'C2o', label = 'accZpick')
    plt.plot(timeabs, AXpick_av, 'C0s', label = 'accXpick_av')
    plt.plot(timeabs, AYpick_av, 'C1s', label = 'accYpick_av')
    plt.plot(timeabs, AZpick_av, 'C2s', label = 'accZpick_av') 
    
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))],accXcount[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C3')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))],accYcount[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C4')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))],accZcount[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C5')
    for n in range(len(meas_point)): 
        plt.annotate(n, xy=(timeabs[n],AXpick[n]+50))
        plt.annotate(n, xy=(timeabs[n],AYpick[n]+50))
        plt.annotate(n, xy=(timeabs[n],AZpick[n]+50))
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    
   
    plt.figure(109, figsize=(12,9))
    plt.plot(time, GX1count, label = 'GX1count')

    plt.plot(time, GY1count, label = 'GY1count')
    plt.plot(time, GX2count, label = 'GX2count')
    plt.plot(time, GZ2count, label = 'GZ2count')
    
    plt.plot(timeabs, GX1pick, 'C0o', label = 'GX1pick')
    plt.plot(timeabs, GY1pick, 'C1o', label = 'GY1pick')
    plt.plot(timeabs, GX2pick, 'C2o', label = 'GX2pick')
    plt.plot(timeabs, GZ2pick, 'C3o', label = 'GZ2pick')
 
    plt.plot(timeabs, GX1pick_av, 'C0s', label = 'GX1pick_av')
    plt.plot(timeabs, GY1pick_av, 'C1s', label = 'GY1pick_av')
    plt.plot(timeabs, GX2pick_av, 'C2s', label = 'GX2pick_av')
    plt.plot(timeabs, GZ2pick_av, 'C3s', label = 'GZ2pick_av')
    
    
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GX1count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C4')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GY1count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C5')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GX2count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C6')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GZ2count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C7')
    for n in range(len(meas_point)): 
        plt.annotate(n, xy=(timeabs[n],GX1pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GY1pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GX2pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GZ2pick[n]+50))
    
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    
    plt.savefig(toolSN + ' prim '+ str(abstemp)+ 'C plot 109 timing control.png', bbox_inches='tight', dpi=200)
    
#    plt.show()
    
    
    
    
    plt.figure(110, figsize=(12,9))
    plt.plot(meas_point, AXpick_std, label = 'AXstd')
  
    plt.plot(meas_point, AYpick_std, label = 'AYstd')
    plt.plot(meas_point, AZpick_std, label = 'AZstd')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    
   
    return




def acc_gsens_pick(meas_point, av_time, timeabs, f_s_acc, accXcount, accYcount, accZcount, TAcount, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count):
    
    AXpick = np.zeros(len(meas_point))
    AYpick = np.zeros(len(meas_point))
    AZpick = np.zeros(len(meas_point))
    
    AXpick_av = np.zeros(len(meas_point))
    AYpick_av = np.zeros(len(meas_point))
    AZpick_av = np.zeros(len(meas_point))
    TApick_av = np.zeros(len(meas_point))
    
    AXpick_std = np.zeros(len(meas_point))
    AYpick_std = np.zeros(len(meas_point))
    AZpick_std = np.zeros(len(meas_point))
    
    GX1pick = np.zeros(len(meas_point))
    GY1pick = np.zeros(len(meas_point))
    GX2pick = np.zeros(len(meas_point))
    GZ2pick = np.zeros(len(meas_point))
       
    
    GX1pick_av = np.zeros(len(meas_point))
    GY1pick_av = np.zeros(len(meas_point))
    GX2pick_av = np.zeros(len(meas_point))
    GZ2pick_av = np.zeros(len(meas_point))
    TG1pick_av = np.zeros(len(meas_point))
    TG2pick_av = np.zeros(len(meas_point))
    
    
    GX1pick_std = np.zeros(len(meas_point))
    GY1pick_std = np.zeros(len(meas_point))
    GX2pick_std = np.zeros(len(meas_point))
    GZ2pick_std = np.zeros(len(meas_point))
    
    av_points = av_time * f_s_acc
    
    
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
        AXpick[n] = accXcount[k]
        AYpick[n] = accYcount[k]
        AZpick[n] = accZcount[k]
        AXpick_av[n] = np.mean(accXcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AYpick_av[n] = np.mean(accYcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AZpick_av[n] = np.mean(accZcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TApick_av[n] = np.mean(TAcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AXpick_std[n] = np.std(accXcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AYpick_std[n] = np.std(accYcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AZpick_std[n] = np.std(accZcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        
        GX1pick[n] = GX1count[k]
        GY1pick[n] = GY1count[k]
        GX2pick[n] = GX2count[k]
        GZ2pick[n] = GZ2count[k]
        GX1pick_av[n] = np.mean(GX1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY1pick_av[n] = np.mean(GY1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GX2pick_av[n] = np.mean(GX2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ2pick_av[n] = np.mean(GZ2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TG1pick_av[n] = np.mean(TG1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TG2pick_av[n] = np.mean(TG2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        
        GX1pick_std[n] = np.std(GX1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY1pick_std[n] = np.std(GY1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GX2pick_std[n] = np.std(GX2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ2pick_std[n] = np.std(GZ2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
    
    return  AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, \
            GY1pick_av, GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, GX2pick_std, \
            GZ2pick_std, TG1pick_av, TG2pick_av, TApick_av


def acc_control_pick(meas_point, av_time, timeabs, f_s_acc, accXcount, accYcount, accZcount, TAcount, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count):
    
    AXpick = np.zeros(len(meas_point))
    AYpick = np.zeros(len(meas_point))
    AZpick = np.zeros(len(meas_point))
    
    AXpick_av = np.zeros(len(meas_point))
    AYpick_av = np.zeros(len(meas_point))
    AZpick_av = np.zeros(len(meas_point))
    TApick_av = np.zeros(len(meas_point))
    
    AXpick_std = np.zeros(len(meas_point))
    AYpick_std = np.zeros(len(meas_point))
    AZpick_std = np.zeros(len(meas_point))
    
    GX1pick = np.zeros(len(meas_point))
    GY1pick = np.zeros(len(meas_point))
    GX2pick = np.zeros(len(meas_point))
    GZ2pick = np.zeros(len(meas_point))
       
    
    GX1pick_av = np.zeros(len(meas_point))
    GY1pick_av = np.zeros(len(meas_point))
    GX2pick_av = np.zeros(len(meas_point))
    GZ2pick_av = np.zeros(len(meas_point))
    TG1pick_av = np.zeros(len(meas_point))
    TG2pick_av = np.zeros(len(meas_point))
    
    
    GX1pick_std = np.zeros(len(meas_point))
    GY1pick_std = np.zeros(len(meas_point))
    GX2pick_std = np.zeros(len(meas_point))
    GZ2pick_std = np.zeros(len(meas_point))
    
    av_points = av_time * f_s_acc
    
    
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
        AXpick[n] = accXcount[k]
        AYpick[n] = accYcount[k]
        AZpick[n] = accZcount[k]
        AXpick_av[n] = np.mean(accXcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AYpick_av[n] = np.mean(accYcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AZpick_av[n] = np.mean(accZcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TApick_av[n] = np.mean(TAcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AXpick_std[n] = np.std(accXcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AYpick_std[n] = np.std(accYcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        AZpick_std[n] = np.std(accZcount[(int(k-av_points/2))-1:(int(k+av_points/2))])
        
        GX1pick[n] = GX1count[k]
        GY1pick[n] = GY1count[k]
        GX2pick[n] = GX2count[k]
        GZ2pick[n] = GZ2count[k]
        GX1pick_av[n] = np.mean(GX1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY1pick_av[n] = np.mean(GY1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GX2pick_av[n] = np.mean(GX2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ2pick_av[n] = np.mean(GZ2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TG1pick_av[n] = np.mean(TG1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TG2pick_av[n] = np.mean(TG2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        
        GX1pick_std[n] = np.std(GX1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY1pick_std[n] = np.std(GY1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GX2pick_std[n] = np.std(GX2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ2pick_std[n] = np.std(GZ2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
    
    return  AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, \
            GY1pick_av, GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, GX2pick_std, \
            GZ2pick_std, TG1pick_av, TG2pick_av, TApick_av


def save_file_ps1(pickfile_write, picksheet_write, meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, \
                  GX2pick_av, GZ2pick_av, TG1pick_av, TG2pick_av, TApick_av, AXRef, AYRef, AZRef):
    
    l0 = meas_point
    l1 = AXpick_av
    l2 = AYpick_av
    l3 = AZpick_av
    l4 = GX1pick_av
    l5 = GY1pick_av
    l6 = GX2pick_av
    l7 = GZ2pick_av
    l8 = TApick_av
    l9 = TG1pick_av
    l10 = TG2pick_av
    l11 = AXRef
    l12 = AYRef
    l13 = AZRef
    
    
    df = DataFrame({'00 Meas point':l0, '01 AXpick_av':l1, '02 AYpick_av':l2, '03 AZpick_av':l3, '04 GX1pick_av': l4, '05 GY1pick_av': l5, 
                    '06 GX2pick_av': l6, '07 GZ2pick_av':l7, '08 TApick_av': l8, '09 TG1pick_av': l9, '10 TG2pick_av': l10,
                    '11 AXRef': l11, '12 AYRef': l12, '13 AZRef': l13})
    df.to_excel(pickfile_write, picksheet_write, index = False)
    
    return


def save_file_ps1AC(pickfile_write, picksheet_write, meas_point, AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef):
    
    l0 = meas_point
    l1 = AXpick_av
    l2 = AYpick_av
    l3 = AZpick_av
    l4 = TApick_av
    l5 = AXRef
    l6 = AYRef
    l7 = AZRef
    
    
    df = DataFrame({'00 Meas point':l0, '01 AXpick_av':l1, '02 AYpick_av':l2, '03 AZpick_av':l3, '04 TApick_av': l4, 
                    '05 AXRef': l5, '06 AYRef': l6, '07 AZRef': l7})
    df.to_excel(pickfile_write, picksheet_write, index = False)
    
    return



def read_file_ps1(file_path, sheetname_data):

    # read data file 
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheetname_data)
       
    # extract index data
    offset = 1
    col = worksheet.col_values(0)
    meas_point = col[offset:]
    meas_point = np.array(meas_point)
    
    # extract AX data
    col = worksheet.col_values(1)
    AXpick_av = col[offset:]
    AXpick_av = np.array(AXpick_av)
    
    # extract AY data
    col = worksheet.col_values(2)
    AYpick_av = col[offset:]
    AYpick_av = np.array(AYpick_av)  
    
    # extract AZ data
    col = worksheet.col_values(3)
    AZpick_av = col[offset:]
    AZpick_av = np.array(AZpick_av)
    
    # extract GX1 data
    col = worksheet.col_values(4)
    GX1pick_av = col[offset:]
    GX1pick_av = np.array(GX1pick_av)
    
    # extract GY1 data
    col = worksheet.col_values(5)
    GY1pick_av = col[offset:]
    GY1pick_av = np.array(GY1pick_av)    

    # extract GX2 data
    col = worksheet.col_values(6)
    GX2pick_av = col[offset:]
    GX2pick_av = np.array(GX2pick_av)     
    
    # extract GZ2 data
    col = worksheet.col_values(7)
    GZ2pick_av = col[offset:]
    GZ2pick_av = np.array(GZ2pick_av) 
    
    # extract TA data
    col = worksheet.col_values(8)
    TApick_av = col[offset:]
    TApick_av = np.array(TApick_av) 
    
    # extract TG1 data
    col = worksheet.col_values(9)
    TG1pick_av = col[offset:]
    TG1pick_av = np.array(TG1pick_av) 
    
    # extract TG2 data
    col = worksheet.col_values(10)
    TG2pick_av = col[offset:]
    TG2pick_av = np.array(TG2pick_av) 
    
    # extract AXRef data
    col = worksheet.col_values(11)
    AXRef = col[offset:]
    AXRef = np.array(AXRef)
    
    # extract AYRef data
    col = worksheet.col_values(12)
    AYRef = col[offset:]
    AYRef = np.array(AYRef) 
    
    # extract AZRef data
    col = worksheet.col_values(13)
    AZRef = col[offset:]
    AZRef = np.array(AZRef)
    
    
    return meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef   



def read_file_ps1AC(file_path, sheetname_data):

    # read data file 
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheetname_data)
       
    # extract index data
    offset = 1
    col = worksheet.col_values(0)
    meas_point = col[offset:]
    meas_point = np.array(meas_point)
    
    # extract AX data
    col = worksheet.col_values(1)
    AXpick_av = col[offset:]
    AXpick_av = np.array(AXpick_av)
    
    # extract AY data
    col = worksheet.col_values(2)
    AYpick_av = col[offset:]
    AYpick_av = np.array(AYpick_av)  
    
    # extract AZ data
    col = worksheet.col_values(3)
    AZpick_av = col[offset:]
    AZpick_av = np.array(AZpick_av)
       
    # extract TA data
    col = worksheet.col_values(4)
    TApick_av = col[offset:]
    TApick_av = np.array(TApick_av) 
    
    # extract AXRef data
    col = worksheet.col_values(5)
    AXRef = col[offset:]
    AXRef = np.array(AXRef)
    
    # extract AYRef data
    col = worksheet.col_values(6)
    AYRef = col[offset:]
    AYRef = np.array(AYRef) 
    
    # extract AZRef data
    col = worksheet.col_values(7)
    AZRef = col[offset:]
    AZRef = np.array(AZRef)
    
    return meas_point, AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef





def acc_cal_classic(AX, AY, AZ, TA, AXRef, AYRef, AZRef):
    
    rows = len(AX)
    M = np.empty((rows,4))
    MT = np.empty((4, rows))
    MTM = np.empty((4, 4))
    invMTM = np.empty((4, 4))
    invMTM_MT = np.empty((4, rows))
    F = np.empty((rows,3))
    OT = np.empty((4, 3))
    O = np.empty((3, 3))
    invO = np.empty((3, 3))
    
    M[:,0] = 1
    M[:,1] = AX
    M[:,2] = AY
    M[:,3] = AZ
      
    F[:,0] = AXRef
    F[:,1] = AYRef
    F[:,2] = AZRef
 
    MT = np.transpose(M)
    MTM = np.matmul(MT, M)
    invMTM = np.linalg.inv(MTM)
    invMTM_MT = np.matmul(invMTM, MT)
    OT = np.matmul(invMTM_MT, F)
    xfactors = OT[:, 0]
    yfactors = OT[:, 1]
    zfactors = OT[:, 2]
    
    O = OT[1:4,:]
    O = np.transpose(O)

    invO = np.linalg.inv(O)
    offset = OT[0,:]
    offsetT = np.transpose(offset)
    offset = -np.matmul(invO, offsetT)
    xoffset = offset[0]
    yoffset = offset[1]
    zoffset = offset[2] 
        
    T = np.average(TA)
    
    return xfactors, yfactors, zfactors, xoffset, yoffset, zoffset, T



def acc_cal_3rd_order(AX, AY, AZ, TA, AXRef, AYRef, AZRef):
    
    rows = len(AX)
    M = np.empty((rows,10))
    MT = np.empty((10, rows))
    MTM = np.empty((10, 10))
    invMTM = np.empty((10, 10))
    invMTM_MT = np.empty((10, rows))
    F = np.empty((rows,3))
    OT = np.empty((10, 3))
    
    M[:,0] = 1
    M[:,1] = AX
    M[:,2] = AY
    M[:,3] = AZ
    M[:,4] = AX**2
    M[:,5] = AY**2
    M[:,6] = AZ**2
    M[:,7] = AX**3
    M[:,8] = AY**3
    M[:,9] = AZ**3
    
   
    F[:,0] = AXRef
    F[:,1] = AYRef
    F[:,2] = AZRef
 
    MT = np.transpose(M)
    MTM = np.matmul(MT, M)
    invMTM = np.linalg.inv(MTM)
    invMTM_MT = np.matmul(invMTM, MT)
    OT = np.matmul(invMTM_MT, F)
    xfactors = OT[:, 0]
    yfactors = OT[:, 1]
    zfactors = OT[:, 2]
                   
    T = np.average(TA)
    
    return xfactors, yfactors, zfactors, T


def save_file_cal_constants(filename, sheetname, AXoffset, AYoffset, AZoffset, Xconstants_lin, Yconstants_lin, Zconstants_lin, \
                  Xconstants_3rd, Yconstants_3rd, Zconstants_3rd, T, abstemp):
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet(sheetname)
    
    worksheet.write(0,0,'AX offset')
    worksheet.write(1,0, AXoffset)
    worksheet.write(0,1,'AY offset')
    worksheet.write(1,1, AYoffset)
    worksheet.write(0,2,'AZ offset')
    worksheet.write(1,2, AZoffset)
    worksheet.write(0,3,'AX constants')
    worksheet.write(0,4,'AY constants')
    worksheet.write(0,5,'AZ constants')
    worksheet.write(0,6,'AX constants')
    worksheet.write(0,7,'AY constants')
    worksheet.write(0,8,'AZ constants') 
    worksheet.write(0,9,'TA count')
    worksheet.write(1,9, T)
    worksheet.write(0, 10,'Cal temp (degC)')
    worksheet.write(1, 10, abstemp)
    
    for k in range(4):
        worksheet.write(k+1, 3, Xconstants_lin[k])
        worksheet.write(k+1, 4, Yconstants_lin[k]) 
        worksheet.write(k+1, 5, Zconstants_lin[k]) 
    
    for k in range(10):
        worksheet.write(k+1, 6, Xconstants_3rd[k])
        worksheet.write(k+1, 7, Yconstants_3rd[k]) 
        worksheet.write(k+1, 8, Zconstants_3rd [k]) 
    
    workbook.close()    
    return



def read_file_cal_constants(filename, sheetname):
    
    # read data file 
    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_name(sheetname)
       
    # extract cal data
    offset = 1
    col = worksheet.col_values(0)
    AXoffset = col[1]
    AXoffset = np.array(AXoffset)
    
    col = worksheet.col_values(1)
    AYoffset = col[1]
    AYoffset = np.array(AYoffset)

    col = worksheet.col_values(2)
    AZoffset = col[1]
    AZoffset = np.array(AZoffset)
    
    col = worksheet.col_values(3)
    Xconstants_lin = col[1:5]
    Xconstants_lin = np.array(Xconstants_lin)
    
    col = worksheet.col_values(4)
    Yconstants_lin = col[1:5]
    Yconstants_lin = np.array(Yconstants_lin)

    col = worksheet.col_values(5)
    Zconstants_lin = col[1:5]
    Zconstants_lin = np.array(Zconstants_lin)
    
    col = worksheet.col_values(6)
    Xconstants_3rd = col[offset:]
    Xconstants_3rd = np.array(Xconstants_3rd)
    
    col = worksheet.col_values(7)
    Yconstants_3rd = col[offset:]
    Yconstants_3rd = np.array(Yconstants_3rd)

    col = worksheet.col_values(8)
    Zconstants_3rd = col[offset:]
    Zconstants_3rd = np.array(Zconstants_3rd)
    
        
    return AXoffset, AYoffset, AZoffset, Xconstants_lin, Yconstants_lin, Zconstants_lin, Xconstants_3rd, Yconstants_3rd, Zconstants_3rd


def calibrate_acc_classical(AXcount, AYcount, AZcount, xfactor, yfactor, zfactor):
    
    AXcal = np.zeros(len(AXcount))
    AYcal = np.zeros(len(AXcount))
    AZcal = np.zeros(len(AXcount))    
    
    for n in range(len(AXcount)):
        AXcal[n] = xfactor[0] + xfactor[1] * AXcount[n] + xfactor[2] * AYcount[n] + xfactor[3] * AZcount[n]
        AYcal[n] = yfactor[0] + yfactor[1] * AXcount[n] + yfactor[2] * AYcount[n] + yfactor[3] * AZcount[n]
        AZcal[n] = zfactor[0] + zfactor[1] * AXcount[n] + zfactor[2] * AYcount[n] + zfactor[3] * AZcount[n]
    
    return AXcal, AYcal, AZcal

def calibrate_acc_3rd_order(AXcount, AYcount, AZcount, xfactor, yfactor, zfactor):
    
    AXcal = np.zeros(len(AXcount))
    AYcal = np.zeros(len(AXcount))
    AZcal = np.zeros(len(AXcount))    
    
    for n in range(len(AXcount)):
        AXcal[n] = xfactor[0] + xfactor[1] * AXcount[n] + xfactor[2] * AYcount[n] + xfactor[3] * AZcount[n] \
                + xfactor[4] * AXcount[n]**2 + xfactor[5] * AYcount[n]**2 + xfactor[6] * AZcount[n]**2  \
                + xfactor[7] * AXcount[n]**3 + xfactor[8] * AYcount[n]**3 + xfactor[9] * AZcount[n]**3  

        AYcal[n] = yfactor[0] + yfactor[1] * AXcount[n] + yfactor[2] * AYcount[n] + yfactor[3] * AZcount[n]  \
                + yfactor[4] * AXcount[n]**2 + yfactor[5] * AYcount[n]**2 + yfactor[6] * AZcount[n]**2  \
                + yfactor[7] * AXcount[n]**3 + yfactor[8] * AYcount[n]**3 + yfactor[9] * AZcount[n]**3  
        
        AZcal[n] = zfactor[0] + zfactor[1] * AXcount[n] + zfactor[2] * AYcount[n] + zfactor[3] * AZcount[n]   \
                + zfactor[4] * AXcount[n]**2 + zfactor[5] * AYcount[n]**2 + zfactor[6] * AZcount[n]**2        \
                + zfactor[7] * AXcount[n]**3 + zfactor[8] * AYcount[n]**3 + zfactor[9] * AZcount[n]**3  
    
    return AXcal, AYcal, AZcal



def acc_calibrationQC(AXcal_lin, AYcal_lin, AZcal_lin, AXcal_3rd, AYcal_3rd, AZcal_3rd, AXRef, AYRef, AZRef):
    
    plt.close(111)
    plt.close(112)
    plt.close(113)
    
    AX_error_lin = np.zeros(len(AXcal_lin))
    AY_error_lin = np.zeros(len(AXcal_lin))
    AZ_error_lin = np.zeros(len(AXcal_lin))
    AX_error_3rd = np.zeros(len(AXcal_lin))
    AY_error_3rd = np.zeros(len(AXcal_lin))
    AZ_error_3rd = np.zeros(len(AXcal_lin))
    
    inc_ref = np.zeros(len(AXcal_lin))
    inc_lin = np.zeros(len(AXcal_lin))
    inc_3rd = np.zeros(len(AXcal_lin))
    tf_ref = np.zeros(len(AXcal_lin))
    tf_lin = np.zeros(len(AXcal_lin))
    tf_3rd = np.zeros(len(AXcal_lin))
    
    inc_error_lin = np.zeros(len(AXcal_lin))
    inc_error_3rd = np.zeros(len(AXcal_lin))
    tf_error_lin = np.zeros(len(AXcal_lin))
    tf_error_3rd = np.zeros(len(AXcal_lin))      

    
    AX_error_lin = AXcal_lin - AXRef
    AY_error_lin = AYcal_lin - AYRef
    AZ_error_lin = AZcal_lin - AZRef
    
    AX_error_3rd = AXcal_3rd - AXRef
    AY_error_3rd = AYcal_3rd - AYRef
    AZ_error_3rd = AZcal_3rd - AZRef

    for n in range(len(AXpick_av)):
        inc_ref[n] = -np.degrees(np.arctan(AXRef[n]/np.sqrt((AYRef[n]**2+AZRef[n]**2)))) 
        inc_lin[n] = -np.degrees(np.arctan(AXcal_lin[n]/np.sqrt((AYcal_lin[n]**2+AZcal_lin[n]**2)))) 
        inc_3rd[n] = -np.degrees(np.arctan(AXcal_3rd[n]/np.sqrt((AYcal_3rd[n]**2+AZcal_3rd[n]**2))))
        
        tf_ref[n] = np.degrees(np.arctan2(AYRef[n], AZRef[n])) + 180
        if tf_ref[n] >= 360:
            tf_ref[n] = tf_ref[n] - 360
        if tf_ref[n] < 0:
            tf_ref[n] = tf_ref[n] + 360
        
        tf_lin[n] = np.degrees(np.arctan2(AYcal_lin[n], AZcal_lin[n])) + 180
        if tf_lin[n] >= 360:
            tf_lin[n] = tf_lin[n] - 360
        if tf_lin[n] < 0:
            tf_lin[n] = tf_lin[n] + 360
            
        tf_3rd[n] = np.degrees(np.arctan2(AYcal_3rd[n], AZcal_3rd[n])) + 180
        if tf_3rd[n] >= 360:
            tf_3rd[n] = tf_3rd[n] - 360
        if tf_3rd[n] < 0:
            tf_3rd[n] = tf_3rd[n] + 360
    
        inc_error_lin[n] = inc_lin[n] - inc_ref[n]
        inc_error_3rd[n] = inc_3rd[n] - inc_ref[n]
        
        tf_error_lin[n] = tf_lin[n] - tf_ref[n]
        if tf_error_lin[n] >= 180:
             tf_error_lin[n] =  tf_error_lin[n] - 360
        if tf_error_lin[n] < -180:
             tf_error_lin[n] =  tf_error_lin[n] + 360
             
        tf_error_3rd[n] = tf_3rd[n] - tf_ref[n]
        if tf_error_3rd[n] >= 180:
             tf_error_3rd[n] =  tf_error_3rd[n] - 360
        if tf_error_3rd[n] < -180:
             tf_error_3rd[n] =  tf_error_3rd[n] + 360
    
    # points were rig is vertical, so TF undefined 
    for k in(3,9,15,21,27,33,75,81,87,93,99,105):
        tf_error_lin[k] = 0
        tf_error_3rd[k] = 0
    
    
    plt.figure(111, figsize=(12,9))
    plt.plot(meas_point, AX_error_lin, label = 'AX lin')

    plt.plot(meas_point, AY_error_lin, label = 'AY lin')
    plt.plot(meas_point, AZ_error_lin, label = 'AZ lin')
    
    plt.plot(meas_point, AX_error_3rd, label = 'AX 3rd order')
    plt.plot(meas_point, AY_error_3rd, label = 'AY 3rd order')
    plt.plot(meas_point, AZ_error_3rd, label = 'AZ 3rd order')
    
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' prim '+ str(abstemp)+'C plot 111 acc result.png')
#    plt.show()
    
    plt.figure(112, figsize=(12,9))
    plt.plot(meas_point, inc_error_lin, label = 'Inc error lin')

    plt.plot(meas_point, inc_error_3rd, label = 'Inc error 3rd')

    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' prim '+ str(abstemp)+'C plot 112 acc inc result.png')
#    plt.show()
    
    plt.figure(113, figsize=(12,9))
    plt.plot(meas_point, tf_error_lin, label = 'TF error lin')

    plt.plot(meas_point, tf_error_3rd, label = 'TF error 3rd')

    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' prim '+ str(abstemp)+ 'C plot 113 acc tf result.png')
#    plt.show()

    return

def acc_calibration_control(meas_point, AXcal_lin, AYcal_lin, AZcal_lin, AXRef, AYRef, AZRef):
    
    plt.close(115)
    plt.close(116)
    plt.close(117)
    
    AX_error_lin = np.zeros(len(AXcal_lin))
    AY_error_lin = np.zeros(len(AXcal_lin))
    AZ_error_lin = np.zeros(len(AXcal_lin))

    inc_ref = np.zeros(len(AXcal_lin))
    inc_lin = np.zeros(len(AXcal_lin))
    tf_ref = np.zeros(len(AXcal_lin))
    tf_lin = np.zeros(len(AXcal_lin))
    
    inc_error_lin = np.zeros(len(AXcal_lin))
    tf_error_lin = np.zeros(len(AXcal_lin))

    
    AX_error_lin = AXcal_lin - AXRef
    AY_error_lin = AYcal_lin - AYRef
    AZ_error_lin = AZcal_lin - AZRef
    

    for n in range(len(AXcal_lin)):
        inc_ref[n] = -np.degrees(np.arctan(AXRef[n]/np.sqrt((AYRef[n]**2+AZRef[n]**2)))) 
        inc_lin[n] = -np.degrees(np.arctan(AXcal_lin[n]/np.sqrt((AYcal_lin[n]**2+AZcal_lin[n]**2)))) 
        
        tf_ref[n] = np.degrees(np.arctan2(AYRef[n], AZRef[n])) + 180
        if tf_ref[n] >= 360:
            tf_ref[n] = tf_ref[n] - 360
        if tf_ref[n] < 0:
            tf_ref[n] = tf_ref[n] + 360
            
        
        tf_lin[n] = np.degrees(np.arctan2(AYcal_lin[n], AZcal_lin[n])) + 180
        if tf_lin[n] >= 360:
            tf_lin[n] = tf_lin[n] - 360
        if tf_lin[n] < 0:
            tf_lin[n] = tf_lin[n] + 360
            
        
        inc_error_lin[n] = inc_lin[n] - inc_ref[n]
        
        tf_error_lin[n] = tf_lin[n] - tf_ref[n]
        if tf_error_lin[n] >= 180:
             tf_error_lin[n] =  tf_error_lin[n] - 360
        if tf_error_lin[n] < -180:
             tf_error_lin[n] =  tf_error_lin[n] + 360
             
    
   
    
    plt.figure(115, figsize=(12,9))
    plt.plot(meas_point, AX_error_lin, label = 'AX lin')

    plt.plot(meas_point, AY_error_lin, label = 'AY lin')
    plt.plot(meas_point, AZ_error_lin, label = 'AZ lin')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    plt.savefig(toolSN +' prim '+ str(abstemp)+ 'C plot 115 acc control.png')
    
    
    fig, ax1 = plt.subplots(num=116, figsize=(12,9))
    plt.title('Inc error')
    plt.plot(meas_point, inc_error_lin, 'r', label = 'Inc error')
    plt.legend(loc = 'upper left')
    plt.ylabel('error [deg]')
    plt.grid()
    ax2 = ax1.twinx()
    plt.plot(meas_point, inc_ref, label = 'Inc')
    plt.plot(meas_point, tf_ref, label = 'TF')
    plt.legend(loc = 'upper right')
    plt.ylabel('angle [deg]')
    plt.savefig(toolSN + ' prim '+ str(abstemp)+ 'C plot 116 acc inc control.png')


    fig, ax1 = plt.subplots(num=117, figsize=(12,9))
    plt.title('Toolface error')
    plt.plot(meas_point, tf_error_lin, 'b', label = 'TF error')
    plt.legend(loc = 'upper left')
    plt.ylabel('error [deg]')
    plt.grid()
    ax2 = ax1.twinx()
    plt.plot(meas_point, inc_ref, label = 'Inc')
    plt.plot(meas_point, tf_ref, label = 'TF')
    plt.legend(loc = 'upper right')
    plt.ylabel('angle [deg]')
    plt.savefig(toolSN + ' prim '+ str(abstemp)+ 'C plot 117 acc tf control.png')

    return


def find_gyro_temp_poly(index, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, f_s):
    
    plt.close(121)
    
    
    TcoeffGX1 = np.zeros(6)
    TcoeffGY1 = np.zeros(6)
    TcoeffGX2 = np.zeros(6)
    TcoeffGZ2 = np.zeros(6)
    
    TcoeffGX1lin = np.zeros(2)
    TcoeffGY1lin = np.zeros(2)
    TcoeffGX2lin = np.zeros(2)
    TcoeffGZ2lin = np.zeros(2)
       
    TcoeffGX1 = np.polyfit(TG1count, GX1count, 5)
    TcoeffGY1 = np.polyfit(TG1count, GY1count, 5)
    TcoeffGX2 = np.polyfit(TG2count, GX2count, 5)
    TcoeffGZ2 = np.polyfit(TG2count, GZ2count, 5)
    
    TG1min = np.min(TG1count)
    TG1max = np.max(TG1count)
    TG2min = np.min(TG2count)
    TG2max = np.max(TG2count)
    
    GX1poly = np.zeros(int(TG1max-TG1min+1))
    GY1poly = np.zeros(int(TG1max-TG1min+1))
    TG1poly = np.zeros(int(TG1max-TG1min+1))
    GX2poly = np.zeros(int(TG2max-TG2min+1))
    GZ2poly = np.zeros(int(TG2max-TG2min+1))
    TG2poly = np.zeros(int(TG2max-TG2min+1))
    
    GX1lin = np.zeros(int(TG1max-TG1min+1)+100)
    GY1lin = np.zeros(int(TG1max-TG1min+1)+100)
    TG1lin = np.zeros(int(TG1max-TG1min+1)+100)
    GX2lin = np.zeros(int(TG2max-TG2min+1)+100)
    GZ2lin = np.zeros(int(TG2max-TG2min+1)+100)
    TG2lin = np.zeros(int(TG2max-TG2min+1)+100)
    
    for n in range(int(1+TG1max) - int(TG1min)):
        TG1poly[n] = TG1min + n
        GX1poly[n] = TcoeffGX1[5] + TcoeffGX1[4]*TG1poly[n] + TcoeffGX1[3]*TG1poly[n]**2 + TcoeffGX1[2]*TG1poly[n]**3 + \
                     TcoeffGX1[1]*TG1poly[n]**4 + TcoeffGX1[0]*TG1poly[n]**5
        GY1poly[n] = TcoeffGY1[5] + TcoeffGY1[4]*TG1poly[n] + TcoeffGY1[3]*TG1poly[n]**2 + TcoeffGY1[2]*TG1poly[n]**3 +  \
                     TcoeffGY1[1]*TG1poly[n]**4 + TcoeffGY1[0]*TG1poly[n]**5
    
    for n in range(int(1+TG2max) - int(TG2min)):
        TG2poly[n] = TG2min + n
        GX2poly[n] = TcoeffGX2[5] + TcoeffGX2[4]*TG2poly[n] + TcoeffGX2[3]*TG2poly[n]**2 + TcoeffGX2[2]*TG2poly[n]**3  \
                    + TcoeffGX2[1]*TG2poly[n]**4 + TcoeffGX2[0]*TG2poly[n]**5
        GZ2poly[n] = TcoeffGZ2[5] + TcoeffGZ2[4]*TG2poly[n] + TcoeffGZ2[3]*TG2poly[n]**2 + TcoeffGZ2[2]*TG2poly[n]**3   \
                     + TcoeffGZ2[1]*TG2poly[n]**4 + TcoeffGZ2[0]*TG2poly[n]**5
    
    # find linear fit to be used if temperature counter values outside calibrated range
    
    GX1Tmin = TcoeffGX1[5] + TcoeffGX1[4]*TG1min + TcoeffGX1[3]*TG1min**2 + TcoeffGX1[2]*TG1min**3 + \
                     TcoeffGX1[1]*TG1min**4 + TcoeffGX1[0]*TG1min**5
    GX1Tmax = TcoeffGX1[5] + TcoeffGX1[4]*TG1max + TcoeffGX1[3]*TG1max**2 + TcoeffGX1[2]*TG1max**3 + \
                     TcoeffGX1[1]*TG1max**4 + TcoeffGX1[0]*TG1max**5
    
    GY1Tmin = TcoeffGY1[5] + TcoeffGY1[4]*TG1min + TcoeffGY1[3]*TG1min**2 + TcoeffGY1[2]*TG1min**3 + \
                     TcoeffGY1[1]*TG1min**4 + TcoeffGY1[0]*TG1min**5
    GY1Tmax = TcoeffGY1[5] + TcoeffGY1[4]*TG1max + TcoeffGY1[3]*TG1max**2 + TcoeffGY1[2]*TG1max**3 + \
                     TcoeffGY1[1]*TG1max**4 + TcoeffGY1[0]*TG1max**5
    
    GX2Tmin = TcoeffGX2[5] + TcoeffGX2[4]*TG2min + TcoeffGX2[3]*TG2min**2 + TcoeffGX2[2]*TG2min**3 + \
                     TcoeffGX2[1]*TG2min**4 + TcoeffGX2[0]*TG2min**5
    GX2Tmax = TcoeffGX2[5] + TcoeffGX2[4]*TG2max + TcoeffGX2[3]*TG2max**2 + TcoeffGX2[2]*TG2max**3 + \
                     TcoeffGX2[1]*TG2max**4 + TcoeffGX2[0]*TG2max**5
    
    GZ2Tmin = TcoeffGZ2[5] + TcoeffGZ2[4]*TG2min + TcoeffGZ2[3]*TG2min**2 + TcoeffGZ2[2]*TG2min**3 + \
                     TcoeffGZ2[1]*TG2min**4 + TcoeffGZ2[0]*TG2min**5
    GZ2Tmax = TcoeffGZ2[5] + TcoeffGZ2[4]*TG2max + TcoeffGZ2[3]*TG2max**2 + TcoeffGZ2[2]*TG2max**3 + \
                     TcoeffGZ2[1]*TG2max**4 + TcoeffGZ2[0]*TG2max**5
    
    TcoeffGX1lin[0] = (GX1Tmax - GX1Tmin)/(TG1max - TG1min)  
    TcoeffGX1lin[1] = GX1Tmin - TcoeffGX1lin[0] * TG1min
    
    TcoeffGY1lin[0] = (GY1Tmax - GY1Tmin)/(TG1max - TG1min)  
    TcoeffGY1lin[1] = GY1Tmin - TcoeffGY1lin[0] * TG1min
    
    TcoeffGX2lin[0] = (GX2Tmax - GX2Tmin)/(TG2max - TG2min)  
    TcoeffGX2lin[1] = GX2Tmin - TcoeffGX2lin[0] * TG2min
    
    TcoeffGZ2lin[0] = (GZ2Tmax - GZ2Tmin)/(TG2max - TG2min)  
    TcoeffGZ2lin[1] = GZ2Tmin - TcoeffGZ2lin[0] * TG2min
    
    for n in range(int(51+TG1max) - int(TG1min-50)):
        TG1lin[n] = TG1min + n - 50
        GX1lin[n] = TcoeffGX1lin[1] + TcoeffGX1lin[0]*TG1lin[n] 
        GY1lin[n] = TcoeffGY1lin[1] + TcoeffGY1lin[0]*TG1lin[n] 
    
    for n in range(int(51+TG2max) - int(TG2min-50)):
        TG2lin[n] = TG2min + n - 50
        GX2lin[n] = TcoeffGX2lin[1] + TcoeffGX2lin[0]*TG2lin[n] 
        GZ2lin[n] = TcoeffGZ2lin[1] + TcoeffGZ2lin[0]*TG2lin[n] 
    
    plt.figure(121, figsize=(12,9))
    plt.plot(TG1count, GX1count, label = 'GX1count')

    plt.plot(TG1count, GY1count, label = 'GY1count')
    plt.plot(TG2count, GX2count, label = 'GX2count')
    plt.plot(TG2count, GZ2count, label = 'GZ2count')
    
    plt.plot(TG1poly, GX1poly, label = 'GX1poly', lw = 3)
    plt.plot(TG1poly, GY1poly, label = 'GY1poly', lw = 3)
    plt.plot(TG2poly, GX2poly, label = 'GX2poly', lw = 3)
    plt.plot(TG2poly, GZ2poly, label = 'GZ2poly', lw = 3)
    
    plt.plot(TG1lin, GX1lin, label = 'GX1lin', lw = 1)
    plt.plot(TG1lin, GY1lin, label = 'GY1lin', lw = 1)
    plt.plot(TG2lin, GX2lin, label = 'GX2lin', lw = 1)
    plt.plot(TG2lin, GZ2lin, label = 'GZ2lin', lw = 1)
    
    
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' prim '+ str(abstemp)+ 'C plot 121 gyro Tbias.png')
#    plt.show()
    
#    print TcoeffGX1
#    print TcoeffGY1
#    print TcoeffGX2
#    print TcoeffGZ2
    
    return TcoeffGX1, TcoeffGY1, TcoeffGX2,TcoeffGZ2, TcoeffGX1lin, TcoeffGY1lin, TcoeffGX2lin, TcoeffGZ2lin, TG1min, TG1max, TG2min, TG2max

def save_file_gyroT_constants(filename, sheetname, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2, TcoeffGX1lin, TcoeffGY1lin, \
                              TcoeffGX2lin, TcoeffGZ2lin, TG1min, TG1max, TG2min, TG2max):
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet(sheetname)
    
    worksheet.write(0, 0, 'GX1 coeffs')
    worksheet.write(0, 1, 'GY1 coeffs')
    worksheet.write(0, 2, 'GX2 coeffs')
    worksheet.write(0, 3, 'GZ2 coeffs')
    worksheet.write(0, 4, 'GX1 lin coeffs')
    worksheet.write(0, 5, 'GY1 lin coeffs')
    worksheet.write(0, 6, 'GX2 lin coeffs')
    worksheet.write(0, 7, 'GZ2 lin coeffs')
    worksheet.write(0, 8, 'TG1min')
    worksheet.write(0, 9, 'TG1max')
    worksheet.write(0, 10, 'TG2min')
    worksheet.write(0, 11, 'TG2max')
    worksheet.write(1, 8, TG1min)
    worksheet.write(1, 9, TG1max)
    worksheet.write(1, 10, TG2min)
    worksheet.write(1, 11, TG2max)


    for k in range(6):
        worksheet.write(k+1, 0, TcoeffGX1[k])
        worksheet.write(k+1, 1, TcoeffGY1[k])
        worksheet.write(k+1, 2, TcoeffGX2[k])
        worksheet.write(k+1, 3, TcoeffGZ2[k])
    for k in range(2):
        worksheet.write(k+1, 4, TcoeffGX1lin[k])
        worksheet.write(k+1, 5, TcoeffGY1lin[k]) 
        worksheet.write(k+1, 6, TcoeffGX2lin[k]) 
        worksheet.write(k+1, 7, TcoeffGZ2lin[k])
    
    workbook.close() 
    return


def gyro_bias_comp(index, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, ss1_length, ss2_length, GX1_Tconst, GY1_Tconst, GX2_Tconst, GZ2_Tconst):
    
    GX1_TP = np.zeros(len(index))              # temperature polynomials
    GY1_TP = np.zeros(len(index))
    GX2_TP = np.zeros(len(index))
    GZ2_TP = np.zeros(len(index))
    
    GX1_Tbias = np.zeros(len(index))
    GY1_Tbias = np.zeros(len(index))
    GX2_Tbias = np.zeros(len(index))
    GZ2_Tbias = np.zeros(len(index))
    
    GX1bias = np.zeros(len(index))
    GY1bias = np.zeros(len(index))
    GX2bias = np.zeros(len(index))
    GZ2bias = np.zeros(len(index))
        
    for k in range(len(index)):
        GX1_TP[k] = GX1_Tconst[5] + GX1_Tconst[4]*TG1count[k] + GX1_Tconst[3]*TG1count[k]**2 + GX1_Tconst[2]*TG1count[k]**3 \
                    + GX1_Tconst[1]*TG1count[k]**4 + GX1_Tconst[0]*TG1count[k]**5
        GY1_TP[k] = GY1_Tconst[5] + GY1_Tconst[4]*TG1count[k] + GY1_Tconst[3]*TG1count[k]**2 + GY1_Tconst[2]*TG1count[k]**3 \
                    + GY1_Tconst[1]*TG1count[k]**4 + GY1_Tconst[0]*TG1count[k]**5
        GX2_TP[k] = GX2_Tconst[5] + GX2_Tconst[4]*TG2count[k] + GX2_Tconst[3]*TG2count[k]**2 + GX2_Tconst[2]*TG2count[k]**3  \
                    + GX2_Tconst[1]*TG2count[k]**4 + GX2_Tconst[0]*TG2count[k]**5
        GZ2_TP[k] = GZ2_Tconst[5] + GZ2_Tconst[4]*TG2count[k] + GZ2_Tconst[3]*TG2count[k]**2 + GZ2_Tconst[2]*TG2count[k]**3  \
                    + GZ2_Tconst[1]*TG2count[k]**4 + GZ2_Tconst[0]*TG2count[k]**5
    
    GX1_TPss1 = np.average(GX1_TP[0:ss1_length])              # Temperature polynomial average value in first standstill period 
    GY1_TPss1 = np.average(GY1_TP[0:ss1_length])
    GX2_TPss1 = np.average(GX2_TP[0:ss1_length])
    GZ2_TPss1 = np.average(GZ2_TP[0:ss1_length])
    
    GX1_GCss1 = np.average(GX1count[0:ss1_length])            # Gravity compensated gyro counter value average value in first standstill period 
    GY1_GCss1 = np.average(GY1count[0:ss1_length])
    GX2_GCss1 = np.average(GX2count[0:ss1_length])
    GZ2_GCss1 = np.average(GZ2count[0:ss1_length])
    
    GX1_GCss2 = np.average(GX1count[(len(index)-ss2_length):len(index)])    # Gravity compensated gyro counter value average value in second standstill period
    GY1_GCss2 = np.average(GY1count[(len(index)-ss2_length):len(index)])
    GX2_GCss2 = np.average(GX2count[(len(index)-ss2_length):len(index)])
    GZ2_GCss2 = np.average(GZ2count[(len(index)-ss2_length):len(index)]) 
        
    for k in range(len(index)):
        GX1_Tbias[k] = GX1_GCss1 + GX1_TP[k] - GX1_TPss1        # This is temperature polynomial adjusted in counter value to match gravity 
        GY1_Tbias[k] = GY1_GCss1 + GY1_TP[k] - GY1_TPss1        # compensated bias in the first standstill period 
        GX2_Tbias[k] = GX2_GCss1 + GX2_TP[k] - GX2_TPss1
        GZ2_Tbias[k] = GZ2_GCss1 + GZ2_TP[k] - GZ2_TPss1
    
    GX1_Tbias_ss2 = np.average(GX1_Tbias[(len(index)-ss2_length):len(index)])     # Temperature commpensated bias value, average at second standstill 
    GY1_Tbias_ss2 = np.average(GY1_Tbias[(len(index)-ss2_length):len(index)]) 
    GX2_Tbias_ss2 = np.average(GX2_Tbias[(len(index)-ss2_length):len(index)])
    GZ2_Tbias_ss2 = np.average(GZ2_Tbias[(len(index)-ss2_length):len(index)])
    
    aX1 = (GX1_GCss2 - GX1_Tbias_ss2)/(len(index) - ss1_length - ss2_length)       # slope for adjusting remaining bias (after g and temp comp) linearly
    aY1 = (GY1_GCss2 - GY1_Tbias_ss2)/(len(index) - ss1_length - ss2_length)
    aX2 = (GX2_GCss2 - GX2_Tbias_ss2)/(len(index) - ss1_length - ss2_length)
    aZ2 = (GZ2_GCss2 - GZ2_Tbias_ss2)/(len(index) - ss1_length - ss2_length)
    
    for k in range(ss1_length):
        GX1bias[k] = GX1count[k]
        GY1bias[k] = GY1count[k]
        GX2bias[k] = GX2count[k]
        GZ2bias[k] = GZ2count[k]
    
    for k in range(len(index) - ss1_length - ss2_length):
        GX1bias[k+ss1_length] = GX1_Tbias[k+ss1_length] + aX1*k
        GY1bias[k+ss1_length] = GY1_Tbias[k+ss1_length] + aY1*k
        GX2bias[k+ss1_length] = GX2_Tbias[k+ss1_length] + aX2*k
        GZ2bias[k+ss1_length] = GZ2_Tbias[k+ss1_length] + aZ2*k
    
    for k in range (ss2_length):
        GX1bias[k + len(index) - ss2_length] = GX1count[k + len(index) - ss2_length]
        GY1bias[k + len(index) - ss2_length] = GY1count[k + len(index) - ss2_length]
        GX2bias[k + len(index) - ss2_length] = GX2count[k + len(index) - ss2_length]
        GZ2bias[k + len(index) - ss2_length] = GZ2count[k + len(index) - ss2_length]
    
    return GX1bias, GY1bias, GX2bias, GZ2bias


def Tcomp_gyro_pick_average(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time, f_s_acc, timeabs, GX1count, \
                            GY1count, GX2count, GZ2count, TG1count, TG2count, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2):
    
    GX1pick_TCav = np.zeros(len(meas_point))
    GY1pick_TCav = np.zeros(len(meas_point))
    GX2pick_TCav = np.zeros(len(meas_point))
    GZ2pick_TCav = np.zeros(len(meas_point))
          
    startindex = (starttime_Tcomp-standstill_time) * f_s_acc
    stopindex = (stoptime_Tcomp+standstill_time) * f_s_acc
    standstill_samples = standstill_time * f_s_acc 
    
    indexforTC = index[startindex:stopindex]
    GX1countforTC = GX1count[startindex:stopindex]
    GY1countforTC = GY1count[startindex:stopindex]
    GX2countforTC = GX2count[startindex:stopindex]
    GZ2countforTC = GZ2count[startindex:stopindex]
    TG1countforTC = TG1count[startindex:stopindex]
    TG2countforTC = TG2count[startindex:stopindex]
    
    GX1bias, GY1bias, GX2bias, GZ2bias = gyro_bias_comp(indexforTC, GX1countforTC, GY1countforTC, GX2countforTC, GZ2countforTC, \
                TG1countforTC, TG2countforTC, standstill_samples, standstill_samples, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2) 
   
#    print GX1bias[5000]
#    print GY1bias[5000]
#    print GX2bias[5000]
    
    GX1countTC = GX1count
    GY1countTC = GY1count
    GX2countTC = GX2count
    GZ2countTC = GZ2count
    
    GX1countTC[startindex:stopindex] = GX1countTC[startindex:stopindex] - GX1bias
    GY1countTC[startindex:stopindex] = GY1countTC[startindex:stopindex] - GY1bias
    GX2countTC[startindex:stopindex] = GX2countTC[startindex:stopindex] - GX2bias
    GZ2countTC[startindex:stopindex] = GZ2countTC[startindex:stopindex] - GZ2bias
    
#    print GX1countTC
#    print GY1countTC
    
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
      
        GX1pick_TCav[n] = np.mean(GX1countTC[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY1pick_TCav[n] = np.mean(GY1countTC[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GX2pick_TCav[n] = np.mean(GX2countTC[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ2pick_TCav[n] = np.mean(GZ2countTC[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TG1pick_av[n] = np.mean(TG1count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TG2pick_av[n] = np.mean(TG2count[(int(k-av_points/2))-1:(int(k+av_points/2))])
    
#    print GX1pick_TCav
#    print GY1pick_TCav
    
    return GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav



def gsens_calc(AX, AY, AZ, GX1, GY1, GX2, GZ2, TG1pick_av, TG2pick_av, figno1, figno2, figno3, figno4, plotfigs):
    
    if plotfigs ==True:
        plt.close(figno1)
        plt.close(figno2)
        plt.close(figno3)
        plt.close(figno4)
    
    
    rows = len(AX)
    M = np.empty((rows,3))
    MT = np.empty((3, rows))
    MTM = np.empty((3, 3))
    invMTM = np.empty((3, 3))
    invMTM_MT = np.empty((3, rows))
    F = np.empty((rows,4))
    OT = np.empty((3, 4))
    
    M[:,0] = AX
    M[:,1] = AY
    M[:,2] = AZ
    
    F[:,0] = GX1
    F[:,1] = GY1
    F[:,2] = GX2
    F[:,3] = GZ2
 
    MT = np.transpose(M)
    MTM = np.matmul(MT, M)
    invMTM = np.linalg.inv(MTM)
    invMTM_MT = np.matmul(invMTM, MT)
    OT = np.matmul(invMTM_MT, F)
    GX1factors = OT[:, 0]
    GY1factors = OT[:, 1]
    GX2factors = OT[:, 2]
    GZ2factors = OT[:, 3]
    
    GX1gsens = np.zeros(rows)
    GY1gsens = np.zeros(rows)
    GX2gsens = np.zeros(rows)
    GZ2gsens = np.zeros(rows)   
    
    TG1 = np.average(TG1pick_av)
    TG2 = np.average(TG2pick_av)
    
    for k in range(rows):
        GX1gsens[k] = GX1factors[0]*AX[k] + GX1factors[1]*AY[k] + GX1factors[2]*AZ[k]
        GY1gsens[k] = GY1factors[0]*AX[k] + GY1factors[1]*AY[k] + GY1factors[2]*AZ[k]
        GX2gsens[k] = GX2factors[0]*AX[k] + GX2factors[1]*AY[k] + GX2factors[2]*AZ[k]
        GZ2gsens[k] = GZ2factors[0]*AX[k] + GZ2factors[1]*AY[k] + GZ2factors[2]*AZ[k]
    
    abstempher = int(abstemp)
    
    if plotfigs == True:
        plt.figure(figno1, figsize=(12,9))
        plt.plot(meas_point, GX1, label = 'GX1')
  
        plt.plot(meas_point, GX1gsens, label = 'GX1poly')
        plt.legend(loc = 'best')
        plt.legend(loc = 'best')
        plt.grid()
        plt.savefig(toolSN + ' prim '+ str(abstempher)+ 'C plot' + str(figno1) + ' GX1 gsens.png')
    #    plt.show()    
        
        plt.figure(figno2, figsize=(12,9))
        plt.plot(meas_point, GY1, label = 'GY1')

        plt.plot(meas_point, GY1gsens, label = 'GY1poly')
        plt.legend(loc = 'best')
        plt.legend(loc = 'best')
        plt.grid()
        plt.savefig(toolSN + ' prim '+ str(abstempher)+ 'C plot' + str(figno2) + ' GY1 gsens.png')
    #    plt.show() 
        
        plt.figure(figno3, figsize=(12,9))
        plt.plot(meas_point, GX2, label = 'GX2')

        plt.plot(meas_point, GX2gsens, label = 'GX2poly')
        plt.legend(loc = 'best')
        plt.legend(loc = 'best')
        plt.grid()
        plt.savefig(toolSN + ' prim '+ str(abstempher)+ 'C plot' + str(figno3) + ' GX2 gsens.png')
    #    plt.show() 
        
        plt.figure(figno4, figsize=(12,9))
        plt.plot(meas_point, GZ2, label = 'GZ2')
 
        plt.plot(meas_point, GZ2gsens, label = 'GZ2poly')
        plt.legend(loc = 'best')
        plt.legend(loc = 'best')
        plt.grid()
        plt.savefig(toolSN +' prim '+ str(abstempher)+ 'C plot' + str(figno4) + ' GZ2 gsens.png')
    #    plt.show() 
       
    return GX1factors, GY1factors, GX2factors, GZ2factors, TG1, TG2

def save_file_gsens(filename, sheetname, gsensGX1, gsensGY1, gsensGX2, gsensGZ2, TG1, TG2, abstemp):
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet(sheetname)
    
    worksheet.write(0,0,'GX1 coeffs')
    worksheet.write(0,1,'GY1 coeffs')
    worksheet.write(0,2,'GX2 coeffs')
    worksheet.write(0,3,'GZ2 coeffs')
    worksheet.write(0,4,'TG1 count')
    worksheet.write(0,5,'TG2 count')
    worksheet.write(0,6, 'Cal temp (degC)')
    worksheet.write(1,4, TG1)
    worksheet.write(1,5, TG2)
    worksheet.write(1,6, abstemp)
    
    
    for k in range(3):
        worksheet.write(k+1, 0, gsensGX1[k])
        worksheet.write(k+1, 1, gsensGY1[k]) 
        worksheet.write(k+1, 2, gsensGX2[k]) 
        worksheet.write(k+1, 3, gsensGZ2[k])
    
    workbook.close()     
    return

def read_file_gsens(filename, sheetname):
    
    # read data file 
    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_name(sheetname)
       
    # extract cal data
    
    col = worksheet.col_values(0)
    GX1_g_coeff = col[1:4]
    GX1_g_coeff = np.array(GX1_g_coeff)
    
    col = worksheet.col_values(1)
    GY1_g_coeff = col[1:4]
    GY1_g_coeff = np.array(GY1_g_coeff)

    col = worksheet.col_values(2)
    GX2_g_coeff = col[1:4]
    GX2_g_coeff = np.array(GX2_g_coeff)
    
    col = worksheet.col_values(3)
    GZ2_g_coeff = col[1:4]
    GZ2_g_coeff = np.array(GZ2_g_coeff)
    
    return GX1_g_coeff, GY1_g_coeff, GX2_g_coeff, GZ2_g_coeff


def calculate_gyro_gbias(GX1count, GY1count, GX2count, GZ2count, GX1factors, GY1factors, GX2factors, GZ2factors, AXcal, AYcal, AZcal):
    
    GX1g_bias = np.zeros(len(GX1count))
    GY1g_bias = np.zeros(len(GX1count))
    GX2g_bias = np.zeros(len(GX1count))
    GZ2g_bias = np.zeros(len(GX1count))
    
    for k in range(len(GX1count)):
            
        GX1g_bias[k] = GX1factors[0]*AXcal[k] + GX1factors[1]*AYcal[k] + GX1factors[2]*AZcal[k]
        GY1g_bias[k] = GY1factors[0]*AXcal[k] + GY1factors[1]*AYcal[k] + GY1factors[2]*AZcal[k]
        GX2g_bias[k] = GX2factors[0]*AXcal[k] + GX2factors[1]*AYcal[k] + GX2factors[2]*AZcal[k]
        GZ2g_bias[k] = GZ2factors[0]*AXcal[k] + GZ2factors[1]*AYcal[k] + GZ2factors[2]*AZcal[k]
    
    return GX1g_bias, GY1g_bias, GX2g_bias, GZ2g_bias
    
    
    

def plot_figures_ps4(time, GX1count, GY1count, GX2count, GZ2count, meas_point, GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav):
    
    #plt.close(141)
    #plt.close(142)
    
#    plt.figure(141)
#    plt.plot(time, GX1count, label = 'GX1countTC')
#    plt.hold(True)
#    plt.plot(time, GY1count, label = 'GY1countTC')
#    plt.plot(time, GX2count, label = 'GX2countTC')
#    plt.plot(time, GZ2count, label = 'GZ2countTC')
#    plt.legend(loc = 'best')
#    plt.legend(loc = 'best')
#    plt.grid()
#    plt.show()
    
#    plt.figure(142)
#    plt.plot(meas_point, GX1pick_TCav, label = 'GX1avTC')
#    plt.hold(True)
#    plt.plot(meas_point, GY1pick_TCav, label = 'GY1avTC')
#    plt.plot(meas_point, GX2pick_TCav, label = 'GX2avTC')
#    plt.plot(meas_point, GZ2pick_TCav, label = 'GZ2avTC')
#    plt.legend(loc = 'best')
#    plt.legend(loc = 'best')
#    plt.grid()
#    plt.show()
    return

def read_file_gyroTcoeff(filename, sheetname):
    
    # read data file 
    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_name(sheetname)
    
    TcoeffGX1 = np.zeros(6)
    TcoeffGY1 = np.zeros(6)
    TcoeffGX2 = np.zeros(6)
    TcoeffGZ2 = np.zeros(6)
    
    offset = 1
    # extract index
    col = worksheet.col_values(0)
    TcoeffGX1 = col[offset:]
    TcoeffGX1 = np.array(TcoeffGX1)
    
    col = worksheet.col_values(1)
    TcoeffGY1 = col[offset:]
    TcoeffGY1 = np.array(TcoeffGY1)
    
    col = worksheet.col_values(2)
    TcoeffGX2 = col[offset:]
    TcoeffGX2 = np.array(TcoeffGX2)
    
    col = worksheet.col_values(3)
    TcoeffGZ2 = col[offset:]
    TcoeffGZ2 = np.array(TcoeffGZ2)
    
    return TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2


def read_rot_timefile(file_path, sheetname_data):

    # read data file 
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheetname_data)
       
    # extract index data
    offset = 1
    col = worksheet.col_values(0)
    meas_point = col[offset:]
    meas_point = np.array(meas_point)
    
    # extract time gap data
    col = worksheet.col_values(1)
    timegap = col[offset:]
    timegap = np.array(timegap)
    
    # extract absolute timing data
    col = worksheet.col_values(2)
    timeabs = col[offset:]
    timeabs = np.array(timeabs)   
    
    # extract rotation rate data
    col = worksheet.col_values(3)
    rotrate = col[offset:]
    rotrate = np.array(rotrate)   
       
        
    return meas_point, timegap, timeabs, rotrate


def gyro_rot_pick(meas_point, timeabs, rotrate, av_time_ss, av_time_rot, margin_ss , margin_rot, f_s, \
                  accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, rot):
    
    
    if rot == 1:
        plt.close(151)
        figno = 151
#        plt.close(154)
#        figno2 = 154
    if rot == 2:
        plt.close(152)
        figno = 152
#        plt.close(155)
#        figno2 = 155
    if rot == 3:
        plt.close(153)
        figno = 153
#        plt.close(156)
#        figno2 = 156
    
    
    
    GX1TP = np.zeros(len(accXcount))
    GY1TP = np.zeros(len(accXcount))
    GX2TP = np.zeros(len(accXcount))
    GZ2TP = np.zeros(len(accXcount))
    
    time = np.zeros(len(accXcount))
    
    GX1av = np.zeros(len(meas_point)-1)
    GY1av = np.zeros(len(meas_point)-1)
    GX2av = np.zeros(len(meas_point)-1)
    GZ2av = np.zeros(len(meas_point)-1)
    TG1av = np.zeros(len(meas_point)-1)
    TG2av = np.zeros(len(meas_point)-1)
    
    
    # calibrate accelerometer data
    xoff, yoff, zoff, xfactor, yfactor, zfactor, x3rd, y3rd, z3rd = read_file_cal_constants(acc_cal_file, acc_cal_sheet)
    AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(accXcount, accYcount, accZcount, xfactor, yfactor, zfactor)
    
    # read g sensitivity constants from file
    GX1_g_coeff, GY1_g_coeff, GX2_g_coeff, GZ2_g_coeff = read_file_gsens(gyrogsenscal_file, gyrogsenscal_sheet)
    
    # calculate gyro g bias
    GX1_gbias, GY1_gbias, GX2_gbias, GZ2_gbias = calculate_gyro_gbias(GX1count, GY1count, GX2count, GZ2count, \
                                        GX1_g_coeff, GY1_g_coeff, GX2_g_coeff, GZ2_g_coeff, AXcal_lin, AYcal_lin, AZcal_lin)
    
    # subtract bias 
    GX1count = GX1count - GX1_gbias
    GY1count = GY1count - GY1_gbias
    GX2count = GX2count - GX2_gbias
    GZ2count = GZ2count - GZ2_gbias
    
    # read temperature coefficients from file 
    TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2 = read_file_gyroTcoeff(gyroTcal_file, gyroTcal_sheet)
    
    for k in range(len(accXcount)):
        
        GX1TP[k] = TcoeffGX1[5] + TcoeffGX1[4]*TG1count[k] + TcoeffGX1[3]*TG1count[k]**2 + TcoeffGX1[2]*TG1count[k]**3  \
                    + TcoeffGX1[1]*TG1count[k]**4 + TcoeffGX1[0]*TG1count[k]**5
        GY1TP[k] = TcoeffGY1[5] + TcoeffGY1[4]*TG1count[k] + TcoeffGY1[3]*TG1count[k]**2 + TcoeffGY1[2]*TG1count[k]**3  \
                    + TcoeffGY1[1]*TG1count[k]**4 + TcoeffGY1[0]*TG1count[k]**5
        GX2TP[k] = TcoeffGX2[5] + TcoeffGX2[4]*TG2count[k] + TcoeffGX2[3]*TG2count[k]**2 + TcoeffGX2[2]*TG2count[k]**3  \
                    + TcoeffGX2[1]*TG2count[k]**4 + TcoeffGX2[0]*TG2count[k]**5
        GZ2TP[k] = TcoeffGZ2[5] + TcoeffGZ2[4]*TG2count[k] + TcoeffGZ2[3]*TG2count[k]**2 + TcoeffGZ2[2]*TG2count[k]**3  \
                    + TcoeffGZ2[1]*TG2count[k]**4 + TcoeffGZ2[0]*TG2count[k]**5
    
    
    GX1TP0 = np.average(GX1count[int(f_s*(timeabs[1]-15)):int(f_s*(timeabs[1]-5))])
    GY1TP0 = np.average(GY1count[int(f_s*(timeabs[1]-15)):int(f_s*(timeabs[1]-5))])
    GX2TP0 = np.average(GX2count[int(f_s*(timeabs[1]-15)):int(f_s*(timeabs[1]-5))])
    GZ2TP0 = np.average(GZ2count[int(f_s*(timeabs[1]-15)):int(f_s*(timeabs[1]-5))])
    
    for k in range(len(accXcount)):
        GX1count[k] = GX1count[k] -(GX1TP[k]-GX1TP0)
        GY1count[k] = GY1count[k] -(GY1TP[k]-GY1TP0)
        GX2count[k] = GX2count[k] -(GX2TP[k]-GX2TP0)
        GZ2count[k] = GZ2count[k] -(GZ2TP[k]-GZ2TP0)
    
    for k in range(len(meas_point)-1):
        
        if abs(rotrate[k]) < 0.01:  # standstill
            GX1av[k] = np.average(GX1count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])
            GY1av[k] = np.average(GY1count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])
            GX2av[k] = np.average(GX2count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])
            GZ2av[k] = np.average(GZ2count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])
            TG1av[k] = np.average(TG1count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])
            TG2av[k] = np.average(TG2count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])
        
        if abs(rotrate[k]) > 0.01:  # rotation
            GX1av[k] = np.average(GX1count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])
            GY1av[k] = np.average(GY1count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])
            GX2av[k] = np.average(GX2count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])
            GZ2av[k] = np.average(GZ2count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])
            TG1av[k] = np.average(TG1count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])
            TG2av[k] = np.average(TG2count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])  
   
    
    for n in range(len(GX1count)):
        time[n] = n/f_s
    
    plotstartindex = int(f_s*(timeabs[1] - 30))
    plotstopindex = int(f_s*(timeabs[5] + 30))
    
    plt.figure(figno, figsize=(12,9))
    plt.plot(time[plotstartindex:plotstopindex], GX1count[plotstartindex:plotstopindex], 'C0', label = 'GX1count')

    plt.plot(time[plotstartindex:plotstopindex], GY1count[plotstartindex:plotstopindex], 'C1', label = 'GY1count')
    plt.plot(time[plotstartindex:plotstopindex], GX2count[plotstartindex:plotstopindex], 'C2', label = 'GX2count')
    plt.plot(time[plotstartindex:plotstopindex], GZ2count[plotstartindex:plotstopindex], 'C3', label = 'GZ2count')
    

    for k in range(len(meas_point)-1):
        
        if abs(rotrate[k]) < 0.01:  # standstill
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], \
                GX1count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], 'C4')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], \
                GY1count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], 'C5')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], \
                GX2count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], 'C6')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], \
                GZ2count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], 'C7')
            
            
            
            
        if abs(rotrate[k]) > 0.01:  # rotation
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], \
                GX1count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], 'C4')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], \
                GY1count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], 'C5')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], \
                GX2count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], 'C6')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], \
                GZ2count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], 'C7')
 
    plt.ylim((30000, 36000)) 
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    abstempher = int(abstemp)
    if rot == 1:
        plt.savefig(toolSN + ' prim '+ str(abstempher) + 'C plot' + str(figno) + ' gyro Xrot.png')
    if rot == 2:
        plt.savefig(toolSN + ' prim '+ str(abstempher) + 'C plot' + str(figno) + ' gyro Yrot.png')
    if rot == 3:
        plt.savefig(toolSN + ' prim '+ str(abstempher) + 'C plot' + str(figno) + ' gyro Zrot.png')
#    plt.show()
    
#    plt.figure(figno2)
#    plt.plot(GX1av, 'C0o', label = 'GX1av')
#    plt.hold(True)
#    plt.plot(GY1av, 'C1o', label = 'GY1av')
#    plt.plot(GX2av, 'C2o', label = 'GX2av')
#    plt.plot(GZ2av, 'C3o', label = 'GZ2av')
#    plt.legend(loc = 'best')
#    plt.legend(loc = 'best')
#    plt.grid()
#    plt.show()
#    
    return GX1av, GY1av, GX2av, GZ2av, TG1av, TG2av


def calc_gyro_ortomatrix(GX1avXr, GY1avXr, GX2avXr, GZ2avXr, TG1avXr, TG2avXr, rotrateXr, GX1avYr, GY1avYr, GX2avYr, \
                         GZ2avYr, TG1avYr, TG2avYr, rotrateYr, GX1avZr, GY1avZr, GX2avZr, GZ2avZr, TG1avZr, TG2avZr, rotrateZr):
    
    xelements = 0
    yelements = 0
    zelements = 0
    
    xindex = np.zeros(len(GX1avXr))
    yindex = np.zeros(len(GX1avYr))
    zindex = np.zeros(len(GX1avZr))
    
    n=0
    for k in range(len(GX1avXr)):
        if abs(rotrateXr[k]) > 0.1:
            xelements = xelements + 1
            xindex[n] = k
            n = n+1
    n=0        
    for k in range(len(GY1avYr)):
        if abs(rotrateYr[k]) > 0.1:
            yelements = yelements + 1
            yindex[n] = k
            n = n+1
    
    n=0         
    for k in range(len(GZ2avZr)):
        if abs(rotrateZr[k]) > 0.1:
            zelements = zelements + 1
            zindex[n] = k
            n = n+1
    
#    print xindex
    rows = xelements + yelements + zelements 
    
    for GXoption in (0,1,2):
        M = np.empty((rows, 3))
        F = np.empty((rows, 3))
        MT = np.empty((3, rows))
        MTM = np.empty((3, 3))
        invMTM = np.empty((3, 3))
        invMTM_MT = np.empty((3, rows))
        OT = np.empty((3, 3))
        
        for k in range(xelements):                                                  # X rotation                   
            ind = int(xindex[k])   
            if GXoption == 0:
                M[k,0] = (GX1avXr[ind]+GX2avXr[ind])/2 -((GX1avXr[ind-1]+GX2avXr[ind-1])+(GX1avXr[ind+1]+GX2avXr[ind+1]))/4
            if GXoption == 1:
                M[k,0] = GX1avXr[ind] - (GX1avXr[ind-1] + GX1avXr[ind+1])/2
            if GXoption == 2:
                M[k,0] = GX2avXr[ind] - (GX2avXr[ind-1] + GX2avXr[ind+1])/2
            M[k,1] = GY1avXr[ind] - (GY1avXr[ind-1] + GY1avXr[ind+1])/2
            M[k,2] = GZ2avXr[ind] - (GZ2avXr[ind-1] + GZ2avXr[ind+1])/2
            
            F[k,0] = rotrateXr[ind]
            F[k,1] = 0
            F[k,2] = 0
        
        
                
        for k in range(yelements): 
            ind = int(yindex[k])                                          # Y rotation                        
            if GXoption == 0:
                M[k+xelements, 0] = (GX1avYr[ind]+GX2avYr[ind])/2 -((GX1avYr[ind-1]+GX2avYr[ind-1])+(GX1avYr[ind+1]+GX2avYr[ind+1]))/4
            if GXoption == 1:
                M[k+xelements, 0] = GX1avYr[ind] - (GX1avYr[ind-1] + GX1avYr[ind+1])/2
            if GXoption == 2:
                M[k+xelements, 0] = GX2avYr[ind] - (GX2avYr[ind-1] + GX2avYr[ind+1])/2
            
            M[k + xelements, 1] = GY1avYr[ind] - (GY1avYr[ind-1] + GY1avYr[ind+1])/2
            M[k + xelements, 2] = GZ2avYr[ind] - (GZ2avYr[ind-1] + GZ2avYr[ind+1])/2
            
            F[k + xelements, 0] = 0
            F[k + xelements, 1] = rotrateYr[ind]
            F[k + xelements, 2] = 0
        
        for k in range(zelements): 
            ind = int(zindex[k])                                          # Z rotation                        
            if GXoption == 0:
                M[k+xelements + yelements, 0] = (GX1avZr[ind]+GX2avZr[ind])/2 -((GX1avZr[ind-1]+GX2avZr[ind-1])+(GX1avZr[ind+1]+GX2avZr[ind+1]))/4
            if GXoption == 1:
                M[k+xelements + yelements, 0] = GX1avZr[ind] - (GX1avZr[ind-1] + GX1avZr[ind+1])/2
            if GXoption == 2:
                M[k+xelements + yelements, 0] = GX2avZr[ind] - (GX2avZr[ind-1] + GX2avZr[ind+1])/2
            
            M[k + xelements + yelements, 1] = GY1avZr[ind] - (GY1avZr[ind-1] + GY1avZr[ind+1])/2
            M[k + xelements + yelements, 2] = GZ2avZr[ind] - (GZ2avZr[ind-1] + GZ2avZr[ind+1])/2
            
            F[k + xelements + yelements, 0] = 0
            F[k + xelements + yelements, 1] = 0
            F[k + xelements + yelements, 2] = rotrateZr[ind]
        
        np.set_printoptions(precision=2, suppress = True)
        np.set_printoptions(formatter={'float': '{: 10.2f}'.format})
        print (M)
        np.set_printoptions(edgeitems=3,infstr='inf', linewidth=75, nanstr='nan', \
                            precision=8, suppress=False, threshold=1000, formatter=None)            # set back to default
        MT = np.transpose(M)
        MTM = np.matmul(MT, M)
        invMTM = np.linalg.inv(MTM)
        invMTM_MT = np.matmul(invMTM, MT)
        OT = np.matmul(invMTM_MT, F)
        if GXoption == 0:
            xfactorsGX0 = OT[:, 0]
            yfactorsGX0 = OT[:, 1]
            zfactorsGX0 = OT[:, 2]
        if GXoption == 1:
            xfactorsGX1 = OT[:, 0]
            yfactorsGX1 = OT[:, 1]
            zfactorsGX1 = OT[:, 2]    
        if GXoption == 2:
            xfactorsGX2 = OT[:, 0]
            yfactorsGX2 = OT[:, 1]
            zfactorsGX2 = OT[:, 2]
            
    TG1 = (np.average(TG1avXr) + np.average(TG1avYr) + np.average(TG1avZr))/3
    TG2 = (np.average(TG2avXr) + np.average(TG2avYr) + np.average(TG2avZr))/3
       
    return xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, TG1, TG2


def save_file_gyro_orto(filename, sheetname, xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, \
                        xfactorsGX2, yfactorsGX2, zfactorsGX2, bestGX, TG1, TG2, abstemp):
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet(sheetname)
    
    worksheet.write(0,0,'bestGX')
    worksheet.write(0,1, bestGX)
        
    worksheet.write(1, 0, 'X coeffs GX0')
    worksheet.write(1, 1, 'Y coeffs GX0')
    worksheet.write(1, 2, 'Z coeffs GX0')
    worksheet.write(1, 3, 'X coeffs GX1')
    worksheet.write(1, 4, 'Y coeffs GX1')
    worksheet.write(1, 5, 'Z coeffs GX1')
    worksheet.write(1, 6, 'X coeffs GX2')
    worksheet.write(1, 7, 'Y coeffs GX2')
    worksheet.write(1, 8, 'Z coeffs GX2')
    worksheet.write(1, 9, 'TG1 count')
    worksheet.write(1, 10, 'TG2 count')
    worksheet.write(1, 11, 'Cal temp (deg C)')
    worksheet.write(2, 9, TG1)
    worksheet.write(2, 10, TG2)
    worksheet.write(2, 11, abstemp)


    for k in range(3):
        worksheet.write(k+2, 0, xfactorsGX0[k])
        worksheet.write(k+2, 1, yfactorsGX0[k]) 
        worksheet.write(k+2, 2, zfactorsGX0[k]) 
        
        worksheet.write(k+2, 3, xfactorsGX1[k])
        worksheet.write(k+2, 4, yfactorsGX1[k]) 
        worksheet.write(k+2, 5, zfactorsGX1[k]) 
        
        worksheet.write(k+2, 6, xfactorsGX2[k])
        worksheet.write(k+2, 7, yfactorsGX2[k]) 
        worksheet.write(k+2, 8, zfactorsGX2[k]) 

    workbook.close() 
    return

def read_file_gyro_orto(filename, sheetname):
    
    # read data file 
    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_name(sheetname)
    
    xfactorsGX0 = np.zeros(3)
    yfactorsGX0 = np.zeros(3)
    zfactorsGX0 = np.zeros(3)
    xfactorsGX1 = np.zeros(3)
    yfactorsGX1 = np.zeros(3)
    zfactorsGX1 = np.zeros(3)
    xfactorsGX2 = np.zeros(3)
    yfactorsGX2 = np.zeros(3)
    zfactorsGX2 = np.zeros(3)
    
    offset = 2
    # extract index
    col = worksheet.col_values(0)
    xfactorsGX0 = col[offset:]
    xfactorsGX0 = np.array(xfactorsGX0)
    
    col = worksheet.col_values(1)
    yfactorsGX0 = col[offset:]
    yfactorsGX0 = np.array(yfactorsGX0)
    
    bestGX = col[0]
    bestGX = int(bestGX)
    
    col = worksheet.col_values(2)
    zfactorsGX0 = col[offset:]
    zfactorsGX0 = np.array(zfactorsGX0)
    
    col = worksheet.col_values(3)
    xfactorsGX1 = col[offset:]
    xfactorsGX1 = np.array(xfactorsGX1)
    
    col = worksheet.col_values(4)
    yfactorsGX1 = col[offset:]
    yfactorsGX1 = np.array(yfactorsGX1)
    
    col = worksheet.col_values(5)
    zfactorsGX1 = col[offset:]
    zfactorsGX1 = np.array(zfactorsGX1)
    
    col = worksheet.col_values(6)
    xfactorsGX2 = col[offset:]
    xfactorsGX2 = np.array(xfactorsGX2)
    
    col = worksheet.col_values(7)
    yfactorsGX2 = col[offset:]
    yfactorsGX2 = np.array(yfactorsGX2)
    
    col = worksheet.col_values(8)
    zfactorsGX2 = col[offset:]
    zfactorsGX2 = np.array(zfactorsGX2)
    
    col = worksheet.col_values(9)
    TG1 = col[2]
    TG1 = float(TG1)
    
    col = worksheet.col_values(10)
    TG2 = col[2]
    TG2 = float(TG2)
    
    col = worksheet.col_values(11)
    abstemp = col[2]
    abstemp = float(abstemp)
    
    return xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, \
            xfactorsGX2, yfactorsGX2, zfactorsGX2, bestGX, TG1, TG2, abstemp
            

# returns unit vectors for each gyro sensor along sensitive axis in tool body coordinate system. Also returns scale
# factor for each gyro sensor in units of (counts/(deg/s))
            
def gyro_body_vec(xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2):
    
    O_GX1 = np.empty((3,3))
    O_GX2 = np.empty((3,3))
    Oinv_GX1 = np.empty((3,3))
    Oinv_GX2 = np.empty((3,3))
    
    GX1point = np.zeros(3)
    GY1point = np.zeros(3)
    GX2point = np.zeros(3)
    GZ2point = np.zeros(3)    
    
    O_GX1[0, :] = xfactorsGX1
    O_GX1[1, :] = yfactorsGX1
    O_GX1[2, :] = zfactorsGX1
    
    O_GX2[0, :] = xfactorsGX2
    O_GX2[1, :] = yfactorsGX2
    O_GX2[2, :] = zfactorsGX2
    
    Oinv_GX1 = np.linalg.inv(O_GX1)
    Oinv_GX2 = np.linalg.inv(O_GX2)
    
    Xrot = [1,0,0]
    Yrot = [0,1,0]
    Zrot = [0,0,1]
    
    GX1Xrot = np.matmul(Oinv_GX1, Xrot)
    GX1Yrot = np.matmul(Oinv_GX1, Yrot)
    GX1Zrot = np.matmul(Oinv_GX1, Zrot)
       
    GX2Xrot = np.matmul(Oinv_GX2, Xrot)
    GX2Yrot = np.matmul(Oinv_GX2, Yrot)
    GX2Zrot = np.matmul(Oinv_GX2, Zrot)
          
    GX1point = [GX1Xrot[0], GX1Yrot[0], GX1Zrot[0]]
    GX1point = np.array(GX1point)
    GX1scale = np.sqrt((GX1point[0])**2 + (GX1point[1])**2 + (GX1point[2])**2)
    GX1point = GX1point/GX1scale
    
    GY1point = [GX1Xrot[1], GX1Yrot[1], GX1Zrot[1]]
    GY1point = np.array(GY1point)
    GY1scale = np.sqrt((GY1point[0])**2 + (GY1point[1])**2 + (GY1point[2])**2)
    GY1point = GY1point/GY1scale

    GX2point = [GX2Xrot[0], GX2Yrot[0], GX2Zrot[0]]
    GX2point = np.array(GX2point)
    GX2scale = np.sqrt((GX2point[0])**2 + (GX2point[1])**2 + (GX2point[2])**2)
    GX2point = GX2point/GX2scale
    
    GZ2point = [GX1Xrot[2], GX1Yrot[2], GX1Zrot[2]]
    GZ2point = np.array(GZ2point)
    GZ2scale = np.sqrt((GZ2point[0])**2 + (GZ2point[1])**2 + (GZ2point[2])**2)
    GZ2point = GZ2point/GZ2scale
        
    return GX1point, GY1point, GX2point, GZ2point, GX1scale, GY1scale, GX2scale, GZ2scale


# shall return earth rotation vector in tool body coordinate system
def earth_rot_vec(latitude, azimuth, inclination, toolface):
    
    er_point_ref = np.zeros(3)          # earth rotation pointing vector in NWU system
    er_point_tool = np.zeros(3)         # earth rotation pointing vector in tool body system
    
    # in local reference system (NWU), earth rotation vector is given by
    er_point_ref = [np.cos(np.radians(latitude)), 0, np.sin(np.radians(latitude))]
    # then rotate vector from reference system (NWU) to tool body system
    
    Cazi = np.cos(np.radians(-azimuth))            # -sign in order to adapt difference between Devico convetions and standard Euler angles
    Sazi = np.sin(np.radians(-azimuth))
    Cinc = np.cos(np.radians(-inclination))
    Sinc = np.sin(np.radians(-inclination))
    Ctool = np.cos(np.radians(toolface))
    Stool = np.sin(np.radians(toolface))
    
    
    er_point_tool[0] = Cinc*Cazi*er_point_ref[0] + Cinc*Sazi*er_point_ref[1] - Sinc*er_point_ref[2]
    
    er_point_tool[1] = (Stool*Sinc*Cazi - Ctool*Sazi)*er_point_ref[0] + (Stool*Sinc*Sazi + Ctool*Cazi)*er_point_ref[1] \
                        + Cinc*Stool*er_point_ref[2]
                        
    er_point_tool[2] = (Ctool*Sinc*Cazi + Stool*Sazi)*er_point_ref[0] + (Ctool*Sinc*Sazi - Stool*Cazi)*er_point_ref[1] \
                        + Cinc*Ctool*er_point_ref[2]

    return er_point_tool



# returns counter values caused by earth rotation for each individual gyro sensor
    
def er_comp_gyros(GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav, azimuth, inclination, toolface, \
                  GX1point, GY1point, GX2point, GZ2point, GX1scale, GY1scale, GX2scale, GZ2scale):
    
    GX1_er_comp = np.zeros(len(GX1pick_TCav))
    GY1_er_comp = np.zeros(len(GX1pick_TCav))
    GX2_er_comp = np.zeros(len(GX1pick_TCav))
    GZ2_er_comp = np.zeros(len(GX1pick_TCav))
    
    for k in range(len(GX1pick_TCav)):
        earthrot =  earth_rot_vec(latitude, azimuth[k], inclination[k], toolface[k])
        GX1_er_comp[k] = np.dot(earthrot, GX1point) * (earthrate/3600) * GX1scale
        GY1_er_comp[k] = np.dot(earthrot, GY1point) * (earthrate/3600) * GY1scale
        GX2_er_comp[k] = np.dot(earthrot, GX2point) * (earthrate/3600) * GX2scale
        GZ2_er_comp[k] = np.dot(earthrot, GZ2point) * (earthrate/3600) * GZ2scale
#        print k, earthrot
    return GX1_er_comp, GY1_er_comp, GX2_er_comp, GZ2_er_comp




# main program starts here

if sequencemode == 0: 

    if processtep == 1:
        
        #read data from raw data file 
        index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, \
            TG1count, TG2count, TG3count = readrawdata(logfile_acc_gsens, sheetname_acc_gsens)
        
        time = index / f_s_acc
        av_points = av_time * f_s_acc
        av_pointsAC = av_timeAC * f_s_acc
        
        
        #read data from timing file
        meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)

        
        # read data from timing file for control signal     

        if controlseq == True:
            meas_pointAC, timegapAC, timeabsAC, AXRefAC, AYRefAC, AZRefAC = readtimefile(file_path_timefile, sheetname_control)
        
                
 
        # pick correctly timed value from counter sequence. Adjust timing file as necessary   
        AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, GY1pick_av, \
        GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, GX2pick_std, GZ2pick_std, TG1pick_av, \
        TG2pick_av, TApick_av \
        = acc_gsens_pick(meas_point, av_time, timeabs, f_s_acc, accXcount, accYcount, accZcount, TAcount, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count)
        
        if controlseq == True:
            AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, GX1pickAC, GY1pickAC, GX2pickAC, \
                GZ2pickAC, GX1pick_avAC, GY1pick_avAC, GX2pick_avAC, GZ2pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, \
                GX1pick_stdAC, GY1pick_stdAC, GX2pick_stdAC, GZ2pick_stdAC, TG1pick_avAC, TG2pick_avAC, TApick_avAC \
                = acc_control_pick(meas_pointAC, av_timeAC, timeabsAC, f_s_acc, accXcount, accYcount, accZcount, TAcount, \
                                 GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count)


        
        # QC plots for this sequence
        plot_figures_ps1(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                         timeabs, av_points, GX1count, GY1count, GX2count, GZ2count, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, \
                         GY1pick_av, GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, \
                         GX2pick_std, GZ2pick_std, TG1pick_av, TG2pick_av, TApick_av)     
        
        
        if controlseq == True:
            plot_figures_ps1AC(time, accXcount, accYcount, accZcount, AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, meas_pointAC, \
                     timeabsAC, av_pointsAC, GX1count, GY1count, GX2count, GZ2count, GX1pickAC, GY1pickAC, GX2pickAC, GZ2pickAC, GX1pick_avAC, \
                     GY1pick_avAC, GX2pick_avAC, GZ2pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, GX1pick_stdAC, GY1pick_stdAC, \
                     GX2pick_stdAC, GZ2pick_stdAC, TG1pick_avAC, TG2pick_avAC, TApick_avAC)
    
        
        # write selected and averaged data to file
        save_file_ps1(pickfile_write, picksheet_write, meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, \
                      GX2pick_av, GZ2pick_av, TG1pick_av, TG2pick_av, TApick_av, AXRef, AYRef, AZRef)
            
        if controlseq == True:  
            save_file_ps1AC(pickfile_writeAC, picksheet_writeAC, meas_pointAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, TApick_avAC, AXRefAC, AYRefAC, AZRefAC)
        
        
        if autostarttime == True:
            starttime_Tcomp = int(timeabs[0] - 1520)
            print ('starttime_Tcomp = ', starttime_Tcomp)
            
            stoptime_Tcomp = int(timeabs[107] + 100)
            if abstemp == 28:
                stoptime_Tcomp = int(timeabs[107] + 450)
            print ('stoptime_Tcomp = ', stoptime_Tcomp)
        
    
    if processtep == 2:
        
        # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
        meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef \
        = read_file_ps1(pickfile_read , picksheet_read)
        
        # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
        xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)    
        
        # 3rd order accelerometer calibration constants
        xfactors_3rd, yfactors_3rd, zfactors_3rd, T_3rd = acc_cal_3rd_order(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
       
        # calibrate picked accelerometer data, classical method
        AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
        
        # calibrate picked accelerometer data, 3rd order method
        AXcal_3rd, AYcal_3rd, AZcal_3rd = calibrate_acc_3rd_order(AXpick_av, AYpick_av, AZpick_av, xfactors_3rd, yfactors_3rd, zfactors_3rd)
        
        # calculate and plot QC parameters for the accelerometer calibration
        acc_calibrationQC(AXcal_lin, AYcal_lin, AZcal_lin, AXcal_3rd, AYcal_3rd, AZcal_3rd, AXRef, AYRef, AZRef)
        
        # save calibration constants to file
        save_file_cal_constants(acc_cal_file, acc_cal_sheet, xoffset, yoffset, zoffset, xfactors_lin, yfactors_lin, zfactors_lin, \
                      xfactors_3rd, yfactors_3rd, zfactors_3rd, T_lin, abstemp)
        
        
        if controlseq == True: 
            # read control signals
            meas_pointAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, TApick_avAC, AXRefAC, AYRefAC, AZRefAC = read_file_ps1AC(pickfile_readAC, picksheet_readAC)
            # calibrate picked accelerometer data, classical method
            AXcal_linAC, AYcal_linAC, AZcal_linAC = calibrate_acc_classical(AXpick_avAC, AYpick_avAC, AZpick_avAC, xfactors_lin, yfactors_lin, zfactors_lin)
            # plot control plots
            acc_calibration_control(meas_pointAC, AXcal_linAC, AYcal_linAC, AZcal_linAC, AXRefAC, AYRefAC, AZRefAC)

        
        
    
    
    if processtep == 3:
       
        #read data from raw data file - needed for temperature compensation
        index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, \
        TG1count, TG2count, TG3count = readrawdata(tempfile_read, tempsheet_read)
     
        # calculate 5th order polynominal
        TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2, TcoeffGX1lin, TcoeffGY1lin, TcoeffGX2lin, TcoeffGZ2lin, TG1min, TG1max, TG2min, TG2max = \
            find_gyro_temp_poly(index, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, f_s_temp)
        
        save_file_gyroT_constants(gyroTcal_file, gyroTcal_sheet, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2, \
                                  TcoeffGX1lin, TcoeffGY1lin, TcoeffGX2lin, TcoeffGZ2lin,TG1min, TG1max, TG2min, TG2max)

        
    
    if processtep == 4:
        
        # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
        meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef \
            = read_file_ps1(pickfile_read, picksheet_read)
        
        # read data from raw data file - needed for temperature compensation
        index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, \
            TG1count, TG2count, TG3count = readrawdata(logfile_acc_gsens, sheetname_acc_gsens)
        
        # read data from timing file
        meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
        
        # read temperature coefficients from file 
        TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2 = read_file_gyroTcoeff(gyroTcal_file, gyroTcal_sheet)
        
        time = index / f_s_acc
        av_points = av_time * f_s_acc
        
        
        
        # temperature compensate gyro counters and pick same time instants as before
        GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav = Tcomp_gyro_pick_average(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
            f_s_acc, timeabs, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2)
        

        # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
        xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
        
        # calibrate picked accelerometer data, classical method
        AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
        
        # calculate gravity sensitivity, ignoring earth rotation. Use in gyro scale factor calibration. Plots 
        GX1gsens, GY1gsens, GX2gsens, GZ2gsens , TG1, TG2 = gsens_calc(AXcal_lin, AYcal_lin, AZcal_lin, GX1pick_TCav, GY1pick_TCav, \
                                                GX2pick_TCav, GZ2pick_TCav, TG1pick_av, TG2pick_av, 131, 132, 133, 134, plotfigs1)        
        
        # plot QC data, 
   
        plot_figures_ps4(time, GX1count, GY1count, GX2count, GZ2count, meas_point, GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav)
            

           
        save_file_gsens(gyrogsenscal_file, gyrogsenscal_sheet, GX1gsens, GY1gsens, GX2gsens, GZ2gsens, TG1, TG2, abstemp)
        
    
    if processtep == 5:
        
        # read data file for X rotation
        indexXr, accXcountXr, accYcountXr, accZcountXr, GX1countXr, GY1countXr, GX2countXr, GZ2countXr, GX3countXr, GY3countXr, GZ3countXr, \
        TAcountXr, TG1countXr, TG2countXr, TG3countXr = readrawdata(Xrotfile, Xrotsheet)
        
        meas_pointXr, timegapXr, timeabsXr, rotrateXr = read_rot_timefile(rot_timefile, Xrot_timesheet)
        
        rot = 1 # X rotation
        
        # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
       
        GX1avXr, GY1avXr, GX2avXr, GZ2avXr, TG1avXr, TG2avXr = gyro_rot_pick(meas_pointXr, timeabsXr, rotrateXr, av_time_standstill, av_time_rot, \
                            margin_standstill ,margin_rotation, f_s_rot, accXcountXr, accYcountXr, accZcountXr, GX1countXr, \
                            GY1countXr, GX2countXr, GZ2countXr, TG1countXr, TG2countXr, rot)
        
        
       
        # read data file for Y rotation
        indexYr, accXcountYr, accYcountYr, accZcountYr, GX1countYr, GY1countYr, GX2countYr, GZ2countYr, GX3countYr, GY3countYr, GZ3countYr, \
        TAcountYr, TG1countYr, TG2countYr, TG3countYr  = readrawdata(Yrotfile, Yrotsheet)
        
        meas_pointYr, timegapYr, timeabsYr, rotrateYr = read_rot_timefile(rot_timefile, Yrot_timesheet)
        
        rot = 2 # Y rotation
        
        # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
        
   
        GX1avYr, GY1avYr, GX2avYr, GZ2avYr, TG1avYr, TG2avYr = gyro_rot_pick(meas_pointYr, timeabsYr, rotrateYr, av_time_standstill, av_time_rot, \
                                margin_standstill ,margin_rotation, f_s_rot, accXcountYr, accYcountYr, accZcountYr, GX1countYr, \
                                GY1countYr, GX2countYr, GZ2countYr, TG1countYr, TG2countYr, rot)
        

        
        # read data file for Z rotation
        indexZr, accXcountZr, accYcountZr, accZcountZr, GX1countZr, GY1countZr, GX2countZr, GZ2countZr, GX3countZr, GY3countZr, GZ3countZr, \
        TAcountZr, TG1countZr, TG2countZr, TG3countZr = readrawdata(Zrotfile, Zrotsheet)
        
        meas_pointZr, timegapZr, timeabsZr, rotrateZr = read_rot_timefile(rot_timefile, Zrot_timesheet)
        
        rot = 3 # Z rotation
        
        # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile 

        GX1avZr, GY1avZr, GX2avZr, GZ2avZr, TG1avZr, TG2avZr = gyro_rot_pick(meas_pointZr, timeabsZr, rotrateZr, av_time_standstill, av_time_rot, \
                                margin_standstill ,margin_rotation, f_s_rot, accXcountZr, accYcountZr, accZcountZr, GX1countZr, \
                                GY1countZr, GX2countZr, GZ2countZr, TG1countZr, TG2countZr, rot)
            

        
        
        # calculate ortomatrix, returns three versions of ortomatrix, one for each GXoption
        
        xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, TG1, TG2 = calc_gyro_ortomatrix(GX1avXr, GY1avXr, GX2avXr, GZ2avXr, TG1avXr, TG2avXr, rotrateXr, GX1avYr, GY1avYr, GX2avYr, \
                                  GZ2avYr, TG1avYr, TG2avYr, rotrateYr, GX1avZr, GY1avZr, GX2avZr, GZ2avZr, TG1avZr, TG2avZr, rotrateZr)
    
        # write gyro ortomatrix to file. Write all three versions to same file (GX0 = average of GX1 and GX2 counts, GX1 count and GX2 counts)
        
        save_file_gyro_orto(gyro_orto_file, gyro_orto_sheet, xfactorsGX0, yfactorsGX0, zfactorsGX0, \
                            xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, bestGX, TG1, TG2, abstemp)
        
    
    if processtep == 6:
        
        # read gyro scale factor matrix
        xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, \
                bestGX, TG1, TG2, abstemp = read_file_gyro_orto(gyro_orto_file, gyro_orto_sheet)
        
        # find gyro sensors unit pointing vectors (along sensitive axis) from gyro orto matrix
        GX1point, GY1point, GX2point, GZ2point, GX1scale, GY1scale, GX2scale, GZ2scale = \
                    gyro_body_vec(xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2)
        
        # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
        meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef \
            = read_file_ps1(pickfile_read, picksheet_read)
        
        # read data from raw data file - needed for temperature compensation
        index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, TG1count, TG2count, TG3count \
            = readrawdata(logfile_acc_gsens, sheetname_acc_gsens)
        
        # read data from timing file
        meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
        
        # read temperature coefficients from file 
        TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2 = read_file_gyroTcoeff(gyroTcal_file, gyroTcal_sheet)
        
        time = index / f_s_acc
        av_points = av_time * f_s_acc
        
        # temperature compensate gyro counters and pick same time instants as before
     
        GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav = Tcomp_gyro_pick_average(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                f_s_acc, timeabs, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2)
        
  
        # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
        xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
        
        # calibrate picked accelerometer data, classical method
        AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
        
        # read data from angle file (same as timing file)
        azimuth, inclination, toolface = readanglefile(file_path_timefile, sheetname_time)
        
        # calculate gyro counter values resulting from earth rotation
        GX1_er_comp, GY1_er_comp, GX2_er_comp, GZ2_er_comp = er_comp_gyros(GX1pick_TCav, GY1pick_TCav, GX2pick_TCav,\
            GZ2pick_TCav, azimuth, inclination, toolface, GX1point, GY1point, GX2point, GZ2point, GX1scale, GY1scale, GX2scale, GZ2scale)
        
        # adjust gyro counter values for earth rotation
        GX1pick_TCav = GX1pick_TCav - GX1_er_comp
        GY1pick_TCav = GY1pick_TCav - GY1_er_comp
        GX2pick_TCav = GX2pick_TCav - GX2_er_comp
        GZ2pick_TCav = GZ2pick_TCav - GZ2_er_comp
        
        
        # calculate gravity sensitivity, compensating for earth rotation. Use in later data calibration. Plots 
        GX1gsens_erc, GY1gsens_erc, GX2gsens_erc, GZ2gsens_erc , TG1, TG2 = gsens_calc(AXcal_lin, AYcal_lin, AZcal_lin, GX1pick_TCav, \
                                    GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav, TG1pick_av, TG2pick_av, 171, 172, 173, 174, plotfigs2)        
        
        save_file_gsens(gyrogsenscal_erc_file, gyrogsenscal_erc_sheet, GX1gsens_erc, GY1gsens_erc, GX2gsens_erc, GZ2gsens_erc, TG1, TG2, abstemp)



if sequencemode ==1:
    
    processtep = 1
            
    if processtep == 1:
        #read data from raw data file 
        index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, \
            TG1count, TG2count, TG3count = readrawdata(logfile_acc_gsens, sheetname_acc_gsens)
        
        time = index / f_s_acc
        av_points = av_time * f_s_acc
        av_pointsAC = av_timeAC * f_s_acc
        
        
        #read data from timing file
        meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)

        
        # read data from timing file for control signal     

        if controlseq == True:
            meas_pointAC, timegapAC, timeabsAC, AXRefAC, AYRefAC, AZRefAC = readtimefile(file_path_timefile, sheetname_control)
        
                
 
        # pick correctly timed value from counter sequence. Adjust timing file as necessary   
        AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, GY1pick_av, \
        GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, GX2pick_std, GZ2pick_std, TG1pick_av, \
        TG2pick_av, TApick_av \
        = acc_gsens_pick(meas_point, av_time, timeabs, f_s_acc, accXcount, accYcount, accZcount, TAcount, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count)
        
        if controlseq == True:
            AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, GX1pickAC, GY1pickAC, GX2pickAC, \
                GZ2pickAC, GX1pick_avAC, GY1pick_avAC, GX2pick_avAC, GZ2pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, \
                GX1pick_stdAC, GY1pick_stdAC, GX2pick_stdAC, GZ2pick_stdAC, TG1pick_avAC, TG2pick_avAC, TApick_avAC \
                = acc_control_pick(meas_pointAC, av_timeAC, timeabsAC, f_s_acc, accXcount, accYcount, accZcount, TAcount, \
                                 GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count)


        
        # QC plots for this sequence
        plot_figures_ps1(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                         timeabs, av_points, GX1count, GY1count, GX2count, GZ2count, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, \
                         GY1pick_av, GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, \
                         GX2pick_std, GZ2pick_std, TG1pick_av, TG2pick_av, TApick_av)     
        
        
        if controlseq == True:
            plot_figures_ps1AC(time, accXcount, accYcount, accZcount, AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, meas_pointAC, \
                     timeabsAC, av_pointsAC, GX1count, GY1count, GX2count, GZ2count, GX1pickAC, GY1pickAC, GX2pickAC, GZ2pickAC, GX1pick_avAC, \
                     GY1pick_avAC, GX2pick_avAC, GZ2pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, GX1pick_stdAC, GY1pick_stdAC, \
                     GX2pick_stdAC, GZ2pick_stdAC, TG1pick_avAC, TG2pick_avAC, TApick_avAC)
    
        
        # write selected and averaged data to file
        save_file_ps1(pickfile_write, picksheet_write, meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, \
                      GX2pick_av, GZ2pick_av, TG1pick_av, TG2pick_av, TApick_av, AXRef, AYRef, AZRef)
            
        if controlseq == True:  
            save_file_ps1AC(pickfile_writeAC, picksheet_writeAC, meas_pointAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, TApick_avAC, AXRefAC, AYRefAC, AZRefAC)
        
        
        if autostarttime == True:
            starttime_Tcomp = int(timeabs[0] - 1520)
            print ('starttime_Tcomp = ', starttime_Tcomp)
            
            stoptime_Tcomp = int(timeabs[107] + 100)
            if abstemp == 28:
                stoptime_Tcomp = int(timeabs[107] + 450)
            print ('stoptime_Tcomp = ', stoptime_Tcomp)


if sequencemode == 2:            
    
    for processtep in (2,3,4,5,6):    
    
        if processtep == 2:
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef \
            = read_file_ps1(pickfile_read , picksheet_read)
            
            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)    
            
            # 3rd order accelerometer calibration constants
            xfactors_3rd, yfactors_3rd, zfactors_3rd, T_3rd = acc_cal_3rd_order(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
           
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # calibrate picked accelerometer data, 3rd order method
            AXcal_3rd, AYcal_3rd, AZcal_3rd = calibrate_acc_3rd_order(AXpick_av, AYpick_av, AZpick_av, xfactors_3rd, yfactors_3rd, zfactors_3rd)
            
            # calculate and plot QC parameters for the accelerometer calibration
            acc_calibrationQC(AXcal_lin, AYcal_lin, AZcal_lin, AXcal_3rd, AYcal_3rd, AZcal_3rd, AXRef, AYRef, AZRef)
            
            # save calibration constants to file
            save_file_cal_constants(acc_cal_file, acc_cal_sheet, xoffset, yoffset, zoffset, xfactors_lin, yfactors_lin, zfactors_lin, \
                          xfactors_3rd, yfactors_3rd, zfactors_3rd, T_lin, abstemp)
            
            if controlseq == True: 
                # read control signals
                meas_pointAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, TApick_avAC, AXRefAC, AYRefAC, AZRefAC = read_file_ps1AC(pickfile_readAC, picksheet_readAC)
                # calibrate picked accelerometer data, classical method
                AXcal_linAC, AYcal_linAC, AZcal_linAC = calibrate_acc_classical(AXpick_avAC, AYpick_avAC, AZpick_avAC, xfactors_lin, yfactors_lin, zfactors_lin)
                # plot control plots
                acc_calibration_control(meas_pointAC, AXcal_linAC, AYcal_linAC, AZcal_linAC, AXRefAC, AYRefAC, AZRefAC)
            
        
        
        if processtep == 3:
    
            #read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, \
            TG1count, TG2count, TG3count = readrawdata(tempfile_read, tempsheet_read)
         
            # calculate 5th order polynominal
            TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2, TcoeffGX1lin, TcoeffGY1lin, TcoeffGX2lin, TcoeffGZ2lin, TG1min, TG1max, TG2min, TG2max = \
                find_gyro_temp_poly(index, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, f_s_temp)
            
            save_file_gyroT_constants(gyroTcal_file, gyroTcal_sheet, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2, \
                                      TcoeffGX1lin, TcoeffGY1lin, TcoeffGX2lin, TcoeffGZ2lin,TG1min, TG1max, TG2min, TG2max)
                

            
            
        
        if processtep == 4:
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef \
                = read_file_ps1(pickfile_read, picksheet_read)
            
            # read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, \
                TG1count, TG2count, TG3count = readrawdata(logfile_acc_gsens, sheetname_acc_gsens)
            
            # read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
            
            # read temperature coefficients from file 
            TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2 = read_file_gyroTcoeff(gyroTcal_file, gyroTcal_sheet)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            
            
          
                # temperature compensate gyro counters and pick same time instants as before
            GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav = Tcomp_gyro_pick_average(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                    f_s_acc, timeabs, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2)
            

            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
            
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # calculate gravity sensitivity, ignoring earth rotation. Use in gyro scale factor calibration. Plots 
            GX1gsens, GY1gsens, GX2gsens, GZ2gsens , TG1, TG2 = gsens_calc(AXcal_lin, AYcal_lin, AZcal_lin, GX1pick_TCav, GY1pick_TCav, \
                                                    GX2pick_TCav, GZ2pick_TCav, TG1pick_av, TG2pick_av, 131, 132, 133, 134, plotfigs1)        
            
            # plot QC data, 
           
            plot_figures_ps4(time, GX1count, GY1count, GX2count, GZ2count, meas_point, GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav)
                
                
            save_file_gsens(gyrogsenscal_file, gyrogsenscal_sheet, GX1gsens, GY1gsens, GX2gsens, GZ2gsens, TG1, TG2, abstemp)
            
        
        if processtep == 5:
            
            # read data file for X rotation
            indexXr, accXcountXr, accYcountXr, accZcountXr, GX1countXr, GY1countXr, GX2countXr, GZ2countXr, GX3countXr, GY3countXr, GZ3countXr, \
            TAcountXr, TG1countXr, TG2countXr, TG3countXr = readrawdata(Xrotfile, Xrotsheet)
            
            meas_pointXr, timegapXr, timeabsXr, rotrateXr = read_rot_timefile(rot_timefile, Xrot_timesheet)
            
            rot = 1 # X rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
     
            GX1avXr, GY1avXr, GX2avXr, GZ2avXr, TG1avXr, TG2avXr = gyro_rot_pick(meas_pointXr, timeabsXr, rotrateXr, av_time_standstill, av_time_rot, \
                                    margin_standstill ,margin_rotation, f_s_rot, accXcountXr, accYcountXr, accZcountXr, GX1countXr, \
                                    GY1countXr, GX2countXr, GZ2countXr, TG1countXr, TG2countXr, rot)
            
            
           
            # read data file for Y rotation
            indexYr, accXcountYr, accYcountYr, accZcountYr, GX1countYr, GY1countYr, GX2countYr, GZ2countYr, GX3countYr, GY3countYr, GZ3countYr, \
            TAcountYr, TG1countYr, TG2countYr, TG3countYr  = readrawdata(Yrotfile, Yrotsheet)
            
            meas_pointYr, timegapYr, timeabsYr, rotrateYr = read_rot_timefile(rot_timefile, Yrot_timesheet)
            
            rot = 2 # Y rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
            
        
            GX1avYr, GY1avYr, GX2avYr, GZ2avYr, TG1avYr, TG2avYr = gyro_rot_pick(meas_pointYr, timeabsYr, rotrateYr, av_time_standstill, av_time_rot, \
                                    margin_standstill ,margin_rotation, f_s_rot, accXcountYr, accYcountYr, accZcountYr, GX1countYr, \
                                    GY1countYr, GX2countYr, GZ2countYr, TG1countYr, TG2countYr, rot)
            

            
            # read data file for Z rotation
            indexZr, accXcountZr, accYcountZr, accZcountZr, GX1countZr, GY1countZr, GX2countZr, GZ2countZr, GX3countZr, GY3countZr, GZ3countZr, \
            TAcountZr, TG1countZr, TG2countZr, TG3countZr = readrawdata(Zrotfile, Zrotsheet)
            
            meas_pointZr, timegapZr, timeabsZr, rotrateZr = read_rot_timefile(rot_timefile, Zrot_timesheet)
            
            rot = 3 # Z rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile 
          
            GX1avZr, GY1avZr, GX2avZr, GZ2avZr, TG1avZr, TG2avZr = gyro_rot_pick(meas_pointZr, timeabsZr, rotrateZr, av_time_standstill, av_time_rot, \
                                    margin_standstill ,margin_rotation, f_s_rot, accXcountZr, accYcountZr, accZcountZr, GX1countZr, \
                                    GY1countZr, GX2countZr, GZ2countZr, TG1countZr, TG2countZr, rot)
                
            
            
            # calculate ortomatrix, returns three versions of ortomatrix, one for each GXoption
            
            xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, TG1, TG2 = calc_gyro_ortomatrix(GX1avXr, GY1avXr, GX2avXr, GZ2avXr, TG1avXr, TG2avXr, rotrateXr, GX1avYr, GY1avYr, GX2avYr, \
                                      GZ2avYr, TG1avYr, TG2avYr, rotrateYr, GX1avZr, GY1avZr, GX2avZr, GZ2avZr, TG1avZr, TG2avZr, rotrateZr)
        
            # write gyro ortomatrix to file. Write all three versions to same file (GX0 = average of GX1 and GX2 counts, GX1 count and GX2 counts)
            
            save_file_gyro_orto(gyro_orto_file, gyro_orto_sheet, xfactorsGX0, yfactorsGX0, zfactorsGX0, \
                                xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, bestGX, TG1, TG2, abstemp)
            
        
        if processtep == 6:
            
            # read gyro scale factor matrix
            xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, \
                    bestGX, TG1, TG2, abstemp = read_file_gyro_orto(gyro_orto_file, gyro_orto_sheet)
            
            # find gyro sensors unit pointing vectors (along sensitive axis) from gyro orto matrix
            GX1point, GY1point, GX2point, GZ2point, GX1scale, GY1scale, GX2scale, GZ2scale = \
                        gyro_body_vec(xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2)
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef \
                = read_file_ps1(pickfile_read, picksheet_read)
            
            # read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, TG1count, TG2count, TG3count \
                = readrawdata(logfile_acc_gsens, sheetname_acc_gsens)
            
            # read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
            
            # read temperature coefficients from file 
            TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2 = read_file_gyroTcoeff(gyroTcal_file, gyroTcal_sheet)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            
            # temperature compensate gyro counters and pick same time instants as before
         
            GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav = Tcomp_gyro_pick_average(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                    f_s_acc, timeabs, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2)
            
  
            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
            
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # read data from angle file (same as timing file)
            azimuth, inclination, toolface = readanglefile(file_path_timefile, sheetname_time)
            
            # calculate gyro counter values resulting from earth rotation
            GX1_er_comp, GY1_er_comp, GX2_er_comp, GZ2_er_comp = er_comp_gyros(GX1pick_TCav, GY1pick_TCav, GX2pick_TCav,\
                GZ2pick_TCav, azimuth, inclination, toolface, GX1point, GY1point, GX2point, GZ2point, GX1scale, GY1scale, GX2scale, GZ2scale)
            
            # adjust gyro counter values for earth rotation
            GX1pick_TCav = GX1pick_TCav - GX1_er_comp
            GY1pick_TCav = GY1pick_TCav - GY1_er_comp
            GX2pick_TCav = GX2pick_TCav - GX2_er_comp
            GZ2pick_TCav = GZ2pick_TCav - GZ2_er_comp
            
            
            # calculate gravity sensitivity, compensating for earth rotation. Use in later data calibration. Plots 
            GX1gsens_erc, GY1gsens_erc, GX2gsens_erc, GZ2gsens_erc , TG1, TG2 = gsens_calc(AXcal_lin, AYcal_lin, AZcal_lin, GX1pick_TCav, \
                                        GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav, TG1pick_av, TG2pick_av, 171, 172, 173, 174, plotfigs2)        
            
            save_file_gsens(gyrogsenscal_erc_file, gyrogsenscal_erc_sheet, GX1gsens_erc, GY1gsens_erc, GX2gsens_erc, GZ2gsens_erc, TG1, TG2, abstemp)    

if sequencemode == 3:            
    
    for processtep in (1,2,3,4,5,6):   
        if processtep == 1:
        
    #read data from raw data file 
            index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, \
                TG1count, TG2count, TG3count = readrawdata(logfile_acc_gsens, sheetname_acc_gsens)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            av_pointsAC = av_timeAC * f_s_acc
            
            
            #read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
    
            
            # read data from timing file for control signal     
    
            if controlseq == True:
                meas_pointAC, timegapAC, timeabsAC, AXRefAC, AYRefAC, AZRefAC = readtimefile(file_path_timefile, sheetname_control)
            
                    
     
            # pick correctly timed value from counter sequence. Adjust timing file as necessary   
            AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, GY1pick_av, \
            GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, GX2pick_std, GZ2pick_std, TG1pick_av, \
            TG2pick_av, TApick_av \
            = acc_gsens_pick(meas_point, av_time, timeabs, f_s_acc, accXcount, accYcount, accZcount, TAcount, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count)
            
            if controlseq == True:
                AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, GX1pickAC, GY1pickAC, GX2pickAC, \
                    GZ2pickAC, GX1pick_avAC, GY1pick_avAC, GX2pick_avAC, GZ2pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, \
                    GX1pick_stdAC, GY1pick_stdAC, GX2pick_stdAC, GZ2pick_stdAC, TG1pick_avAC, TG2pick_avAC, TApick_avAC \
                    = acc_control_pick(meas_pointAC, av_timeAC, timeabsAC, f_s_acc, accXcount, accYcount, accZcount, TAcount, \
                                     GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count)
    
    
            
            # QC plots for this sequence
            plot_figures_ps1(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                             timeabs, av_points, GX1count, GY1count, GX2count, GZ2count, GX1pick, GY1pick, GX2pick, GZ2pick, GX1pick_av, \
                             GY1pick_av, GX2pick_av, GZ2pick_av, AXpick_std, AYpick_std, AZpick_std, GX1pick_std, GY1pick_std, \
                             GX2pick_std, GZ2pick_std, TG1pick_av, TG2pick_av, TApick_av)     
            
            
            if controlseq == True:
                plot_figures_ps1AC(time, accXcount, accYcount, accZcount, AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, meas_pointAC, \
                         timeabsAC, av_pointsAC, GX1count, GY1count, GX2count, GZ2count, GX1pickAC, GY1pickAC, GX2pickAC, GZ2pickAC, GX1pick_avAC, \
                         GY1pick_avAC, GX2pick_avAC, GZ2pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, GX1pick_stdAC, GY1pick_stdAC, \
                         GX2pick_stdAC, GZ2pick_stdAC, TG1pick_avAC, TG2pick_avAC, TApick_avAC)
        
            
            # write selected and averaged data to file
            save_file_ps1(pickfile_write, picksheet_write, meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, \
                          GX2pick_av, GZ2pick_av, TG1pick_av, TG2pick_av, TApick_av, AXRef, AYRef, AZRef)
                
            if controlseq == True:  
                save_file_ps1AC(pickfile_writeAC, picksheet_writeAC, meas_pointAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, TApick_avAC, AXRefAC, AYRefAC, AZRefAC)
            
            
            if autostarttime == True:
                starttime_Tcomp = int(timeabs[0] - 1520)
                print ('starttime_Tcomp = ', starttime_Tcomp)
                
                stoptime_Tcomp = int(timeabs[107] + 100)
                if abstemp == 28:
                    stoptime_Tcomp = int(timeabs[107] + 450)
                print ('stoptime_Tcomp = ', stoptime_Tcomp)
            
        
        if processtep == 2:
        
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef \
            = read_file_ps1(pickfile_read , picksheet_read)
            
            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)    
            
            # 3rd order accelerometer calibration constants
            xfactors_3rd, yfactors_3rd, zfactors_3rd, T_3rd = acc_cal_3rd_order(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
           
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # calibrate picked accelerometer data, 3rd order method
            AXcal_3rd, AYcal_3rd, AZcal_3rd = calibrate_acc_3rd_order(AXpick_av, AYpick_av, AZpick_av, xfactors_3rd, yfactors_3rd, zfactors_3rd)
            
            # calculate and plot QC parameters for the accelerometer calibration
            acc_calibrationQC(AXcal_lin, AYcal_lin, AZcal_lin, AXcal_3rd, AYcal_3rd, AZcal_3rd, AXRef, AYRef, AZRef)
            
            # save calibration constants to file
            save_file_cal_constants(acc_cal_file, acc_cal_sheet, xoffset, yoffset, zoffset, xfactors_lin, yfactors_lin, zfactors_lin, \
                          xfactors_3rd, yfactors_3rd, zfactors_3rd, T_lin, abstemp)
            
            
            if controlseq == True: 
                # read control signals
                meas_pointAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, TApick_avAC, AXRefAC, AYRefAC, AZRefAC = read_file_ps1AC(pickfile_readAC, picksheet_readAC)
                # calibrate picked accelerometer data, classical method
                AXcal_linAC, AYcal_linAC, AZcal_linAC = calibrate_acc_classical(AXpick_avAC, AYpick_avAC, AZpick_avAC, xfactors_lin, yfactors_lin, zfactors_lin)
                # plot control plots
                acc_calibration_control(meas_pointAC, AXcal_linAC, AYcal_linAC, AZcal_linAC, AXRefAC, AYRefAC, AZRefAC)
            
        
        
        if processtep == 3:
            
            #read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, \
            TG1count, TG2count, TG3count = readrawdata(tempfile_read, tempsheet_read)
         
            # calculate 5th order polynominal
            TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2, TcoeffGX1lin, TcoeffGY1lin, TcoeffGX2lin, TcoeffGZ2lin, TG1min, TG1max, TG2min, TG2max = \
                find_gyro_temp_poly(index, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, f_s_temp)
            
            save_file_gyroT_constants(gyroTcal_file, gyroTcal_sheet, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2, \
                                      TcoeffGX1lin, TcoeffGY1lin, TcoeffGX2lin, TcoeffGZ2lin,TG1min, TG1max, TG2min, TG2max)

        
        
    
        if processtep == 4:
        
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef \
                = read_file_ps1(pickfile_read, picksheet_read)
            
            # read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, \
                TG1count, TG2count, TG3count = readrawdata(logfile_acc_gsens, sheetname_acc_gsens)
            
            # read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
            
            # read temperature coefficients from file 
            TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2 = read_file_gyroTcoeff(gyroTcal_file, gyroTcal_sheet)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            
            
                         # temperature compensate gyro counters and pick same time instants as before
            GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav = Tcomp_gyro_pick_average(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                    f_s_acc, timeabs, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2)
            

            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
            
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # calculate gravity sensitivity, ignoring earth rotation. Use in gyro scale factor calibration. Plots 
            GX1gsens, GY1gsens, GX2gsens, GZ2gsens , TG1, TG2 = gsens_calc(AXcal_lin, AYcal_lin, AZcal_lin, GX1pick_TCav, GY1pick_TCav, \
                                                    GX2pick_TCav, GZ2pick_TCav, TG1pick_av, TG2pick_av, 131, 132, 133, 134, plotfigs1)        
            
            # plot QC data, 
            plot_figures_ps4(time, GX1count, GY1count, GX2count, GZ2count, meas_point, GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav)

               
            save_file_gsens(gyrogsenscal_file, gyrogsenscal_sheet, GX1gsens, GY1gsens, GX2gsens, GZ2gsens, TG1, TG2, abstemp)
            
    
        if processtep == 5:
        
            # read data file for X rotation
            indexXr, accXcountXr, accYcountXr, accZcountXr, GX1countXr, GY1countXr, GX2countXr, GZ2countXr, GX3countXr, GY3countXr, GZ3countXr, \
            TAcountXr, TG1countXr, TG2countXr, TG3countXr = readrawdata(Xrotfile, Xrotsheet)
            
            meas_pointXr, timegapXr, timeabsXr, rotrateXr = read_rot_timefile(rot_timefile, Xrot_timesheet)
            
            rot = 1 # X rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
   
            GX1avXr, GY1avXr, GX2avXr, GZ2avXr, TG1avXr, TG2avXr = gyro_rot_pick(meas_pointXr, timeabsXr, rotrateXr, av_time_standstill, av_time_rot, \
                                    margin_standstill ,margin_rotation, f_s_rot, accXcountXr, accYcountXr, accZcountXr, GX1countXr, \
                                    GY1countXr, GX2countXr, GZ2countXr, TG1countXr, TG2countXr, rot)
            
            
           
            # read data file for Y rotation
            indexYr, accXcountYr, accYcountYr, accZcountYr, GX1countYr, GY1countYr, GX2countYr, GZ2countYr, GX3countYr, GY3countYr, GZ3countYr, \
            TAcountYr, TG1countYr, TG2countYr, TG3countYr  = readrawdata(Yrotfile, Yrotsheet)
            
            meas_pointYr, timegapYr, timeabsYr, rotrateYr = read_rot_timefile(rot_timefile, Yrot_timesheet)
            
            rot = 2 # Y rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
            
            GX1avYr, GY1avYr, GX2avYr, GZ2avYr, TG1avYr, TG2avYr = gyro_rot_pick(meas_pointYr, timeabsYr, rotrateYr, av_time_standstill, av_time_rot, \
                                    margin_standstill ,margin_rotation, f_s_rot, accXcountYr, accYcountYr, accZcountYr, GX1countYr, \
                                    GY1countYr, GX2countYr, GZ2countYr, TG1countYr, TG2countYr, rot)

            
            # read data file for Z rotation
            indexZr, accXcountZr, accYcountZr, accZcountZr, GX1countZr, GY1countZr, GX2countZr, GZ2countZr, GX3countZr, GY3countZr, GZ3countZr, \
            TAcountZr, TG1countZr, TG2countZr, TG3countZr = readrawdata(Zrotfile, Zrotsheet)
            
            meas_pointZr, timegapZr, timeabsZr, rotrateZr = read_rot_timefile(rot_timefile, Zrot_timesheet)
            
            rot = 3 # Z rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile 

            GX1avZr, GY1avZr, GX2avZr, GZ2avZr, TG1avZr, TG2avZr = gyro_rot_pick(meas_pointZr, timeabsZr, rotrateZr, av_time_standstill, av_time_rot, \
                                    margin_standstill ,margin_rotation, f_s_rot, accXcountZr, accYcountZr, accZcountZr, GX1countZr, \
                                    GY1countZr, GX2countZr, GZ2countZr, TG1countZr, TG2countZr, rot)
            
            
            # calculate ortomatrix, returns three versions of ortomatrix, one for each GXoption
            
            xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, TG1, TG2 = calc_gyro_ortomatrix(GX1avXr, GY1avXr, GX2avXr, GZ2avXr, TG1avXr, TG2avXr, rotrateXr, GX1avYr, GY1avYr, GX2avYr, \
                                      GZ2avYr, TG1avYr, TG2avYr, rotrateYr, GX1avZr, GY1avZr, GX2avZr, GZ2avZr, TG1avZr, TG2avZr, rotrateZr)
        
            # write gyro ortomatrix to file. Write all three versions to same file (GX0 = average of GX1 and GX2 counts, GX1 count and GX2 counts)
            
            save_file_gyro_orto(gyro_orto_file, gyro_orto_sheet, xfactorsGX0, yfactorsGX0, zfactorsGX0, \
                                xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, bestGX, TG1, TG2, abstemp)
        
    
        if processtep == 6:
        
            # read gyro scale factor matrix
            xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, \
                    bestGX, TG1, TG2, abstemp = read_file_gyro_orto(gyro_orto_file, gyro_orto_sheet)
            
            # find gyro sensors unit pointing vectors (along sensitive axis) from gyro orto matrix
            GX1point, GY1point, GX2point, GZ2point, GX1scale, GY1scale, GX2scale, GZ2scale = \
                        gyro_body_vec(xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2)
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX1pick_av, GY1pick_av, GX2pick_av, GZ2pick_av, TApick_av, TG1pick_av, TG2pick_av, AXRef, AYRef, AZRef \
                = read_file_ps1(pickfile_read, picksheet_read)
            
            # read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, TG1count, TG2count, TG3count \
                = readrawdata(logfile_acc_gsens, sheetname_acc_gsens)
            
            # read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
            
            # read temperature coefficients from file 
            TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2 = read_file_gyroTcoeff(gyroTcal_file, gyroTcal_sheet)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            
            # temperature compensate gyro counters and pick same time instants as before
            
            GX1pick_TCav, GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav = Tcomp_gyro_pick_average(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                    f_s_acc, timeabs, GX1count, GY1count, GX2count, GZ2count, TG1count, TG2count, TcoeffGX1, TcoeffGY1, TcoeffGX2, TcoeffGZ2)
            

            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
            
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # read data from angle file (same as timing file)
            azimuth, inclination, toolface = readanglefile(file_path_timefile, sheetname_time)
            
            # calculate gyro counter values resulting from earth rotation
            GX1_er_comp, GY1_er_comp, GX2_er_comp, GZ2_er_comp = er_comp_gyros(GX1pick_TCav, GY1pick_TCav, GX2pick_TCav,\
                GZ2pick_TCav, azimuth, inclination, toolface, GX1point, GY1point, GX2point, GZ2point, GX1scale, GY1scale, GX2scale, GZ2scale)
            
            # adjust gyro counter values for earth rotation
            GX1pick_TCav = GX1pick_TCav - GX1_er_comp
            GY1pick_TCav = GY1pick_TCav - GY1_er_comp
            GX2pick_TCav = GX2pick_TCav - GX2_er_comp
            GZ2pick_TCav = GZ2pick_TCav - GZ2_er_comp
            
            
            # calculate gravity sensitivity, compensating for earth rotation. Use in later data calibration. Plots 
            GX1gsens_erc, GY1gsens_erc, GX2gsens_erc, GZ2gsens_erc , TG1, TG2 = gsens_calc(AXcal_lin, AYcal_lin, AZcal_lin, GX1pick_TCav, \
                                        GY1pick_TCav, GX2pick_TCav, GZ2pick_TCav, TG1pick_av, TG2pick_av, 171, 172, 173, 174, plotfigs2)        
            
            save_file_gsens(gyrogsenscal_erc_file, gyrogsenscal_erc_sheet, GX1gsens_erc, GY1gsens_erc, GX2gsens_erc, GZ2gsens_erc, TG1, TG2, abstemp)
            
            if autostarttime == True:
                starttime_Tcomp = int(timeabs[0] - 1520)
                print ('starttime_Tcomp = ', starttime_Tcomp)
                
                stoptime_Tcomp = int(timeabs[107] + 100)
                if abstemp == 28:
                    stoptime_Tcomp = int(timeabs[107] + 450)
                print ('stoptime_Tcomp = ', stoptime_Tcomp)
