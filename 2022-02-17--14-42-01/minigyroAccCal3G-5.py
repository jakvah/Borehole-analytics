################################################################################
# Company       : Devico AS
# Author        : Geir Bjarte
#
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
# R15 change loacl gravity at Heimdal to 9.821m/s^2, used from 10.06.2020
# R16 for python 3.8.5
# R17 fixed MAA to MAB bug in write constant file script. No change in this script
# R18 introduced gyro result image script, fixed minor error in plot 117

# R99 Commented out plt.show() for automatic processing

from __future__ import division
import matplotlib.pyplot as plt
import numpy as np
import xlrd
import xlsxwriter
from pandas import DataFrame


sequencemode = 3            # 0: run old style one processtep at a time
                            # 1: run processtep 1 only (use for first tool in a batch, where timing files need to be tuned)
                            # 2: run processteps 2-6 (use for first tool in a batch)
                            # 3: run processteps 1-6 (when timing is known in advance)   

processtep = 6              # processtep = 1: read acc and gyro raw data files from calibration, adjust timing until happy (in file_path_timefile)
                            #                 Plot QC figures 101-107. Write selected data to file
                            # processtep = 2: read selected accelerometer data, calculate calibration constants (ortomatrix and 3rd order. )
                            #                 write calibration constants to file and plot QC plots 111-113
                            # processtep = 3: read raw data from gyro bias temperature test. Plot QC plot 121, and write temperature coeffcients to file
                            # processtep = 4: temperature compensate gyro data, calculate g sensitivity. For now, this is done without 
                            #                 taking earth rotation into account
                            # processtep = 5: read three raw data files from gyro calibration, read timing file (external excel document), plot data. Calculate 
                            #                 gyro orto matrix. Option to select GX1, GX2 or average. But shall be 3x3 matrix! 
                            # processtep = 6: calculate earth rotation compensated g sensitivity
                        
#STgyro = True


# parameters processtep 1

abstemp = -5
toolSN = '3331'
g_range_ADXL = 4        # for cal file names only
g_range_ST = 16         # for cal file names only

av_time = 10    
av_timeAC = 5

tempfile_read =  toolSN + ' Tcal cut.xlsx' 
tempsheet_read =  toolSN + ' Tcal cut'


autostarttime = False
standstill_time = 10    

if autostarttime == False: 
	starttime_Tcomp = 80
	stoptime_Tcomp = 4299
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

acc_cal_file = toolSN + ' ST ' + str(abstemp) + 'C ' + str(g_range_ST) + 'g acc cal.xlsx'     # file for writing accelerometer calibration constants
acc_cal_sheet =  'Ark1'


 

# parameters processtep 3



f_s_temp = 10                                                            


gyroTcal_file = toolSN + ' ST gyro Tconst.xlsx'

gyroTcal_sheet = 'Ark1'


# parameters processtep 4
                                            


gyrogsenscal_file = toolSN + ' ST ' + str(abstemp) + 'C' + ' ' + 'gyro gsens ier.xlsx'

    
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

gyro_orto_file = toolSN + ' ST ' + str(abstemp) + 'C' + ' ' + 'gyro orto.xlsx'
gyro_orto_sheet = 'Ark1'


# parameters processtep 6

earthrate = 15.04       # deg/hour
latitude = 63.35        # deg at Heimdal

gyrogsenscal_erc_file = toolSN + ' ST ' + str(abstemp) + 'C' + ' ' + 'gyro gsens erc.xlsx'
gyrogsenscal_erc_sheet = 'Ark1'

# function definitions

          
def readrawdata_3G(file_path, sheetname_data):
          
    # read data file 
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheetname_data)
       
    # extract data from X accelerometer
    offset = 1
    col = worksheet.col_values(4)
    accXcount = col[offset:]
    accXcount = np.array(accXcount)
    # extract data from Y accelerometer
    col = worksheet.col_values(5)
    accYcount = col[offset:]
    accYcount = np.array(accYcount)
    # extract data from Z accelerometer
    col = worksheet.col_values(6)
    accZcount = col[offset:]
    accZcount = np.array(accZcount)   
    
    # extract index
    col = worksheet.col_values(0)
    index = col[offset:]
    index = np.array(index)
          
    
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
    col = worksheet.col_values(17)
    TAcount = col[offset:]
    TAcount = np.array(TAcount)
       
    col = worksheet.col_values(17)
    TG3count = col[offset:]
    TG3count = np.array(TG3count)
    
    index = index - index[0]
    
    return index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount,TG3count 

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


def plot_figures_ps1_3G(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                     timeabs, av_points, GX3count, GY3count, GZ4count, GX3pick, GY3pick, GZ3pick, GX3pick_av, GY3pick_av, GZ3pick_av, \
                     AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, GZ3pick_std, TG3pick_av, TApick_av):
    
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
    plt.plot(time, GX3count, label = 'GX3count')

    plt.plot(time, GY3count, label = 'GY3count')
    plt.plot(time, GZ3count, label = 'GZ3count')
    
    plt.plot(timeabs, GX3pick, 'C0o', label = 'GX3pick')
    plt.plot(timeabs, GY3pick, 'C1o', label = 'GY3pick')
    plt.plot(timeabs, GZ3pick, 'C3o', label = 'GZ3pick')
 
    plt.plot(timeabs, GX3pick_av, 'C0s', label = 'GX3pick_av')
    plt.plot(timeabs, GY3pick_av, 'C1s', label = 'GY3pick_av')
    plt.plot(timeabs, GZ3pick_av, 'C3s', label = 'GZ3pick_av')
    
    
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GX3count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C4')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GY3count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C5')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GZ3count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C7')
    for n in range(len(meas_point)): 
        plt.annotate(n, xy=(timeabs[n],GX3pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GY3pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GZ3pick[n]+50))
    
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    
    
    plt.figure(103,figsize=(12,9))
    plt.plot(meas_point, AXpick_std, label = 'AXstd')

    plt.plot(meas_point, AYpick_std, label = 'AYstd')
    plt.plot(meas_point, AZpick_std, label = 'AZstd')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' secondary '+ str(abstemp)+ 'C plot 103 acc stdev.png')
#    plt.show()
    
    plt.figure(104, figsize=(12,9))
    plt.plot(meas_point, GX3pick_std, label = 'GX3std')

    plt.plot(meas_point, GY3pick_std, label = 'GY3std')
    plt.plot(meas_point, GZ3pick_std, label = 'GZ3std')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' secondary '+ str(abstemp)+ 'C plot 104 gyro stdev.png')
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
    plt.plot(meas_point, GX3pick_av, label = 'GX3av')

    plt.plot(meas_point, GY3pick_av, label = 'GY3av')
    plt.plot(meas_point, GZ3pick_av, label = 'GZ3av')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    
    plt.figure(107, figsize=(12,9))
    plt.plot(meas_point, TG3pick_av, label = 'TG3av')
  
    plt.plot(meas_point, TApick_av, label = 'TAav')
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
#    plt.show()
    return



def plot_figures_ps1AC(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                     timeabs, av_points, GX3count, GY3count, GZ3count, GX3pick, GY3pick, GZ3pick, GX3pick_av, \
                     GY3pick_av, GZ3pick_av, AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, \
                     GZ3pick_std, TG3pick_av, TApick_av):
    
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
    plt.plot(time, GX3count, label = 'GX3count')

    plt.plot(time, GY3count, label = 'GY3count')
    plt.plot(time, GZ3count, label = 'GZ3count')
    
    plt.plot(timeabs, GX3pick, 'C0o', label = 'GX3pick')
    plt.plot(timeabs, GY3pick, 'C1o', label = 'GY3pick')
    plt.plot(timeabs, GZ3pick, 'C3o', label = 'GZ3pick')
 
    plt.plot(timeabs, GX3pick_av, 'C0s', label = 'GX3pick_av')
    plt.plot(timeabs, GY3pick_av, 'C1s', label = 'GY3pick_av')
    plt.plot(timeabs, GZ3pick_av, 'C3s', label = 'GZ3pick_av')
    
    
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GX3count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C4')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GY3count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C5')
        plt.plot(time[(int(k-av_points/2))-1:(int(k+av_points/2))], GZ3count[(int(k-av_points/2))-1:(int(k+av_points/2))], 'C7')
    for n in range(len(meas_point)): 
        plt.annotate(n, xy=(timeabs[n],GX3pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GY3pick[n]+50))
        plt.annotate(n, xy=(timeabs[n],GZ3pick[n]+50))
    
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
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


def acc_gsens_pick_3G(meas_point, av_time, f_s_acc, timeabs, accXcount, accYcount, accZcount, TAcount, GX3count, GY3count, GZ3count, TG3count):
    
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
    
    GX3pick = np.zeros(len(meas_point))
    GY3pick = np.zeros(len(meas_point))
    GZ3pick = np.zeros(len(meas_point))
       
    
    GX3pick_av = np.zeros(len(meas_point))
    GY3pick_av = np.zeros(len(meas_point))
    GZ3pick_av = np.zeros(len(meas_point))
    TG3pick_av = np.zeros(len(meas_point))
    
    
    GX3pick_std = np.zeros(len(meas_point))
    GY3pick_std = np.zeros(len(meas_point))
    GZ3pick_std = np.zeros(len(meas_point))
    
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
        
        GX3pick[n] = GX3count[k]
        GY3pick[n] = GY3count[k]
        GZ3pick[n] = GZ3count[k]
        GX3pick_av[n] = np.mean(GX3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY3pick_av[n] = np.mean(GY3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ3pick_av[n] = np.mean(GZ3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TG3pick_av[n] = np.mean(TG3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        
        GX3pick_std[n] = np.std(GX3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY3pick_std[n] = np.std(GY3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ3pick_std[n] = np.std(GZ3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
    
    return  AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX3pick, GY3pick, GZ3pick, GX3pick_av, \
            GY3pick_av, GZ3pick_av, AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, \
            GZ3pick_std, TG3pick_av, TApick_av
            
            
def acc_control_pick(meas_point, av_time, timeabs, f_s_acc, accXcount, accYcount, accZcount, TAcount, GX3count, GY3count, GZ3count, TG3count):
    
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
    
    GX3pick = np.zeros(len(meas_point))
    GY3pick = np.zeros(len(meas_point))
    GZ3pick = np.zeros(len(meas_point))
       
    
    GX3pick_av = np.zeros(len(meas_point))
    GY3pick_av = np.zeros(len(meas_point))
    GZ3pick_av = np.zeros(len(meas_point))
    TG3pick_av = np.zeros(len(meas_point))

    
    
    GX3pick_std = np.zeros(len(meas_point))
    GY3pick_std = np.zeros(len(meas_point))
    GZ3pick_std = np.zeros(len(meas_point))
    
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
        
        GX3pick[n] = GX3count[k]
        GY3pick[n] = GY3count[k]
        GZ3pick[n] = GZ3count[k]
        GX3pick_av[n] = np.mean(GX3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY3pick_av[n] = np.mean(GY3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ3pick_av[n] = np.mean(GZ3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TG3pick_av[n] = np.mean(TG3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        
        GX3pick_std[n] = np.std(GX3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY3pick_std[n] = np.std(GY3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ3pick_std[n] = np.std(GZ3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
    
    return  AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX3pick, GY3pick, GZ3pick, GX3pick_av, \
            GY3pick_av, GZ3pick_av, AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, \
            GZ3pick_std, TG3pick_av, TApick_av

            



def save_file_ps1_3G(pickfile_write, picksheet_write, meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, \
                  GZ3pick_av, TG3pick_av, TApick_av, AXRef, AYRef, AZRef):
    
    l0 = meas_point
    l1 = AXpick_av
    l2 = AYpick_av
    l3 = AZpick_av
    l4 = GX3pick_av
    l5 = GY3pick_av
    l6 = GZ3pick_av
    l7 = TApick_av
    l8 = TG3pick_av
    l9 = AXRef
    l10 = AYRef
    l11 = AZRef
    
    
    df = DataFrame({'00 Meas point':l0, '01 AXpick_av':l1, '02 AYpick_av':l2, '03 AZpick_av':l3, '04 GX3pick_av': l4, '05 GY3pick_av': l5, 
                    '06 GZ3pick_av':l6, '07 TApick_av': l7, '08 TG3pick_av': l8, '09 AXRef': l9, '10 AYRef': l10, '11 AZRef': l11})
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


def read_file_ps1_3G(file_path, sheetname_data):

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
    GX3pick_av = col[offset:]
    GX3pick_av = np.array(GX3pick_av)
    
    # extract GY1 data
    col = worksheet.col_values(5)
    GY3pick_av = col[offset:]
    GY3pick_av = np.array(GY3pick_av)    
     
    # extract GZ2 data
    col = worksheet.col_values(6)
    GZ3pick_av = col[offset:]
    GZ3pick_av = np.array(GZ3pick_av) 
    
    # extract TA data
    col = worksheet.col_values(7)
    TApick_av = col[offset:]
    TApick_av = np.array(TApick_av) 
    
    # extract TG1 data
    col = worksheet.col_values(8)
    TG3pick_av = col[offset:]
    TG3pick_av = np.array(TG3pick_av) 
       
    # extract AXRef data
    col = worksheet.col_values(9)
    AXRef = col[offset:]
    AXRef = np.array(AXRef)
    
    # extract AYRef data
    col = worksheet.col_values(10)
    AYRef = col[offset:]
    AYRef = np.array(AYRef) 
    
    # extract AZRef data
    col = worksheet.col_values(11)
    AZRef = col[offset:]
    AZRef = np.array(AZRef)
    
    
    return meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef   


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
    plt.savefig(toolSN + ' secondary '+ str(abstemp)+'C plot 111 acc result.png')
#    plt.show()
    
    plt.figure(112, figsize=(12,9))
    plt.plot(meas_point, inc_error_lin, label = 'Inc error lin')

    plt.plot(meas_point, inc_error_3rd, label = 'Inc error 3rd')

    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' secondary '+ str(abstemp)+'C plot 112 acc inc result.png')
#    plt.show()
    
    plt.figure(113, figsize=(12,9))
    plt.plot(meas_point, tf_error_lin, label = 'TF error lin')

    plt.plot(meas_point, tf_error_3rd, label = 'TF error 3rd')

    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' secondary '+ str(abstemp)+ 'C plot 113 acc tf result.png')
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
    plt.savefig(toolSN +' secondary '+ str(abstemp)+ 'C plot 115 acc control.png')
#    plt.show()
    
    
    
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
    plt.savefig(toolSN + ' secondary '+ str(abstemp)+ 'C plot 116 acc inc control.png')


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
    plt.savefig(toolSN + ' secondary '+ str(abstemp)+ 'C plot 117 acc tf control.png')

    return


def find_gyro_temp_poly_3G(index, GX3count, GY3count, GZ3count, TG3count, f_s):
    
    plt.close(121)
    
    
    TcoeffGX3 = np.zeros(6)
    TcoeffGY3 = np.zeros(6)
    TcoeffGZ3 = np.zeros(6)
    
    TcoeffGX3lin = np.zeros(2)
    TcoeffGY3lin = np.zeros(2)
    TcoeffGZ3lin = np.zeros(2)
       
    TcoeffGX3 = np.polyfit(TG3count, GX3count, 5)
    TcoeffGY3 = np.polyfit(TG3count, GY3count, 5)
    TcoeffGZ3 = np.polyfit(TG3count, GZ3count, 5)
    
    TG3min = np.min(TG3count)
    TG3max = np.max(TG3count)
    
    GX3poly = np.zeros(int(TG3max-TG3min+1))
    GY3poly = np.zeros(int(TG3max-TG3min+1))
    TG3poly = np.zeros(int(TG3max-TG3min+1))
    GZ3poly = np.zeros(int(TG3max-TG3min+1))

    
    GX3lin = np.zeros(int(TG3max-TG3min+1)+100)
    GY3lin = np.zeros(int(TG3max-TG3min+1)+100)
    TG3lin = np.zeros(int(TG3max-TG3min+1)+100)
    GZ3lin = np.zeros(int(TG3max-TG3min+1)+100)

    
    for n in range(int(1+TG3max) - int(TG3min)):
        TG3poly[n] = TG3min + n
        GX3poly[n] = TcoeffGX3[5] + TcoeffGX3[4]*TG3poly[n] + TcoeffGX3[3]*TG3poly[n]**2 + TcoeffGX3[2]*TG3poly[n]**3 + \
                     TcoeffGX3[1]*TG3poly[n]**4 + TcoeffGX3[0]*TG3poly[n]**5
        GY3poly[n] = TcoeffGY3[5] + TcoeffGY3[4]*TG3poly[n] + TcoeffGY3[3]*TG3poly[n]**2 + TcoeffGY3[2]*TG3poly[n]**3 +  \
                     TcoeffGY3[1]*TG3poly[n]**4 + TcoeffGY3[0]*TG3poly[n]**5
        GZ3poly[n] = TcoeffGZ3[5] + TcoeffGZ3[4]*TG3poly[n] + TcoeffGZ3[3]*TG3poly[n]**2 + TcoeffGZ3[2]*TG3poly[n]**3   \
                     + TcoeffGZ3[1]*TG3poly[n]**4 + TcoeffGZ3[0]*TG3poly[n]**5
    


    
    # find linear fit to be used if temperature counter values outside calibrated range
    
    GX3Tmin = TcoeffGX3[5] + TcoeffGX3[4]*TG3min + TcoeffGX3[3]*TG3min**2 + TcoeffGX3[2]*TG3min**3 + \
                     TcoeffGX3[1]*TG3min**4 + TcoeffGX3[0]*TG3min**5
    GX3Tmax = TcoeffGX3[5] + TcoeffGX3[4]*TG3max + TcoeffGX3[3]*TG3max**2 + TcoeffGX3[2]*TG3max**3 + \
                     TcoeffGX3[1]*TG3max**4 + TcoeffGX3[0]*TG3max**5
    
    GY3Tmin = TcoeffGY3[5] + TcoeffGY3[4]*TG3min + TcoeffGY3[3]*TG3min**2 + TcoeffGY3[2]*TG3min**3 + \
                     TcoeffGY3[1]*TG3min**4 + TcoeffGY3[0]*TG3min**5
    GY3Tmax = TcoeffGY3[5] + TcoeffGY3[4]*TG3max + TcoeffGY3[3]*TG3max**2 + TcoeffGY3[2]*TG3max**3 + \
                     TcoeffGY3[1]*TG3max**4 + TcoeffGY3[0]*TG3max**5
      
    GZ3Tmin = TcoeffGZ3[5] + TcoeffGZ3[4]*TG3min + TcoeffGZ3[3]*TG3min**2 + TcoeffGZ3[2]*TG3min**3 + \
                     TcoeffGZ3[1]*TG3min**4 + TcoeffGZ3[0]*TG3min**5
    GZ3Tmax = TcoeffGZ3[5] + TcoeffGZ3[4]*TG3max + TcoeffGZ3[3]*TG3max**2 + TcoeffGZ3[2]*TG3max**3 + \
                     TcoeffGZ3[1]*TG3max**4 + TcoeffGZ3[0]*TG3max**5
    
    TcoeffGX3lin[0] = (GX3Tmax - GX3Tmin)/(TG3max - TG3min)  
    TcoeffGX3lin[1] = GX3Tmin - TcoeffGX3lin[0] * TG3min
    
    TcoeffGY3lin[0] = (GY3Tmax - GY3Tmin)/(TG3max - TG3min)  
    TcoeffGY3lin[1] = GY3Tmin - TcoeffGY3lin[0] * TG3min
    
    TcoeffGZ3lin[0] = (GZ3Tmax - GZ3Tmin)/(TG3max - TG3min)  
    TcoeffGZ3lin[1] = GZ3Tmin - TcoeffGZ3lin[0] * TG3min
    
    for n in range(int(51+TG3max) - int(TG3min-50)):
        TG3lin[n] = TG3min + n - 50
        GX3lin[n] = TcoeffGX3lin[1] + TcoeffGX3lin[0]*TG3lin[n] 
        GY3lin[n] = TcoeffGY3lin[1] + TcoeffGY3lin[0]*TG3lin[n] 
        GZ3lin[n] = TcoeffGZ3lin[1] + TcoeffGZ3lin[0]*TG3lin[n] 

    
    plt.figure(121, figsize=(12,9))
    plt.plot(TG3count, GX3count, label = 'GX3count')

    plt.plot(TG3count, GY3count, label = 'GY3count')
    plt.plot(TG3count, GZ3count, label = 'GZ3count')
    
    plt.plot(TG3poly, GX3poly, label = 'GX3poly', lw = 3)
    plt.plot(TG3poly, GY3poly, label = 'GY3poly', lw = 3)
    plt.plot(TG3poly, GZ3poly, label = 'GZ3poly', lw = 3)
    
    plt.plot(TG3lin, GX3lin, label = 'GX3lin', lw = 1)
    plt.plot(TG3lin, GY3lin, label = 'GY3lin', lw = 1)
    plt.plot(TG3lin, GZ3lin, label = 'GZ3lin', lw = 1)
    
    
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    plt.savefig(toolSN + ' secondary '+ str(abstemp)+ 'C plot 121 gyro Tbias.png')
#    plt.show()
    
    print (TcoeffGX3)
    print (TcoeffGY3)
    print (TcoeffGZ3)
    
    print ('TG3min ', TG3min)
    
    return TcoeffGX3, TcoeffGY3, TcoeffGZ3, TcoeffGX3lin, TcoeffGY3lin, TcoeffGZ3lin, TG3min, TG3max


def save_file_gyroT_constants_3G(filename, sheetname, TcoeffGX3, TcoeffGY3, TcoeffGZ3, TcoeffGX3lin, TcoeffGY3lin, \
                              TcoeffGZ3lin, TG3min, TG3max):
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet(sheetname)
    
    worksheet.write(0, 0, 'GX3 coeffs')
    worksheet.write(0, 1, 'GY3 coeffs')
    worksheet.write(0, 2, 'GZ3 coeffs')
    worksheet.write(0, 3, 'GX3 lin coeffs')
    worksheet.write(0, 4, 'GY3 lin coeffs')
    worksheet.write(0, 5, 'GZ3 lin coeffs')
    worksheet.write(0, 6, 'TG3min')
    worksheet.write(0, 7, 'TG3max')
    worksheet.write(1, 6, TG3min)
    worksheet.write(1, 7, TG3max)


    for k in range(6):
        worksheet.write(k+1, 0, TcoeffGX3[k])
        worksheet.write(k+1, 1, TcoeffGY3[k])
        worksheet.write(k+1, 2, TcoeffGZ3[k])
    for k in range(2):
        worksheet.write(k+1, 3, TcoeffGX3lin[k])
        worksheet.write(k+1, 4, TcoeffGY3lin[k]) 
        worksheet.write(k+1, 5, TcoeffGZ3lin[k])
    
    workbook.close() 
    return


def gyro_bias_comp_3G(index, GX3count, GY3count, GZ3count, TG3count, ss1_length, ss2_length, GX3_Tconst, GY3_Tconst, GZ3_Tconst):
    
    GX3_TP = np.zeros(len(index))              # temperature polynomials
    GY3_TP = np.zeros(len(index))
    GZ3_TP = np.zeros(len(index))
    
    GX3_Tbias = np.zeros(len(index))
    GY3_Tbias = np.zeros(len(index))
    GZ3_Tbias = np.zeros(len(index))
    
    GX3bias = np.zeros(len(index))
    GY3bias = np.zeros(len(index))
    GZ3bias = np.zeros(len(index))
        
    for k in range(len(index)):
        GX3_TP[k] = GX3_Tconst[5] + GX3_Tconst[4]*TG3count[k] + GX3_Tconst[3]*TG3count[k]**2 + GX3_Tconst[2]*TG3count[k]**3 \
                    + GX3_Tconst[1]*TG3count[k]**4 + GX3_Tconst[0]*TG3count[k]**5
        GY3_TP[k] = GY3_Tconst[5] + GY3_Tconst[4]*TG3count[k] + GY3_Tconst[3]*TG3count[k]**2 + GY3_Tconst[2]*TG3count[k]**3 \
                    + GY3_Tconst[1]*TG3count[k]**4 + GY3_Tconst[0]*TG3count[k]**5
        GZ3_TP[k] = GZ3_Tconst[5] + GZ3_Tconst[4]*TG3count[k] + GZ3_Tconst[3]*TG3count[k]**2 + GZ3_Tconst[2]*TG3count[k]**3  \
                    + GZ3_Tconst[1]*TG3count[k]**4 + GZ3_Tconst[0]*TG3count[k]**5
    
    GX3_TPss1 = np.average(GX3_TP[0:ss1_length])              # Temperature polynomial average value in first standstill period 
    GY3_TPss1 = np.average(GY3_TP[0:ss1_length])
    GZ3_TPss1 = np.average(GZ3_TP[0:ss1_length])
    
    GX3_GCss1 = np.average(GX3count[0:ss1_length])            # Gravity compensated gyro counter value average value in first standstill period 
    GY3_GCss1 = np.average(GY3count[0:ss1_length])
    GZ3_GCss1 = np.average(GZ3count[0:ss1_length])
    
    GX3_GCss2 = np.average(GX3count[(len(index)-ss2_length):len(index)])    # Gravity compensated gyro counter value average value in second standstill period
    GY3_GCss2 = np.average(GY3count[(len(index)-ss2_length):len(index)])
    GZ3_GCss2 = np.average(GZ3count[(len(index)-ss2_length):len(index)]) 
        
    for k in range(len(index)):
        GX3_Tbias[k] = GX3_GCss1 + GX3_TP[k] - GX3_TPss1        # This is temperature polynomial adjusted in counter value to match gravity 
        GY3_Tbias[k] = GY3_GCss1 + GY3_TP[k] - GY3_TPss1        # compensated bias in the first standstill period 
        GZ3_Tbias[k] = GZ3_GCss1 + GZ3_TP[k] - GZ3_TPss1
    
    GX3_Tbias_ss2 = np.average(GX3_Tbias[(len(index)-ss2_length):len(index)])     # Temperature commpensated bias value, average at second standstill 
    GY3_Tbias_ss2 = np.average(GY3_Tbias[(len(index)-ss2_length):len(index)]) 
    GZ3_Tbias_ss2 = np.average(GZ3_Tbias[(len(index)-ss2_length):len(index)])
    
    aX3 = (GX3_GCss2 - GX3_Tbias_ss2)/(len(index) - ss1_length - ss2_length)       # slope for adjusting remaining bias (after g and temp comp) linearly
    aY3 = (GY3_GCss2 - GY3_Tbias_ss2)/(len(index) - ss1_length - ss2_length)
    aZ3 = (GZ3_GCss2 - GZ3_Tbias_ss2)/(len(index) - ss1_length - ss2_length)
    
    for k in range(ss1_length):
        GX3bias[k] = GX3count[k]
        GY3bias[k] = GY3count[k]
        GZ3bias[k] = GZ3count[k]
    
    for k in range(len(index) - ss1_length - ss2_length):
        GX3bias[k+ss1_length] = GX3_Tbias[k+ss1_length] + aX3*k
        GY3bias[k+ss1_length] = GY3_Tbias[k+ss1_length] + aY3*k
        GZ3bias[k+ss1_length] = GZ3_Tbias[k+ss1_length] + aZ3*k
    
    for k in range (ss2_length):
        GX3bias[k + len(index) - ss2_length] = GX3count[k + len(index) - ss2_length]
        GY3bias[k + len(index) - ss2_length] = GY3count[k + len(index) - ss2_length]
        GZ3bias[k + len(index) - ss2_length] = GZ3count[k + len(index) - ss2_length]
    
    return GX3bias, GY3bias, GZ3bias


def Tcomp_gyro_pick_average_3G(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time, f_s_acc, timeabs, \
                               GX3count, GY3count, GZ3count, TG3count, TcoeffGX3, TcoeffGY3, TcoeffGZ3):
    
    GX3pick_TCav = np.zeros(len(meas_point))
    GY3pick_TCav = np.zeros(len(meas_point))
    GZ3pick_TCav = np.zeros(len(meas_point))
          
    startindex = (starttime_Tcomp-standstill_time) * f_s_acc
    stopindex = (stoptime_Tcomp+standstill_time) * f_s_acc
    standstill_samples = standstill_time * f_s_acc 
    
    indexforTC = index[startindex:stopindex]
    GX3countforTC = GX3count[startindex:stopindex]
    GY3countforTC = GY3count[startindex:stopindex]
    GZ3countforTC = GZ3count[startindex:stopindex]
    TG3countforTC = TG3count[startindex:stopindex]
    
    GX3bias, GY3bias, GZ3bias = gyro_bias_comp_3G(indexforTC, GX3countforTC, GY3countforTC, GZ3countforTC, \
                TG3countforTC, standstill_samples, standstill_samples, TcoeffGX3, TcoeffGY3, TcoeffGZ3) 
   
    GX3countTC = GX3count
    GY3countTC = GY3count
    GZ3countTC = GZ3count
    
    GX3countTC[startindex:stopindex] = GX3countTC[startindex:stopindex] - GX3bias
    GY3countTC[startindex:stopindex] = GY3countTC[startindex:stopindex] - GY3bias
    GZ3countTC[startindex:stopindex] = GZ3countTC[startindex:stopindex] - GZ3bias
        
    for n in range(len(meas_point)):
        k = int(f_s_acc * timeabs[n])
      
        GX3pick_TCav[n] = np.mean(GX3countTC[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GY3pick_TCav[n] = np.mean(GY3countTC[(int(k-av_points/2))-1:(int(k+av_points/2))])
        GZ3pick_TCav[n] = np.mean(GZ3countTC[(int(k-av_points/2))-1:(int(k+av_points/2))])
        TG3pick_av[n] = np.mean(TG3count[(int(k-av_points/2))-1:(int(k+av_points/2))])
    
    
    return GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav



def gsens_calc_3G(AX, AY, AZ, GX3, GY3, GZ3, TG3pick_av, figno1, figno2, figno3, plotfigs):
    
    if plotfigs == True:
        plt.close(figno1)
        plt.close(figno2)
        plt.close(figno3)
    
    
    rows = len(AX)
    M = np.empty((rows,3))
    MT = np.empty((3, rows))
    MTM = np.empty((3, 3))
    invMTM = np.empty((3, 3))
    invMTM_MT = np.empty((3, rows))
    F = np.empty((rows,3))
    OT = np.empty((3, 3))
    
    M[:,0] = AX
    M[:,1] = AY
    M[:,2] = AZ
    
    F[:,0] = GX3
    F[:,1] = GY3
    F[:,2] = GZ3
 
    MT = np.transpose(M)
    MTM = np.matmul(MT, M)
    invMTM = np.linalg.inv(MTM)
    invMTM_MT = np.matmul(invMTM, MT)
    OT = np.matmul(invMTM_MT, F)
    GX3factors = OT[:, 0]
    GY3factors = OT[:, 1]
    GZ3factors = OT[:, 2]
        
    GX3gsens = np.zeros(rows)
    GY3gsens = np.zeros(rows)
    GZ3gsens = np.zeros(rows)
   
    TG3 = np.average(TG3pick_av)
    
    for k in range(rows):
        GX3gsens[k] = GX3factors[0]*AX[k] + GX3factors[1]*AY[k] + GX3factors[2]*AZ[k]
        GY3gsens[k] = GY3factors[0]*AX[k] + GY3factors[1]*AY[k] + GY3factors[2]*AZ[k]
        GZ3gsens[k] = GZ3factors[0]*AX[k] + GZ3factors[1]*AY[k] + GZ3factors[2]*AZ[k]
    
    abstempher = int(abstemp)
    
    if plotfigs == True:
        plt.figure(figno1, figsize=(12,9))
        plt.plot(meas_point, GX3, label = 'GX3')

        plt.plot(meas_point, GX3gsens, label = 'GX3poly')
        plt.legend(loc = 'best')
        plt.legend(loc = 'best')
        plt.grid()
        plt.savefig(toolSN + ' secondary '+ str(abstempher)+ 'C plot' + str(figno1) + ' GX3 gsens.png')
    #    plt.show()    
        
        plt.figure(figno2, figsize=(12,9))
        plt.plot(meas_point, GY3, label = 'GY3')

        plt.plot(meas_point, GY3gsens, label = 'GY3poly')
        plt.legend(loc = 'best')
        plt.legend(loc = 'best')
        plt.grid()
        plt.savefig(toolSN + ' secondary '+ str(abstempher)+ 'C plot' + str(figno2) + ' GY3 gsens.png')
    #    plt.show() 
        
        plt.figure(figno3, figsize=(12,9))
        plt.plot(meas_point, GZ3, label = 'GZ3')

        plt.plot(meas_point, GZ3gsens, label = 'GZ3poly')
        plt.legend(loc = 'best')
        plt.legend(loc = 'best')
        plt.grid()
        plt.savefig(toolSN +' secondary '+ str(abstempher)+ 'C plot' + str(figno3) + ' GZ3 gsens.png')
    #    plt.show() 
       
    return GX3factors, GY3factors, GZ3factors, TG3

def save_file_gsens_3G(filename, sheetname, gsensGX3, gsensGY3, gsensGZ3, TG3, abstemp):
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet(sheetname)
    
    worksheet.write(0,0,'GX3 coeffs')
    worksheet.write(0,1,'GY3 coeffs')
    worksheet.write(0,2,'GZ3 coeffs')
    worksheet.write(0,3,'TG3 count')
    worksheet.write(0,4, 'Cal temp (degC)')
    worksheet.write(1,3, TG3)
    worksheet.write(1,4, abstemp)
    
    
    for k in range(3):
        worksheet.write(k+1, 0, gsensGX3[k])
        worksheet.write(k+1, 1, gsensGY3[k]) 
        worksheet.write(k+1, 2, gsensGZ3[k]) 
    
    workbook.close()     
    return

def read_file_gsens_3G(filename, sheetname):
    
    # read data file 
    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_name(sheetname)
       
    # extract cal data
    
    col = worksheet.col_values(0)
    GX3_g_coeff = col[1:4]
    GX3_g_coeff = np.array(GX3_g_coeff)
    
    col = worksheet.col_values(1)
    GY3_g_coeff = col[1:4]
    GY3_g_coeff = np.array(GY3_g_coeff)

    col = worksheet.col_values(2)
    GZ3_g_coeff = col[1:4]
    GZ3_g_coeff = np.array(GZ3_g_coeff)
    
   
    return GX3_g_coeff, GY3_g_coeff, GZ3_g_coeff


def calculate_gyro_gbias_3G(GX3count, GY3count, GZ3count, GX3factors, GY3factors, GZ3factors, AXcal, AYcal, AZcal):
    
    GX3g_bias = np.zeros(len(GX3count))
    GY3g_bias = np.zeros(len(GX3count))
    GZ3g_bias = np.zeros(len(GX3count))
    
    for k in range(len(GX3count)):
            
        GX3g_bias[k] = GX3factors[0]*AXcal[k] + GX3factors[1]*AYcal[k] + GX3factors[2]*AZcal[k]
        GY3g_bias[k] = GY3factors[0]*AXcal[k] + GY3factors[1]*AYcal[k] + GY3factors[2]*AZcal[k]
        GZ3g_bias[k] = GZ3factors[0]*AXcal[k] + GZ3factors[1]*AYcal[k] + GZ3factors[2]*AZcal[k]
    
    return GX3g_bias, GY3g_bias, GZ3g_bias
    
    
    
def plot_figures_ps4_3G(time, GX3count, GY3count, GZ3count, meas_point, GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav):
    
#    plt.close(141)
#    #plt.close(142)
#    
#    plt.figure(141)
#    plt.plot(time, GX3count, label = 'GX3countTC')

#    plt.plot(time, GY3count, label = 'GY3countTC')
#    plt.plot(time, GZ3count, label = 'GZ3countTC')
#    plt.legend(loc = 'best')
#    plt.legend(loc = 'best')
#    plt.grid()
#    plt.show()
    
#    plt.figure(142)
#    plt.plot(meas_point, GX3pick_TCav, label = 'GX3avTC')

#    plt.plot(meas_point, GY3pick_TCav, label = 'GY3avTC')
#    plt.plot(meas_point, GZ3pick_TCav, label = 'GZ3avTC')
#    plt.legend(loc = 'best')
#    plt.legend(loc = 'best')
#    plt.grid()
#    plt.show()
    return

def read_file_gyroTcoeff_3G(filename, sheetname):
    
    # read data file 
    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_name(sheetname)
    
    TcoeffGX3 = np.zeros(6)
    TcoeffGY3 = np.zeros(6)
    TcoeffGZ3 = np.zeros(6)
    
    offset = 1
    # extract index
    col = worksheet.col_values(0)
    TcoeffGX3 = col[offset:]
    TcoeffGX3 = np.array(TcoeffGX3)
    
    col = worksheet.col_values(1)
    TcoeffGY3 = col[offset:]
    TcoeffGY3 = np.array(TcoeffGY3)
    
    col = worksheet.col_values(2)
    TcoeffGZ3 = col[offset:]
    TcoeffGZ3 = np.array(TcoeffGZ3)
       
    return TcoeffGX3, TcoeffGY3, TcoeffGZ3

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

def gyro_rot_pick_3G(meas_point, timeabs, rotrate, av_time_ss, av_time_rot, margin_ss , margin_rot, f_s, \
                  accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TG3count, rot):
    
    
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
    
    
    
    GX3TP = np.zeros(len(accXcount))
    GY3TP = np.zeros(len(accXcount))
    GZ3TP = np.zeros(len(accXcount))
    
    time = np.zeros(len(accXcount))
    
    GX3av = np.zeros(len(meas_point)-1)
    GY3av = np.zeros(len(meas_point)-1)
    GZ3av = np.zeros(len(meas_point)-1)
    TG3av = np.zeros(len(meas_point)-1)
    
    
    # calibrate accelerometer data
    xoff, yoff, zoff, xfactor, yfactor, zfactor, x3rd, y3rd, z3rd = read_file_cal_constants(acc_cal_file, acc_cal_sheet)
    AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(accXcount, accYcount, accZcount, xfactor, yfactor, zfactor)
    
    # read g sensitivity constants from file
    GX3_g_coeff, GY3_g_coeff, GZ3_g_coeff = read_file_gsens_3G(gyrogsenscal_file, gyrogsenscal_sheet)
    
    # calculate gyro g bias
    GX3_gbias, GY3_gbias, GZ3_gbias = calculate_gyro_gbias_3G(GX3count, GY3count, GZ3count, \
                                        GX3_g_coeff, GY3_g_coeff, GZ3_g_coeff, AXcal_lin, AYcal_lin, AZcal_lin)
    
    # subtract bias 
    GX3count = GX3count - GX3_gbias
    GY3count = GY3count - GY3_gbias
    GZ3count = GZ3count - GZ3_gbias
    
    # read temperature coefficients from file 
    TcoeffGX3, TcoeffGY3, TcoeffGZ3 = read_file_gyroTcoeff_3G(gyroTcal_file, gyroTcal_sheet)
    
    for k in range(len(accXcount)):
        
        GX3TP[k] = TcoeffGX3[5] + TcoeffGX3[4]*TG3count[k] + TcoeffGX3[3]*TG3count[k]**2 + TcoeffGX3[2]*TG3count[k]**3  \
                    + TcoeffGX3[1]*TG3count[k]**4 + TcoeffGX3[0]*TG3count[k]**5
        GY3TP[k] = TcoeffGY3[5] + TcoeffGY3[4]*TG3count[k] + TcoeffGY3[3]*TG3count[k]**2 + TcoeffGY3[2]*TG3count[k]**3  \
                    + TcoeffGY3[1]*TG3count[k]**4 + TcoeffGY3[0]*TG3count[k]**5
        GZ3TP[k] = TcoeffGZ3[5] + TcoeffGZ3[4]*TG3count[k] + TcoeffGZ3[3]*TG3count[k]**2 + TcoeffGZ3[2]*TG3count[k]**3  \
                    + TcoeffGZ3[1]*TG3count[k]**4 + TcoeffGZ3[0]*TG3count[k]**5
    
    
    GX3TP0 = np.average(GX3count[int(f_s*(timeabs[1]-15)):int(f_s*(timeabs[1]-5))])
    GY3TP0 = np.average(GY3count[int(f_s*(timeabs[1]-15)):int(f_s*(timeabs[1]-5))])
    GZ3TP0 = np.average(GZ3count[int(f_s*(timeabs[1]-15)):int(f_s*(timeabs[1]-5))])
    
    for k in range(len(accXcount)):
        GX3count[k] = GX3count[k] -(GX3TP[k]-GX3TP0)
        GY3count[k] = GY3count[k] -(GY3TP[k]-GY3TP0)
        GZ3count[k] = GZ3count[k] -(GZ3TP[k]-GZ3TP0)
    
    for k in range(len(meas_point)-1):
        
        if abs(rotrate[k]) < 0.01:  # standstill
            GX3av[k] = np.average(GX3count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])
            GY3av[k] = np.average(GY3count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])
            GZ3av[k] = np.average(GZ3count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])
            TG3av[k] = np.average(TG3count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)): \
                     int(f_s*(timeabs[k+1] - margin_ss))-1])

        if abs(rotrate[k]) > 0.01:  # rotation
            GX3av[k] = np.average(GX3count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])
            GY3av[k] = np.average(GY3count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])
            GZ3av[k] = np.average(GZ3count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])
            TG3av[k] = np.average(TG3count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)): \
                     int(f_s*(timeabs[k+1] - margin_rot))-1])
   
    
    for n in range(len(GX3count)):
        time[n] = n/f_s
    
    
    plotstartindex = int(f_s*(timeabs[1] - 30))
    plotstopindex = int(f_s*(timeabs[5] + 30))
    
    plt.figure(figno, figsize=(12,9))
    plt.plot(time[plotstartindex:plotstopindex], GX3count[plotstartindex:plotstopindex], 'C0', label = 'GX3count')

    plt.plot(time[plotstartindex:plotstopindex], GY3count[plotstartindex:plotstopindex], 'C1', label = 'GY3count')
    plt.plot(time[plotstartindex:plotstopindex], GZ3count[plotstartindex:plotstopindex], 'C3', label = 'GZ3count')
    

    for k in range(len(meas_point)-1):
        
        if abs(rotrate[k]) < 0.01:  # standstill
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], \
                GX3count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], 'C4')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], \
                GY3count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], 'C5')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], \
                GZ3count[int(f_s*(timeabs[k+1] - margin_ss - av_time_ss)) : int(f_s*(timeabs[k+1] - margin_ss))-1], 'C7')
            
            
            
            
        if abs(rotrate[k]) > 0.01:  # rotation
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], \
                GX3count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], 'C4')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], \
                GY3count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], 'C5')
            plt.plot(time[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], \
                GZ3count[int(f_s*(timeabs[k+1] - margin_rot - av_time_rot)) : int(f_s*(timeabs[k+1] - margin_rot))-1], 'C7')
 
    plt.ylim((32000,33500))  
    plt.legend(loc = 'best')
    plt.legend(loc = 'best')
    plt.grid()
    abstempher = int(abstemp)
    if rot == 1:
        plt.savefig(toolSN + ' secondary '+ str(abstempher) + 'C plot' + str(figno) + ' gyro Xrot.png')
    if rot == 2:
        plt.savefig(toolSN + ' secondary '+ str(abstempher) + 'C plot' + str(figno) + ' gyro Yrot.png')
    if rot == 3:
        plt.savefig(toolSN + ' secondary '+ str(abstempher) + 'C plot' + str(figno) + ' gyro Zrot.png')
#    plt.show()
    
#    plt.figure(figno2)
#    plt.plot(GX3av, 'C0o', label = 'GX3av')

#    plt.plot(GY3av, 'C1o', label = 'GY3av')
#    plt.plot(GZ3av, 'C3o', label = 'GZ3av')
#    plt.legend(loc = 'best')
#    plt.legend(loc = 'best')
#    plt.grid()
#    plt.show()
    
    return GX3av, GY3av, GZ3av, TG3av


def calc_gyro_ortomatrix_3G(GX3avXr, GY3avXr, GZ3avXr, TG3avXr, rotrateXr, GX3avYr, GY3avYr, GZ3avYr, TG3avYr, \
                            rotrateYr, GX3avZr, GY3avZr, GZ3avZr, TG3avZr, rotrateZr):
    
    xelements = 0
    yelements = 0
    zelements = 0
    
    xindex = np.zeros(len(GX3avXr))
    yindex = np.zeros(len(GX3avYr))
    zindex = np.zeros(len(GX3avZr))
    
    n=0
    for k in range(len(GX3avXr)):
        if abs(rotrateXr[k]) > 0.1:
            xelements = xelements + 1
            xindex[n] = k
            n = n+1
    n=0        
    for k in range(len(GY3avYr)):
        if abs(rotrateYr[k]) > 0.1:
            yelements = yelements + 1
            yindex[n] = k
            n = n+1
    
    n=0         
    for k in range(len(GZ3avZr)):
        if abs(rotrateZr[k]) > 0.1:
            zelements = zelements + 1
            zindex[n] = k
            n = n+1
    
    rows = xelements + yelements + zelements 
    

    M = np.empty((rows, 3))
    F = np.empty((rows, 3))
    MT = np.empty((3, rows))
    MTM = np.empty((3, 3))
    invMTM = np.empty((3, 3))
    invMTM_MT = np.empty((3, rows))
    OT = np.empty((3, 3))
    
    for k in range(xelements):                                                  # X rotation                   
        ind = int(xindex[k])   
        M[k,0] = GX3avXr[ind] - (GX3avXr[ind-1] + GX3avXr[ind+1])/2
        M[k,1] = GY3avXr[ind] - (GY3avXr[ind-1] + GY3avXr[ind+1])/2
        M[k,2] = GZ3avXr[ind] - (GZ3avXr[ind-1] + GZ3avXr[ind+1])/2
        
        F[k,0] = rotrateXr[ind]
        F[k,1] = 0
        F[k,2] = 0
    
    
            
    for k in range(yelements): 
        ind = int(yindex[k])                                          # Y rotation                        

        M[k+xelements, 0] = GX3avYr[ind] - (GX3avYr[ind-1] + GX3avYr[ind+1])/2
        
        M[k + xelements, 1] = GY3avYr[ind] - (GY3avYr[ind-1] + GY3avYr[ind+1])/2
        M[k + xelements, 2] = GZ3avYr[ind] - (GZ3avYr[ind-1] + GZ3avYr[ind+1])/2
        
        F[k + xelements, 0] = 0
        F[k + xelements, 1] = rotrateYr[ind]
        F[k + xelements, 2] = 0
    
    for k in range(zelements): 
        ind = int(zindex[k])                                          # Z rotation                        

        M[k+xelements + yelements, 0] = GX3avZr[ind] - (GX3avZr[ind-1] + GX3avZr[ind+1])/2
        M[k + xelements + yelements, 1] = GY3avZr[ind] - (GY3avZr[ind-1] + GY3avZr[ind+1])/2
        M[k + xelements + yelements, 2] = GZ3avZr[ind] - (GZ3avZr[ind-1] + GZ3avZr[ind+1])/2
        
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

    xfactors = OT[:, 0]
    yfactors = OT[:, 1]
    zfactors = OT[:, 2]    

            
    TG3 = (np.average(TG3avXr) + np.average(TG3avYr) + np.average(TG3avZr))/3

       
    return xfactors, yfactors, zfactors, TG3


def save_file_gyro_orto_3G(filename, sheetname, xfactors, yfactors, bestGX, TG3, abstemp):
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet(sheetname)
    
      
    worksheet.write(0, 0, 'X coeffs')
    worksheet.write(0, 1, 'Y coeffs')
    worksheet.write(0, 2, 'Z coeffs')

    worksheet.write(0, 3, 'TG3 count')
    worksheet.write(0, 4, 'Cal temp (deg C)')
    worksheet.write(1, 3, TG3)
    worksheet.write(1, 4, abstemp)


    for k in range(3):
        worksheet.write(k+1, 0, xfactors[k])
        worksheet.write(k+1, 1, yfactors[k]) 
        worksheet.write(k+1, 2, zfactors[k]) 
    
    workbook.close()     
    return

def read_file_gyro_orto_3G(filename, sheetname):
    
    # read data file 
    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_name(sheetname)
    
    xfactors = np.zeros(3)
    yfactors = np.zeros(3)
    zfactors = np.zeros(3)
    
    offset = 1
    # extract index
    col = worksheet.col_values(0)
    xfactors = col[offset:]
    xfactors = np.array(xfactors)
    
    col = worksheet.col_values(1)
    yfactors = col[offset:]
    yfactors = np.array(yfactors)   
   
    col = worksheet.col_values(2)
    zfactors = col[offset:]
    zfactors = np.array(zfactors)
      
    col = worksheet.col_values(3)
    TG3 = col[1]
    TG3 = float(TG3)
      
    col = worksheet.col_values(4)
    abstemp = col[1]
    abstemp = float(abstemp)
    
    return xfactors, yfactors, zfactors, TG3, abstemp
            

# returns unit vectors for each gyro sensor along sensitive axis in tool body coordinate system. Also returns scale
# factor for each gyro sensor in units of (counts/(deg/s))
            
def gyro_body_vec_3G(xfactors, yfactors, zfactors):
    
    O = np.empty((3,3))
    Oinv = np.empty((3,3))
    
    GX3point = np.zeros(3)
    GY3point = np.zeros(3)
    GZ3point = np.zeros(3)    
    
    
    O[0, :] = xfactors
    O[1, :] = yfactors
    O[2, :] = zfactors
    
    Oinv = np.linalg.inv(O)

    
    Xrot = [1,0,0]
    Yrot = [0,1,0]
    Zrot = [0,0,1]
    
    GX1Xrot = np.matmul(Oinv, Xrot)
    GX1Yrot = np.matmul(Oinv, Yrot)
    GX1Zrot = np.matmul(Oinv, Zrot)
                 
    GX3point = [GX1Xrot[0], GX1Yrot[0], GX1Zrot[0]]
    GX3point = np.array(GX3point)
    GX3scale = np.sqrt((GX3point[0])**2 + (GX3point[1])**2 + (GX3point[2])**2)
    GX3point = GX3point/GX3scale
    
    GY3point = [GX1Xrot[1], GX1Yrot[1], GX1Zrot[1]]
    GY3point = np.array(GY3point)
    GY3scale = np.sqrt((GY3point[0])**2 + (GY3point[1])**2 + (GY3point[2])**2)
    GY3point = GY3point/GY3scale
   
    GZ3point = [GX1Xrot[2], GX1Yrot[2], GX1Zrot[2]]
    GZ3point = np.array(GZ3point)
    GZ3scale = np.sqrt((GZ3point[0])**2 + (GZ3point[1])**2 + (GZ3point[2])**2)
    GZ3point = GZ3point/GZ3scale
        
    return GX3point, GY3point, GZ3point, GX3scale, GY3scale, GZ3scale


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
    
def er_comp_gyros_3G(GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav, azimuth, inclination, toolface, \
                  GX3point, GY3point, GZ3point, GX3scale, GY3scale, GZ3scale):
    
    GX3_er_comp = np.zeros(len(GX3pick_TCav))
    GY3_er_comp = np.zeros(len(GX3pick_TCav))
    GZ3_er_comp = np.zeros(len(GX3pick_TCav))
    
    for k in range(len(GX3pick_TCav)):
        earthrot =  earth_rot_vec(latitude, azimuth[k], inclination[k], toolface[k])
        GX3_er_comp[k] = np.dot(earthrot, GX3point) * (earthrate/3600) * GX3scale
        GY3_er_comp[k] = np.dot(earthrot, GY3point) * (earthrate/3600) * GY3scale
        GZ3_er_comp[k] = np.dot(earthrot, GZ3point) * (earthrate/3600) * GZ3scale
        #print k, earthrot
    return GX3_er_comp, GY3_er_comp, GZ3_er_comp




# main program starts here

if sequencemode == 0: 

    if processtep == 1:
        
        #read data from raw data file 
        index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count \
            = readrawdata_3G(logfile_acc_gsens, sheetname_acc_gsens)
        
        time = index / f_s_acc
        av_points = av_time * f_s_acc
        av_pointsAC = av_timeAC * f_s_acc
        
        #read data from timing file
        meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
        
     
        # read data from timing file for control signal
        
        if controlseq == True:
            meas_pointAC, timegapAC, timeabsAC, AXRefAC, AYRefAC, AZRefAC = readtimefile(file_path_timefile, sheetname_control)
        
        
        # pick correctly timed value from counter sequence. Adjust timing file as necessary   
    
        
        AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX3pick, GY3pick, GZ3pick, GX3pick_av, GY3pick_av, GZ3pick_av, \
        AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, GZ3pick_std, TG3pick_av, TApick_av \
        = acc_gsens_pick_3G(meas_point, av_time,  f_s_acc, timeabs, accXcount, accYcount, accZcount, TAcount, GX3count, GY3count, GZ3count, TG3count)
    
        if controlseq == True:
            AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, GX3pickAC, GY3pickAC, \
                GZ3pickAC, GX3pick_avAC, GY3pick_avAC, GZ3pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, \
                GX3pick_stdAC, GY3pick_stdAC, GZ3pick_stdAC, TG3pick_avAC, TApick_avAC \
                = acc_control_pick(meas_pointAC, av_timeAC, timeabsAC, f_s_acc, accXcount, accYcount, accZcount, TAcount, \
                                 GX3count, GY3count, GZ3count, TG3count)
    
        # QC plots for this sequence
        plot_figures_ps1_3G(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                         timeabs, av_points, GX3count, GY3count, GZ3count, GX3pick, GY3pick, GZ3pick, GX3pick_av, GY3pick_av, GZ3pick_av, \
                         AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, GZ3pick_std, TG3pick_av, TApick_av)       
        
        if controlseq == True:
            plot_figures_ps1AC(time, accXcount, accYcount, accZcount, AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, meas_pointAC, \
                     timeabsAC, av_pointsAC, GX3count, GY3count, GZ3count, GX3pickAC, GY3pickAC, GZ3pickAC, GX3pick_avAC, \
                     GY3pick_avAC, GZ3pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, GX3pick_stdAC, GY3pick_stdAC, \
                     GZ3pick_stdAC, TG3pick_avAC, TApick_avAC)
    
        # write selected and averaged data to file
        save_file_ps1_3G(pickfile_write, picksheet_write, meas_point, AXpick_av, AYpick_av, AZpick_av, \
                         GX3pick_av, GY3pick_av, GZ3pick_av, TG3pick_av, TApick_av, AXRef, AYRef, AZRef)
        
        if controlseq == True:  
            save_file_ps1AC(pickfile_writeAC, picksheet_writeAC, meas_pointAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, TApick_avAC, AXRefAC, AYRefAC, AZRefAC)
        
    
        
        
        if autostarttime == True:
            starttime_Tcomp = int(timeabs[0] - 1520)
            print ('starttime_Tcomp =', starttime_Tcomp)
            
            
            stoptime_Tcomp = int(timeabs[107] + 100)
            if abstemp == 28:
                stoptime_Tcomp = int(timeabs[107] + 450)
            print ('stoptime_Tcomp =', stoptime_Tcomp)
            
        
    
    if processtep == 2:
        
        # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
        meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef \
        = read_file_ps1_3G(pickfile_read , picksheet_read)
        
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
        index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count = readrawdata_3G(tempfile_read, tempsheet_read)
     
        # calculate 5th order polynominal
        TcoeffGX3, TcoeffGY3, TcoeffGZ3, TcoeffGX3lin, TcoeffGY3lin, TcoeffGZ3lin, TG3min, TG3max = \
            find_gyro_temp_poly_3G(index, GX3count, GY3count, GZ3count, TG3count, f_s_temp)
        
        save_file_gyroT_constants_3G(gyroTcal_file, gyroTcal_sheet, TcoeffGX3, TcoeffGY3, TcoeffGZ3, \
                                  TcoeffGX3lin, TcoeffGY3lin, TcoeffGZ3lin,TG3min, TG3max)
        
        
    
    if processtep == 4:
        
        # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
        meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef \
            = read_file_ps1_3G(pickfile_read, picksheet_read)
        
        # read data from raw data file - needed for temperature compensation
        index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count \
            = readrawdata_3G(logfile_acc_gsens, sheetname_acc_gsens)
        
        # read data from timing file
        meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
        
        # read temperature coefficients from file 
        TcoeffGX3, TcoeffGY3, TcoeffGZ3 = read_file_gyroTcoeff_3G(gyroTcal_file, gyroTcal_sheet)
        
        time = index / f_s_acc
        av_points = av_time * f_s_acc
        
        
        # temperature compensate gyro counters and pick same time instants as before
        GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav = Tcomp_gyro_pick_average_3G(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
            f_s_acc, timeabs, GX3count, GY3count, GZ3count, TG3count, TcoeffGX3, TcoeffGY3, TcoeffGZ3)
        
    
        # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
        xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin \
            = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
        
        # calibrate picked accelerometer data, classical method
        AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
        
        # calculate gravity sensitivity, ignoring earth rotation. Use in gyro scale factor calibration. Plots 
        GX3gsens, GY3gsens, GZ3gsens, TG3 = gsens_calc_3G(AXcal_lin, AYcal_lin, AZcal_lin, GX3pick_TCav, GY3pick_TCav, \
                                                GZ3pick_TCav, TG3pick_av, 131, 132, 133, plotfigs1)        
        
        # plot QC data, 
       
        plot_figures_ps4_3G(time, GX3count, GY3count, GZ3count, meas_point, GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav)
       
           
        save_file_gsens_3G(gyrogsenscal_file, gyrogsenscal_sheet, GX3gsens, GY3gsens, GZ3gsens, TG3, abstemp)
        
    
    if processtep == 5:
        
        # read data file for X rotation
        indexXr, accXcountXr, accYcountXr, accZcountXr, GX3countXr, GY3countXr, GZ3countXr, TAcountXr, TG3countXr \
            = readrawdata_3G(Xrotfile, Xrotsheet)
        
        meas_pointXr, timegapXr, timeabsXr, rotrateXr = read_rot_timefile(rot_timefile, Xrot_timesheet)
        
        rot = 1 # X rotation
        
        # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
    
        GX3avXr, GY3avXr, GZ3avXr, TG3avXr = gyro_rot_pick_3G(meas_pointXr, timeabsXr, rotrateXr, av_time_standstill, av_time_rot, \
                                margin_standstill, margin_rotation, f_s_rot, accXcountXr, accYcountXr, accZcountXr, GX3countXr, \
                                GY3countXr, GZ3countXr, TG3countXr, rot)
        
        
       
        # read data file for Y rotation
        indexYr, accXcountYr, accYcountYr, accZcountYr, GX3countYr, GY3countYr, GZ3countYr, TAcountYr, TG3countYr \
                    = readrawdata_3G(Yrotfile, Yrotsheet)
        
        meas_pointYr, timegapYr, timeabsYr, rotrateYr = read_rot_timefile(rot_timefile, Yrot_timesheet)
        
        rot = 2 # Y rotation
        
        # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
        
    
        GX3avYr, GY3avYr, GZ3avYr, TG3avYr = gyro_rot_pick_3G(meas_pointYr, timeabsYr, rotrateYr, av_time_standstill, av_time_rot, \
                                margin_standstill, margin_rotation, f_s_rot, accXcountYr, accYcountYr, accZcountYr, GX3countYr, \
                                GY3countYr, GZ3countYr, TG3countYr, rot)
        
        
        # read data file for Z rotation
        indexZr, accXcountZr, accYcountZr, accZcountZr, GX3countZr, GY3countZr, GZ3countZr, TAcountZr, TG3countZr \
            = readrawdata_3G(Zrotfile, Zrotsheet)
        
        meas_pointZr, timegapZr, timeabsZr, rotrateZr = read_rot_timefile(rot_timefile, Zrot_timesheet)
        
        rot = 3 # Z rotation
        
        # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile 
     
        GX3avZr, GY3avZr, GZ3avZr, TG3avZr = gyro_rot_pick_3G(meas_pointZr, timeabsZr, rotrateZr, av_time_standstill, av_time_rot, \
                            margin_standstill, margin_rotation, f_s_rot, accXcountZr, accYcountZr, accZcountZr, GX3countZr, \
                            GY3countZr, GZ3countZr, TG3countZr, rot)
            
        
        
        
        # calculate gyro ortomatrix
        
        xfactors, yfactors, zfactors, TG3 = calc_gyro_ortomatrix_3G(GX3avXr, GY3avXr, GZ3avXr, TG3avXr, rotrateXr, GX3avYr, GY3avYr, \
                                  GZ3avYr, TG3avYr, rotrateYr, GX3avZr, GY3avZr, GZ3avZr, TG3avZr, rotrateZr)
    
        # write gyro ortomatrix to file. 
        
        save_file_gyro_orto_3G(gyro_orto_file, gyro_orto_sheet, xfactors, yfactors, zfactors, TG3, abstemp)
        
    
    if processtep == 6:
        
        # read gyro scale factor matrix
        xfactors, yfactors, zfactors, TG3, abstemp = read_file_gyro_orto_3G(gyro_orto_file, gyro_orto_sheet)
        
        # find gyro sensors unit pointing vectors (along sensitive axis) from gyro orto matrix
        GX3point, GY3point, GZ3point, GX3scale, GY3scale, GZ3scale = \
                    gyro_body_vec_3G(xfactors, yfactors, zfactors)
        
        # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
        meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef \
            = read_file_ps1_3G(pickfile_read, picksheet_read)
        
        # read data from raw data file - needed for temperature compensation
        index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count \
            = readrawdata_3G(logfile_acc_gsens, sheetname_acc_gsens)
        
        # read data from timing file
        meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
        
        # read temperature coefficients from file 
        TcoeffGX3, TcoeffGY3, TcoeffGZ3 = read_file_gyroTcoeff_3G(gyroTcal_file, gyroTcal_sheet)
        
        time = index / f_s_acc
        av_points = av_time * f_s_acc
        
        # temperature compensate gyro counters and pick same time instants as before
    
        GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav = Tcomp_gyro_pick_average_3G(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                f_s_acc, timeabs, GX3count, GY3count, GZ3count, TG3count, TcoeffGX3, TcoeffGY3, TcoeffGZ3)
        
    
        # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
        xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin \
        = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
        
        # calibrate picked accelerometer data, classical method
        AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
        
        # read data from angle file (same as timing file)
        azimuth, inclination, toolface = readanglefile(file_path_timefile, sheetname_time)
        
        # calculate gyro counter values resulting from earth rotation
        GX3_er_comp, GY3_er_comp, GZ3_er_comp = er_comp_gyros_3G(GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav, \
                        azimuth, inclination, toolface, GX3point, GY3point, GZ3point, GX3scale, GY3scale, GZ3scale)
        
        # adjust gyro counter values for earth rotation
        GX3pick_TCav = GX3pick_TCav - GX3_er_comp
        GY3pick_TCav = GY3pick_TCav - GY3_er_comp
        GZ3pick_TCav = GZ3pick_TCav - GZ3_er_comp
        
        
        # calculate gravity sensitivity, compensating for earth rotation. Use in later data calibration. Plots 
        GX3gsens_erc, GY3gsens_erc, GZ3gsens_erc , TG3 = gsens_calc_3G(AXcal_lin, AYcal_lin, AZcal_lin, GX3pick_TCav, \
                                    GY3pick_TCav, GZ3pick_TCav, TG3pick_av, 171, 172, 173, plotfigs2)        
        
        save_file_gsens_3G(gyrogsenscal_erc_file, gyrogsenscal_erc_sheet, GX3gsens_erc, GY3gsens_erc, GZ3gsens_erc, TG3, abstemp)



if sequencemode ==1:
    
    processtep = 1
            
    if processtep == 1:
        
         #read data from raw data file 
        index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count \
            = readrawdata_3G(logfile_acc_gsens, sheetname_acc_gsens)
        
        time = index / f_s_acc
        av_points = av_time * f_s_acc
        av_pointsAC = av_timeAC * f_s_acc
        
        #read data from timing file
        meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
        
     
        # read data from timing file for control signal
        
        if controlseq == True:
            meas_pointAC, timegapAC, timeabsAC, AXRefAC, AYRefAC, AZRefAC = readtimefile(file_path_timefile, sheetname_control)
        
        
        # pick correctly timed value from counter sequence. Adjust timing file as necessary   
    
        
        AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX3pick, GY3pick, GZ3pick, GX3pick_av, GY3pick_av, GZ3pick_av, \
        AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, GZ3pick_std, TG3pick_av, TApick_av \
        = acc_gsens_pick_3G(meas_point, av_time,  f_s_acc, timeabs, accXcount, accYcount, accZcount, TAcount, GX3count, GY3count, GZ3count, TG3count)
    
        if controlseq == True:
            AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, GX3pickAC, GY3pickAC, \
                GZ3pickAC, GX3pick_avAC, GY3pick_avAC, GZ3pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, \
                GX3pick_stdAC, GY3pick_stdAC, GZ3pick_stdAC, TG3pick_avAC, TApick_avAC \
                = acc_control_pick(meas_pointAC, av_timeAC, timeabsAC, f_s_acc, accXcount, accYcount, accZcount, TAcount, \
                                 GX3count, GY3count, GZ3count, TG3count)
    
        # QC plots for this sequence
        plot_figures_ps1_3G(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                         timeabs, av_points, GX3count, GY3count, GZ3count, GX3pick, GY3pick, GZ3pick, GX3pick_av, GY3pick_av, GZ3pick_av, \
                         AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, GZ3pick_std, TG3pick_av, TApick_av)       
        
        if controlseq == True:
            plot_figures_ps1AC(time, accXcount, accYcount, accZcount, AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, meas_pointAC, \
                     timeabsAC, av_pointsAC, GX3count, GY3count, GZ3count, GX3pickAC, GY3pickAC, GZ3pickAC, GX3pick_avAC, \
                     GY3pick_avAC, GZ3pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, GX3pick_stdAC, GY3pick_stdAC, \
                     GZ3pick_stdAC, TG3pick_avAC, TApick_avAC)
    
        # write selected and averaged data to file
        save_file_ps1_3G(pickfile_write, picksheet_write, meas_point, AXpick_av, AYpick_av, AZpick_av, \
                         GX3pick_av, GY3pick_av, GZ3pick_av, TG3pick_av, TApick_av, AXRef, AYRef, AZRef)
        
        if controlseq == True:  
            save_file_ps1AC(pickfile_writeAC, picksheet_writeAC, meas_pointAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, TApick_avAC, AXRefAC, AYRefAC, AZRefAC)
        
    
        
        
        if autostarttime == True:
            starttime_Tcomp = int(timeabs[0] - 1520)
            print ('starttime_Tcomp =', starttime_Tcomp)
            
            stoptime_Tcomp = int(timeabs[107] + 100)
            if abstemp == 28:
                stoptime_Tcomp = int(timeabs[107] + 450)
            print ('stoptime_Tcomp =', stoptime_Tcomp)


if sequencemode == 2:            
    
    for processtep in (2,3,4,5,6):    
    
        if processtep == 2:
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef \
            = read_file_ps1_3G(pickfile_read , picksheet_read)
            
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
            index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count = readrawdata_3G(tempfile_read, tempsheet_read)
         
            # calculate 5th order polynominal
            TcoeffGX3, TcoeffGY3, TcoeffGZ3, TcoeffGX3lin, TcoeffGY3lin, TcoeffGZ3lin, TG3min, TG3max = \
                find_gyro_temp_poly_3G(index, GX3count, GY3count, GZ3count, TG3count, f_s_temp)
            
            save_file_gyroT_constants_3G(gyroTcal_file, gyroTcal_sheet, TcoeffGX3, TcoeffGY3, TcoeffGZ3, \
                                      TcoeffGX3lin, TcoeffGY3lin, TcoeffGZ3lin,TG3min, TG3max)
            
            
        
        if processtep == 4:
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef \
                = read_file_ps1_3G(pickfile_read, picksheet_read)
            
            # read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count \
                = readrawdata_3G(logfile_acc_gsens, sheetname_acc_gsens)
            
            # read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
            
            # read temperature coefficients from file 
            TcoeffGX3, TcoeffGY3, TcoeffGZ3 = read_file_gyroTcoeff_3G(gyroTcal_file, gyroTcal_sheet)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            
            
            # temperature compensate gyro counters and pick same time instants as before
            GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav = Tcomp_gyro_pick_average_3G(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                f_s_acc, timeabs, GX3count, GY3count, GZ3count, TG3count, TcoeffGX3, TcoeffGY3, TcoeffGZ3)
            
        
            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin \
                = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
            
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # calculate gravity sensitivity, ignoring earth rotation. Use in gyro scale factor calibration. Plots 
            GX3gsens, GY3gsens, GZ3gsens, TG3 = gsens_calc_3G(AXcal_lin, AYcal_lin, AZcal_lin, GX3pick_TCav, GY3pick_TCav, \
                                                    GZ3pick_TCav, TG3pick_av, 131, 132, 133, plotfigs1)        
            
            # plot QC data, 
           
            plot_figures_ps4_3G(time, GX3count, GY3count, GZ3count, meas_point, GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav)
           
               
            save_file_gsens_3G(gyrogsenscal_file, gyrogsenscal_sheet, GX3gsens, GY3gsens, GZ3gsens, TG3, abstemp)
            
        
        if processtep == 5:
            
            # read data file for X rotation
            indexXr, accXcountXr, accYcountXr, accZcountXr, GX3countXr, GY3countXr, GZ3countXr, TAcountXr, TG3countXr \
                = readrawdata_3G(Xrotfile, Xrotsheet)
            
            meas_pointXr, timegapXr, timeabsXr, rotrateXr = read_rot_timefile(rot_timefile, Xrot_timesheet)
            
            rot = 1 # X rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
        
            GX3avXr, GY3avXr, GZ3avXr, TG3avXr = gyro_rot_pick_3G(meas_pointXr, timeabsXr, rotrateXr, av_time_standstill, av_time_rot, \
                                    margin_standstill, margin_rotation, f_s_rot, accXcountXr, accYcountXr, accZcountXr, GX3countXr, \
                                    GY3countXr, GZ3countXr, TG3countXr, rot)
            
            
           
            # read data file for Y rotation
            indexYr, accXcountYr, accYcountYr, accZcountYr, GX3countYr, GY3countYr, GZ3countYr, TAcountYr, TG3countYr \
                        = readrawdata_3G(Yrotfile, Yrotsheet)
            
            meas_pointYr, timegapYr, timeabsYr, rotrateYr = read_rot_timefile(rot_timefile, Yrot_timesheet)
            
            rot = 2 # Y rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
            
        
            GX3avYr, GY3avYr, GZ3avYr, TG3avYr = gyro_rot_pick_3G(meas_pointYr, timeabsYr, rotrateYr, av_time_standstill, av_time_rot, \
                                    margin_standstill, margin_rotation, f_s_rot, accXcountYr, accYcountYr, accZcountYr, GX3countYr, \
                                    GY3countYr, GZ3countYr, TG3countYr, rot)
            
            
            # read data file for Z rotation
            indexZr, accXcountZr, accYcountZr, accZcountZr, GX3countZr, GY3countZr, GZ3countZr, TAcountZr, TG3countZr \
                = readrawdata_3G(Zrotfile, Zrotsheet)
            
            meas_pointZr, timegapZr, timeabsZr, rotrateZr = read_rot_timefile(rot_timefile, Zrot_timesheet)
            
            rot = 3 # Z rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile 
         
            GX3avZr, GY3avZr, GZ3avZr, TG3avZr = gyro_rot_pick_3G(meas_pointZr, timeabsZr, rotrateZr, av_time_standstill, av_time_rot, \
                                margin_standstill, margin_rotation, f_s_rot, accXcountZr, accYcountZr, accZcountZr, GX3countZr, \
                                GY3countZr, GZ3countZr, TG3countZr, rot)
                
            
            
            
            # calculate gyro ortomatrix
            
            xfactors, yfactors, zfactors, TG3 = calc_gyro_ortomatrix_3G(GX3avXr, GY3avXr, GZ3avXr, TG3avXr, rotrateXr, GX3avYr, GY3avYr, \
                                      GZ3avYr, TG3avYr, rotrateYr, GX3avZr, GY3avZr, GZ3avZr, TG3avZr, rotrateZr)
        
            # write gyro ortomatrix to file. 
            
            save_file_gyro_orto_3G(gyro_orto_file, gyro_orto_sheet, xfactors, yfactors, zfactors, TG3, abstemp)
            
        
        if processtep == 6:
            
            # read gyro scale factor matrix
            xfactors, yfactors, zfactors, TG3, abstemp = read_file_gyro_orto_3G(gyro_orto_file, gyro_orto_sheet)
            
            # find gyro sensors unit pointing vectors (along sensitive axis) from gyro orto matrix
            GX3point, GY3point, GZ3point, GX3scale, GY3scale, GZ3scale = \
                        gyro_body_vec_3G(xfactors, yfactors, zfactors)
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef \
                = read_file_ps1_3G(pickfile_read, picksheet_read)
            
            # read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count \
                = readrawdata_3G(logfile_acc_gsens, sheetname_acc_gsens)
            
            # read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
            
            # read temperature coefficients from file 
            TcoeffGX3, TcoeffGY3, TcoeffGZ3 = read_file_gyroTcoeff_3G(gyroTcal_file, gyroTcal_sheet)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            
            # temperature compensate gyro counters and pick same time instants as before
        
            GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav = Tcomp_gyro_pick_average_3G(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                    f_s_acc, timeabs, GX3count, GY3count, GZ3count, TG3count, TcoeffGX3, TcoeffGY3, TcoeffGZ3)
            
        
            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin \
            = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
            
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # read data from angle file (same as timing file)
            azimuth, inclination, toolface = readanglefile(file_path_timefile, sheetname_time)
            
            # calculate gyro counter values resulting from earth rotation
            GX3_er_comp, GY3_er_comp, GZ3_er_comp = er_comp_gyros_3G(GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav, \
                            azimuth, inclination, toolface, GX3point, GY3point, GZ3point, GX3scale, GY3scale, GZ3scale)
            
            # adjust gyro counter values for earth rotation
            GX3pick_TCav = GX3pick_TCav - GX3_er_comp
            GY3pick_TCav = GY3pick_TCav - GY3_er_comp
            GZ3pick_TCav = GZ3pick_TCav - GZ3_er_comp
            
            
            # calculate gravity sensitivity, compensating for earth rotation. Use in later data calibration. Plots 
            GX3gsens_erc, GY3gsens_erc, GZ3gsens_erc , TG3 = gsens_calc_3G(AXcal_lin, AYcal_lin, AZcal_lin, GX3pick_TCav, \
                                        GY3pick_TCav, GZ3pick_TCav, TG3pick_av, 171, 172, 173, plotfigs2)        
            
            save_file_gsens_3G(gyrogsenscal_erc_file, gyrogsenscal_erc_sheet, GX3gsens_erc, GY3gsens_erc, GZ3gsens_erc, TG3, abstemp) 

if sequencemode == 3:            
    
    for processtep in (1,2,3,4,5,6):   
        
        if processtep == 1:
            
            #read data from raw data file 
            index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count \
                = readrawdata_3G(logfile_acc_gsens, sheetname_acc_gsens)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            av_pointsAC = av_timeAC * f_s_acc
            
            #read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
            
         
            # read data from timing file for control signal
            
            if controlseq == True:
                meas_pointAC, timegapAC, timeabsAC, AXRefAC, AYRefAC, AZRefAC = readtimefile(file_path_timefile, sheetname_control)
            
            
            # pick correctly timed value from counter sequence. Adjust timing file as necessary   
        
            
            AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, GX3pick, GY3pick, GZ3pick, GX3pick_av, GY3pick_av, GZ3pick_av, \
            AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, GZ3pick_std, TG3pick_av, TApick_av \
            = acc_gsens_pick_3G(meas_point, av_time,  f_s_acc, timeabs, accXcount, accYcount, accZcount, TAcount, GX3count, GY3count, GZ3count, TG3count)
        
            if controlseq == True:
                AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, GX3pickAC, GY3pickAC, \
                    GZ3pickAC, GX3pick_avAC, GY3pick_avAC, GZ3pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, \
                    GX3pick_stdAC, GY3pick_stdAC, GZ3pick_stdAC, TG3pick_avAC, TApick_avAC \
                    = acc_control_pick(meas_pointAC, av_timeAC, timeabsAC, f_s_acc, accXcount, accYcount, accZcount, TAcount, \
                                     GX3count, GY3count, GZ3count, TG3count)
        
            # QC plots for this sequence
            plot_figures_ps1_3G(time, accXcount, accYcount, accZcount, AXpick, AYpick, AZpick, AXpick_av, AYpick_av, AZpick_av, meas_point, \
                             timeabs, av_points, GX3count, GY3count, GZ3count, GX3pick, GY3pick, GZ3pick, GX3pick_av, GY3pick_av, GZ3pick_av, \
                             AXpick_std, AYpick_std, AZpick_std, GX3pick_std, GY3pick_std, GZ3pick_std, TG3pick_av, TApick_av)       
            
            if controlseq == True:
                plot_figures_ps1AC(time, accXcount, accYcount, accZcount, AXpickAC, AYpickAC, AZpickAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, meas_pointAC, \
                         timeabsAC, av_pointsAC, GX3count, GY3count, GZ3count, GX3pickAC, GY3pickAC, GZ3pickAC, GX3pick_avAC, \
                         GY3pick_avAC, GZ3pick_avAC, AXpick_stdAC, AYpick_stdAC, AZpick_stdAC, GX3pick_stdAC, GY3pick_stdAC, \
                         GZ3pick_stdAC, TG3pick_avAC, TApick_avAC)
        
            # write selected and averaged data to file
            save_file_ps1_3G(pickfile_write, picksheet_write, meas_point, AXpick_av, AYpick_av, AZpick_av, \
                             GX3pick_av, GY3pick_av, GZ3pick_av, TG3pick_av, TApick_av, AXRef, AYRef, AZRef)
            
            if controlseq == True:  
                save_file_ps1AC(pickfile_writeAC, picksheet_writeAC, meas_pointAC, AXpick_avAC, AYpick_avAC, AZpick_avAC, TApick_avAC, AXRefAC, AYRefAC, AZRefAC)
            
        
            
            
            if autostarttime == True:
                starttime_Tcomp = int(timeabs[0] - 1520)
                print ('starttime_Tcomp =', starttime_Tcomp)
                
                stoptime_Tcomp = int(timeabs[107] + 100)
                if abstemp == 28:
                    stoptime_Tcomp = int(timeabs[107] + 450)
                print ('stoptime_Tcomp =', stoptime_Tcomp)
                
            
        
        if processtep == 2:
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef \
            = read_file_ps1_3G(pickfile_read , picksheet_read)
            
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
            index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count = readrawdata_3G(tempfile_read, tempsheet_read)
         
            # calculate 5th order polynominal
            TcoeffGX3, TcoeffGY3, TcoeffGZ3, TcoeffGX3lin, TcoeffGY3lin, TcoeffGZ3lin, TG3min, TG3max = \
                find_gyro_temp_poly_3G(index, GX3count, GY3count, GZ3count, TG3count, f_s_temp)
            
            save_file_gyroT_constants_3G(gyroTcal_file, gyroTcal_sheet, TcoeffGX3, TcoeffGY3, TcoeffGZ3, \
                                      TcoeffGX3lin, TcoeffGY3lin, TcoeffGZ3lin,TG3min, TG3max)
            
            
        
        if processtep == 4:
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef \
                = read_file_ps1_3G(pickfile_read, picksheet_read)
            
            # read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count \
                = readrawdata_3G(logfile_acc_gsens, sheetname_acc_gsens)
            
            # read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
            
            # read temperature coefficients from file 
            TcoeffGX3, TcoeffGY3, TcoeffGZ3 = read_file_gyroTcoeff_3G(gyroTcal_file, gyroTcal_sheet)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            
            
            # temperature compensate gyro counters and pick same time instants as before
            GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav = Tcomp_gyro_pick_average_3G(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                f_s_acc, timeabs, GX3count, GY3count, GZ3count, TG3count, TcoeffGX3, TcoeffGY3, TcoeffGZ3)
            
        
            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin \
                = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
            
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # calculate gravity sensitivity, ignoring earth rotation. Use in gyro scale factor calibration. Plots 
            GX3gsens, GY3gsens, GZ3gsens, TG3 = gsens_calc_3G(AXcal_lin, AYcal_lin, AZcal_lin, GX3pick_TCav, GY3pick_TCav, \
                                                    GZ3pick_TCav, TG3pick_av, 131, 132, 133, plotfigs1)        
            
            # plot QC data, 
           
            plot_figures_ps4_3G(time, GX3count, GY3count, GZ3count, meas_point, GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav)
           
               
            save_file_gsens_3G(gyrogsenscal_file, gyrogsenscal_sheet, GX3gsens, GY3gsens, GZ3gsens, TG3, abstemp)
            
        
        if processtep == 5:
            
            # read data file for X rotation
            indexXr, accXcountXr, accYcountXr, accZcountXr, GX3countXr, GY3countXr, GZ3countXr, TAcountXr, TG3countXr \
                = readrawdata_3G(Xrotfile, Xrotsheet)
            
            meas_pointXr, timegapXr, timeabsXr, rotrateXr = read_rot_timefile(rot_timefile, Xrot_timesheet)
            
            rot = 1 # X rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
        
            GX3avXr, GY3avXr, GZ3avXr, TG3avXr = gyro_rot_pick_3G(meas_pointXr, timeabsXr, rotrateXr, av_time_standstill, av_time_rot, \
                                    margin_standstill, margin_rotation, f_s_rot, accXcountXr, accYcountXr, accZcountXr, GX3countXr, \
                                    GY3countXr, GZ3countXr, TG3countXr, rot)
            
            
           
            # read data file for Y rotation
            indexYr, accXcountYr, accYcountYr, accZcountYr, GX3countYr, GY3countYr, GZ3countYr, TAcountYr, TG3countYr \
                        = readrawdata_3G(Yrotfile, Yrotsheet)
            
            meas_pointYr, timegapYr, timeabsYr, rotrateYr = read_rot_timefile(rot_timefile, Yrot_timesheet)
            
            rot = 2 # Y rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile
            
        
            GX3avYr, GY3avYr, GZ3avYr, TG3avYr = gyro_rot_pick_3G(meas_pointYr, timeabsYr, rotrateYr, av_time_standstill, av_time_rot, \
                                    margin_standstill, margin_rotation, f_s_rot, accXcountYr, accYcountYr, accZcountYr, GX3countYr, \
                                    GY3countYr, GZ3countYr, TG3countYr, rot)
            
            
            # read data file for Z rotation
            indexZr, accXcountZr, accYcountZr, accZcountZr, GX3countZr, GY3countZr, GZ3countZr, TAcountZr, TG3countZr \
                = readrawdata_3G(Zrotfile, Zrotsheet)
            
            meas_pointZr, timegapZr, timeabsZr, rotrateZr = read_rot_timefile(rot_timefile, Zrot_timesheet)
            
            rot = 3 # Z rotation
            
            # extract averaged, g compensated and temperature compensated gyro counter values for standstill and rotation periods, as described by rot_timefile 
         
            GX3avZr, GY3avZr, GZ3avZr, TG3avZr = gyro_rot_pick_3G(meas_pointZr, timeabsZr, rotrateZr, av_time_standstill, av_time_rot, \
                                margin_standstill, margin_rotation, f_s_rot, accXcountZr, accYcountZr, accZcountZr, GX3countZr, \
                                GY3countZr, GZ3countZr, TG3countZr, rot)
                
            
            
            
            # calculate gyro ortomatrix
            
            xfactors, yfactors, zfactors, TG3 = calc_gyro_ortomatrix_3G(GX3avXr, GY3avXr, GZ3avXr, TG3avXr, rotrateXr, GX3avYr, GY3avYr, \
                                      GZ3avYr, TG3avYr, rotrateYr, GX3avZr, GY3avZr, GZ3avZr, TG3avZr, rotrateZr)
        
            # write gyro ortomatrix to file. 
            
            save_file_gyro_orto_3G(gyro_orto_file, gyro_orto_sheet, xfactors, yfactors, zfactors, TG3, abstemp)
            
        
        if processtep == 6:
            
            # read gyro scale factor matrix
            xfactors, yfactors, zfactors, TG3, abstemp = read_file_gyro_orto_3G(gyro_orto_file, gyro_orto_sheet)
            
            # find gyro sensors unit pointing vectors (along sensitive axis) from gyro orto matrix
            GX3point, GY3point, GZ3point, GX3scale, GY3scale, GZ3scale = \
                        gyro_body_vec_3G(xfactors, yfactors, zfactors)
            
            # read picked and averaged accelerometer and gyro counter values from file (file written in processtep 1)
            meas_point, AXpick_av, AYpick_av, AZpick_av, GX3pick_av, GY3pick_av, GZ3pick_av, TApick_av, TG3pick_av, AXRef, AYRef, AZRef \
                = read_file_ps1_3G(pickfile_read, picksheet_read)
            
            # read data from raw data file - needed for temperature compensation
            index, accXcount, accYcount, accZcount, GX3count, GY3count, GZ3count, TAcount, TG3count \
                = readrawdata_3G(logfile_acc_gsens, sheetname_acc_gsens)
            
            # read data from timing file
            meas_point, timegap, timeabs, AXRef, AYRef, AZRef = readtimefile(file_path_timefile, sheetname_time)
            
            # read temperature coefficients from file 
            TcoeffGX3, TcoeffGY3, TcoeffGZ3 = read_file_gyroTcoeff_3G(gyroTcal_file, gyroTcal_sheet)
            
            time = index / f_s_acc
            av_points = av_time * f_s_acc
            
            # temperature compensate gyro counters and pick same time instants as before
        
            GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav = Tcomp_gyro_pick_average_3G(meas_point, starttime_Tcomp, stoptime_Tcomp, standstill_time,  \
                    f_s_acc, timeabs, GX3count, GY3count, GZ3count, TG3count, TcoeffGX3, TcoeffGY3, TcoeffGZ3)
            
        
            # find accelerometer calibration constants, classical calibration (ortomatrix  and offset)
            xfactors_lin, yfactors_lin, zfactors_lin, xoffset, yoffset, zoffset, T_lin \
            = acc_cal_classic(AXpick_av, AYpick_av, AZpick_av, TApick_av, AXRef, AYRef, AZRef)
            
            # calibrate picked accelerometer data, classical method
            AXcal_lin, AYcal_lin, AZcal_lin = calibrate_acc_classical(AXpick_av, AYpick_av, AZpick_av, xfactors_lin, yfactors_lin, zfactors_lin)
            
            # read data from angle file (same as timing file)
            azimuth, inclination, toolface = readanglefile(file_path_timefile, sheetname_time)
            
            # calculate gyro counter values resulting from earth rotation
            GX3_er_comp, GY3_er_comp, GZ3_er_comp = er_comp_gyros_3G(GX3pick_TCav, GY3pick_TCav, GZ3pick_TCav, \
                            azimuth, inclination, toolface, GX3point, GY3point, GZ3point, GX3scale, GY3scale, GZ3scale)
            
            # adjust gyro counter values for earth rotation
            GX3pick_TCav = GX3pick_TCav - GX3_er_comp
            GY3pick_TCav = GY3pick_TCav - GY3_er_comp
            GZ3pick_TCav = GZ3pick_TCav - GZ3_er_comp
            
            
            # calculate gravity sensitivity, compensating for earth rotation. Use in later data calibration. Plots 
            GX3gsens_erc, GY3gsens_erc, GZ3gsens_erc , TG3 = gsens_calc_3G(AXcal_lin, AYcal_lin, AZcal_lin, GX3pick_TCav, \
                                        GY3pick_TCav, GZ3pick_TCav, TG3pick_av, 171, 172, 173, plotfigs2)        
            
            save_file_gsens_3G(gyrogsenscal_erc_file, gyrogsenscal_erc_sheet, GX3gsens_erc, GY3gsens_erc, GZ3gsens_erc, TG3, abstemp)
