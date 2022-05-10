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
# R2 adapt to ST sensor included, FW A110
# R3 adapt to FW A111, include accelerometer from ST sensor
# R11 adapt to FW A112, with RUN L/DUMP L command, including all sensors from ST sensor, ADXL355, ADXRS 290 and magnetometer
# R12 adapt to calibration file in 
# R14 include support for control signal at 28C, write figures to file
# R15 change loacl gravity at Heimdal to 9.821m/s^2
# R16 for python 3.8.5
# R17 fixed MAA to MAB bug in write constant file script. No change in this script
# R18 introduced gyro result image script, fixed minor error in plot 117. No change in this script

# R99 Commented out plt.show() for automatic processing

from __future__ import division
import matplotlib.pyplot as plt
import numpy as np
import xlrd
from pandas import DataFrame
from scipy import signal


processtep = 1
                            #           
processtep = 1



# parameters processtep 1

toolSN = '3331'

abstemp1 = 28           # temperature in deg C as set in the temperature bath (not neccessarily identical to actual tool temperature)
abstemp2 = -5
abstemp3 = 60
            
offset =   0            # do not use, shall be set to 0
offset2 =  119         # offset in seconds, compared to first tool in calibration batch
splice =  30            # extension of each overlapping temperature section

starttime1cal = 1153
stoptime1cal = 5965

starttime2cal = 8670
stoptime2cal = 13059

starttime3cal = 16993
stoptime3cal = 21430


starttime1temp = 6145
stoptime1temp = 8012

starttime2temp = 13760
stoptime2temp = 15867

starttime3temp = 22073
stoptime3temp = 23321



logfile_3T = toolSN + ' rawcal.xlsx' 
sheetname_3T = toolSN + ' rawcal'


f_s_acc = 10.0    
av_time = 10                                                           # averaging time around each selected (half before and half after selected time ) 

tempfile = toolSN + ' tempcal.xlsx' 
tempsheet = 'tempcal'

logfile1 = toolSN + ' ' + str(abstemp1) + 'C' + ' ' + 'rotgsens.xlsx'
logsheet1 = toolSN + ' ' + str(abstemp1) + 'C' + ' ' + 'rotgsens'

logfile2 = toolSN + ' ' + str(abstemp2) + 'C' + ' ' + 'rotgsens.xlsx'
logsheet2 = toolSN + ' ' + str(abstemp2) + 'C' + ' ' + 'rotgsens'

logfile3 = toolSN + ' ' + str(abstemp3) + 'C' + ' ' + 'rotgsens.xlsx'
logsheet3 = toolSN + ' ' + str(abstemp3) + 'C' + ' ' + 'rotgsens'

logfile_TCout = toolSN + ' Tcal cut.xlsx' 
sheetname_TCout = toolSN + ' Tcal cut' 




# function definitions

          
def readrawdata(file_path, sheetname_data):
    

    
    # read data file (flex data)
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheetname_data)
    
    # extract data from ST accelerometer
    offset = 1
    col = worksheet.col_values(15)
    accXcountST = col[offset:]
    accXcountST = np.array(accXcountST)
    # extract data from Y accelerometer
    col = worksheet.col_values(16)
    accYcountST = col[offset:]
    accYcountST = np.array(accYcountST)
    # extract data from Z accelerometer
    col = worksheet.col_values(17)
    accZcountST = col[offset:]
    accZcountST = np.array(accZcountST)   
    
    
    # extract data from ADXL accelerometer
    offset = 1
    col = worksheet.col_values(5)
    accXcount = col[offset:]
    accXcount = np.array(accXcount)
    # extract data from Y accelerometer
    col = worksheet.col_values(6)
    accYcount = col[offset:]
    accYcount = np.array(accYcount)
    # extract data from Z accelerometer
    col = worksheet.col_values(7)
    accZcount = col[offset:]
    accZcount = np.array(accZcount)   
    
    # extract index
    col = worksheet.col_values(1)
    index = col[offset:]
    index = np.array(index)
    
    # extract gyro counter values ADXRS290
    col = worksheet.col_values(8)
    GX1count = col[offset:]
    GX1count = np.array(GX1count)
    
    col = worksheet.col_values(9)
    GY1count = col[offset:]
    GY1count = np.array(GY1count)
    
    col = worksheet.col_values(10)
    GX2count = col[offset:]
    GX2count = np.array(GX2count)    
    
    col = worksheet.col_values(11)
    GZ2count = col[offset:]
    GZ2count = np.array(GZ2count)           
    
    # extract gyro counter values ST sensor
    col = worksheet.col_values(12)
    GX3count = col[offset:]
    GX3count = np.array(GX3count)
    
    col = worksheet.col_values(13)
    GY3count = col[offset:]
    GY3count = np.array(GY3count)
    
    col = worksheet.col_values(14)
    GZ3count = col[offset:]
    GZ3count = np.array(GZ3count)  
    
    
    
    TAcount = np.zeros(len(index))
    TG1count = np.zeros(len(index))
    TG2count = np.zeros(len(index))
    TG3count = np.zeros(len(index))
    
    # extract temperature data
    col = worksheet.col_values(19)
    TAcount = col[offset:]
    col = worksheet.col_values(20)
    TG1count = col[offset:]
    col = worksheet.col_values(21)
    TG2count = col[offset:]
    col = worksheet.col_values(22)
    TG3count = col[offset:]
    
 
    
    lastTA = 0
    lastTG1 = 0
    lastTG2 = 0
    lastTG3 = 0
    
    
    for k in range(len(index)):
        if (TAcount[k]) == '':
            TAcount[k] = lastTA
            TG1count [k] = lastTG1
            TG2count [k] = lastTG2
            TG3count [k] = lastTG3
        else:
            lastTA = TAcount[k]
            lastTG1 = TG1count[k]
            lastTG2 = TG2count[k]
            lastTG3 = TG3count[k]
            
    for k in range(15):
        TAcount[k] = TAcount[16]
        TG1count[k] = TG1count[16]
        TG2count[k] = TG2count[16]
        TG3count[k] = TG3count[16]
    
    TAcount = np.array(TAcount)    
    TG1count = np.array(TG1count)
    TG2count = np.array(TG2count)
    TG3count = np.array(TG3count)
    
    return index, accXcount, accYcount, accZcount, accXcountST, accYcountST, accZcountST, GX1count, GY1count, GX2count, \
            GZ2count, GX3count, GY3count, GZ3count, TAcount, TG1count, TG2count, TG3count 

def save_cal_file(filename, sheetname, index, AX, AY, AZ, AX2, AY2, AZ2, GX1, GY1, GX2, GZ2, GX3, GY3, GZ3, TA, TG1, TG2, TG3):
    
    l0 = index
    l1 = AX
    l2 = AY
    l3 = AZ
    l4 = AX2
    l5 = AY2
    l6 = AZ2
    l7 = GX1
    l8 = GY1
    l9 = GX2
    l10 = GZ2
    l11 = GX3
    l12 = GY3
    l13 = GZ3
    l14 = TA
    l15 = TG1
    l16 = TG2
    l17 = TG3

    
    df = DataFrame({'00 Index':l0, '01 AX':l1, '02 AY':l2, '03 AZ':l3, '04 AX_ST':l4, '05 AY_ST':l5, '06 AZ_ST':l6,'07 GX1': l7, '08 GY1': l8, '09 GX2': l9, '10 GZ2':l10, 
                    '11 GX3':l11, '12 GY3':l12, '13 GZ3':l13, '14 TA':l14, '15 TG1':l15, '16 TG2':l16, '17 TG3':l17})
    
    df.to_excel(filename, sheetname, index = False)
    
    return

def readtempdata(file_path, sheetname_data):
    
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheetname_data)
    
    # extract data from ST accelerometer
    offset = 1
    col = worksheet.col_values(0)
    index = col[offset:]
    index = np.array(index)

    col = worksheet.col_values(1)
    accXcount = col[offset:]
    accXcount = np.array(accXcount)    
    
    col = worksheet.col_values(2)
    accYcount = col[offset:]
    accYcount = np.array(accYcount)    
    
    col = worksheet.col_values(3)
    accZcount = col[offset:]
    accZcount = np.array(accZcount)    
    
    col = worksheet.col_values(4)
    accXcountST = col[offset:]
    accXcountST = np.array(accXcountST)    
    
    col = worksheet.col_values(5)
    accYcountST = col[offset:]
    accYcountST = np.array(accYcountST)    
    
    col = worksheet.col_values(6)
    accZcountST = col[offset:]
    accZcountST = np.array(accZcountST)    
    
    
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
    
    col = worksheet.col_values(11)
    GX3count = col[offset:]
    GX3count = np.array(GX3count)    
    
    col = worksheet.col_values(12)
    GY3count = col[offset:]
    GY3count = np.array(GY3count)   
    
    col = worksheet.col_values(13)
    GZ3count = col[offset:]
    GZ3count = np.array(GZ3count) 
    
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
    
    return index, accXcount, accYcount, accZcount, accXcountST, accYcountST, accZcountST, GX1count, GY1count, GX2count, \
            GZ2count, GX3count, GY3count, GZ3count, TAcount, TG1count, TG2count, TG3count 




#read data from raw data file 

if processtep == 1:

    index, accXcount, accYcount, accZcount, accXcountST, accYcountST, accZcountST, GX1count, GY1count, GX2count, GZ2count, \
            GX3count, GY3count, GZ3count, TAcount, TG1count, TG2count, TG3count  = readrawdata(logfile_3T, sheetname_3T)
    
    time = index / f_s_acc
    
    
    fig, ax1 = plt.subplots()
    plt.plot((time-offset), GX1count, label = 'GX1')

    plt.plot((time-offset), GY1count, label = 'GY1')
    plt.plot((time-offset), GX2count, label = 'GX2')
    plt.plot((time-offset), GZ2count, label = 'GZ2')
   
    plt.legend(loc = 'best')
    
    ax2 = ax1.twinx()
    plt.plot((time-offset), TG1count, 'r', label = 'TG1count')
    plt.plot((time-offset), TG2count, 'b', label = 'TG2count')
    plt.legend(loc = 'best')
        
    plt.grid()
    plt.savefig(toolSN + ' split step 1 plot 1.png', bbox_inches='tight', dpi=250)
#    plt.show()
    
    
    
   
    
    
    fig, ax1 = plt.subplots()
    plt.plot((time-offset), GX3count, label = 'GX3')
  
    plt.plot((time-offset), GY3count, label = 'GY3')
    plt.plot((time-offset), GZ3count, label = 'GZ3')
    
    plt.legend(loc = 'best')
    
    ax2 = ax1.twinx()
    plt.plot((time-offset), TG3count, 'r', label = 'TG3count')
    plt.legend(loc = 'best')
        
    plt.grid()
    plt.savefig(toolSN + ' split step 1 plot 2.png', bbox_inches='tight', dpi=250)
#    plt.show()
    
       
    
    


    fig, ax1 = plt.subplots()
    plt.plot((time-offset), accXcount, label = 'AX')
  
    plt.plot((time-offset), accYcount, label = 'AY')
    plt.plot((time-offset), accZcount, label = 'AZ')
    plt.plot((time-offset), accXcountST, label = 'STAX')
    plt.plot((time-offset), accYcountST, label = 'STAY')
    plt.plot((time-offset), accZcountST, label = 'STAZ')
    
    plt.legend(loc = 'best')
    

        
    plt.grid()
    plt.savefig(toolSN + ' split step 1 plot 3.png', bbox_inches='tight', dpi=250)
#    plt.show()    

    save_cal_file(tempfile, tempsheet, index, accXcount, accYcount, accZcount, accXcountST, accYcountST, accZcountST,\
                  GX1count, GY1count, GX2count, GZ2count, GX3count, GY3count, GZ3count, TAcount, TG1count, TG2count, TG3count)
    
    
if processtep == 2:
    
    starttime1cal = starttime1cal - offset
    stoptime1cal = stoptime1cal - offset

    starttime2cal = starttime2cal - offset
    stoptime2cal = stoptime2cal - offset

    starttime3cal = starttime3cal - offset
    stoptime3cal = stoptime3cal - offset
    
    startindex1cal = int(f_s_acc * starttime1cal)
    startindex2cal = int(f_s_acc * starttime2cal)
    startindex3cal = int(f_s_acc * starttime3cal)
    
    stopindex1cal = int(f_s_acc * stoptime1cal)
    stopindex2cal = int(f_s_acc * stoptime2cal)
    stopindex3cal = int(f_s_acc * stoptime3cal)
    
    
    index, accXcount, accYcount, accZcount, accXcountST, accYcountST, accZcountST, GX1count, GY1count, GX2count, \
        GZ2count, GX3count, GY3count, GZ3count, TAcount, TG1count, TG2count, TG3count = \
                readtempdata(tempfile, tempsheet)
    
    
    
    index_1 = index[startindex1cal:stopindex1cal]
    AX_1 = accXcount[startindex1cal:stopindex1cal]
    AY_1 = accYcount[startindex1cal:stopindex1cal]
    AZ_1 = accZcount[startindex1cal:stopindex1cal]
    AXST_1 = accXcountST[startindex1cal:stopindex1cal]
    AYST_1 = accYcountST[startindex1cal:stopindex1cal]
    AZST_1 = accZcountST[startindex1cal:stopindex1cal]    
    GX1_1 = GX1count[startindex1cal:stopindex1cal]
    GY1_1 = GY1count[startindex1cal:stopindex1cal]
    GX2_1 = GX2count[startindex1cal:stopindex1cal]
    GZ2_1 = GZ2count[startindex1cal:stopindex1cal]
    GX3_1 = GX3count[startindex1cal:stopindex1cal]
    GY3_1 = GY3count[startindex1cal:stopindex1cal]
    GZ3_1 = GZ3count[startindex1cal:stopindex1cal]
    TA_1 = TAcount[startindex1cal:stopindex1cal]
    TG1_1 = TG1count[startindex1cal:stopindex1cal]
    TG2_1 = TG2count[startindex1cal:stopindex1cal]
    TG3_1 = TG3count[startindex1cal:stopindex1cal]
      
    index_2 = index[startindex2cal:stopindex2cal]
    AX_2 = accXcount[startindex2cal:stopindex2cal]
    AY_2 = accYcount[startindex2cal:stopindex2cal]
    AZ_2 = accZcount[startindex2cal:stopindex2cal]
    AXST_2 = accXcountST[startindex2cal:stopindex2cal]
    AYST_2 = accYcountST[startindex2cal:stopindex2cal]
    AZST_2 = accZcountST[startindex2cal:stopindex2cal]  
    GX1_2 = GX1count[startindex2cal:stopindex2cal]
    GY1_2 = GY1count[startindex2cal:stopindex2cal]
    GX2_2 = GX2count[startindex2cal:stopindex2cal]
    GZ2_2 = GZ2count[startindex2cal:stopindex2cal]
    GX3_2 = GX3count[startindex2cal:stopindex2cal]
    GY3_2 = GY3count[startindex2cal:stopindex2cal]
    GZ3_2 = GZ3count[startindex2cal:stopindex2cal]
    TA_2 = TAcount[startindex2cal:stopindex2cal]
    TG1_2 = TG1count[startindex2cal:stopindex2cal]
    TG2_2 = TG2count[startindex2cal:stopindex2cal]
    TG3_2 = TG3count[startindex2cal:stopindex2cal]
    
    index_3 = index[startindex3cal:stopindex3cal]
    AX_3 = accXcount[startindex3cal:stopindex3cal]
    AY_3 = accYcount[startindex3cal:stopindex3cal]
    AZ_3 = accZcount[startindex3cal:stopindex3cal]
    AXST_3 = accXcountST[startindex3cal:stopindex3cal]
    AYST_3 = accYcountST[startindex3cal:stopindex3cal]
    AZST_3 = accZcountST[startindex3cal:stopindex3cal]  
    GX1_3 = GX1count[startindex3cal:stopindex3cal]
    GY1_3 = GY1count[startindex3cal:stopindex3cal]
    GX2_3 = GX2count[startindex3cal:stopindex3cal]
    GZ2_3 = GZ2count[startindex3cal:stopindex3cal]
    GX3_3 = GX3count[startindex3cal:stopindex3cal]
    GY3_3 = GY3count[startindex3cal:stopindex3cal]
    GZ3_3 = GZ3count[startindex3cal:stopindex3cal]
    TA_3 = TAcount[startindex3cal:stopindex3cal]
    TG1_3 = TG1count[startindex3cal:stopindex3cal]
    TG2_3 = TG2count[startindex3cal:stopindex3cal]
    TG3_3 = TG3count[startindex3cal:stopindex3cal]
    
    save_cal_file(logfile1, logsheet1, index_1, AX_1, AY_1, AZ_1, AXST_1, AYST_1, AZST_1, \
                  GX1_1, GY1_1, GX2_1, GZ2_1, GX3_1, GY3_1, GZ3_1, TA_1, TG1_1, TG2_1, TG3_1)    
    save_cal_file(logfile2, logsheet2, index_2, AX_2, AY_2, AZ_2, AXST_2, AYST_2, AZST_2, \
                  GX1_2, GY1_2, GX2_2, GZ2_2, GX3_2, GY3_2, GZ3_2, TA_2, TG1_2, TG2_2, TG3_2)
    save_cal_file(logfile3, logsheet3, index_3, AX_3, AY_3, AZ_3, AXST_3, AYST_3, AZST_3, \
                  GX1_3, GY1_3, GX2_3, GZ2_3, GX3_3, GY3_3, GZ3_3, TA_3, TG1_3, TG2_3, TG3_3) 
    
    
  
    starttime1temp = starttime1temp - offset - splice
    stoptime1temp = stoptime1temp - offset + splice

    starttime2temp = starttime2temp - offset - splice
    stoptime2temp = stoptime2temp - offset

    starttime3temp = starttime3temp - offset
    stoptime3temp = stoptime3temp - offset + splice
        
    startindex1temp = int(f_s_acc * starttime1temp)
    startindex2temp = int(f_s_acc * starttime2temp)
    startindex3temp = int(f_s_acc * starttime3temp)
    
    stopindex1temp = int(f_s_acc * stoptime1temp)
    stopindex2temp = int(f_s_acc * stoptime2temp)
    stopindex3temp = int(f_s_acc * stoptime3temp)
    
    spliceindex = int(splice * f_s_acc)
    
    index_1t = index[startindex1temp:stopindex1temp]
      
    index_2t = index[startindex2temp:stopindex2temp]
    
    index_3t = index[startindex3temp:stopindex3temp]

    
    indextemp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    AXtemp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    AYtemp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    AZtemp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    AXSTtemp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    AYSTtemp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    AZSTtemp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    GX1temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    GY1temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    GX2temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    GZ2temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    GX3temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    GY3temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    GZ3temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    TAtemp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    TG1temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    TG2temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    TG3temp = np.zeros(len(index_1t)+len(index_2t)+len(index_3t))
    
    indextemp[0:(stopindex1temp - startindex1temp)] = index[startindex1temp:stopindex1temp]
    indextemp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = index[startindex2temp:stopindex2temp]
    indextemp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = index[startindex3temp:stopindex3temp]
    
    AXtemp[0:(stopindex1temp - startindex1temp)] = accXcount[startindex1temp:stopindex1temp]
    AXtemp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = accXcount[startindex2temp:stopindex2temp]
    AXtemp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = accXcount[startindex3temp:stopindex3temp]
    
    AYtemp[0:(stopindex1temp - startindex1temp)] = accYcount[startindex1temp:stopindex1temp]
    AYtemp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = accYcount[startindex2temp:stopindex2temp]
    AYtemp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = accYcount[startindex3temp:stopindex3temp]
    
    AZtemp[0:(stopindex1temp - startindex1temp)] = accZcount[startindex1temp:stopindex1temp]
    AZtemp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = accZcount[startindex2temp:stopindex2temp]
    AZtemp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = accZcount[startindex3temp:stopindex3temp]
    
    AXSTtemp[0:(stopindex1temp - startindex1temp)] = accXcountST[startindex1temp:stopindex1temp]
    AXSTtemp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = accXcountST[startindex2temp:stopindex2temp]
    AXSTtemp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = accXcountST[startindex3temp:stopindex3temp]
    
    AYSTtemp[0:(stopindex1temp - startindex1temp)] = accYcountST[startindex1temp:stopindex1temp]
    AYSTtemp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = accYcountST[startindex2temp:stopindex2temp]
    AYSTtemp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = accYcountST[startindex3temp:stopindex3temp]
    
    AZSTtemp[0:(stopindex1temp - startindex1temp)] = accZcountST[startindex1temp:stopindex1temp]
    AZSTtemp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = accZcountST[startindex2temp:stopindex2temp]
    AZSTtemp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = accZcountST[startindex3temp:stopindex3temp]
    
    GX1temp[0:(stopindex1temp - startindex1temp)] = GX1count[startindex1temp:stopindex1temp]
    GX1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = GX1count[startindex2temp:stopindex2temp]
    GX1temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = GX1count[startindex3temp:stopindex3temp]
        
    GY1temp[0:(stopindex1temp - startindex1temp)] = GY1count[startindex1temp:stopindex1temp]
    GY1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = GY1count[startindex2temp:stopindex2temp]
    GY1temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = GY1count[startindex3temp:stopindex3temp]
    
    GX2temp[0:(stopindex1temp - startindex1temp)] = GX2count[startindex1temp:stopindex1temp]
    GX2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = GX2count[startindex2temp:stopindex2temp]
    GX2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = GX2count[startindex3temp:stopindex3temp]
    
    GZ2temp[0:(stopindex1temp - startindex1temp)] = GZ2count[startindex1temp:stopindex1temp]
    GZ2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = GZ2count[startindex2temp:stopindex2temp]
    GZ2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = GZ2count[startindex3temp:stopindex3temp]
        
    GX3temp[0:(stopindex1temp - startindex1temp)] = GX3count[startindex1temp:stopindex1temp]
    GX3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = GX3count[startindex2temp:stopindex2temp]
    GX3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = GX3count[startindex3temp:stopindex3temp]
        
    GY3temp[0:(stopindex1temp - startindex1temp)] = GY3count[startindex1temp:stopindex1temp]
    GY3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = GY3count[startindex2temp:stopindex2temp]
    GY3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = GY3count[startindex3temp:stopindex3temp]
    
    GZ3temp[0:(stopindex1temp - startindex1temp)] = GZ3count[startindex1temp:stopindex1temp]
    GZ3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = GZ3count[startindex2temp:stopindex2temp]
    GZ3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = GZ3count[startindex3temp:stopindex3temp]
    
    TAtemp[0:(stopindex1temp - startindex1temp)] = TAcount[startindex1temp:stopindex1temp]
    TAtemp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = TAcount[startindex2temp:stopindex2temp]
    TAtemp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = TAcount[startindex3temp:stopindex3temp]
    
    TG1temp[0:(stopindex1temp - startindex1temp)] = TG1count[startindex1temp:stopindex1temp]
    TG1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = TG1count[startindex2temp:stopindex2temp]
    TG1temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = TG1count[startindex3temp:stopindex3temp]
    
    TG2temp[0:(stopindex1temp - startindex1temp)] = TG2count[startindex1temp:stopindex1temp]
    TG2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = TG2count[startindex2temp:stopindex2temp]
    TG2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = TG2count[startindex3temp:stopindex3temp]
    
    TG3temp[0:(stopindex1temp - startindex1temp)] = TG3count[startindex1temp:stopindex1temp]
    TG3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ \
              (stopindex2temp - startindex2temp)] = TG3count[startindex2temp:stopindex2temp]
    TG3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = TG3count[startindex3temp:stopindex3temp]
    
    
    # at end of first section
    
    a1GX1, b1GX1 = np.polyfit(TG1temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp) ], \
                             GX1temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp)], 1)
    
    print ('a1GX1 ', a1GX1)
    print ('b1GX1 ', b1GX1)

    
    a1GY1, b1GY1 = np.polyfit(TG1temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp) ], \
                             GY1temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp)], 1)
    
    print ('a1GY1 ', a1GY1)
    print ('b1GY1 ', b1GY1)
    
    a1GX2, b1GX2 = np.polyfit(TG2temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp) ], \
                             GX2temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp)], 1)
    
    print ('a1GX2 ', a1GX2)
    print ('b1GX2 ', b1GX2)
    
    a1GZ2, b1GZ2 = np.polyfit(TG2temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp) ], \
                             GZ2temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp)], 1)
    
    print ('a1GZ2 ', a1GZ2)
    print ('b1GZ2 ', b1GZ2)
    
    a1GX3, b1GX3 = np.polyfit(TG3temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp) ], \
                             GX3temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp)], 1)
    
    print ('a1GX3 ', a1GX3)
    print ('b1GX3 ', b1GX3)
    
    a1GY3, b1GY3 = np.polyfit(TG3temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp) ], \
                             GY3temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp)], 1)
    
    print ('a1GY3 ', a1GY3)
    print ('b1GY3 ', b1GY3)
    
    a1GZ3, b1GZ3 = np.polyfit(TG3temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp) ], \
                             GZ3temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp)], 1)
    
    print ('a1GZ3 ', a1GZ3)
    print ('b1GZ3 ', b1GZ3)
    
    
    # at beginning of second section
    a2GX1, b2GX1 = np.polyfit(TG1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex)], \
                             GX1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex)], 1)
    
    print ('a22GX1 ', a2GX1)
    print ('b22GX1 ', b2GX1)
    
    a2GY1, b2GY1 = np.polyfit(TG1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex) ], \
                             GY1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex)], 1)
    
    print ('a22GY1 ', a2GY1)
    print ('b22GY1 ', b2GY1)
    
    a2GX2, b2GX2 = np.polyfit(TG2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex) ], \
                             GX2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex)], 1)
    
    print ('a2GX2 ', a2GX2)
    print ('b2GX2 ', b2GX2)
    
    a2GZ2, b2GZ2 = np.polyfit(TG2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex) ], \
                             GZ2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex)], 1)
    
    print ('a2GZ2 ', a2GZ2)
    print ('b2GZ2 ', b2GZ2)
    
    a2GX3, b2GX3 = np.polyfit(TG3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex) ], \
                             GX3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex)], 1)
    
    print ('a2GX3 ', a2GX3)
    print ('b2GX3 ', b2GX3)
    
    a2GY3, b2GY3 = np.polyfit(TG3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex) ], \
                             GY3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex)], 1)
    
    print ('a2GY3 ', a2GY3)
    print ('b2GY3 ', b2GY3)
    
    a2GZ3, b2GZ3 = np.polyfit(TG3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex) ], \
                             GZ3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp + spliceindex)], 1)
    
    print ('a2GZ3 ', a2GZ3)
    print ('b2GZ3 ', b2GZ3)
    
    # find middle temperature count in first splice
    TG1splice1 = np.average(TG1temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp + spliceindex)])
    TG2splice1 = np.average(TG2temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp + spliceindex)])
    TG3splice1 = np.average(TG3temp[(stopindex1temp - startindex1temp - spliceindex):(stopindex1temp - startindex1temp + spliceindex)])

    # calculate drift from end of first section to beginning of second section
    
    GX1drift1 = a2GX1 * TG1splice1 + b2GX1 - a1GX1 * TG1splice1 - b1GX1
    GY1drift1 = a2GY1 * TG1splice1 + b2GY1 - a1GY1 * TG1splice1 - b1GY1
    GX2drift1 = a2GX2 * TG2splice1 + b2GX2 - a1GX2 * TG2splice1 - b1GX2
    GZ2drift1 = a2GZ2 * TG2splice1 + b2GZ2 - a1GZ2 * TG2splice1 - b1GZ2
    
    GX3drift1 = a2GX3 * TG3splice1 + b2GX3 - a1GX3 * TG3splice1 - b1GX3
    GY3drift1 = a2GY3 * TG3splice1 + b2GY3 - a1GY3 * TG3splice1 - b1GY3
    GZ3drift1 = a2GZ3 * TG3splice1 + b2GZ3 - a1GZ3 * TG3splice1 - b1GZ3
    
    # subtract the calculated drift from the data in the second section
    GX1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)] = \
            GX1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)] - GX1drift1 
    
    GY1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)] = \
            GY1temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)] - GY1drift1
            
    GX2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)] = \
            GX2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)] - GX2drift1
    
    GZ2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)] = \
            GZ2temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)] - GZ2drift1
    
    
    GX3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)] = \
            GX3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)] - GX3drift1 
    
    GY3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)] = \
            GY3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)] - GY3drift1
            
    GZ3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)] = \
            GZ3temp[(stopindex1temp - startindex1temp):(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)] - GZ3drift1
    
    (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)
    
    
    
    # at end of second section TO BE REMOVED/Change to end of THIRD section: DONE
           
    a3GX1, b3GX1 = np.polyfit(TG1temp[(stopindex1temp - startindex1temp + stopindex2temp - startindex2temp \
                                       + stopindex3temp - startindex3temp - spliceindex):\
                                      (stopindex1temp - startindex1temp + stopindex2temp - startindex2temp + \
                                       stopindex3temp - startindex3temp)], \
                                GX1temp[(stopindex1temp - startindex1temp + stopindex2temp - startindex2temp \
                                       + stopindex3temp - startindex3temp - spliceindex):\
                                      (stopindex1temp - startindex1temp + stopindex2temp - startindex2temp \
                                       + stopindex3temp - startindex3temp)], 1)
    
    print ('a3GX1 ', a3GX1)
    print ('b3GX1 ', b3GX1)
    
    a3GY1, b3GY1 = np.polyfit(TG1temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], \
                             GY1temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], 1)
    
    print ('a3GY1 ', a3GY1)
    print ('b3GY1 ', b3GY1)
    
    a3GX2, b3GX2 = np.polyfit(TG2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], \
                             GX2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], 1)
    print ('a3GX2 ', a3GX2)
    print ('b3GX2 ', b3GX2)
    
    a3GZ2, b3GZ2 = np.polyfit(TG2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], \
                             GZ2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], 1)
    
    print ('a3GZ2 ', a3GZ2)
    print ('b3GZ2 ', b3GZ2)
    
    a3GX3, b3GX3 = np.polyfit(TG3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], \
                             GX3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], 1)
    
    print ('a3GX3 ', a3GX3)
    print ('b3GX3 ', b3GX3)
    
    a3GY3, b3GY3 = np.polyfit(TG3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], \
                             GY3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], 1)
    
    print ('a3GY3 ', a3GY3)
    print ('b4GY3 ', b3GY3)
    
    a3GZ3, b3GZ3 = np.polyfit(TG3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], \
                             GZ3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp)- spliceindex):\
                                      ((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) \
                                       + (stopindex3temp - startindex3temp))], 1)
    
    print ('a3GZ3 ', a3GZ3)
    print ('b3GZ3 ', b3GZ3)
    
    
    # at beginning of third section. REMOVE/CHANGE to beginning of first section. DONE
    
    a4GX1, b4GX1 = np.polyfit(TG1temp[0 : spliceindex], GX1temp[0 : spliceindex], 1)
    
    print ('a42GX1 ', a4GX1)
    print ('b42GX1 ', b4GX1)
    
    
    a4GY1, b4GY1 = np.polyfit(TG1temp[0 : spliceindex], GY1temp[0 : spliceindex], 1)
    
    print ('a42GY1 ', a4GY1)
    print ('b42GY1 ', b4GY1)
    
    
    a4GX2, b4GX2 = np.polyfit(TG2temp[0 : spliceindex], GX2temp[0 : spliceindex], 1)
    
    print ('a4GX2 ', a4GX2)
    print ('b4GX2 ', b4GX2)
    
    
    a4GZ2, b4GZ2 =  np.polyfit(TG2temp[0 : spliceindex], GZ2temp[0 : spliceindex], 1)
    
    print ('a4GZ2 ', a4GZ2)
    print ('b4GZ2 ', b4GZ2)
    
    
    a4GX3, b4GX3 = np.polyfit(TG3temp[0 : spliceindex], GX3temp[0 : spliceindex], 1)
    
    
    print ('a4GX3 ', a4GX3)
    print ('b4GX3 ', b4GX3)
    
    
    a4GY3, b4GY3 = np.polyfit(TG3temp[0 : spliceindex], GY3temp[0 : spliceindex], 1)
    
    print ('a4GY3 ', a4GY3)
    print ('b4GY3 ', b4GY3)
    
    
    a4GZ3, b4GZ3 = np.polyfit(TG3temp[0 : spliceindex], GZ3temp[0 : spliceindex], 1)
    
    print ('a4GZ3 ', a4GZ3)
    print ('b4GZ3 ', b4GZ3)
    
    
    # find middle temperature count in second splice
      
    TG1splice2part1 = np.average(TG1temp[0:spliceindex])
    TG1splice2part2 = np.average(TG1temp[(stopindex1temp - startindex1temp + stopindex2temp - startindex2temp \
                                     + stopindex3temp - startindex3temp - spliceindex):\
                                    (stopindex1temp - startindex1temp + stopindex2temp - startindex2temp \
                                     + stopindex3temp - startindex3temp)])
    
    TG1splice2 = (TG1splice2part1 + TG1splice2part2)/2
    
    print ('TG1splice2', TG1splice2part1, TG1splice2part2)
    
    TG2splice2part1 = np.average(TG2temp[0:spliceindex])
    TG2splice2part2 = np.average(TG2temp[(stopindex1temp - startindex1temp + stopindex2temp - startindex2temp \
                                     + stopindex3temp - startindex3temp - spliceindex):\
                                    (stopindex1temp - startindex1temp + stopindex2temp - startindex2temp \
                                     + stopindex3temp - startindex3temp)])
    
    TG2splice2 = (TG2splice2part1 + TG2splice2part2)/2
    
    
    TG3splice2part1 = np.average(TG3temp[0:spliceindex])
    TG3splice2part2 = np.average(TG3temp[(stopindex1temp - startindex1temp + stopindex2temp - startindex2temp \
                                     + stopindex3temp - startindex3temp - spliceindex):\
                                    (stopindex1temp - startindex1temp + stopindex2temp - startindex2temp \
                                     + stopindex3temp - startindex3temp)])
    
    TG3splice2 = (TG3splice2part1 + TG3splice2part2)/2
    

    # calculate drift from end of second section to beginning of third section
    # CHANGE to calculate drift from end of third section to beginning of first section
    
    GX1drift2 = a4GX1 * TG1splice2 + b4GX1 - a3GX1 * TG1splice2 - b3GX1
    GY1drift2 = a4GY1 * TG1splice2 + b4GY1 - a3GY1 * TG1splice2 - b3GY1
    GX2drift2 = a4GX2 * TG2splice2 + b4GX2 - a3GX2 * TG2splice2 - b3GX2
    GZ2drift2 = a4GZ2 * TG2splice2 + b4GZ2 - a3GZ2 * TG2splice2 - b3GZ2
    GX3drift2 = a4GX3 * TG3splice2 + b4GX3 - a3GX3 * TG3splice2 - b3GX3
    GY3drift2 = a4GY3 * TG3splice2 + b4GY3 - a3GY3 * TG3splice2 - b3GY3
    GZ3drift2 = a4GZ3 * TG3splice2 + b4GZ3 - a3GZ3 * TG3splice2 - b3GZ3
    
    print (GX1drift2)
    print (GY1drift2)
    print (GX2drift2)
    print (GZ2drift2)
    print (GX3drift2)
    print (GY3drift2)
    print (GZ3drift2)
    
    
    # subtract the calculated drift from the data in the third section
    GX1temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = \
            GX1temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] + GX1drift2
                    
                    
    
    GY1temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)]  =\
            GY1temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] + GY1drift2
            
    GX2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] =  \
            GX2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] + GX2drift2
    
    GZ2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = \
            GZ2temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] + GZ2drift2
    
    
    GX3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = \
            GX3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] + GX3drift2
    
    GY3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = \
            GY3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] + GY3drift2
            
    GZ3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] = \
            GZ3temp[((stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp)] + GZ3drift2
    
    
    # remove overlapped data in splices
    
    indextemp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    AXtemp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    AYtemp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    AZtemp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    AXSTtemp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    AYSTtemp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    AZSTtemp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    GX1temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    GY1temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    GX2temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    GZ2temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    GX3temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    GY3temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    GZ3temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    TAtemp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    TG1temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    TG2temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    TG3temp2 = np.zeros(len(index_1t) + len(index_2t) + len(index_3t) - 4 * spliceindex)
    
    indextemp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = indextemp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    indextemp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = indextemp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    indextemp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = indextemp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]     
    
    
    AXtemp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = AXtemp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    AXtemp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = AXtemp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    AXtemp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = AXtemp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex] 
    
    
    AYtemp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = AYtemp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    AYtemp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = AYtemp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    AYtemp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = AYtemp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex] 
    
    
    AZtemp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = AYtemp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    AZtemp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = AZtemp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    AZtemp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = AZtemp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex] 

    
    AXSTtemp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = AXSTtemp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    AXSTtemp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = AXSTtemp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    AXSTtemp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = AXSTtemp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex] 
    
    
    AYSTtemp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = AYSTtemp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    AYSTtemp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = AYSTtemp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    AYSTtemp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = AYSTtemp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex] 
    
    
    AZSTtemp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = AYSTtemp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    AZSTtemp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = AZSTtemp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    AZSTtemp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = AZSTtemp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex] 
    
    
    
    GX1temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = GX1temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    GX1temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = GX1temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    GX1temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = GX1temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    GY1temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = GY1temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    GY1temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = GY1temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    GY1temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = GY1temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    GX2temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = GX2temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    GX2temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = GX2temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    GX2temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = GX2temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    GZ2temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = GZ2temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    GZ2temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = GZ2temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    GZ2temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = GZ2temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    
    GX3temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = GX3temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    GX3temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = GX3temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    GX3temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = GX3temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    GY3temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = GY3temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    GY3temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = GY3temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    GY3temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = GY3temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    GZ3temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = GZ3temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    GZ3temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = GZ3temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    GZ3temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = GZ3temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
   
    
    TAtemp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = TAtemp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    TAtemp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = TAtemp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    TAtemp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = TAtemp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    
    TG1temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = TG1temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    TG1temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = TG1temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    TG1temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = TG1temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    TG2temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = TG2temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    TG2temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = TG2temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    TG2temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = TG2temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    TG3temp2[0:(stopindex1temp - startindex1temp - 2*spliceindex)] \
    = TG3temp[spliceindex:(stopindex1temp - startindex1temp - spliceindex)]
    
    TG3temp2[(stopindex1temp - startindex1temp - 2*spliceindex):\
               (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex)] \
    = TG3temp[(stopindex1temp - startindex1temp + spliceindex): \
                (stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp)]
    
    TG3temp2[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp - 3*spliceindex): \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+ \
              (stopindex3temp - startindex3temp) - 4*spliceindex] \
    = TG3temp[(stopindex1temp - startindex1temp) + (stopindex2temp - startindex2temp) : \
              (stopindex1temp - startindex1temp)+ (stopindex2temp - startindex2temp)+(stopindex3temp - startindex3temp) - spliceindex]
    
    
    
    save_cal_file(logfile_TCout, sheetname_TCout, indextemp2, AXtemp2, AYtemp2, AZtemp2, AXSTtemp2, AYSTtemp2, AZSTtemp2, \
                  GX1temp2, GY1temp2, GX2temp2, GZ2temp2, GX3temp2, GY3temp2, GZ3temp2, TAtemp2, TG1temp2, TG2temp2, TG3temp2) 
    
    
    fig, ax1 = plt.subplots()
    plt.plot(indextemp2, GX1temp2, label = 'GX1')
    plt.plot(indextemp2, GY1temp2, label = 'GY1')
    plt.plot(indextemp2, GX2temp2, label = 'GX2')
    plt.plot(indextemp2, GZ2temp2, label = 'GZ2')
    
    plt.legend(loc = 'best')
    
    ax2 = ax1.twinx()
    plt.plot(indextemp2, TG1temp2, 'r', label = 'TG1count')
    plt.legend(loc = 'best')
        
    plt.grid()
    
    plt.savefig(toolSN + ' split step 2 plot 1.png', bbox_inches='tight', dpi=250)
    
#    plt.show()
    
    
    
    fig, ax1 = plt.subplots()
    plt.plot(TG1temp2, GX1temp2, label = 'GX1')
    plt.plot(TG1temp2, GY1temp2, label = 'GY1')
    plt.plot(TG2temp2, GX2temp2, label = 'GX2')
    plt.plot(TG2temp2, GZ2temp2, label = 'GZ2')
        
    plt.legend(loc = 'best')
    plt.grid()
    
    plt.savefig(toolSN + ' split step 2 plot 2.png', bbox_inches='tight', dpi=250)
    
#    plt.show()
    
    
            
    fig, ax1 = plt.subplots()
    plt.plot(TG3temp2, GX3temp2, label = 'GX3')

    plt.plot(TG3temp2, GY3temp2, label = 'GY3')
    plt.plot(TG3temp2, GZ3temp2, label = 'GZ3')
        
    plt.legend(loc = 'best')
    plt.grid()
    
    plt.savefig(toolSN + ' split step 2 plot 3.png', bbox_inches='tight', dpi=250)
    
#    plt.show()
    
