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

# R11 adapt to FW A112, with RUN L/DUMP L command, including all sensors from ST sensor, ADXL355, ADXRS 290 and magnetometer
# R14 include support for control signal at 28C, write figures to file
# R15 change loacl gravity at Heimdal to 9.821m/s^2
# R16 for python 3.8.5
# R17 fixed MAA to MAB bug in write constant file script
# R18 introduced gyro result image script, fixed minor error in plot 117. No change in this script

from __future__ import division
import xlrd
import xlsxwriter


toolSN = '3331'

abstemp1 = -5           # temperature in deg C as set in the temperature bath (not neccessarily identical to actual tool temperature)
abstemp2 = 28
abstemp3 = 60

g_range_ADXL = 4        # for cal file names only
g_range_ST = 16          # for cal file names only



# build cal file names

ac_T1_ADXL_name = toolSN + ' ' + str(abstemp1) + 'C' + ' ' + str(g_range_ADXL) + 'g acc cal.xlsx'  
ac_T2_ADXL_name = toolSN + ' ' + str(abstemp2) + 'C' + ' ' + str(g_range_ADXL) + 'g acc cal.xlsx'
ac_T3_ADXL_name = toolSN + ' ' + str(abstemp3) + 'C' + ' ' + str(g_range_ADXL) + 'g acc cal.xlsx'

ac_T1_ST_name = toolSN + ' ST ' + str(abstemp1) + 'C' + ' ' + str(g_range_ST) + 'g acc cal.xlsx'  
ac_T2_ST_name = toolSN + ' ST ' + str(abstemp2) + 'C' + ' ' + str(g_range_ST) + 'g acc cal.xlsx'
ac_T3_ST_name = toolSN + ' ST ' + str(abstemp3) + 'C' + ' ' + str(g_range_ST) + 'g acc cal.xlsx'

gyro_orto_file_T1_ADXRS = toolSN + ' ' + str(abstemp1) + 'C' + ' ' + 'gyro orto.xlsx'
gyro_orto_file_T2_ADXRS = toolSN + ' ' + str(abstemp2) + 'C' + ' ' + 'gyro orto.xlsx'
gyro_orto_file_T3_ADXRS = toolSN + ' ' + str(abstemp3) + 'C' + ' ' + 'gyro orto.xlsx'

gyro_orto_file_T1_ST = toolSN + ' ST ' + str(abstemp1) + 'C' + ' ' + 'gyro orto.xlsx'
gyro_orto_file_T2_ST = toolSN + ' ST ' + str(abstemp2) + 'C' + ' ' + 'gyro orto.xlsx'
gyro_orto_file_T3_ST = toolSN + ' ST ' + str(abstemp3) + 'C' + ' ' + 'gyro orto.xlsx'

gyrogsenscal_file_T1_ADXRS = toolSN + ' ' + str(abstemp1) + 'C' + ' ' + 'gyro gsens erc.xlsx'
gyrogsenscal_file_T2_ADXRS = toolSN + ' ' + str(abstemp2) + 'C' + ' ' + 'gyro gsens erc.xlsx'
gyrogsenscal_file_T3_ADXRS = toolSN + ' ' + str(abstemp3) + 'C' + ' ' + 'gyro gsens erc.xlsx'


gyrogsenscal_file_T1_ST = toolSN + ' ST ' + str(abstemp1) + 'C' + ' ' + 'gyro gsens erc.xlsx'
gyrogsenscal_file_T2_ST = toolSN + ' ST ' + str(abstemp2) + 'C' + ' ' + 'gyro gsens erc.xlsx'
gyrogsenscal_file_T3_ST = toolSN + ' ST ' + str(abstemp3) + 'C' + ' ' + 'gyro gsens erc.xlsx'

gyroTcal_file_ADXRS = toolSN + ' gyro Tconst.xlsx'
gyroTcal_file_ST = toolSN + ' ST gyro Tconst.xlsx'

# output file name
constant_file_name = toolSN + ' calibration constants MAB edit.xlsx'


# read gyro scale factor T1 file 
workbook = xlrd.open_workbook(gyro_orto_file_T1_ADXRS)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
faa = col[2]
fab = col[3]
fac = col[4]

col = worksheet.col_values(1)
fad = col[2]
fae = col[3]
faf = col[4]

col = worksheet.col_values(2)
fag = col[2]
fah = col[3]
fai = col[4]

col = worksheet.col_values(3)
fba = col[2]
fbb = col[3]
fbc = col[4]

col = worksheet.col_values(4)
fbd = col[2]
fbe = col[3]
fbf = col[4]

col = worksheet.col_values(5)
fbg = col[2]
fbh = col[3]
fbi = col[4]

col = worksheet.col_values(6)
fca = col[2]
fcb = col[3]
fcc = col[4]

col = worksheet.col_values(7)
fcd = col[2]
fce = col[3]
fcf = col[4]

col = worksheet.col_values(8)
fcg = col[2]
fch = col[3]
fci = col[4]

col = worksheet.col_values(9)
fcj = col[2]

col = worksheet.col_values(10)
fck = col[2]

col = worksheet.col_values(11)
fcl = col[2]

# read gyro g sensitivty T1 file 
workbook = xlrd.open_workbook(gyrogsenscal_file_T1_ADXRS)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
fda = col[1]
fdb = col[2]
fdc = col[3]

col = worksheet.col_values(1)
fdd = col[1]
fde = col[2]
fdf = col[3]

col = worksheet.col_values(2)
fdg = col[1]
fdh = col[2]
fdi = col[3]

col = worksheet.col_values(3)
fdj = col[1]
fdk = col[2]
fdl = col[3]

col = worksheet.col_values(4)
fdm = col[1]

col = worksheet.col_values(5)
fdn = col[1]

col = worksheet.col_values(6)
fdo = col[1]

# read gyro scale factor T2 file 
workbook = xlrd.open_workbook(gyro_orto_file_T2_ADXRS)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
fea = col[2]
feb = col[3]
fec = col[4]

col = worksheet.col_values(1)
fed = col[2]
fee = col[3]
fef = col[4]

col = worksheet.col_values(2)
feg = col[2]
feh = col[3]
fei = col[4]

col = worksheet.col_values(3)
ffa = col[2]
ffb = col[3]
ffc = col[4]

col = worksheet.col_values(4)
ffd = col[2]
ffe = col[3]
fff = col[4]

col = worksheet.col_values(5)
ffg = col[2]
ffh = col[3]
ffi = col[4]

col = worksheet.col_values(6)
fga = col[2]
fgb = col[3]
fgc = col[4]

col = worksheet.col_values(7)
fgd = col[2]
fge = col[3]
fgf = col[4]

col = worksheet.col_values(8)
fgg = col[2]
fgh = col[3]
fgi = col[4]

col = worksheet.col_values(9)
fgj = col[2]

col = worksheet.col_values(10)
fgk = col[2]

col = worksheet.col_values(11)
fgl = col[2]

# read gyro g sensitivty T2 file 
workbook = xlrd.open_workbook(gyrogsenscal_file_T2_ADXRS)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
fha = col[1]
fhb = col[2]
fhc = col[3]

col = worksheet.col_values(1)
fhd = col[1]
fhe = col[2]
fhf = col[3]

col = worksheet.col_values(2)
fhg = col[1]
fhh = col[2]
fhi = col[3]

col = worksheet.col_values(3)
fhj = col[1]
fhk = col[2]
fhl = col[3]

col = worksheet.col_values(4)
fhm = col[1]

col = worksheet.col_values(5)
fhn = col[1]

col = worksheet.col_values(6)
fho = col[1]

# read gyro scale factor T3 file 
workbook = xlrd.open_workbook(gyro_orto_file_T3_ADXRS)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
fia = col[2]
fib = col[3]
fic = col[4]

col = worksheet.col_values(1)
fid = col[2]
fie = col[3]
fif = col[4]

col = worksheet.col_values(2)
fig = col[2]
fih = col[3]
fii = col[4]

col = worksheet.col_values(3)
fja = col[2]
fjb = col[3]
fjc = col[4]

col = worksheet.col_values(4)
fjd = col[2]
fje = col[3]
fjf = col[4]

col = worksheet.col_values(5)
fjg = col[2]
fjh = col[3]
fji = col[4]

col = worksheet.col_values(6)
fka = col[2]
fkb = col[3]
fkc = col[4]

col = worksheet.col_values(7)
fkd = col[2]
fke = col[3]
fkf = col[4]

col = worksheet.col_values(8)
fkg = col[2]
fkh = col[3]
fki = col[4]

col = worksheet.col_values(9)
fkj = col[2]

col = worksheet.col_values(10)
fkk = col[2]

col = worksheet.col_values(11)
fkl = col[2]

# read gyro g sensitivty T3 file 
workbook = xlrd.open_workbook(gyrogsenscal_file_T3_ADXRS)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
fla = col[1]
flb = col[2]
flc = col[3]

col = worksheet.col_values(1)
fld = col[1]
fle = col[2]
flf = col[3]

col = worksheet.col_values(2)
flg = col[1]
flh = col[2]
fli = col[3]

col = worksheet.col_values(3)
flj = col[1]
flk = col[2]
fll = col[3]

col = worksheet.col_values(4)
flm = col[1]

col = worksheet.col_values(5)
fln = col[1]

col = worksheet.col_values(6)
flo = col[1]

# read gyro temperature response file 
workbook = xlrd.open_workbook(gyroTcal_file_ADXRS)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
fma = col[1]
fmb = col[2]
fmc = col[3]
fmd = col[4]
fme = col[5]
fmf = col[6]

col = worksheet.col_values(1)
fmg = col[1]
fmh = col[2]
fmi = col[3]
fmj = col[4]
fmk = col[5]
fml = col[6]

col = worksheet.col_values(2)
fmm = col[1]
fmn = col[2]
fmo = col[3]
fmp = col[4]
fmq = col[5]
fmr = col[6]

col = worksheet.col_values(3)
fms = col[1]
fmt = col[2]
fmu = col[3]
fmv = col[4]
fmw = col[5]
fmx = col[6]

col = worksheet.col_values(4)
fna = col[1]
fnb = col[2]

col = worksheet.col_values(5)
fnc = col[1]
fnd = col[2]

col = worksheet.col_values(6)
fne = col[1]
fnf = col[2]

col = worksheet.col_values(7)
fng = col[1]
fnh = col[2]

col = worksheet.col_values(8)
foa = col[1]

col = worksheet.col_values(9)
fob = col[1]

col = worksheet.col_values(10)
foc = col[1]

col = worksheet.col_values(11)
fod = col[1]

# read ADXL lowest temperature scal factor file 
workbook = xlrd.open_workbook(ac_T1_ADXL_name)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
maa = col[1]

col = worksheet.col_values(1)
mab = col[1]

col = worksheet.col_values(2)
mac = col[1]

col = worksheet.col_values(3)
mad = col[2]
mae = col[3]
maf = col[4]

col = worksheet.col_values(4)
mag = col[2]
mah = col[3]
mai = col[4]

col = worksheet.col_values(5)
maj = col[2]
mak = col[3]
mal = col[4]

col = worksheet.col_values(9)
mam = col[1]

col = worksheet.col_values(10)
man = col[1]

# read ADXL middle temperature scale factor file 
workbook = xlrd.open_workbook(ac_T2_ADXL_name)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
mba = col[1]

col = worksheet.col_values(1)
mbb = col[1]

col = worksheet.col_values(2)
mbc = col[1]

col = worksheet.col_values(3)
mbd = col[2]
mbe = col[3]
mbf = col[4]

col = worksheet.col_values(4)
mbg = col[2]
mbh = col[3]
mbi = col[4]

col = worksheet.col_values(5)
mbj = col[2]
mbk = col[3]
mbl = col[4]

col = worksheet.col_values(9)
mbm = col[1]

col = worksheet.col_values(10)
mbn = col[1]

# read ADXL high temperature scale factor file 
workbook = xlrd.open_workbook(ac_T3_ADXL_name)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
mca = col[1]

col = worksheet.col_values(1)
mcb = col[1]

col = worksheet.col_values(2)
mcc = col[1]

col = worksheet.col_values(3)
mcd = col[2]
mce = col[3]
mcf = col[4]

col = worksheet.col_values(4)
mcg = col[2]
mch = col[3]
mci = col[4]

col = worksheet.col_values(5)
mcj = col[2]
mck = col[3]
mcl = col[4]

col = worksheet.col_values(9)
mcm = col[1]

col = worksheet.col_values(10)
mcn = col[1]


# read ST gyro scale factor T1 file 
workbook = xlrd.open_workbook(gyro_orto_file_T1_ST)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
saa = col[1]
sab = col[2]
sac = col[3]

col = worksheet.col_values(1)
sad = col[1]
sae = col[2]
saf = col[3]

col = worksheet.col_values(2)
sag = col[1]
sah = col[2]
sai = col[3]

col = worksheet.col_values(3)
saj = col[1]

col = worksheet.col_values(4)
sak = col[1]

# read ST gyro g sensitivity T1 file 
workbook = xlrd.open_workbook(gyrogsenscal_file_T1_ST)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
sda = col[1]
sdb = col[2]
sdc = col[3]

col = worksheet.col_values(1)
sdd = col[1]
sde = col[2]
sdf = col[3]

col = worksheet.col_values(2)
sdg = col[1]
sdh = col[2]
sdi = col[3]

col = worksheet.col_values(3)
sdj = col[1]

col = worksheet.col_values(4)
sdk = col[1]

# read ST gyro scale factor T2 file 
workbook = xlrd.open_workbook(gyro_orto_file_T2_ST)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
sea = col[1]
seb = col[2]
sec = col[3]

col = worksheet.col_values(1)
sed = col[1]
see = col[2]
sef = col[3]

col = worksheet.col_values(2)
seg = col[1]
seh = col[2]
sei = col[3]

col = worksheet.col_values(3)
sej = col[1]

col = worksheet.col_values(4)
sek = col[1]

# read ST gyro g sensitivity T2 file 
workbook = xlrd.open_workbook(gyrogsenscal_file_T2_ST)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
sha = col[1]
shb = col[2]
shc = col[3]

col = worksheet.col_values(1)
shd = col[1]
she = col[2]
shf = col[3]

col = worksheet.col_values(2)
shg = col[1]
shh = col[2]
shi = col[3]

col = worksheet.col_values(3)
shj = col[1]

col = worksheet.col_values(4)
shk = col[1]

# read ST gyro scale factor T3 file 
workbook = xlrd.open_workbook(gyro_orto_file_T3_ST)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
sia = col[1]
sib = col[2]
sic = col[3]

col = worksheet.col_values(1)
sid = col[1]
sie = col[2]
sif = col[3]

col = worksheet.col_values(2)
sig = col[1]
sih = col[2]
sii = col[3]

col = worksheet.col_values(3)
sij = col[1]

col = worksheet.col_values(4)
sik = col[1]

# read ST gyro g sensitivity T3 file 
workbook = xlrd.open_workbook(gyrogsenscal_file_T3_ST)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
sla = col[1]
slb = col[2]
slc = col[3]

col = worksheet.col_values(1)
sld = col[1]
sle = col[2]
slf = col[3]

col = worksheet.col_values(2)
slg = col[1]
slh = col[2]
sli = col[3]

col = worksheet.col_values(3)
slj = col[1]

col = worksheet.col_values(4)
slk = col[1]

# read ST gyro temperature response file 
workbook = xlrd.open_workbook(gyroTcal_file_ST)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
sma = col[1]
smb = col[2]
smc = col[3]
smd = col[4]
sme = col[5]
smf = col[6]

col = worksheet.col_values(1)
smg = col[1]
smh = col[2]
smi = col[3]
smj = col[4]
smk = col[5]
sml = col[6]

col = worksheet.col_values(2)
smm = col[1]
smn = col[2]
smo = col[3]
smp = col[4]
smq = col[5]
smr = col[6]

col = worksheet.col_values(3)
sna = col[1]
snb = col[2]

col = worksheet.col_values(4)
snc = col[1]
snd = col[2]

col = worksheet.col_values(5)
sne = col[1]
snf = col[2]

col = worksheet.col_values(6)
soa = col[1]

col = worksheet.col_values(7)
sob = col[1]


# read ST lowest temperature scal factor file 
workbook = xlrd.open_workbook(ac_T1_ST_name)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
taa = col[1]

col = worksheet.col_values(1)
tab = col[1]

col = worksheet.col_values(2)
tac = col[1]

col = worksheet.col_values(3)
tad = col[2]
tae = col[3]
taf = col[4]

col = worksheet.col_values(4)
tag = col[2]
tah = col[3]
tai = col[4]

col = worksheet.col_values(5)
taj = col[2]
tak = col[3]
tal = col[4]

col = worksheet.col_values(9)
tam = col[1]

col = worksheet.col_values(10)
tan = col[1]

# read ST middle temperature scale factor file 
workbook = xlrd.open_workbook(ac_T2_ST_name)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
tba = col[1]

col = worksheet.col_values(1)
tbb = col[1]

col = worksheet.col_values(2)
tbc = col[1]

col = worksheet.col_values(3)
tbd = col[2]
tbe = col[3]
tbf = col[4]

col = worksheet.col_values(4)
tbg = col[2]
tbh = col[3]
tbi = col[4]

col = worksheet.col_values(5)
tbj = col[2]
tbk = col[3]
tbl = col[4]

col = worksheet.col_values(9)
tbm = col[1]

col = worksheet.col_values(10)
tbn = col[1]

# read ST high temperature scale factor file 
workbook = xlrd.open_workbook(ac_T3_ST_name)
worksheet = workbook.sheet_by_name('Ark1')

col = worksheet.col_values(0)
tca = col[1]

col = worksheet.col_values(1)
tcb = col[1]

col = worksheet.col_values(2)
tcc = col[1]

col = worksheet.col_values(3)
tcd = col[2]
tce = col[3]
tcf = col[4]

col = worksheet.col_values(4)
tcg = col[2]
tch = col[3]
tci = col[4]

col = worksheet.col_values(5)
tcj = col[2]
tck = col[3]
tcl = col[4]

col = worksheet.col_values(9)
tcm = col[1]

col = worksheet.col_values(10)
tcn = col[1]

# write file
workbook = xlsxwriter.Workbook(constant_file_name)
worksheet = workbook.add_worksheet('Ark1')

worksheet.write(0,0,'FAA:')
worksheet.write(1,0,'FAB:')
worksheet.write(2,0,'FAC:')
worksheet.write(3,0,'FAD:')
worksheet.write(4,0,'FAE:')
worksheet.write(5,0,'FAF:')
worksheet.write(6,0,'FAG:')
worksheet.write(7,0,'FAH:')
worksheet.write(8,0,'FAI:')
worksheet.write(9,0,'FBA:')
worksheet.write(10,0,'FBB:')

worksheet.write(11,0,'FBC:')
worksheet.write(12,0,'FBD:')
worksheet.write(13,0,'FBE:')
worksheet.write(14,0,'FBF:')
worksheet.write(15,0,'FBG:')
worksheet.write(16,0,'FBH:')
worksheet.write(17,0,'FBI:')
worksheet.write(18,0,'FCA:')
worksheet.write(19,0,'FCB:')
worksheet.write(20,0,'FCC:')

worksheet.write(21,0,'FCD:')
worksheet.write(22,0,'FCE:')
worksheet.write(23,0,'FCF:')
worksheet.write(24,0,'FCG:')
worksheet.write(25,0,'FCH:')
worksheet.write(26,0,'FCI:')
worksheet.write(27,0,'FCJ:')
worksheet.write(28,0,'FCK:')
worksheet.write(29,0,'FCL:')
worksheet.write(30,0,'FDA:')

worksheet.write(31,0,'FDB:')
worksheet.write(32,0,'FDC:')
worksheet.write(33,0,'FDD:')
worksheet.write(34,0,'FDE:')
worksheet.write(35,0,'FDF:')
worksheet.write(36,0,'FDG:')
worksheet.write(37,0,'FDH:')
worksheet.write(38,0,'FDI:')
worksheet.write(39,0,'FDJ:')
worksheet.write(40,0,'FDK:')

worksheet.write(41,0,'FDL:')
worksheet.write(42,0,'FDM:')
worksheet.write(43,0,'FDN:')
worksheet.write(44,0,'FDO:')
worksheet.write(45,0,'FEA:')
worksheet.write(46,0,'FEB:')
worksheet.write(47,0,'FEC:')
worksheet.write(48,0,'FED:')
worksheet.write(49,0,'FEE:')
worksheet.write(50,0,'FEF:')

worksheet.write(51,0,'FEG:')
worksheet.write(52,0,'FEH:')
worksheet.write(53,0,'FEI:')
worksheet.write(54,0,'FFA:')
worksheet.write(55,0,'FFB:')
worksheet.write(56,0,'FFC:')
worksheet.write(57,0,'FFD:')
worksheet.write(58,0,'FFE:')
worksheet.write(59,0,'FFF:')
worksheet.write(60,0,'FFG:')

worksheet.write(61,0,'FFH:')
worksheet.write(62,0,'FFI:')
worksheet.write(63,0,'FGA:')
worksheet.write(64,0,'FGB:')
worksheet.write(65,0,'FGC:')
worksheet.write(66,0,'FGD:')
worksheet.write(67,0,'FGE:')
worksheet.write(68,0,'FGF:')
worksheet.write(69,0,'FGG:')
worksheet.write(70,0,'FGH:')
worksheet.write(71,0,'FGI:')

worksheet.write(72,0,'FGJ:')
worksheet.write(73,0,'FGK:')
worksheet.write(74,0,'FGL:')
worksheet.write(75,0,'FHA:')
worksheet.write(76,0,'FHB:')
worksheet.write(77,0,'FHC:')
worksheet.write(78,0,'FHD:')
worksheet.write(79,0,'FHE:')
worksheet.write(80,0,'FHF:')
worksheet.write(81,0,'FHG:')

worksheet.write(82,0,'FHH:')
worksheet.write(83,0,'FHI:')
worksheet.write(84,0,'FHJ:')
worksheet.write(85,0,'FHK:')
worksheet.write(86,0,'FHL:')
worksheet.write(87,0,'FHM:')
worksheet.write(88,0,'FHN:')
worksheet.write(89,0,'FHO:')
worksheet.write(90,0,'FIA:')
worksheet.write(91,0,'FIB:')

worksheet.write(92,0,'FIC:')
worksheet.write(93,0,'FID:')
worksheet.write(94,0,'FIE:')
worksheet.write(95,0,'FIF:')
worksheet.write(96,0,'FIG:')
worksheet.write(97,0,'FIH:')
worksheet.write(98,0,'FII:')
worksheet.write(99,0,'FJA:')
worksheet.write(100,0,'FJB:')
worksheet.write(101,0,'FJC:')

worksheet.write(102,0,'FJD:')
worksheet.write(103,0,'FJE:')
worksheet.write(104,0,'FJF:')
worksheet.write(105,0,'FJG:')
worksheet.write(106,0,'FJH:')
worksheet.write(107,0,'FJI:')
worksheet.write(108,0,'FKA:')
worksheet.write(109,0,'FKB:')
worksheet.write(110,0,'FKC:')
worksheet.write(111,0,'FKD:')

worksheet.write(112,0,'FKE:')
worksheet.write(113,0,'FKF:')
worksheet.write(114,0,'FKG:')
worksheet.write(115,0,'FKH:')
worksheet.write(116,0,'FKI:')
worksheet.write(117,0,'FKJ:')
worksheet.write(118,0,'FKK:')
worksheet.write(119,0,'FKL:')
worksheet.write(120,0,'FLA:')
worksheet.write(121,0,'FLB:')

worksheet.write(122,0,'FLC:')
worksheet.write(123,0,'FLD:')
worksheet.write(124,0,'FLE:')
worksheet.write(125,0,'FLF:')
worksheet.write(126,0,'FLG:')
worksheet.write(127,0,'FLH:')
worksheet.write(128,0,'FLI:')
worksheet.write(129,0,'FLJ:')
worksheet.write(130,0,'FLK:')
worksheet.write(131,0,'FLL:')

worksheet.write(132,0,'FLM:')
worksheet.write(133,0,'FLN:')
worksheet.write(134,0,'FLO:')
worksheet.write(135,0,'FMA:')
worksheet.write(136,0,'FMB:')
worksheet.write(137,0,'FMC:')
worksheet.write(138,0,'FMD:')
worksheet.write(139,0,'FME:')
worksheet.write(140,0,'FMF:')
worksheet.write(141,0,'FMG:')

worksheet.write(142,0,'FMH:')
worksheet.write(143,0,'FMI:')
worksheet.write(144,0,'FMJ:')
worksheet.write(145,0,'FMK:')
worksheet.write(146,0,'FML:')
worksheet.write(147,0,'FMM:')
worksheet.write(148,0,'FMN:')
worksheet.write(149,0,'FMO:')
worksheet.write(150,0,'FMP:')
worksheet.write(151,0,'FMQ:')

worksheet.write(152,0,'FMR:')
worksheet.write(153,0,'FMS:')
worksheet.write(154,0,'FMT:')
worksheet.write(155,0,'FMU:')
worksheet.write(156,0,'FMV:')
worksheet.write(157,0,'FMW:')
worksheet.write(158,0,'FMX:')
worksheet.write(159,0,'FNA:')
worksheet.write(160,0,'FNB:')
worksheet.write(161,0,'FNC:')

worksheet.write(162,0,'FND:')
worksheet.write(163,0,'FNE:')
worksheet.write(164,0,'FNF:')
worksheet.write(165,0,'FNG:')
worksheet.write(166,0,'FNH:')
worksheet.write(167,0,'FOA:')
worksheet.write(168,0,'FOB:')
worksheet.write(169,0,'FOC:')
worksheet.write(170,0,'FOD:')
worksheet.write(171,0,'MAA:')

worksheet.write(172,0,'MAB:')
worksheet.write(173,0,'MAC:')
worksheet.write(174,0,'MAD:')
worksheet.write(175,0,'MAE:')
worksheet.write(176,0,'MAF:')
worksheet.write(177,0,'MAG:')
worksheet.write(178,0,'MAH:')
worksheet.write(179,0,'MAI:')
worksheet.write(180,0,'MAJ:')
worksheet.write(181,0,'MAK:')

worksheet.write(182,0,'MAL:')
worksheet.write(183,0,'MAM:')
worksheet.write(184,0,'MAN:')
worksheet.write(185,0,'MBA:')
worksheet.write(186,0,'MBB:')
worksheet.write(187,0,'MBC:')
worksheet.write(188,0,'MBD:')
worksheet.write(189,0,'MBE:')
worksheet.write(190,0,'MBF:')
worksheet.write(191,0,'MBG:')

worksheet.write(192,0,'MBH:')
worksheet.write(193,0,'MBI:')
worksheet.write(194,0,'MBJ:')
worksheet.write(195,0,'MBK:')
worksheet.write(196,0,'MBL:')
worksheet.write(197,0,'MBM:')
worksheet.write(198,0,'MBN:')
worksheet.write(199,0,'MCA:')
worksheet.write(200,0,'MCB:')
worksheet.write(201,0,'MCC:')

worksheet.write(202,0,'MCD:')
worksheet.write(203,0,'MCE:')
worksheet.write(204,0,'MCF:')
worksheet.write(205,0,'MCG:')
worksheet.write(206,0,'MCH:')
worksheet.write(207,0,'MCI:')
worksheet.write(208,0,'MCJ:')
worksheet.write(209,0,'MCK:')
worksheet.write(210,0,'MCL:')
worksheet.write(211,0,'MCM:')

worksheet.write(212,0,'MCN:')
worksheet.write(213,0,'SAA:')
worksheet.write(214,0,'SAB:')
worksheet.write(215,0,'SAC:')
worksheet.write(216,0,'SAD:')
worksheet.write(217,0,'SAE:')
worksheet.write(218,0,'SAF:')
worksheet.write(219,0,'SAG:')
worksheet.write(220,0,'SAH:')
worksheet.write(221,0,'SAI:')

worksheet.write(222,0,'SAJ:')
worksheet.write(223,0,'SAK:')
worksheet.write(224,0,'SDA:')
worksheet.write(225,0,'SDB:')
worksheet.write(226,0,'SDC:')
worksheet.write(227,0,'SDD:')
worksheet.write(228,0,'SDE:')
worksheet.write(229,0,'SDF:')
worksheet.write(230,0,'SDG:')
worksheet.write(231,0,'SDH:')

worksheet.write(232,0,'SDI:')
worksheet.write(233,0,'SDJ:')
worksheet.write(234,0,'SDK:')
worksheet.write(235,0,'SEA:')
worksheet.write(236,0,'SEB:')
worksheet.write(237,0,'SEC:')
worksheet.write(238,0,'SED:')
worksheet.write(239,0,'SEE:')
worksheet.write(240,0,'SEF:')
worksheet.write(241,0,'SEG:')

worksheet.write(242,0,'SEH:')
worksheet.write(243,0,'SEI:')
worksheet.write(244,0,'SEJ:')
worksheet.write(245,0,'SEK:')
worksheet.write(246,0,'SHA:')
worksheet.write(247,0,'SHB:')
worksheet.write(248,0,'SHC:')
worksheet.write(249,0,'SHD:')
worksheet.write(250,0,'SHE:')
worksheet.write(251,0,'SHF:')

worksheet.write(252,0,'SHG:')
worksheet.write(253,0,'SHH:')
worksheet.write(254,0,'SHI:')
worksheet.write(255,0,'SHJ:')
worksheet.write(256,0,'SHK:')
worksheet.write(257,0,'SIA:')
worksheet.write(258,0,'SIB:')
worksheet.write(259,0,'SIC:')
worksheet.write(260,0,'SID:')
worksheet.write(261,0,'SIE:')

worksheet.write(262,0,'SIF:')
worksheet.write(263,0,'SIG:')
worksheet.write(264,0,'SIH:')
worksheet.write(265,0,'SII:')
worksheet.write(266,0,'SIJ:')
worksheet.write(267,0,'SIK:')
worksheet.write(268,0,'SLA:')
worksheet.write(269,0,'SLB:')
worksheet.write(270,0,'SLC:')
worksheet.write(271,0,'SLD:')

worksheet.write(272,0,'SLE:')
worksheet.write(273,0,'SLF:')
worksheet.write(274,0,'SLG:')
worksheet.write(275,0,'SLH:')
worksheet.write(276,0,'SLI:')
worksheet.write(277,0,'SLJ:')
worksheet.write(278,0,'SLK:')
worksheet.write(279,0,'SMA:')
worksheet.write(280,0,'SMB:')
worksheet.write(281,0,'SMC:')

worksheet.write(282,0,'SMD:')
worksheet.write(283,0,'SME:')
worksheet.write(284,0,'SMF:')
worksheet.write(285,0,'SMG:')
worksheet.write(286,0,'SMH:')
worksheet.write(287,0,'SMI:')
worksheet.write(288,0,'SMJ:')
worksheet.write(289,0,'SMK:')
worksheet.write(290,0,'SML:')
worksheet.write(291,0,'SMM:')

worksheet.write(292,0,'SMN:')
worksheet.write(293,0,'SMO:')
worksheet.write(294,0,'SMP:')
worksheet.write(295,0,'SMQ:')
worksheet.write(296,0,'SMR:')
worksheet.write(297,0,'SNA:')
worksheet.write(298,0,'SNB:')
worksheet.write(299,0,'SNC:')
worksheet.write(300,0,'SND:')
worksheet.write(301,0,'SNE:')

worksheet.write(302,0,'SNF:')
worksheet.write(303,0,'SOA:')
worksheet.write(304,0,'SOB:')
worksheet.write(305,0,'TAA:')
worksheet.write(306,0,'TAB:')
worksheet.write(307,0,'TAC:')
worksheet.write(308,0,'TAD:')
worksheet.write(309,0,'TAE:')
worksheet.write(310,0,'TAF:')
worksheet.write(311,0,'TAG:')

worksheet.write(312,0,'TAH:')
worksheet.write(313,0,'TAI:')
worksheet.write(314,0,'TAJ:')
worksheet.write(315,0,'TAK:')
worksheet.write(316,0,'TAL:')
worksheet.write(317,0,'TAM:')
worksheet.write(318,0,'TAN:')
worksheet.write(319,0,'TBA:')
worksheet.write(320,0,'TBB:')
worksheet.write(321,0,'TBC:')

worksheet.write(322,0,'TBD:')
worksheet.write(323,0,'TBE:')
worksheet.write(324,0,'TBF:')
worksheet.write(325,0,'TBG:')
worksheet.write(326,0,'TBH:')
worksheet.write(327,0,'TBI:')
worksheet.write(328,0,'TBJ:')
worksheet.write(329,0,'TBK:')
worksheet.write(330,0,'TBL:')
worksheet.write(331,0,'TBM:')

worksheet.write(332,0,'TBN:')
worksheet.write(333,0,'TCA:')
worksheet.write(334,0,'TCB:')
worksheet.write(335,0,'TCC:')
worksheet.write(336,0,'TCD:')
worksheet.write(337,0,'TCE:')
worksheet.write(338,0,'TCF:')
worksheet.write(339,0,'TCG:')
worksheet.write(340,0,'TCH:')
worksheet.write(341,0,'TCI:')

worksheet.write(342,0,'TCJ:')
worksheet.write(343,0,'TCK:')
worksheet.write(344,0,'TCL:')
worksheet.write(345,0,'TCM:')
worksheet.write(346,0,'TCN:')
worksheet.write(347,0,'ZZZ:')
worksheet.write(348,0,'/END')

worksheet.write(0,1,faa)
worksheet.write(1,1,fab)
worksheet.write(2,1,fac)
worksheet.write(3,1,fad)
worksheet.write(4,1,fae)
worksheet.write(5,1,faf)
worksheet.write(6,1,fag)
worksheet.write(7,1,fah)
worksheet.write(8,1,fai)
worksheet.write(9,1,fba)
worksheet.write(10,1,fbb)

worksheet.write(11,1,fbc)
worksheet.write(12,1,fbd)
worksheet.write(13,1,fbe)
worksheet.write(14,1,fbf)
worksheet.write(15,1,fbg)
worksheet.write(16,1,fbh)
worksheet.write(17,1,fbi)
worksheet.write(18,1,fca)
worksheet.write(19,1,fcb)
worksheet.write(20,1,fcc)

worksheet.write(21,1,fcd)
worksheet.write(22,1,fce)
worksheet.write(23,1,fcf)
worksheet.write(24,1,fcg)
worksheet.write(25,1,fch)
worksheet.write(26,1,fci)
worksheet.write(27,1,fcj)
worksheet.write(28,1,fck)
worksheet.write(29,1,fcl)
worksheet.write(30,1,fda)

worksheet.write(31,1,fdb)
worksheet.write(32,1,fdc)
worksheet.write(33,1,fdd)
worksheet.write(34,1,fde)
worksheet.write(35,1,fdf)
worksheet.write(36,1,fdg)
worksheet.write(37,1,fdh)
worksheet.write(38,1,fdi)
worksheet.write(39,1,fdj)
worksheet.write(40,1,fdk)

worksheet.write(41,1,fdl)
worksheet.write(42,1,fdm)
worksheet.write(43,1,fdn)
worksheet.write(44,1,fdo)
worksheet.write(45,1,fea)
worksheet.write(46,1,feb)
worksheet.write(47,1,fec)
worksheet.write(48,1,fed)
worksheet.write(49,1,fee)
worksheet.write(50,1,fef)

worksheet.write(51,1,feg)
worksheet.write(52,1,feh)
worksheet.write(53,1,fei)
worksheet.write(54,1,ffa)
worksheet.write(55,1,ffb)
worksheet.write(56,1,ffc)
worksheet.write(57,1,ffd)
worksheet.write(58,1,ffe)
worksheet.write(59,1,fff)
worksheet.write(60,1,ffg)

worksheet.write(61,1,ffh)
worksheet.write(62,1,ffi)
worksheet.write(63,1,fga)
worksheet.write(64,1,fgb)
worksheet.write(65,1,fgc)
worksheet.write(66,1,fgd)
worksheet.write(67,1,fge)
worksheet.write(68,1,fgf)
worksheet.write(69,1,fgg)
worksheet.write(70,1,fgh)
worksheet.write(71,1,fgi)

worksheet.write(72,1,fgj)
worksheet.write(73,1,fgk)
worksheet.write(74,1,fgl)
worksheet.write(75,1,fha)
worksheet.write(76,1,fhb)
worksheet.write(77,1,fhc)
worksheet.write(78,1,fhd)
worksheet.write(79,1,fhe)
worksheet.write(80,1,fhf)
worksheet.write(81,1,fhg)

worksheet.write(82,1,fhh)
worksheet.write(83,1,fhi)
worksheet.write(84,1,fhj)
worksheet.write(85,1,fhk)
worksheet.write(86,1,fhl)
worksheet.write(87,1,fhm)
worksheet.write(88,1,fhn)
worksheet.write(89,1,fho)
worksheet.write(90,1,fia)
worksheet.write(91,1,fib)

worksheet.write(92,1,fic)
worksheet.write(93,1,fid)
worksheet.write(94,1,fie)
worksheet.write(95,1,fif)
worksheet.write(96,1,fig)
worksheet.write(97,1,fih)
worksheet.write(98,1,fii)
worksheet.write(99,1,fja)
worksheet.write(100,1,fjb)
worksheet.write(101,1,fjc)

worksheet.write(102,1,fjd)
worksheet.write(103,1,fje)
worksheet.write(104,1,fjf)
worksheet.write(105,1,fjg)
worksheet.write(106,1,fjh)
worksheet.write(107,1,fji)
worksheet.write(108,1,fka)
worksheet.write(109,1,fkb)
worksheet.write(110,1,fkc)
worksheet.write(111,1,fkd)

worksheet.write(112,1,fke)
worksheet.write(113,1,fkf)
worksheet.write(114,1,fkg)
worksheet.write(115,1,fkh)
worksheet.write(116,1,fki)
worksheet.write(117,1,fkj)
worksheet.write(118,1,fkk)
worksheet.write(119,1,fkl)
worksheet.write(120,1,fla)
worksheet.write(121,1,flb)

worksheet.write(122,1,flc)
worksheet.write(123,1,fld)
worksheet.write(124,1,fle)
worksheet.write(125,1,flf)
worksheet.write(126,1,flg)
worksheet.write(127,1,flh)
worksheet.write(128,1,fli)
worksheet.write(129,1,flj)
worksheet.write(130,1,flk)
worksheet.write(131,1,fll)

worksheet.write(132,1,flm)
worksheet.write(133,1,fln)
worksheet.write(134,1,flo)
worksheet.write(135,1,fma)
worksheet.write(136,1,fmb)
worksheet.write(137,1,fmc)
worksheet.write(138,1,fmd)
worksheet.write(139,1,fme)
worksheet.write(140,1,fmf)
worksheet.write(141,1,fmg)

worksheet.write(142,1,fmh)
worksheet.write(143,1,fmi)
worksheet.write(144,1,fmj)
worksheet.write(145,1,fmk)
worksheet.write(146,1,fml)
worksheet.write(147,1,fmm)
worksheet.write(148,1,fmn)
worksheet.write(149,1,fmo)
worksheet.write(150,1,fmp)
worksheet.write(151,1,fmq)

worksheet.write(152,1,fmr)
worksheet.write(153,1,fms)
worksheet.write(154,1,fmt)
worksheet.write(155,1,fmu)
worksheet.write(156,1,fmv)
worksheet.write(157,1,fmw)
worksheet.write(158,1,fmx)
worksheet.write(159,1,fna)
worksheet.write(160,1,fnb)
worksheet.write(161,1,fnc)

worksheet.write(162,1,fnd)
worksheet.write(163,1,fne)
worksheet.write(164,1,fnf)
worksheet.write(165,1,fng)
worksheet.write(166,1,fnh)
worksheet.write(167,1,foa)
worksheet.write(168,1,fob)
worksheet.write(169,1,foc)
worksheet.write(170,1,fod)
worksheet.write(171,1,maa)

worksheet.write(172,1,mab)
worksheet.write(173,1,mac)
worksheet.write(174,1,mad)
worksheet.write(175,1,mae)
worksheet.write(176,1,maf)
worksheet.write(177,1,mag)
worksheet.write(178,1,mah)
worksheet.write(179,1,mai)
worksheet.write(180,1,maj)
worksheet.write(181,1,mak)

worksheet.write(182,1,mal)
worksheet.write(183,1,mam)
worksheet.write(184,1,man)
worksheet.write(185,1,mba)
worksheet.write(186,1,mbb)
worksheet.write(187,1,mbc)
worksheet.write(188,1,mbd)
worksheet.write(189,1,mbe)
worksheet.write(190,1,mbf)
worksheet.write(191,1,mbg)

worksheet.write(192,1,mbh)
worksheet.write(193,1,mbi)
worksheet.write(194,1,mbj)
worksheet.write(195,1,mbk)
worksheet.write(196,1,mbl)
worksheet.write(197,1,mbm)
worksheet.write(198,1,mbn)
worksheet.write(199,1,mca)
worksheet.write(200,1,mcb)
worksheet.write(201,1,mcc)

worksheet.write(202,1,mcd)
worksheet.write(203,1,mce)
worksheet.write(204,1,mcf)
worksheet.write(205,1,mcg)
worksheet.write(206,1,mch)
worksheet.write(207,1,mci)
worksheet.write(208,1,mcj)
worksheet.write(209,1,mck)
worksheet.write(210,1,mcl)
worksheet.write(211,1,mcm)

worksheet.write(212,1,mcn)
worksheet.write(213,1,saa)
worksheet.write(214,1,sab)
worksheet.write(215,1,sac)
worksheet.write(216,1,sad)
worksheet.write(217,1,sae)
worksheet.write(218,1,saf)
worksheet.write(219,1,sag)
worksheet.write(220,1,sah)
worksheet.write(221,1,sai)

worksheet.write(222,1,saj)
worksheet.write(223,1,sak)
worksheet.write(224,1,sda)
worksheet.write(225,1,sdb)
worksheet.write(226,1,sdc)
worksheet.write(227,1,sdd)
worksheet.write(228,1,sde)
worksheet.write(229,1,sdf)
worksheet.write(230,1,sdg)
worksheet.write(231,1,sdh)

worksheet.write(232,1,sdi)
worksheet.write(233,1,sdj)
worksheet.write(234,1,sdk)
worksheet.write(235,1,sea)
worksheet.write(236,1,seb)
worksheet.write(237,1,sec)
worksheet.write(238,1,sed)
worksheet.write(239,1,see)
worksheet.write(240,1,sef)
worksheet.write(241,1,seg)

worksheet.write(242,1,seh)
worksheet.write(243,1,sei)
worksheet.write(244,1,sej)
worksheet.write(245,1,sek)
worksheet.write(246,1,sha)
worksheet.write(247,1,shb)
worksheet.write(248,1,shc)
worksheet.write(249,1,shd)
worksheet.write(250,1,she)
worksheet.write(251,1,shf)

worksheet.write(252,1,shg)
worksheet.write(253,1,shh)
worksheet.write(254,1,shi)
worksheet.write(255,1,shj)
worksheet.write(256,1,shk)
worksheet.write(257,1,sia)
worksheet.write(258,1,sib)
worksheet.write(259,1,sic)
worksheet.write(260,1,sid)
worksheet.write(261,1,sie)

worksheet.write(262,1,sif)
worksheet.write(263,1,sig)
worksheet.write(264,1,sih)
worksheet.write(265,1,sii)
worksheet.write(266,1,sij)
worksheet.write(267,1,sik)
worksheet.write(268,1,sla)
worksheet.write(269,1,slb)
worksheet.write(270,1,slc)
worksheet.write(271,1,sld)

worksheet.write(272,1,sle)
worksheet.write(273,1,slf)
worksheet.write(274,1,slg)
worksheet.write(275,1,slh)
worksheet.write(276,1,sli)
worksheet.write(277,1,slj)
worksheet.write(278,1,slk)
worksheet.write(279,1,sma)
worksheet.write(280,1,smb)
worksheet.write(281,1,smc)

worksheet.write(282,1,smd)
worksheet.write(283,1,sme)
worksheet.write(284,1,smf)
worksheet.write(285,1,smg)
worksheet.write(286,1,smh)
worksheet.write(287,1,smi)
worksheet.write(288,1,smj)
worksheet.write(289,1,smk)
worksheet.write(290,1,sml)
worksheet.write(291,1,smm)

worksheet.write(292,1,smn)
worksheet.write(293,1,smo)
worksheet.write(294,1,smp)
worksheet.write(295,1,smq)
worksheet.write(296,1,smr)
worksheet.write(297,1,sna)
worksheet.write(298,1,snb)
worksheet.write(299,1,snc)
worksheet.write(300,1,snd)
worksheet.write(301,1,sne)

worksheet.write(302,1,snf)
worksheet.write(303,1,soa)
worksheet.write(304,1,sob)
worksheet.write(305,1,taa)
worksheet.write(306,1,tab)
worksheet.write(307,1,tac)
worksheet.write(308,1,tad)
worksheet.write(309,1,tae)
worksheet.write(310,1,taf)
worksheet.write(311,1,tag)

worksheet.write(312,1,tah)
worksheet.write(313,1,tai)
worksheet.write(314,1,taj)
worksheet.write(315,1,tak)
worksheet.write(316,1,tal)
worksheet.write(317,1,tam)
worksheet.write(318,1,tan)
worksheet.write(319,1,tba)
worksheet.write(320,1,tbb)
worksheet.write(321,1,tbc)

worksheet.write(322,1,tbd)
worksheet.write(323,1,tbe)
worksheet.write(324,1,tbf)
worksheet.write(325,1,tbg)
worksheet.write(326,1,tbh)
worksheet.write(327,1,tbi)
worksheet.write(328,1,tbj)
worksheet.write(329,1,tbk)
worksheet.write(330,1,tbl)
worksheet.write(331,1,tbm)

worksheet.write(332,1,tbn)
worksheet.write(333,1,tca)
worksheet.write(334,1,tcb)
worksheet.write(335,1,tcc)
worksheet.write(336,1,tcd)
worksheet.write(337,1,tce)
worksheet.write(338,1,tcf)
worksheet.write(339,1,tcg)
worksheet.write(340,1,tch)
worksheet.write(341,1,tci)

worksheet.write(342,1,tcj)
worksheet.write(343,1,tck)
worksheet.write(344,1,tcl)
worksheet.write(345,1,tcm)
worksheet.write(346,1,tcn)

workbook.close()  





