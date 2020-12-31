# --------------------------------------------------------------------------
# name       : preprocess_data.py
# description: filter drity data and fill in the blanks, if displacement or
#              acc is empty then fill in the interpolation value.
#
# date:      : 20201224
# author     : garyhsieh
# --------------------------------------------------------------------------


#from openpyxl import Workbook
from scipy.stats import mode
from openpyxl.reader.excel import load_workbook
import pandas as pd
import numpy as np



# define sheet index addr
SHT0_INDEX = 0
SHT1_INDEX = 1
SHT2_INDEX = 2
SHT3_INDEX = 3
SHT4_INDEX = 4
SHT5_INDEX = 5

SHT0_NEW = "free"
SHT1_NEW = "free_b"
SHT2_NEW = "sin"
SHT3_NEW = "El"

# define column name
SHT0_TIME_KEY = "Time (s) Auto"
SHT0_DIS1_KEY = "Position (mm) Monitor Run"
SHT0_DIS2_KEY = "Position (mm) Monitor Run.1"
SHT0_DIS3_KEY = "Position (mm) Monitor Run.2"
SHT0_ACC_KEY  = "Acceleration - y (cm/s²) Monitor Run"

SHT1_TIME_KEY = "Time (s) Auto"
SHT1_DIS1_KEY = "Position (mm) Monitor Run"
SHT1_DIS2_KEY = "Position (mm) Monitor Run.1"
SHT1_DIS3_KEY = "Position (mm) Monitor Run.2"
SHT1_ACC_KEY  = "Acceleration - y (cm/s²) Monitor Run"

SHT2_TIME_KEY = "Time (s) Auto"
SHT2_DIS1_KEY = "Position (mm) Monitor Run"
SHT2_DIS2_KEY = "Position (mm) Monitor Run.1"
SHT2_DIS3_KEY = "Position (mm) Monitor Run.2"
SHT2_ACC_KEY  = "Acceleration - y (cm/s²) Monitor Run"

SHT3_TIME_KEY = "Time (s) Auto"
SHT3_DIS1_KEY = "Position (mm) Monitor Run"
SHT3_DIS2_KEY = "Position (mm) Monitor Run.1"
SHT3_DIS3_KEY = "Position (mm) Monitor Run.2"
SHT3_ACC_KEY  = "Acceleration - y (cm/s²) Monitor Run"


# define threshold value
THRESHOLD_ACC = 60
START_ACC = 2
AVGE_NAN  = 4

def read_vtable_acc():

    # read data from excel file.
    df_sht0 = pd.read_excel("/Users/garyhsieh/Desktop/20201219-1_m.xlsx", sheetname = SHT0_INDEX)
    df_sht1 = pd.read_excel("/Users/garyhsieh/Desktop/20201219-1_m.xlsx", sheetname = SHT1_INDEX)
    df_sht2 = pd.read_excel("/Users/garyhsieh/Desktop/20201219-1_m.xlsx", sheetname = SHT2_INDEX)
    df_sht3 = pd.read_excel("/Users/garyhsieh/Desktop/20201219-1_m.xlsx", sheetname = SHT3_INDEX)
    
    # --------------------------------------------------------------------------------------------
    # deal with sheet 00
    #
    # --------------------------------------------------------------------------------------------
    temp_time = []; time_sht0 = []
    temp_dis1 = []; dis1_sht0 = []
    temp_dis2 = []; dis2_sht0 = []
    temp_dis3 = []; dis3_sht0 = []
    temp_acc  = []; acc_sht0  = []

    temp_time = df_sht0.iloc[:, 0:1]
    temp_dis1 = df_sht0.iloc[:, 1:2]
    temp_dis2 = df_sht0.iloc[:, 2:3]
    temp_dis3 = df_sht0.iloc[:, 3:4]
    temp_acc  = df_sht0.iloc[:, 4:5]


    time_sht0.extend(temp_time[SHT0_TIME_KEY])
    dis1_sht0.extend(temp_dis1[SHT0_DIS1_KEY])
    dis2_sht0.extend(temp_dis2[SHT0_DIS2_KEY])
    dis3_sht0.extend(temp_dis3[SHT0_DIS3_KEY])
    acc_sht0.extend(temp_acc[SHT0_ACC_KEY])
    
    #get the point that acc > 60.
    isSearched_thd = 0
    for i in range(len(acc_sht0)):
        if abs(acc_sht0[i]) >= THRESHOLD_ACC:
            isSearched_thd = i
            break

    # find the first value of acc
    isSearched_start = 0
    for j in range(isSearched_thd, 0, -1):
        if abs(acc_sht0[j]) < START_ACC:
            isSearched_start = j
            break

    # find the last value of acc
    isSearched_end = 0
    for l in range(isSearched_thd, len(acc_sht0)):
        if np.isnan(acc_sht0[l]):
            if np.isnan(acc_sht0[l + 1]):
                #isSearched_end = l - 1
                isSearched_end = l
                break
    
    
    # catch avge to get number of jump nan.
    num_nan = 0
    num_nan_list = []
    for dis1_idx in range(0, isSearched_end):
        if np.isnan(dis1_sht0[dis1_idx]):
            num_nan += 1
        else:
            if num_nan != 0:
                num_nan_list.append(num_nan)
                num_nan = 0


        if len(num_nan_list) >= AVGE_NAN:
            break
    jump_nan = mode(num_nan_list).mode[0]

    # fill in nan to avoid dirty data for dis1.
    isSearched_value = 0
    jump_count = 0
    for dis1_idx in range(0, isSearched_end):
        if not np.isnan(dis1_sht0[dis1_idx]):
            isSearched_value = 1
            continue

        if isSearched_value == 1:
            jump_count += 1
            dis1_sht0[dis1_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0

    # fill in nan to avoid dirty data for dis2
    isSearched_value = 0
    jump_count = 0
    for dis2_idx in range(0, isSearched_end):
        if not np.isnan(dis2_sht0[dis2_idx]):
            isSearched_value = 1
            continue
        
        if isSearched_value == 1:
            jump_count += 1
            dis2_sht0[dis2_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0

    
    # fill in nan to avoid dirty data for dis3
    isSearched_value = 0
    jump_count = 0
    for dis3_idx in range(0, isSearched_end):
        if not np.isnan(dis3_sht0[dis3_idx]):
            isSearched_value = 1
            continue
        
        if isSearched_value == 1:
            jump_count += 1
            dis3_sht0[dis3_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0


    #print(dis1_sht0)
    #print(dis2_sht0)
    #print(dis3_sht0)

    """
    print("nan list len: ", len(num_nan_list))
    print("nan list:", num_nan_list)
    print("show most num:", jump_nan)
    """

    # after filter
    ftime_sht0 = []
    fdis1_sht0 = []
    fdis2_sht0 = []
    fdis3_sht0 = []
    facc_sht0  = []

    for m in range(isSearched_start, isSearched_end):
        ftime_sht0.append(time_sht0[m])
        fdis1_sht0.append(dis1_sht0[m])
        fdis2_sht0.append(dis2_sht0[m])
        fdis3_sht0.append(dis3_sht0[m])
        facc_sht0.append(acc_sht0[m])


    # fill dis1's empty in 0 for free vibration
    for free_idx in range(0, len(fdis1_sht0)):
        if np.isnan(fdis1_sht0[free_idx]):
            fdis1_sht0[free_idx] = 0

    """ 
    # fill dis2's empty in interpolation for free vibration
    svalue = None
    evalue = None
    num = 0
    assign_value = 0
    for dis2_idx in range(0, len(fdis2_sht0)):
        if not np.isnan(fdis2_sht0[dis2_idx]) and num == 0 and svalue == None:
            print("enter 1 -------")
            svalue = fdis2_sht0[dis2_idx]
        if not np.isnan(fdis2_sht0[dis2_idx]) and num != 0 and svalue == None:
            print("enter 2 -------")
            for assign_idx in range(0, num):
                fdis2_sht0[assign_idx - assign_idx -1] = fdis2_sht0[assign_idx]
            num = 0
            svalue = evalue
            evalue = None

        if np.isnan(fdis2_sht0[dis2_idx]):
            print("enter 3 -------")
            num +=1

        if not np.isnan(fdis2_sht0[dis2_idx]) and svalue != None:
            print("enter 4 -------")
            evalue = fis2_sht0[dis2_idx]
            assign_value = (evalue - svalue) / jump_nan
            for assign_idx in range(0, jump_nan):
                fdis2_sht0[dis2_idx - assign_idx + 1] = fdis2_sht0[dis2_idxi - assign_idx] - assign_value
            num = 0
            svalue = evalue 
            evalue = None
            
    print(fdis2_sht0)
    """

    temp_acc = 0
    avge_acc = 0
    num = 1
    # fill acc's empty into value
    for n in range(0, len(facc_sht0)):
        if np.isnan(facc_sht0[n]):
            #print("%12s %12.4f" %("n", n))
            num  += 1
            continue
        else:
            temp_acc += facc_sht0[n]
            #print("%12s %12s %12s %12s" % ("n", "num","temp_acc", "facc_sht0"))
            #print("%12.4f %12.4f %12.4f %12.4f" % (n, num,temp_acc, facc_sht0[n]))
            if num >= 2:
                avge_acc = temp_acc / num
                facc_sht0[n - (num - 1)] = avge_acc
                temp_acc = facc_sht0[n]
                num = 1

    

    """    
    for z in facc_sht0:
        print(z)
    """

    print("isSearched start and time: ", isSearched_start, time_sht0[isSearched_start])
    print("isSearched thd and time: ", isSearched_thd, time_sht0[isSearched_thd])
    print("isSearched end and time; ", isSearched_end, time_sht0[isSearched_end])
    print("-----------------------------------------")
    #print(facc_sht0)

    # --------------------------------------------------------------------------------------------
    # deal with sheet 01
    # step1: read data from excel
    # step2: get the point that acc > 60
    # step3: catch avge to get number of jump nan.
    # step4: fill in nan to avoid dirty data for dis1, dis2, dis3
    # step5: create new list to put data
    # step6: fill in dis1 for initial value
    # --------------------------------------------------------------------------------------------
    temp_time = []; time_sht1 = []
    temp_dis1 = []; dis1_sht1 = []
    temp_dis2 = []; dis2_sht1 = []
    temp_dis3 = []; dis3_sht1 = []
    temp_acc  = []; acc_sht1  = []


    temp_time = df_sht1.iloc[:, 0:1]
    temp_dis1 = df_sht1.iloc[:, 1:2]
    temp_dis2 = df_sht1.iloc[:, 2:3]
    temp_dis3 = df_sht1.iloc[:, 3:4]
    temp_acc  = df_sht1.iloc[:, 4:5]
    
    time_sht1.extend(temp_time[SHT1_TIME_KEY])
    dis1_sht1.extend(temp_dis1[SHT1_DIS1_KEY])
    dis2_sht1.extend(temp_dis2[SHT1_DIS2_KEY])
    dis3_sht1.extend(temp_dis3[SHT1_DIS3_KEY])
    acc_sht1.extend(temp_acc[SHT1_ACC_KEY])

    #print(time_sht1)

    #get the point that acc > 60.
    isSearched_thd = 0
    for i in range(0, len(acc_sht1)):
        if abs(acc_sht1[i]) >= THRESHOLD_ACC:
            isSearched_thd = i
            break

    # find the first value of acc
    isSearched_start = 0
    for j in range(isSearched_thd, 0, -1):
        if abs(acc_sht1[j]) < START_ACC:
            isSearched_start = j
            break

    # find the last value of acc
    isSearched_end = 0
    for l in range(isSearched_thd, len(acc_sht1)):
        if np.isnan(acc_sht1[l]):
            if np.isnan(acc_sht1[l + 1]):
                #isSearched_end = l - 1
                isSearched_end = l
                break
            
    
    print("isSearched start and time: ", isSearched_start, time_sht1[isSearched_start])
    print("isSearched thd and time: ", isSearched_thd, time_sht1[isSearched_thd])
    print("isSearched end and time; ", isSearched_end, time_sht1[isSearched_end])
    print("-----------------------------------------")
     
    # catch avge to get number of jump nan.
    num_nan = 0
    num_nan_list = []
    for dis1_idx in range(0, isSearched_end):
        if np.isnan(dis1_sht1[dis1_idx]):
            num_nan += 1
        else:
            if num_nan != 0:
                num_nan_list.append(num_nan)
                num_nan = 0


        if len(num_nan_list) >= AVGE_NAN:
            break
    jump_nan = mode(num_nan_list).mode[0]

    #print("number nump nan: ", jump_nan)

    # fill in nan to avoid dirty data for dis1.
    isSearched_value = 0
    jump_count = 0
    for dis1_idx in range(0, isSearched_end):
        if not np.isnan(dis1_sht1[dis1_idx]):
            isSearched_value = 1
            continue

        if isSearched_value == 1:
            jump_count += 1
            dis1_sht1[dis1_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0

    # fill in nan to avoid dirty data for dis2
    isSearched_value = 0
    jump_count = 0
    for dis2_idx in range(0, isSearched_end):
        if not np.isnan(dis2_sht1[dis2_idx]):
            isSearched_value = 1
            continue
        
        if isSearched_value == 1:
            jump_count += 1
            dis2_sht1[dis2_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0

    
    # fill in nan to avoid dirty data for dis3
    isSearched_value = 0
    jump_count = 0
    for dis3_idx in range(0, isSearched_end):
        if not np.isnan(dis3_sht1[dis3_idx]):
            isSearched_value = 1
            continue
        
        if isSearched_value == 1:
            jump_count += 1
            dis3_sht1[dis3_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0


    #print(dis1_sht1)
    #print(dis2_sht1)
    #print(dis3_sht1)

    # after filter
    ftime_sht1 = []
    fdis1_sht1 = []
    fdis2_sht1 = []
    fdis3_sht1 = []
    facc_sht1  = []

    for m in range(isSearched_start, isSearched_end):
        ftime_sht1.append(time_sht1[m])
        fdis1_sht1.append(dis1_sht1[m])
        fdis2_sht1.append(dis2_sht1[m])
        fdis3_sht1.append(dis3_sht1[m])
        facc_sht1.append(acc_sht1[m])


    # fill dis1's empty in 0 for free vibration
    """
    for free_idx in range(0, len(fdis1_sht0)):
        if np.isnan(fdis1_sht0[free_idx]):
            fdis1_sht0[free_idx] = 0
    """
    temp_acc = 0
    avge_acc = 0
    num = 1
    # fill acc's empty into value
    for n in range(0, len(facc_sht1)):
        if np.isnan(facc_sht1[n]):
            #print("%12s %12.4f" %("n", n))
            num  += 1
            continue
        else:
            temp_acc += facc_sht1[n]
            #print("%12s %12s %12s %12s" % ("n", "num","temp_acc", "facc_sht0"))
            #print("%12.4f %12.4f %12.4f %12.4f" % (n, num,temp_acc, facc_sht0[n]))
            if num >= 2:
                avge_acc = temp_acc / num
                facc_sht1[n - (num - 1)] = avge_acc
                temp_acc = facc_sht1[n]
                num = 1


    # --------------------------------------------------------------------------------------------
    # deal with sheet 02
    # step1: read data from excel
    # step2: get the point that acc > 60
    # step3: catch avge to get number of jump nan.
    # step4: fill in nan to avoid dirty data for dis1, dis2, dis3
    # step5: create new list to put data
    # step6: fill in dis1 for initial value
    # --------------------------------------------------------------------------------------------
    temp_time = []; time_sht2 = []
    temp_dis1 = []; dis1_sht2 = []
    temp_dis2 = []; dis2_sht2 = []
    temp_dis3 = []; dis3_sht2 = []
    temp_acc  = []; acc_sht2  = []


    temp_time = df_sht2.iloc[:, 0:1]
    temp_dis1 = df_sht2.iloc[:, 1:2]
    temp_dis2 = df_sht2.iloc[:, 2:3]
    temp_dis3 = df_sht2.iloc[:, 3:4]
    temp_acc  = df_sht2.iloc[:, 4:5]
    
    time_sht2.extend(temp_time[SHT2_TIME_KEY])
    dis1_sht2.extend(temp_dis1[SHT2_DIS1_KEY])
    dis2_sht2.extend(temp_dis2[SHT2_DIS2_KEY])
    dis3_sht2.extend(temp_dis3[SHT2_DIS3_KEY])
    acc_sht2.extend(temp_acc[SHT2_ACC_KEY])

    #print(time_sht1)

    #get the point that acc > 60.
    isSearched_thd = 0
    for i in range(0, len(acc_sht2)):
        if abs(acc_sht2[i]) >= THRESHOLD_ACC:
            isSearched_thd = i
            break

    # find the first value of acc
    isSearched_start = 0
    for j in range(isSearched_thd, 0, -1):
        if abs(acc_sht2[j]) < START_ACC:
            isSearched_start = j
            break

    # find the last value of acc
    isSearched_end = 0
    for l in range(isSearched_thd, len(acc_sht2)):
        if np.isnan(acc_sht2[l]):
            if np.isnan(acc_sht2[l + 1]):
                #isSearched_end = l - 1
                isSearched_end = l
                break
            
    
    print("isSearched start and time: ", isSearched_start, time_sht2[isSearched_start])
    print("isSearched thd and time: ", isSearched_thd, time_sht2[isSearched_thd])
    print("isSearched end and time; ", isSearched_end, time_sht2[isSearched_end])
    print("-----------------------------------------")
     
    # catch avge to get number of jump nan.
    num_nan = 0
    num_nan_list = []
    for dis1_idx in range(0, isSearched_end):
        if np.isnan(dis1_sht2[dis1_idx]):
            num_nan += 1
        else:
            if num_nan != 0:
                num_nan_list.append(num_nan)
                num_nan = 0


        if len(num_nan_list) >= AVGE_NAN:
            break
    jump_nan = mode(num_nan_list).mode[0]

    #print("number nump nan: ", jump_nan)

    # fill in nan to avoid dirty data for dis1.
    isSearched_value = 0
    jump_count = 0
    for dis1_idx in range(0, isSearched_end):
        if not np.isnan(dis1_sht2[dis1_idx]):
            isSearched_value = 1
            continue

        if isSearched_value == 1:
            jump_count += 1
            dis1_sht2[dis1_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0

    # fill in nan to avoid dirty data for dis2
    isSearched_value = 0
    jump_count = 0
    for dis2_idx in range(0, isSearched_end):
        if not np.isnan(dis2_sht2[dis2_idx]):
            isSearched_value = 1
            continue
        
        if isSearched_value == 1:
            jump_count += 1
            dis2_sht2[dis2_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0

    
    # fill in nan to avoid dirty data for dis3
    isSearched_value = 0
    jump_count = 0
    for dis3_idx in range(0, isSearched_end):
        if not np.isnan(dis3_sht2[dis3_idx]):
            isSearched_value = 1
            continue
        
        if isSearched_value == 1:
            jump_count += 1
            dis3_sht2[dis3_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0


    #print(dis1_sht2)
    #print(dis2_sht2)
    #print(dis3_sht2)

    # after filter
    ftime_sht2 = []
    fdis1_sht2 = []
    fdis2_sht2 = []
    fdis3_sht2 = []
    facc_sht2  = []

    for m in range(isSearched_start, isSearched_end):
        ftime_sht2.append(time_sht2[m])
        fdis1_sht2.append(dis1_sht2[m])
        fdis2_sht2.append(dis2_sht2[m])
        fdis3_sht2.append(dis3_sht2[m])
        facc_sht2.append(acc_sht2[m])


    # fill dis1's empty in 0 for free vibration
    """
    for free_idx in range(0, len(fdis1_sht0)):
        if np.isnan(fdis1_sht0[free_idx]):
            fdis1_sht0[free_idx] = 0
    """
    temp_acc = 0
    avge_acc = 0
    num = 1
    # fill acc's empty into value
    for n in range(0, len(facc_sht2)):
        if np.isnan(facc_sht2[n]):
            #print("%12s %12.4f" %("n", n))
            num  += 1
            continue
        else:
            temp_acc += facc_sht2[n]
            #print("%12s %12s %12s %12s" % ("n", "num","temp_acc", "facc_sht0"))
            #print("%12.4f %12.4f %12.4f %12.4f" % (n, num,temp_acc, facc_sht0[n]))
            if num >= 2:
                avge_acc = temp_acc / num
                facc_sht2[n - (num - 1)] = avge_acc
                temp_acc = facc_sht2[n]
                num = 1


    # --------------------------------------------------------------------------------------------
    # deal with sheet 03
    # step1: read data from excel
    # step2: get the point that acc > 60
    # step3: catch avge to get number of jump nan.
    # step4: fill in nan to avoid dirty data for dis1, dis2, dis3
    # step5: create new list to put data
    # step6: fill in dis1 for initial value
    # --------------------------------------------------------------------------------------------
    temp_time = []; time_sht3 = []
    temp_dis1 = []; dis1_sht3 = []
    temp_dis2 = []; dis2_sht3 = []
    temp_dis3 = []; dis3_sht3 = []
    temp_acc  = []; acc_sht3  = []


    temp_time = df_sht3.iloc[:, 0:1]
    temp_dis1 = df_sht3.iloc[:, 1:2]
    temp_dis2 = df_sht3.iloc[:, 2:3]
    temp_dis3 = df_sht3.iloc[:, 3:4]
    temp_acc  = df_sht3.iloc[:, 4:5]
    
    time_sht3.extend(temp_time[SHT3_TIME_KEY])
    dis1_sht3.extend(temp_dis1[SHT3_DIS1_KEY])
    dis2_sht3.extend(temp_dis2[SHT3_DIS2_KEY])
    dis3_sht3.extend(temp_dis3[SHT3_DIS3_KEY])
    acc_sht3.extend(temp_acc[SHT3_ACC_KEY])

    #print(time_sht3)

    #get the point that acc > 60.
    isSearched_thd = 0
    for i in range(0, len(acc_sht3)):
        if abs(acc_sht3[i]) >= THRESHOLD_ACC:
            isSearched_thd = i
            break

    # find the first value of acc
    isSearched_start = 0
    for j in range(isSearched_thd, 0, -1):
        if abs(acc_sht3[j]) < START_ACC:
            isSearched_start = j
            break

    # find the last value of acc
    isSearched_end = 0
    for l in range(isSearched_thd, len(acc_sht3)):
        if np.isnan(acc_sht3[l]):
            if np.isnan(acc_sht3[l + 1]):
                #isSearched_end = l - 1
                isSearched_end = l
                break
            
    
    print("isSearched start and time: ", isSearched_start, time_sht3[isSearched_start])
    print("isSearched thd and time: ", isSearched_thd, time_sht3[isSearched_thd])
    print("isSearched end and time; ", isSearched_end, time_sht3[isSearched_end])
    print("-----------------------------------------")
     
    # catch avge to get number of jump nan.
    num_nan = 0
    num_nan_list = []
    for dis1_idx in range(0, isSearched_end):
        if np.isnan(dis1_sht3[dis1_idx]):
            num_nan += 1
        else:
            if num_nan != 0:
                num_nan_list.append(num_nan)
                num_nan = 0


        if len(num_nan_list) >= AVGE_NAN:
            break
    jump_nan = mode(num_nan_list).mode[0]

    #print("number nump nan: ", jump_nan)

    # fill in nan to avoid dirty data for dis1.
    isSearched_value = 0
    jump_count = 0
    for dis1_idx in range(0, isSearched_end):
        if not np.isnan(dis1_sht3[dis1_idx]):
            isSearched_value = 1
            continue

        if isSearched_value == 1:
            jump_count += 1
            dis1_sht3[dis1_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0

    # fill in nan to avoid dirty data for dis2
    isSearched_value = 0
    jump_count = 0
    for dis2_idx in range(0, isSearched_end):
        if not np.isnan(dis2_sht3[dis2_idx]):
            isSearched_value = 1
            continue
        
        if isSearched_value == 1:
            jump_count += 1
            dis2_sht3[dis2_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0

    
    # fill in nan to avoid dirty data for dis3
    isSearched_value = 0
    jump_count = 0
    for dis3_idx in range(0, isSearched_end):
        if not np.isnan(dis3_sht3[dis3_idx]):
            isSearched_value = 1
            continue
        
        if isSearched_value == 1:
            jump_count += 1
            dis3_sht3[dis3_idx] = np.nan
            if jump_count == jump_nan:
                isSearched_value = 0


    #print(dis1_sht3)
    #print(dis2_sht3)
    #print(dis3_sht3)

    # after filter
    ftime_sht3 = []
    fdis1_sht3 = []
    fdis2_sht3 = []
    fdis3_sht3 = []
    facc_sht3  = []

    for m in range(isSearched_start, isSearched_end):
        ftime_sht3.append(time_sht3[m])
        fdis1_sht3.append(dis1_sht3[m])
        fdis2_sht3.append(dis2_sht3[m])
        fdis3_sht3.append(dis3_sht3[m])
        facc_sht3.append(acc_sht3[m])


    # fill dis1's empty in 0 for free vibration
    """
    for free_idx in range(0, len(fdis1_sht0)):
        if np.isnan(fdis1_sht0[free_idx]):
            fdis1_sht0[free_idx] = 0
    """
    temp_acc = 0
    avge_acc = 0
    num = 1
    # fill acc's empty into value
    for n in range(0, len(facc_sht3)):
        if np.isnan(facc_sht3[n]):
            #print("%12s %12.4f" %("n", n))
            num  += 1
            continue
        else:
            temp_acc += facc_sht3[n]
            #print("%12s %12s %12s %12s" % ("n", "num","temp_acc", "facc_sht0"))
            #print("%12.4f %12.4f %12.4f %12.4f" % (n, num,temp_acc, facc_sht0[n]))
            if num >= 2:
                avge_acc = temp_acc / num
                facc_sht3[n - (num - 1)] = avge_acc
                temp_acc = facc_sht3[n]
                num = 1




    # --------------------------------------------------------------------------------------------
    # write data into excel file
    # sheet00, sheet01, sheet02, sheet03, sheet04
    #
    # --------------------------------------------------------------------------------------------
    # new dataframe for sheet0
    data_new_sht0 = {}
    data_new_sht0[SHT0_TIME_KEY] = ftime_sht0 
    data_new_sht0[SHT0_DIS1_KEY] = fdis1_sht0
    data_new_sht0[SHT0_DIS2_KEY] = fdis2_sht0
    data_new_sht0[SHT0_DIS3_KEY] = fdis3_sht0
    data_new_sht0[SHT0_ACC_KEY] = facc_sht0
    
    df_new_sht0 = pd.DataFrame(data_new_sht0)
    #print(df_new_sht0)

    # new dataframe for sheet01
    data_new_sht1 = {}
    data_new_sht1[SHT1_TIME_KEY] = ftime_sht1 
    data_new_sht1[SHT1_DIS1_KEY] = fdis1_sht1
    data_new_sht1[SHT1_DIS2_KEY] = fdis2_sht1
    data_new_sht1[SHT1_DIS3_KEY] = fdis3_sht1
    data_new_sht1[SHT1_ACC_KEY] = facc_sht1
 
    df_new_sht1 = pd.DataFrame(data_new_sht1)
    
    # new dataframe for sheet02
    data_new_sht2 = {}
    data_new_sht2[SHT2_TIME_KEY] = ftime_sht2 
    data_new_sht2[SHT2_DIS1_KEY] = fdis1_sht2
    data_new_sht2[SHT2_DIS2_KEY] = fdis2_sht2
    data_new_sht2[SHT2_DIS3_KEY] = fdis3_sht2
    data_new_sht2[SHT2_ACC_KEY] = facc_sht2
 
    df_new_sht2 = pd.DataFrame(data_new_sht2)

    # new dataframe for sheet03
    data_new_sht3 = {}
    data_new_sht3[SHT3_TIME_KEY] = ftime_sht3 
    data_new_sht3[SHT3_DIS1_KEY] = fdis1_sht3
    data_new_sht3[SHT3_DIS2_KEY] = fdis2_sht3
    data_new_sht3[SHT3_DIS3_KEY] = fdis3_sht3
    data_new_sht3[SHT3_ACC_KEY] = facc_sht3
 
    df_new_sht3 = pd.DataFrame(data_new_sht3)


    # save dataframe to excel file
    writer = pd.ExcelWriter("/Users/garyhsieh/Desktop/20201219-1_NEW.xlsx", engine = "xlsxwriter")

    # sheet00 -> free
    df_new_sht0.to_excel(writer, sheet_name = SHT0_NEW, index = False, \
            columns = [SHT0_TIME_KEY, SHT0_DIS1_KEY, SHT0_DIS2_KEY, SHT0_DIS3_KEY, SHT0_ACC_KEY])
    # sheet01 -> free plus brace
    df_new_sht1.to_excel(writer, sheet_name = SHT1_NEW, index = False, \
            columns = [SHT1_TIME_KEY, SHT1_DIS1_KEY, SHT1_DIS2_KEY, SHT1_DIS3_KEY, SHT1_ACC_KEY])
    # sheet02 -> sin curve
    df_new_sht2.to_excel(writer, sheet_name = SHT2_NEW, index = False, \
            columns = [SHT2_TIME_KEY, SHT2_DIS1_KEY, SHT2_DIS2_KEY, SHT2_DIS3_KEY, SHT2_ACC_KEY])
    # sheet03 -> sin curve
    df_new_sht3.to_excel(writer, sheet_name = SHT3_NEW, index = False, \
            columns = [SHT3_TIME_KEY, SHT3_DIS1_KEY, SHT3_DIS2_KEY, SHT3_DIS3_KEY, SHT3_ACC_KEY])


    writer.save()

if __name__ == "__main__":
    print("preprocessing data ...")

    read_vtable_acc()
