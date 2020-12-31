
FILE_NAME = "TAP024.txt"
#FILE_NAME = "TCU081.txt"
#FILE_NAME = "CHY070.txt"
#FILE_NAME = "TAP098.txt"
#FILE_NAME = "TCU026.txt"
#FILE_NAME = "TCU049.txt"
#FILE_NAME = "TCU026_m.txt"

def read_acc_history():

    f = open(FILE_NAME)
    ACC_HIS = f.readlines()
    
    isSearched = False
    time = []
    up_acc = []
    NS_acc = []
    EW_acc = []
    for line in ACC_HIS:
        line_list = line.split()
        #print("show line list: ", line_list) 
        if isSearched  == False:
            for cell in line_list:
                #print("show cell:", cell)
                if cell == "#Data:" and isSearched == False:
                    isSearched = True
                    break
            continue

        if isSearched == True:
            #print("length line_list: ", len(line_list))
            time.append(line_list[0])
            up_acc.append(line_list[1])
            NS_acc.append(line_list[2])
            EW_acc.append(line_list[3])
    """
    print("time: ", time)
    print("up: ", up_acc)
    print("NS: ", NS_acc)
    print("EW: ", EW_acc)
    """

    return time, up_acc, NS_acc, EW_acc

if __name__ == "__main__":
    read_acc_history()
