

#FILE_NAME = "TCU026.txt"
FILE_NAME = "TCU026_m.txt"

def read_acc_history():

    f = open(FILE_NAME)
    ACC_HIS = f.readlines()
    
    isnumber = False
    time = []
    up_acc = []
    NS_acc = []
    EW_acc = []
    for line in ACC_HIS:
        line_list = line.split()
        print("show line list: ", line_list) 
        if isnumber  == False:
            for cell in line_list:
                #print("show cell:", cell)
                if cell == "#Data:" and isnumber == False:
                    isnumber = True
                    break
            continue

        if isnumber == True:
            print("length line_list: ", len(line_list))
            time.append(line_list[0])
            up_acc.append(line_list[1])
            NS_acc.append(line_list[2])
            EW_acc.append(line_list[3])

    print("time: ", time)
    print("up: ", up_acc)
    print("NS: ", NS_acc)
    print("EW: ", EW_acc)


if __name__ == "__main__":
    read_acc_history()
