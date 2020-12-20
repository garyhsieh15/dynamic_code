

FILE_NAME = "TCU026.txt"

def read_acc_history():

    f = open(FILE_NAME)
    ACC_HIS = f.readlines()
    
    isdigit = True
    for line in ACC_HIS:
        line_list = line.split()
        for cell in line_list:
            print("show cell:", cell)
            cell = cell.replace('.', '', 1)
            cell = cell.replace('-', '', 1)
            if not cell.isdigit():
                isdigit = False

        if isdigit == True:
            print(line_list)




if __name__ == "__main__":
    read_acc_history()
