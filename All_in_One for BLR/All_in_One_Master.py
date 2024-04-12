import Group_Master
import CBI_Master
def run(loc,op_l,l):
    if(l[0]==1):
        print("Group files are being excecuted")
        Group_Master.run(loc,op_l)
    if(l[1]==1):
        print("CBI files are being excecuted")
        CBI_Master.run(loc,op_l)
    return "Completed"