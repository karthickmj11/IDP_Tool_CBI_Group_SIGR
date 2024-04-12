import Group_Berth
import Group_Delhipoint
import Group_Point
import Group_Route
import Group_Signal
import Group_Subroute
import Group_Switch
import Group_Trackcircuit
#import Group_Tracksection
print("S")
def run(loc,op_l):
    print('Started')
    Group_Delhipoint.run(loc,op_l)
    Group_Point.run(loc,op_l)
    Group_Route.run(loc,op_l)
    Group_Signal.run(loc,op_l)
    Group_Subroute.run(loc,op_l)
    Group_Switch.run(loc,op_l)
    Group_Trackcircuit.run(loc,op_l)
    print('completed')

