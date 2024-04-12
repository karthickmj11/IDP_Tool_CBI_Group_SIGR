import Group_ASCV
import Group_Berth
import Group_Cycle
import Group_ESP
import Group_MaintenanceBlock
import Group_PSAPR
import Group_Delhipoint
import Group_Point
import Group_Route
import Group_Signal
import Group_Subroute
import Group_Switch
import Group_Trackcircuit
import Group_Tracksection
import Group_Junction
import Group_Traffic_Direction
def run(loc,op_l):
    Group_ASCV.run(loc,op_l)
    Group_Berth.run(loc,op_l)
    Group_Cycle.run(loc,op_l)
    Group_ESP.run(loc,op_l)
    Group_MaintenanceBlock.run(loc,op_l)
    Group_PSAPR.run(loc,op_l)
    #Group_Berth.run(loc,op_l)
    Group_Delhipoint.run(loc,op_l)
    Group_Point.run(loc,op_l)
    Group_Route.run(loc,op_l)
    Group_Signal.run(loc,op_l)
    Group_Subroute.run(loc,op_l)
    Group_Switch.run(loc,op_l)
    Group_Trackcircuit.run(loc,op_l)
    Group_Tracksection.run(loc,op_l)
    Group_Junction.run(loc,op_l)
    Group_Traffic_Direction.run(loc,op_l)

