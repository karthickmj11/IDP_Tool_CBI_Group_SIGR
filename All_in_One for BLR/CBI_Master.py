import openpyxl
import CBI_Berth
import CBI_Cycle
import CBI_Delhipoint
import CBI_InternalVariables
import CBI_Point
import CBI_Pointend
import CBI_Route
import CBI_Signal
import CBI_Subroute
import CBI_Trackportion
import CBI_Traffic_Direction
import CBI_Tracksection
import CBI_ASCV
import CBI_ESP
import CBI_MaintenanceBlock
import CBI_PSAPR
def run(loc,op_l):
    CBI_ASCV.run(loc,op_l)
    CBI_Berth.run(loc,op_l)
    CBI_Cycle.run(loc,op_l)
    CBI_ESP.run(loc,op_l)
    CBI_MaintenanceBlock.run(loc,op_l)
    CBI_Delhipoint.run(loc,op_l)
    CBI_InternalVariables.run(loc,op_l)
    CBI_Point.run(loc,op_l)
    CBI_Pointend.run(loc,op_l)
    CBI_PSAPR.run(loc,op_l)
    CBI_Route.run(loc,op_l)
    CBI_Signal.run(loc,op_l)
    CBI_Subroute.run(loc,op_l)
    CBI_Trackportion.run(loc,op_l)
    CBI_Traffic_Direction.run(loc,op_l)
    CBI_Tracksection.run(loc,op_l)


