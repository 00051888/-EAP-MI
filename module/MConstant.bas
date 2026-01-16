Attribute VB_Name = "MConstant"
Option Explicit


'equipment type
Public Const PROCESS As String = "PROCESS"
Public Const METROLOGY As String = "METROLOGY"

'equipment run mode
Public Const EQP_RUNMODE_AUTO As String = "AUTO"
Public Const EQP_RUNMODE_HOSTONLINE As String = "HOST_ONLINE"
Public Const EQP_RUNMODE_OFFLINE As String = "OFFLINE"

'equipment modes
Public Const OFFLINE As String = "OFFLINE"
Public Const ONLINE_LOCAL As String = "ONLINE_LOCAL"
Public Const ONLINE_REMOTE As String = "ONLINE_REMOTE"
Public Const READY As String = "READY"

'Secs mode
Public Const SECS_DISABLE As String = "DISABLE"
Public Const SECS_ENABLE_COMM As String = "ENABLE_COMM"
Public Const SECS_ENABLE_NOCOMM As String = "ENABLE_NOCOMM"

'equipment run states
Public Const ESTATE_DOWN As String = "ESTATE_DOWN"
Public Const ESTATE_IDLE As String = "ESTATE_IDLE"
Public Const ESTATE_STOP As String = "ESTATE_STOP"
Public Const ESTATE_PROCESSING As String = "ESTATE_PROCESSING"
Public Const ESTATE_COMDISABLE As String = "ESTATE_COMDISABLE"

'port types
Public Const PORT_INPUT As String = "I"
Public Const PORT_OUTPUT As String = "O"
Public Const PORT_INOUT As String = "B"
Public Const PORT_DOWN As String = "N"
Public Const PORT_UNKNOWN As String = "?"

'port states
Public Const PORT_EMPTY As String = "EMPTY"
Public Const PORT_PRESENT As String = "PRESENT"
Public Const PORT_ACTIVE As String = "ACTIVE"

'chamber states
Public Const CHB_DOWN As String = "DOWN"
Public Const CHB_IDLE As String = "IDLE"
Public Const CHB_PROCESSING As String = "PROCESSING"
Public Const CHB_STOP As String = "STOP"

'lot states
Public Const LOT_IDLE As String = "IDLE"
Public Const LOT_MISPPID As String = "MISPPID"
Public Const LOT_WAITING As String = "WAITING"
Public Const LOT_ARRIVED As String = "ARRIVED"
Public Const LOT_RDYLOAD As String = "RDYLOAD"
Public Const LOT_LOADING As String = "LOADING"
Public Const LOT_LOADED As String = "LOADED"
Public Const LOT_MOVEIN As String = "MOVEIN"
Public Const LOT_MOVEINERR As String = "MOVEINERR"
Public Const LOT_STARTING As String = "STARTING"
Public Const LOT_RUNNING As String = "RUNNING"
Public Const LOT_COMPLETED As String = "COMPLETED"
Public Const LOT_UPLOADPDS As String = "UPLOADPDS"
Public Const LOT_UPLOADPDSERR As String = "UPLOADPDSERR"
Public Const LOT_RDYUNLOAD As String = "RDYUNLOAD"
Public Const LOT_UNLOADING As String = "UNLOADING"
Public Const LOT_UNLOADED As String = "UNLOADED"
Public Const LOT_REMOVE As String = "REMOVE"
Public Const LOT_DATACOMPLETED As String = "DATACOMPLETED"
Public Const LOT_REMDATACOMPLETED As String = "REMDATACOMPLETED"
Public Const LOT_MOVEOUT As String = "MOVEOUT"
Public Const LOT_MOVEOUTERR As String = "MOVEOUTERR"
Public Const LOT_PAUSED As String = "PAUSED"
Public Const LOT_CHECKPPIDERR As String = "CHECKPPIDERR"
Public Const LOT_SELPPIDERR As String = "SELPPIDERR"
Public Const LOT_STARTERR As String = "STARTERR"
Public Const LOT_ABORTED As String = "ABORTED"
Public Const LOT_HOLD As String = "HOLD"
Public Const LOT_HOLDERR As String = "HOLDERR"
Public Const LOT_WAITINGHOLD As String = "WAITINGHOLD"
Public Const LOT_ENABLEDCBUTTON As String = "ENABLEDCBUTTON"
Public Const LOT_MSGERROR As String = "MSGERROR"
Public Const LOT_ARRIVEERR As String = "ARRIVEERR"
Public Const LOT_MAPPINGERR As String = "MAPPINGERR"
Public Const LOT_NOIMPLANT As String = "NOIMPLANT"
Public Const LOT_MISPPBODY As String = "MISPPBODY"
Public Const LOT_MISDOPANT As String = "MISDOPANT"

'batch states
Public Const BATCH_IDLE As String = "IDLE"
Public Const BATCH_FORMING As String = "FORMING"
Public Const BATCH_FORMED As String = "FORMED"
Public Const BATCH_MOVEIN As String = "MOVEIN"
Public Const BATCH_MOVEINERR As String = "MOVEINERR"
Public Const BATCH_RDYSTART As String = "RDYSTART"
Public Const BATCH_STARTING As String = "STARTING"
Public Const BATCH_RUNNING As String = "RUNNING"
Public Const BATCH_COMPLETED As String = "COMPLETED"
Public Const BATCH_UPLOADPDS As String = "UPLOADPDS"
Public Const BATCH_UPLOADPDSERR As String = "UPLOADPDSERR"
Public Const BATCH_REMOVE As String = "REMOVE"
Public Const BATCH_MOVEOUT As String = "MOVEOUT"
Public Const BATCH_MOVEOUTERR As String = "MOVEOUTERR"
Public Const BATCH_PAUSED As String = "PAUSED"
Public Const BATCH_CHECKPPIDERR As String = "CHECKPPIDERR"
Public Const BATCH_SELPPIDERR As String = "SELPPIDERR"
Public Const BATCH_STARTERR As String = "STARTERR"
Public Const BATCH_ABORTED As String = "ABORTED"
Public Const BATCH_HOLD As String = "HOLD"
Public Const BATCH_HOLDERR As String = "HOLDERR"
Public Const BATCH_WAITINGHOLD As String = "WAITNIGHOLD"
Public Const BATCH_MSGERROR As String = "MSGERROR"
Public Const BATCH_MAPPINGERR As String = "MAPPINGERR"
Public Const BATCH_NOIMPLANT As String = "NOIMPLANT"

'run modes
Public Const RUNMODE_PR As String = "PR"
Public Const RUNMODE_TEST As String = "TEST"
Public Const RUNMODE_NS As String = "NS"
Public Const RUNMODE_DMMTEST As String = "DMMTEST"
Public Const RUNMODE_DMMNS As String = "DMMNS"
Public Const RUNMODE_MEPR As String = "MEPR"
Public Const RUNMODE_MAJOR As String = "MAJOR"
Public Const RUNMODE_PRE As String = "PRE"
Public Const RUNMODE_POST As String = "POST"
Public Const RUNMODE_CMEPR As String = "CMEPR"
Public Const RUNMODE_CMAJOR As String = "CMAJOR"


