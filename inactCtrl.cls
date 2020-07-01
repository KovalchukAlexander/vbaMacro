VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "inactCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents AppEvt As Application
Attribute AppEvt.VB_VarHelpID = -1
#If VBA7 Then
    Private Declare PtrSafe Sub GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO)
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Sub GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO)
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
    'Instruction to use:
    '1. Be sure public Reflector procedure are declared in current Project modules
    '2. Make Instance of ThisClass
    '3. Verify that Sub:Reflector calling ThisClassInstance.Reflector
    '4. Switch if needed event Listener mode using ThisClassInstance:setEvtListenerMode
    '5. Set delay of inactivity by editing value of Initialize:evtMode  ( 0 - means do not controll)

Private Type timingBuffer
    onLoadTime As Single
    lastScheduledTriger As Single
    lastActionIdle As Single
    delayTime As Single
    lastActionTime As Single
End Type

Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

Private Enum evtListenerMode
    systemEvt = 1            ' Listne any action on PC
    excelEvt = 2             ' Listne Only Excel events
End Enum

Private evtMode As evtListenerMode

Private trigState As Boolean

Private Timing As timingBuffer 'Use time In Seconds stored as Single
Private cWb As Workbook

'------------------------------EVENTS HANDLING FOR EXCEL WORKBOOK----------------------------------------------
Private Sub AppEvt_SheetCalculate(ByVal Sh As Object)
    Timing.lastActionTime CSng(Time) 'renew lastActionTime
End Sub

Private Sub AppEvt_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Timing.lastActionTime CSng(Time) 'renew lastActionTime
End Sub

Private Sub AppEvt_SheetActivate(ByVal Sh As Object)
    Timing.lastActionTime CSng(Time) 'renew lastActionTime
End Sub

Private Sub AppEvt_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    Timing.lastActionTime CSng(Time) 'renew lastActionTime
End Sub

Private Sub AppEvt_WorkbookActivate(ByVal Wb As Workbook)
    Timing.lastActionTime CSng(Time) 'renew lastActionTime
End Sub

Private Sub AppEvt_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Timing.lastActionTime CSng(Time) 'renew lastActionTime
End Sub

'-----------------------------MAIN----------------------------------------------

Private Sub Class_Initialize()
    Set cWb = ThisWorkbook
    evtMode = systemEvt                                             'Edit this Property manually
    If evtMode = excelEvt Then Set AppEvt = cWb.Application
    Timing.lastActionTime = CSng(Time)
    Timing.onLoadTime = CSng(Time)
    Timing.delayTime = 0
    detectInactivity -1
End Sub


Private Function detectInactivity(Optional trigState As Boolean = False) 'parameter used for detecting whether to restart scheduler (true - init | false - restart)
    On Error GoTo errHandler
    If Timing.delayTime = 0 Then Err.Raise 402, "Inactivity Controll", "Inactivity Controll Is Initilize with Zero Delay"
    If trigState Then
        Timing.lastScheduledTriger = 0
    Else
        If Timing.lastScheduledTriger <> 0 Then
            Scheduler Timing.lastScheduledTriger, "Reflector", False 'dissmiss earlier task
            Timing.lastScheduledTriger = 0
        End If
    End If
    
    If evtMode = systemEvt Then
        Timing.lastActionIdle = getIdleTime
    ElseIf evtMode = excelEvt Then
        Timing.lastActionIdle = CSng(Time) - Timing.lastActionTime
    Else
        Err.Raise 401, "detectInactivity", "Event listener mode has bad state"
        End If
    
    If Timing.lastActionIdle > Timing.delayTime And Timing.delayTime > 0 Then
        cWb.Close SaveChanges:=True, RouteWorkbook:=True
    Else
        Scheduler Timing.delayTime, "Reflector", True
    End If
Exit Function
errHandler:
errLog
End Function

Private Function Scheduler(ByVal tInterval As Single, ByVal appProc As String, ByVal Sched As Boolean) As Boolean
    'method scheduling (set or dismiss) tasks where delay time is not zero
    'set or dismiss depends by value of (boolean)Scheduled parameter
    '("true" - planing(set), "false" - switch off(clear))
    
    On Error GoTo errHandler
    If tInterval = 0 Then Err.Raise 400, "Scheduler", "Inactivity Control = Off (delay time = " & tInterval & " s)"
    Dim shT As Date
    
    If Sched Then
        shT = Now() + TimeValue(CStr(CDate(tInterval * 0.00001)))
    Else
        shT = CDate(tInterval)
    End If
    
    Application.OnTime _
                        EarliestTime:=shT _
                        , Procedure:=appProc _
                        , Schedule:=Sched
                        
    Timing.lastScheduledTriger = CSng(shT) 'store last scheduled time as Single
    Scheduler = True
Exit Function
errHandler:
errLog
Scheduler = False
End Function

'-----------------------------------ADDS-------------------------------
Public Sub Reflector()
    detectInactivity -1 'not visible for scheduler try to use absolute path
End Sub

Public Property Let Delay(ByVal t As Single)
    Timing.delayTime = t
    Log "Delay set to " & t & " s"
    detectInactivity 0 'restart inactivity control
End Property

Public Property Get Delay() As Single
   Delay = Timing.delayTime
End Property

Private Function getIdleTime() As Single
  Dim a As LASTINPUTINFO
  a.cbSize = LenB(a)
  GetLastInputInfo a
  getIdleTime = CSng((GetTickCount() - a.dwTime) / 1000) 'seconds
End Function

Private Function Log(ByVal msg As String)
    Debug.Print "Time: " & VBA.Time & VBA.Chr(&H9) & msg & VBA.vbCr
End Function

Private Function errLog()
    Debug.Print "Time: " & VBA.Time & VBA.Chr(&H9) & "_ErrNum:" & Err.Number & VBA.Chr(&H9) & "_errSource:" & Err.Source & VBA.Chr(&H9) & "_errDesc:" & Err.Description & VBA.vbCr
End Function

