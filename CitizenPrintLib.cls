VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CitizenPrintLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class intended for make receipts with correct syntax.
'SIT - Serial Interface Type
'PIT - Paralel Interface Type

#If VBA7 Then ' Excel 2010 or later
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else ' Excel 2007 or earlier
    Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If

Public Enum COMInput
    HearInput = True
    SilentInput = False
End Enum

Public Enum SIT_Character
    International = 1
    NotSet = -1
End Enum

Public Enum RegisterState 'if registered position outside an enum then sentence will print immediately if so, everything will be printed according schedule (position in register)
    NormalPrint = 0
    s1 = 1
    s2 = 2
    s3 = 3
    s4 = 4
    s5 = 5
    s6 = 6
    s7 = 7
    s8 = 8
End Enum

Private Enum RegisterStatus
    Ready = 0
    Waiting = 1
    Printing = 2
    Errors = 3
    Initialized = 4
End Enum

Public Enum SIT_CR 'must correspond to ControlCodes Collection
    CR_Disregarded = 3
    CR_LF = 4
End Enum

Public Enum SIT_BitLen
    b8 = 8
    b7 = 7
End Enum

Public Enum SIT_Parity
    pProvided = 2
    pNotProvided = 1
    pEven = 3
    pOdd = 4
    pSpace = 5
    pMarker = 6
End Enum

Public Enum SIT_Conditions
    Odd = 1
    Even = 2
End Enum

Public Enum SIT_BaudRate
    br1200 = 1200
    br2400 = 2400
    br4800 = 4800
    br9600 = 9600
End Enum

Public Enum SIT_DataFormat
    col24 = 24
    col40 = 40
End Enum

Public Enum SIT_StopBit
    b1 = 1
    b1_5 = 1.5
    b2 = 2
End Enum

Public Enum SIT_Country
    US = 0
    France = 1
    Germany = 2
    England = 3
    Denmark = 4
    Sweden = 5
    Italy = 6
    Spain = 7
    Japan = 9
End Enum

Public Enum InterfaceType
    Serial = 1                      'SIT
    Paralel = 2                     'PIT
End Enum

Public Enum ParamList
    Character = 1
    CR = 2
    BitLength = 3
    Parity = 4
    Conditions = 5
    DataFormat = 6
    BaudRate = 7
    IFaceType = 8
    COM_Number = 9
    StopBit = 10
End Enum
'Make Public in any Module for Reflection
Private Type SIT
    Character As SIT_Character      'default 1
    CR As SIT_CR                    'default 1
    BitLength As SIT_BitLen         'default 8
    Parity As SIT_Parity            'default 1
    Conditions As SIT_Conditions    'default 1
    DataFormat As SIT_DataFormat    'default 40
    BaudRate As SIT_BaudRate        'default 4800
    IFaceType As InterfaceType      'default 1
    COM_Number As Integer
    StopBit As SIT_StopBit
End Type

Private Type PIT
    Character As SIT_Character
    CR As SIT_CR
    DataFormat As SIT_DataFormat
    IFaceType As InterfaceType
End Type

Private Type RegisterBuffer
    Sentence() As String
    RegPos() As RegisterState
    RegStatus() As RegisterStatus
End Type

Public ControlCodes As Collection
Private Register As RegisterBuffer
Private IFaceConfig As SIT
Private COM As WinSerialApi

Private Sub Class_Initialize()
Set ControlCodes = New Collection
ControlCodes.Add "OK", "TEST"
ControlCodes.Add VBA.Chr(&HA), "LF"
ControlCodes.Add VBA.Chr(&HD), "CR"
ControlCodes.Add VBA.Chr(&HD) & VBA.Chr(&HA), "CRLF"
ControlCodes.Add VBA.Chr(&HF), "SI"
ControlCodes.Add VBA.Chr(&HE), "SO"
ControlCodes.Add VBA.Chr(&H1F), "US"
ControlCodes.Add VBA.Chr(&H1E), "RS"
ControlCodes.Add VBA.Chr(&H18), "CAN"
ControlCodes.Add VBA.Chr(&H12), "DC2"
ControlCodes.Add VBA.Chr(&H11), "DC1"
ControlCodes.Add VBA.Chr(&H1B), "ESC"
ControlCodes.Add VBA.Chr(&H1B) & VBA.Chr(&H42), "ESC_B"
ControlCodes.Add VBA.Chr(&H1B) & VBA.Chr(&H52), "ESC_R"
ControlCodes.Add VBA.Chr(&H1B) & VBA.Chr(&H2F), "ESC_/"
ControlCodes.Add VBA.Chr(&H1B) & VBA.Chr(&H21), "ESC_!"
ControlCodes.Add VBA.Chr(&H1B) & VBA.Chr(&H26), "ESC_&"
ControlCodes.Add VBA.Chr(&H1B) & VBA.Chr(&H25), "ESC_%"
ControlCodes.Add VBA.Chr(&H1B) & VBA.Chr(&H4B), "ESC_K"
ControlCodes.Add VBA.Chr(&H1C) & VBA.Chr(&H57) & VBA.Chr(&H31), "FS_W_1"
ControlCodes.Add VBA.Chr(&H1C) & VBA.Chr(&H57) & VBA.Chr(&H30), "FS_W_0"

IFaceConfig.BaudRate = br4800
IFaceConfig.BitLength = b7
IFaceConfig.Character = International
IFaceConfig.Conditions = Odd
IFaceConfig.CR = CR_LF
IFaceConfig.DataFormat = col40
IFaceConfig.Parity = pNotProvided
IFaceConfig.IFaceType = Serial
IFaceConfig.COM_Number = 4
IFaceConfig.StopBit = b1

'Set COM = New WinSerialApi  'WSA.CommOpen(4, "COM4", "4800N71")

' Not used yet ----------------- page 28-29 of UserManual
ReDim Register.RegPos(9)
ReDim Register.RegStatus(9)
ReDim Register.Sentence(9)
Register.RegPos(0) = s1                 'First free register
Register.RegStatus(0) = Initialized
Register.Sentence(0) = vbNullString

'--------------------------------------------------

End Sub

Public Function errReport(Optional ByRef eObj As ErrObject)
If Not (eObj Is Nothing) Then Debug.Print Err.Number & VBA.Chr(&H3A) & Err.Description & " (Source: " & Err.Source & ")" & VBA.vbCr
If Not (Err Is Nothing) Then Debug.Print Err.Number & VBA.Chr(&H3A) & Err.Description & " (Source: " & Err.Source & ")" & VBA.vbCr
End Function

Private Function FilterCode(ByVal SourceString As String, ByVal Code As String) As String
                    On Error GoTo errH
                    
    FilterCode = Replace(SourceString, ControlCodes.Item(Code), vbNullString, compare:=vbBinaryCompare)
    
Exit Function
errH:
errReport
FilterCode = vbNullString
End Function

Public Function setIFaceParam(Optional ByVal IF_Type As InterfaceType, _
                                Optional ByVal IF_BaudRate As SIT_BaudRate, _
                                Optional ByVal IF_DataFormat As SIT_DataFormat, _
                                Optional ByVal IF_Char As SIT_Character, _
                                Optional ByVal IF_CR As SIT_CR, _
                                Optional ByVal IF_BitLength As SIT_BitLen, _
                                Optional IF_Parity As SIT_Parity, _
                                Optional IF_Conditions As SIT_Conditions, _
                                Optional IF_COM_Number As Integer, _
                                Optional IF_StopBit As SIT_StopBit _
                                ) As Boolean
                    On Error GoTo errH
'Need validate values
If IF_Type Then IFaceConfig.IFaceType = IF_Type
If IF_BaudRate Then IFaceConfig.BaudRate = IF_BaudRate
If IF_DataFormat Then IFaceConfig.DataFormat = IF_DataFormat
If IF_Char Then IFaceConfig.Character = IF_Char
If IF_CR Then IFaceConfig.CR = IF_CR
If IF_BitLength Then IFaceConfig.BitLength = IF_BitLength
If IF_Parity Then IFaceConfig.Parity = IF_Parity
If IF_Conditions Then IFaceConfig.Conditions = IF_Conditions
If IF_COM_Number Then IFaceConfig.COM_Number = IF_COM_Number
If IF_StopBit Then IFaceConfig.StopBit = IF_StopBit

setifaceconfig = True
Exit Function
errH:
errReport
setifaceconfig = False
End Function

Public Function COMSetUp()
    Dim COMSet As New COMSetUp
    Dim Pare As SIT_Parity
    
    COMSet.initForm IFaceConfig.COM_Number, IFaceConfig.BaudRate, IFaceConfig.StopBit, IFaceConfig.Parity, IFaceConfig.BitLength
    
    If COMSet.getItem("Parity") = "N" Then Pare = pNotProvided
    
        setIFaceParam _
                                IF_Type:=Serial, _
                                IF_BaudRate:=COMSet.getItem("BaudRate"), _
                                IF_BitLength:=COMSet.getItem("DataBits"), _
                                IF_Parity:=Pare, _
                                IF_COM_Number:=CInt(COMSet.getItem("COMNumber")), _
                                IF_StopBit:=COMSet.getItem("StopBits")
    
    
End Function

Public Function getIFaceParam(ByVal pName As ParamList) As Variant
    Select Case pName
        Case BaudRate
            getIFaceParam = IFaceConfig.BaudRate
        Case BitLength
            getIFaceParam = IFaceConfig.BitLength
        Case Character
            getIFaceParam = IFaceConfig.Character
        Case COM_Number
            getIFaceParam = IFaceConfig.COM_Number
        Case Conditions
            getIFaceParam = IFaceConfig.Conditions
        Case CR
            getIFaceParam = IFaceConfig.CR
        Case DataFormat
            getIFaceParam = IFaceConfig.DataFormat
        Case IFaceType
            getIFaceParam = IFaceConfig.IFaceType
        Case StopBit
            getIFaceParam = IFaceConfig.StopBit
        Case Parity
            If (IFaceConfig.Parity = pNotProvided) Then
                getIFaceParam = "N"
                Else
                getIFaceParam = vbNullString
            End If
        Case Else
        getIFaceParam = Empty
    End Select
End Function

Private Function addRegister()
 'not used page 28-29 of UserManual
End Function

Private Function freeRegister()
'not used  page 28-29 of UserManual
End Function


Public Function showPreview(ByVal Data As String)
Dim Prev As Visualisation
If Data = vbNullString Then Exit Function
Set Prev = New Visualisation
Prev.Content = "~~~~~~~~~~~~~~~~~~~~~~~~" & ControlCodes("CR") & Replace(Data, ControlCodes("ESC_/") & "0", vbNullString) & ControlCodes("CR") & "~~~~~~~~~~~~~~~~~~~~~~~~"
Prev.Height = Math.Round(VBA.Len(Data) / IFaceConfig.DataFormat, 0) * 30
Prev.Content.Height = Prev.Height - 10
Prev.Show
End Function

'------------------------------------------LIBRARY OF ACTIONS------------------------------------------------------------------------
'All methods return a character of special action. For more details see "Citizen CBM-910 user manual".
'Actions has unified and predefined by hardware microprogram.
'

Public Function LineFeedAfterPrint(Optional ByVal CR_Type As SIT_CR = CR_Disregarded) As String
                    On Error GoTo errH
                    
If Not (CR_Type = CR_Disregarded Or CR_Type = CR_LF) And Not (IsEmpty(IFaceConfig)) Then CR_Type = IFaceConfig.CR
If Not (CR_Type = CR_Disregarded Or CR_Type = CR_LF) Then Err.Raise 400, "CitizenPrintApi::LineFeedAfterPrint", "Wrong argument. Choose from enum or leave not set."
    
If CR_Type = CR_LF Then
    LineFeedAfterPrint = ControlCodes("CR") & ControlCodes("LF")
Else
    LineFeedAfterPrint = ControlCodes("CR")
End If

Exit Function
errH:
errReport
LineFeedAfterPrint = vbNullString
End Function

Public Function ContinuousPaperFeed(ByVal n As Integer) As String
                On Error GoTo errH

If (n < 4 Or n > 255) Then Err.Raise 400, "CitizenPrintApi::ContinuousPaperFeed", "Wrong number of lines(dots) - " & n & ", value must into the range from 4 to 255."
ContinuousPaperFeed = ControlCodes("ESC_B") & n

Exit Function
errH:
errReport
ContinuousPaperFeed = vbNullString
End Function

Public Function SetInternationalCharacter(ByVal n As SIT_Country) As String
If Not (0 <= n And n <= 8) Then n = US
    SetInternationalCharacter = ControlCodes("ESC_R") & n
End Function

Public Function SentenceRegistration(ByVal Sentence As String, Optional ByVal n As RegisterState = NormalPrint) As String
     On Error GoTo errH
    'not working proper . Sentence print but not use register
    Sentence = Sentence & ControlCodes(IFaceConfig.CR)
    If VBA.Len(Sentence) <= IFaceConfig.DataFormat Then
        SentenceRegistration = ControlCodes("ESC_/") & n & Sentence
    Else
        Err.Raise 200, "CitizenPrintApi::SentenceRegistration", "Could not register Sentence. Data cannot exceed one line (" & IFaceConfig.DataFormat & " characters) which should be ended with CR or LF."
    End If
Exit Function
errH:
errReport
SentenceRegistration = vbNullString
End Function

Public Function RegisteregSentencePrint(ByVal n As RegisterState) As String
    RegisteregSentencePrint = ControlCodes("ESC_!") & n
End Function

Public Function DataCancel() As String
    DataCancel = ControlCodes("CAN")
End Function

Public Function InitialSetting() As String
    InitialSetting = ControlCodes("DC1")
End Function

Public Function InvertedLetter() As String
    InvertedLetter = ControlCodes("DC2")
End Function

Public Function StandardLetter() As String
    If IFaceConfig.IFaceType = Serial Then
        StandardLetter = ControlCodes("SI")
    ElseIf (IFaceConfig.IFaceType = Paralel) Then
        StandardLetter = ControlCodes("US")
    Else
        Err.Raise 100, "CitizenPrintApi::StandardLetter", "Interface not set. Use setIFaceConfig(InterfaceType) to fix."
    End If
End Function

Public Function EnlargedLetter() As String
     If IFaceConfig.IFaceType = Serial Then
        EnlargedLetter = ControlCodes("SO")
    ElseIf (IFaceConfig.IFaceType = Paralel) Then
        EnlargedLetter = ControlCodes("RS")
    Else
        Err.Raise 100, "CitizenPrintApi::StandardLetter", "Interface not set. Use setIFaceConfig(InterfaceType) to fix."
    End If
End Function

Public Function CharRegistration_notImp(ByVal PatternData As Byte, ByVal A1 As Byte, ByVal A2 As Byte) As String
On Error GoTo errH
    Debug.Print "Method not implemented yet."
    Exit Function
    'has a difference for usage between 24 or 40 column modes
    MsgBox CByte(PatternData)
    If &H20 > A1 Or &HFF < A1 Then Err.Raise 200, "CitizenPrintApi::CharRegistration", "A1 and A2 must be in range from 20H to FFH."
    If &H20 > A2 Or &HFF < A1 Then Err.Raise 200, "CitizenPrintApi::CharRegistration", "A1 and A2 must be in range from 20H to FFH."
    If A2 <= A1 Then Err.Raise 200, "CitizenPrintApi::CharRegistration", "A1 must be smaler than A2."
    If A2 - A1 > 7 Then Err.Raise 200, "CitizenPrintApi::CharRegistration", "A2 - A1 must be positive and smaler or equal 7."
    If Len(PatternData) < 7 Then CharRegistration = ControlCodes("ESC_&") & A1 & A2 & PatternData
Exit Function
errH:
errReport
CharRegistration = vbNullString
End Function

Public Function bitImagePrint_notImp() As String
Debug.Print "Method not implemented yet. See UserManual p.36"
    Exit Function
End Function
