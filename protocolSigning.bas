Attribute VB_Name = "protocolSigning"
Private Wb As Workbook
Private App As Application

Public Type onSheetData
    xFormula As String
    xValue As String
End Type

Public Type formulaCollection
    Formula As Collection
    Key As Collection
End Type

Public Type headerFooterContent
    lHeader As String
    cHeader As String
    rHeader As String
    lFooter As String
    cFooter As String
    rFooter As String
End Type

Public Enum isValidEnum
    xValid = 1
    xOverride = 2
    xNotValid = 3
End Enum

Public mailInfoWbPath As String
Public mailInfoFormulaHash As String
Public mailInfoUserPasswd As String
Public mailInfoIsNew As Boolean
Public mailInfoReadyToSend As Boolean
Public mailInfoWbSaved As Boolean

Public uPasswd, refPasswd As String
Public validatePassword As Boolean
Public HFDump As headerFooterContent

Const defaultTarget = "A1:AZ500" 'don't know exactly how to define this value automatically
Const formulaVar = "formulaHash"
Const passwdVar = "userPasswd"

Public Function detectEnvObj(Optional ByVal Sh As Worksheet = Nothing) As Worksheet
    On Error GoTo errH
    If Sh Is Nothing Then
        If Not (ActiveCell Is Nothing) Then Set Sh = ActiveCell.Parent
        If Sh Is Nothing Then Set Sh = ActiveSheet
    End If

    If Not (Sh Is Nothing) Then Set Wb = Sh.Parent
    If (Wb Is Nothing) And (Not (ActiveWorkbook Is Nothing)) Then Set Wb = ActiveWorkbook
    If Wb Is Nothing Then Set Wb = ThisWorkbook
    If (Sh Is Nothing) And (Not (Wb Is Nothing)) Then Set Sh = Wb.ActiveSheet
    If App Is Nothing Then Set App = Wb.Application
    
    Set detectEnvObj = Sh
Exit Function
errH:
    Debug.Print Err.Number & ":" & Err.Description
    Resume Next
End Function

'----------------------------------------WORK WITH HEADERS AND FOOTERS INFORMATION----------------------------------------------------
Public Function getHeaderFooterDump(Optional ByVal Sh As Worksheet) As Boolean
    Set Sh = detectEnvObj(Sh)

App.PrintCommunication = False
With Sh.PageSetup
    HFDump.cFooter = .CenterFooter
    HFDump.rFooter = .RightFooter
    HFDump.lFooter = .LeftFooter
    HFDump.lHeader = .LeftHeader
    HFDump.cHeader = .CenterHeader
    HFDump.rHeader = .RightHeader
End With
App.PrintCommunication = True
getHeaderFooterDump = True
End Function

Public Function setHeaderFooter(Optional ByRef Sh As Worksheet _
                            , Optional ByVal lHeader As String = "EmptyNullOff" _
                            , Optional ByVal cHeader As String = "EmptyNullOff" _
                            , Optional ByVal rHeader As String = "EmptyNullOff" _
                            , Optional ByVal lFooter As String = "EmptyNullOff" _
                            , Optional ByVal cFooter As String = "EmptyNullOff" _
                            , Optional ByVal rFooter As String = "EmptyNullOff" _
                            , Optional ByVal fFormat As String = vbNullString _
                            ) As Boolean
Set Sh = detectEnvObj(Sh)

App.PrintCommunication = False
With Sh.PageSetup
    If (lFooter <> "EmptyNullOff") Then .LeftFooter = fFormat & lFooter
    If (rFooter <> "EmptyNullOff") Then .RightFooter = fFormat & rFooter 'page count can get using &N and current page number &P ' &P of &N
    If (cFooter <> "EmptyNullOff") Then .CenterFooter = fFormat & cFooter
    If (lHeader <> "EmptyNullOff") Then .LeftHeader = fFormat & lHeader
    If (cHeader <> "EmptyNullOff") Then .CenterHeader = fFormat & cHeader
    If (rHeader <> "EmptyNullOff") Then .RightHeader = fFormat & rHeader
End With
App.PrintCommunication = True
setHeaderFooter = True
End Function

Public Function getWbName() As String
    If Wb Is Nothing Then detectEnvObj
    getWbName = Wb.Name
End Function

Public Sub clearAllFooters()
    Dim Sh As Worksheet
    Set Sh = detectEnvObj(Sh)
    
    If Sh Is Nothing Then Exit Sub
        
    For Each Sh In Wb.Sheets
        setHeaderFooter Sh, lFooter:=vbNullString
    Next Sh
End Sub
'---------------------------------------------------------END HEADERS AND FOOTERS METHODS---------------------------------------------
'-------------------------------------------------------PUBLIC METHODS---------------------------------------------------------------
Public Function getOnSheetData(Optional ByRef Target As Range, Optional Sh As Worksheet = Nothing) As onSheetData 'scan target range and collect formulas and values
    On Error Resume Next

    Dim Str As String
    Dim xCell, xEl As Variant
    Dim xData As onSheetData
    Dim uFormulas As formulaCollection
    
    Set uFormulas.Formula = New Collection
    Set uFormulas.Key = New Collection
    
    Set Sh = detectEnvObj(Sh)
    
    Sh.Unprotect
    
    
    If (Target Is Nothing) Then Set Target = Sh.Range(defaultTarget)
    
    For Each xCell In Sh.Range(Target.Address).Cells
        If (xCell.HasFormula = True) Then  'collect formulas to collection
            uFormulas.Formula.Add xCell.Formula, xCell.Address
            uFormulas.Key.Add xCell.Address
        End If
        
       xData.xValue = xData.xValue & CStr(Trim(xCell.Value)) 'collect values in string
    Next xCell
    
     For Each xEl In uFormulas.Key
            Str = Str & Trim(xEl) & Trim(uFormulas.Formula.Item(xEl)) 'Convert collection of formulas to string
     Next xEl

            xData.xFormula = Str
    getOnSheetData = xData
    shProtect Sh
End Function

Public Function isValidProtocol() As isValidEnum
    On Error GoTo errH
    Dim xData As onSheetData
    Dim Store As String
    
    xData = getOnSheetData
    Store = getPersistVariable(formulaVar)
    
    If Store = vbNullString Then
        isValidProtocol = xOverride 'Err.Raise 201, "isValidProtocol", "Расчеты не подписаны (подпишите используя процедуру setCalculationSign на вкладке Сервис/Разработчик -> Макросы)"
        Exit Function
    End If
    If (Store = getHashOfString(xData.xFormula)) Then
        If (getPersistVariable(passwdVar) <> vbNullString) Then isValidProtocol = xValid
    Else
        isValidProtocol = xNotValid
    End If
    Exit Function
errH:
    Debug.Print Err.Description
    MsgBox Err.Description
    isValidProtocol = xNotValid
End Function

Public Sub setCalculationSign()
    On Error GoTo errH
    Dim xData As onSheetData
    Dim Hash As String
    
    If Not AuthorizeMe(True) Then Exit Sub
    
    xData = getOnSheetData
    Hash = getHashOfString(xData.xFormula)
    setPersistVariable formulaVar, Hash
    
    mailInfoIsNew = True
    mailInfoReadyToSend = False
    mailInfoWbSaved = False
    mailInfoFormulaHash = "Поточний хеш формул: " & Hash
    mailInfoUserPasswd = "Пароль: " & uPasswd
        
    Exit Sub
errH:
    Debug.Print Err.Description
    MsgBox Err.Description
End Sub

Public Sub getCalculationSign()
    On Error GoTo errH
    
    If Not AuthorizeMe Then Exit Sub
    MsgBox "Valid Formula Hash:" & vbCr & getPersistVariable(formulaVar)

    Exit Sub
errH:
    Debug.Print Err.Description
    MsgBox Err.Description
End Sub

Private Function packData(ByVal Str As String, Optional ByVal Sugar As String = passwdVar) As String
    packData = VBA.Len(Str) & Str & Sugar
End Function

'-------------------------------------------------------PRIVATE METHODS----------------------------------------------
Private Function AuthorizeMe(Optional ByVal allowStoring As Boolean = False) As Boolean
    On Error GoTo errH
    Dim Sugar, storePasswd, refPasswd As String
    
    If uPasswd = vbNullString Then
        Sugar = getPasswd
    Else
        Sugar = uPasswd
    End If
    
    If Sugar = vbNullString Then GoTo Reject
    storePasswd = getPersistVariable(passwdVar)
    
    If storePasswd = vbNullString And allowStoring Then
        refPasswd = getPasswd(True)
        If (Not (refPasswd = Sugar)) Then GoTo Reject
        setPersistVariable passwdVar, getHashOfString(packData(Sugar, passwdVar))
        AuthorizeMe = True
    ElseIf storePasswd = getHashOfString(packData(Sugar, passwdVar)) Then
        AuthorizeMe = True
    Else
Reject:
        uPasswd = vbNullString
        AuthorizeMe = False
        Err.Raise 200, "protocolSigning", "Не правильний пароль!"
    End If
Exit Function
errH:
MsgBox Err.Description
AuthorizeMe = False
End Function

Private Function getPasswd(Optional ByVal pType As Boolean = False) As String
On Error GoTo errH
    Dim pRF As New requestPassword
    validatePassword = pType
    If pType = False Then
        uPasswd = vbNullString
        pRF.Show
        getPasswd = uPasswd
    Else
        refPasswd = vbNullString
        pRF.Caption = "Реєстрація"
        pRF.Label1 = "Повторіть пароль:"
        pRF.Show
        getPasswd = refPasswd
    End If
Exit Function
errH:
MsgBox Err.Description
End Function

Private Sub resetAllPasswd()
'make public only if you know that it's vulnerability and for single use only.
' strongly recommended store this procedure in private scope in protected macro
    Dim Sh As Worksheet
    
    Set Sh = detectEnvObj(Sh)
    
    If Sh Is Nothing Then Exit Sub
    
    For Each Sh In Wb.Sheets
        setPersistVariable VarName:=passwdVar, Sh:=Sh
    Next Sh
End Sub

Private Function pregMaskNonAlfa(ByVal Str As String) As String
On Error Resume Next
Dim Preg As Object
Dim xCh As Variant
Set Preg = New RegExp 'CreateObject("VBScript.RegExp")
Preg.Pattern = "([\x20\x21\x22\x23\x24\x25\x26\x27\x28\x29\x2a\x2b\x2c\x2d\x2e\x2f\x3a\x3b\x3c\x3d\x3e\x3f\x40\x5b\x5c\x5d\x5e\x5f\x60\x7b\x7c\x7d\x7e]{1})?"
Preg.Global = True
Preg.MultiLine = False
Preg.IgnoreCase = True
    For Each xCh In Preg.Execute(Str)
        
      If xCh <> vbNullString Then
        xCh = Hex(VBA.Asc(xCh))
        Preg.Pattern = "[\x" & xCh & "]"
        Str = Preg.Replace(Str, "x" & xCh)
        End If
    Next xCh
pregMaskNonAlfa = Str
End Function

'--------------------------------------------PERSITENT STORAGE OF VARIABLES------------------------------------------
'use only unique name because values with same key will rewrites
'sending empty values removes key from storage
Private Function setPersistVariable(ByVal VarName As String, Optional ByVal VarValue As String = vbNullString, Optional ByRef Sh As Worksheet) As Boolean
    Dim XMLParts As CustomXMLPart
    
    Set Sh = detectEnvObj(Sh)
    
    Dim Preg As New RegExp
    Dim TagName As String
    
    Preg.Global = True
    Preg.IgnoreCase = True
    Preg.MultiLine = True
    
    TagName = pregMaskNonAlfa(VarName & "_on_" & Sh.Name)
    
    Preg.Pattern = "<" & TagName & ">(.*)</" & TagName & ">" 'provide limitation for storing values
    
    For Each XMLParts In Wb.CustomXMLParts
        If Preg.Test(XMLParts.XML) Then
            XMLParts.Delete
        End If
    Next XMLParts
    
    If VarValue <> vbNullString Then Set XMLParts = Wb.CustomXMLParts.Add("<" & TagName & ">" & VarValue & "</" & TagName & ">")
    setPersistVariable = True
End Function

Private Function getPersistVariable(ByVal VarName As String, Optional ByRef Sh As Worksheet) As String
    Dim XMLPart As CustomXMLPart
    Dim Preg As New RegExp
    Dim TagName As String
    
    Set Sh = detectEnvObj(Sh)
    
    Preg.Global = True
    Preg.IgnoreCase = True
    Preg.MultiLine = True
    
    TagName = pregMaskNonAlfa(VarName & "_on_" & Sh.Name)
    
    Preg.Pattern = "<" & TagName & ">(.*)</" & TagName & ">"
    
    For Each XMLParts In Wb.CustomXMLParts
        If Preg.Test(XMLParts.XML) Then
            getPersistVariable = Preg.Execute(XMLParts.XML)(0).SubMatches(0)
            Exit Function
        End If
    Next XMLParts
    getPersistVariable = vbNullString
End Function
'---------------------------------------------END STORAGE METHODS (TESTED)-------------------------------------------
'recommended to provide validation and limitation of values for storing by changing RegExp pattern

'-------------------------------------------AUTOMATION---------------------------------------------------------------

Private Sub Auto_Open()
Dim btnCollection As New Collection
Dim Index As Long

btnCollection.Add Array("Встановити Оновити підпис", "protocolSigning.setCalculationSign", "setCalculationSign")
btnCollection.Add Array("Переглянути підпис", "protocolSigning.getCalculationSign", "getCalculationSign")


For Index = 1 To btnCollection.Count
     remBtn btnCollection.Item(Index)(0)
     mkBtn btnCollection.Item(Index)(0), btnCollection.Item(Index)(1), btnCollection.Item(Index)(2)
Next Index

End Sub

Private Function mkBtn(ByVal btnName As String, ByVal btnClbck As String, Optional ByVal Descript As String) As Boolean
    Dim cmdBtn As CommandBarButton
    On Error Resume Next
        With Application
            Set cmdBtn = .CommandBars("Cell").Controls.Add(Temporary:=True)
        End With
 
        With cmdBtn
           .Caption = btnName
           .DescriptionText = Descript
           .Style = msoButtonCaption
           .OnAction = btnClbck
        End With
        mkBtn = True
End Function

Private Function remBtn(ByVal btnName As String) As Boolean
    Dim cmdBtn As CommandBarButton
    On Error Resume Next
        With Application
            .CommandBars("Cell").Controls(btnName).Delete
        End With
        
        remBtn = True
End Function

'---------------------------------------------OTHER METHODS----------------------------------------------------------
Private Function shProtect(ByRef Sh As Worksheet) As Boolean
    Sh.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
End Function

