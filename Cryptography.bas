Attribute VB_Name = "Cryptography"
'Author of changes: A.Kovalchuk
'Default result string: MD5 in Hex

Public Enum b64
    ConvToBase64 = 1
    ConvToHex = 0
End Enum

Public Enum CryptoAlg
    MD5 = 1
    SHA1 = 2
    SHA256 = 3
    HMACSHA512 = 4
End Enum

Public Function getHashOfString(Optional ByVal ResourceString As String = vbNullString _
                                , Optional ByVal CriptoAlgorithm As CryptoAlg = 1 _
                                , Optional ByVal Base As b64 = 0 _
                                ) As String
On Error GoTo errHandler
If ResourceString = vbNullString Then Exit Function
Dim Crypto, UTF As Object
Dim Result() As Byte
Dim criptoAlgResourceString(5) As String

criptoAlgResourceString(1) = "System.Security.Cryptography.MD5CryptoServiceProvider"
criptoAlgResourceString(2) = "System.Security.Cryptography.SHA1Managed"
criptoAlgResourceString(3) = "System.Security.Cryptography.SHA256Managed"
criptoAlgResourceString(4) = "System.Security.Cryptography.HMACSHA512"

Set UTF = CreateObject("System.Text.UTF8Encoding")
Set Crypto = CreateObject(criptoAlgResourceString(CriptoAlgorithm))
   Result = Crypto.ComputeHash_2(UTF.GetBytes_4(ResourceString))
   If Base = ConvToBase64 Then
        getHashOfString = ConvToBase64String(Result)
    Else
        getHashOfString = ConvToHexString(Result)
    End If
Exit Function
errHandler:
Debug.Print Err.Number & ": " & Err.Description & " _ON:" & Err.Source & "(Cryptography)"
getHashOfString = vbNullString
End Function

Private Function getHashOfFile(ByVal FilePath As String) As String
'File could be read with:
'Dim FSO As Object
'Set FSO = CreateObject("Scripting.FileSystemObject")
'See FSO manpage
'OR Native VBA
'Open FilePath For Binary Input Lock Write As #FF 'FF = FreeFile
'
'not implemented yet
End Function

'additional methods (for servicing of main function only)
Private Function getBitesFromString(Optional ByVal xStr As String = vbNullString) As Variant
    Exit Function 'for this task we use built in function System.Text.UTF8Encoding.GetBytes_4(String)
    
    Dim i, r(), b() As Byte
    b = xStr
    ReDim r(0)
    For Each i In b
        If (i = 0 Or IsEmpty(i)) Then GoTo 1
        ReDim Preserve r(UBound(r) + 1)
        r(UBound(r) - 1) = i
1:
    Next
    ReDim Preserve r(UBound(r) - 1)
    getBitesFromString = r
End Function


Private Function ConvToHexString(vIn As Variant) As Variant
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing
End Function

Private Function ConvToBase64String(vIn As Variant) As Variant
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing
End Function
