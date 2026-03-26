Attribute VB_Name = "General"
Option Explicit

Private aDecTab(255)        As Integer
Private aEncTab(63)         As Byte

'''''''''''''''23112010
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const FLOAT = 1, SINK = 0
Public Type VERSION_TYPE
    Version                 As String * 16
End Type
Public Rutecode_new As String
Public min_fare As Single
Public BustypeID As Long

' ------------------------------------------------------------
'                   Base 64 Radix functions
' ------------------------------------------------------------

Private Function PadString(strData As String) As String
Dim nLen As Long
Dim sPad As String
Dim nPad As Integer
nLen = Len(strData)
nPad = ((nLen \ 8) + 1) * 8 - nLen
sPad = String(nPad, Chr(nPad))
PadString = strData & sPad
End Function


Public Function selectMatchDropdown(ByVal objComboBox As ComboBox, Index As Long)
On Error GoTo erromod
Dim I As Integer
For I = 0 To objComboBox.ListCount - 1
    If objComboBox.ItemData(I) = Index Then
       objComboBox.ListIndex = I
       Exit For
    End If
Next
Exit Function
erromod:
MsgBox "Error due to " & err.Description, vbInformation
End Function


Public Function WriteVersionInfo()

Dim WriteString As String
Dim FreeFileHandle As Integer
Dim I As Integer

For I = 1 To Len(App.FileDescription)
    WriteString = WriteString & "|" & Hex(Asc(Mid(App.FileDescription, I, 1)))
Next I

WriteString = App.Major & "." & App.Minor & "." & App.Revision & "_" & WriteString

FreeFileHandle = FreeFile()
Open App.Path & "\SilVerDt.cfg" For Output As #FreeFileHandle
    Print #FreeFileHandle, WriteString
Close #FreeFileHandle

End Function
Private Function UnpadString(strData As String) As String
Dim nLen As Long
Dim nPad As Long
nLen = Len(strData)
If nLen = 0 Then Exit Function
nPad = Asc(Right(strData, 1))
If nPad > 8 Then nPad = 0
UnpadString = Left(strData, nLen - nPad)
End Function



Public Sub ValidateKeyPress(ValidateTxtBox As Object, KeyPressAscii As Integer, Optional UseUpperCase As Boolean = False, Optional IntegerDigitCount As Byte = 8, Optional DecimalDigitCount As Byte = 2)
On Error GoTo ErrorMod
Dim DotPos As Byte
'''''''''''''''''''''''''''''''''''''''''''
'Author  : Subodh [ Copy Enum For Validation mod for using this code in another project ]
'Purpose : KeyPress Validation in text boxes
'Modifed : 28-04-2015 Added Two Optional Arguments Decimal Point Limiting Maximun Integer Number
'''''''''''''''''''''''''''''''''''''''''''

    If KeyPressAscii = 13 Then
        SendKeys "{TAB}" 'Error may occur in Vista , Win 7 or greater while debugging but when compiled to Exe there will be no issue
''    ElseIf KeyPressAscii = 22 Then 'To Block Paste Text Content using Ctrl+V uncommnet these two lines .
''        KeyPressAscii = 0
    ElseIf KeyPressAscii < 32 Then 'Used to Allow Delete, Backspace, Copy etc.
        Exit Sub
    End If
    
    If ValidationMode = FloatingPointValue Then
        If InStr(1, "0123456789.", Chr(KeyPressAscii)) = 0 Then
            KeyPressAscii = 0
        ElseIf InStr(1, ".", Chr(KeyPressAscii)) = 1 Then
            If InStr(1, ValidateTxtBox.Text, ".", vbBinaryCompare) > 0 Then
                KeyPressAscii = 0
            End If
        ElseIf InStr(1, "0123456789", Chr(KeyPressAscii)) > 0 Then
            If InStr(1, ValidateTxtBox.Text, ".", vbBinaryCompare) > 0 Then
                DotPos = InStr(1, ValidateTxtBox.Text, ".", vbBinaryCompare)
                If Len(Right(ValidateTxtBox.Text, Len(ValidateTxtBox.Text) - DotPos)) > (DecimalDigitCount - 1) Then
                    If ValidateTxtBox.SelStart > DotPos - 1 Then
                        KeyPressAscii = 0
                    End If
                End If
            ElseIf Len(ValidateTxtBox.Text) >= IntegerDigitCount Then
                KeyPressAscii = 0
            End If
        End If
    ElseIf ValidationMode = IntegerValue Then
        If InStr(1, "0123456789", Chr(KeyPressAscii)) = 0 Then
            KeyPressAscii = 0
        End If
    ElseIf ValidationMode = StrictAlphaNumeric Then
        If InStr(1, "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyPressAscii)) = 0 Then
            KeyPressAscii = 0
        End If
        If UseUpperCase Then KeyPressAscii = Asc(UCase(Chr(KeyPressAscii)))
        
    ElseIf ValidationMode = AlphaNumeric Then
        If InStr(1, "0123456789.abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ !#*_:/?-@=,", Chr(KeyPressAscii)) = 0 Then
        '[special charecters can be added to the above string if you wish to bypas that charecter]
            KeyPressAscii = 0
        End If
        
       If UseUpperCase Then KeyPressAscii = Asc(UCase(Chr(KeyPressAscii)))
    ElseIf ValidationMode = Other Then 'Used For SMS Txt Only
        If InStr(1, "0123456789.abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ !#%&*()+=-/.?,:;", Chr(KeyPressAscii)) = 0 Then
        '[special charecters can be added to the above string if you wish to bypas that charecter]
            KeyPressAscii = 0
        End If
        
       If UseUpperCase Then KeyPressAscii = Asc(UCase(Chr(KeyPressAscii)))
    Else
        ' You can add any type of validation here
    End If

Exit Sub
ErrorMod:
MsgBox "Error Occured due to " & err.Description, vbExclamation, gblstrPrjTitle

End Sub
Public Function DatabaseADOB_Connection() As Boolean
 On Error GoTo CatchError
    If Dir(App.Path & "\Pvt.mdb", vbNormal) = "" Then
        MsgBox "Database file not found!", vbExclamation
        
    Exit Function
    End If
    If adoc.State = 1 Then adoc.Close
    If adoc.State <> 1 Then
        adoc.Provider = "Microsoft.Jet.OLEDB.4.0"
         adoc.Properties("Jet OLEDB:Database Password") = "silbus"
        adoc.ConnectionString = App.Path & "\Pvt.mdb"
        adoc.Open
    End If
    
    If adoc.State = adStateOpen Then
        DatabaseADOB_Connection = True
    Else
         DatabaseADOB_Connection = False
    End If
    Exit Function
CatchError:
    MsgBox "Database Error! " & vbCrLf & "Error Number : " & err.Number & vbTab & "Description : " & err.Description, vbExclamation
End Function

Public Function EncodeStr64(encString As String, ByVal MaxPerLine As Integer) As String
' Return radix64 of string
Dim abOutput()  As Byte
Dim sLast       As String
Dim b(3)        As Byte
Dim j           As Integer
Dim CharCount   As Integer
Dim iIndex      As Long
Dim Umax        As Long
Dim I As Long, nLen As Long, nQuants As Long
EncodeStr64 = ""
nLen = Len(encString)
nQuants = nLen \ 3
iIndex = 0
If MaxPerLine < 10 Then MaxPerLine = 10
Umax = nQuants + 1
Call MakeEncTab
If (nQuants > 0) Then
    ReDim abOutput(nQuants * 4 - 1)
    For I = 0 To nQuants - 1
        For j = 0 To 2
            b(j) = Asc(Mid(encString, (I * 3) + j + 1, 1))
        Next
        Call EncodeQuantumB(b)
        abOutput(iIndex) = b(0)
        abOutput(iIndex + 1) = b(1)
        abOutput(iIndex + 2) = b(2)
        abOutput(iIndex + 3) = b(3)
        CharCount = CharCount + 4
        ' insert CRLF if max char per line is reached
        If CharCount >= MaxPerLine Then
            ReDim Preserve abOutput(UBound(abOutput) + 2)
            CharCount = 0
            abOutput(iIndex + 4) = 13
            abOutput(iIndex + 5) = 10
            iIndex = iIndex + 6
            Else
            iIndex = iIndex + 4
            End If
    Next
    EncodeStr64 = StrConv(abOutput, vbUnicode)
End If
Select Case nLen Mod 3
Case 0
    sLast = ""
Case 1
    b(0) = Asc(Mid(encString, nLen, 1))
    b(1) = 0
    b(2) = 0
    Call EncodeQuantumB(b)
    sLast = StrConv(b(), vbUnicode)
    sLast = Left(sLast, 2) & "=="
Case 2
    b(0) = Asc(Mid(encString, nLen - 1, 1))
    b(1) = Asc(Mid(encString, nLen, 1))
    b(2) = 0
    Call EncodeQuantumB(b)
    sLast = StrConv(b(), vbUnicode)
    sLast = Left(sLast, 3) & "="
End Select
EncodeStr64 = EncodeStr64 & sLast
End Function

Public Function DecodeStr64(decString As String) As String
' Return string of decoded values from radix64
Dim abDecoded() As Byte
Dim d(3)    As Byte
Dim c       As Integer
Dim di      As Integer
Dim I       As Long
Dim nLen    As Long
Dim iIndex  As Long
Dim Umax    As Long
nLen = Len(decString)
If nLen < 4 Then
    Exit Function
End If
ReDim abDecoded(((nLen \ 4) * 3) - 1)
Umax = nLen
iIndex = 0
di = 0
Call MakeDecTab
For I = 1 To Len(decString)
    c = CByte(Asc(Mid(decString, I, 1)))
    c = aDecTab(c)
    If c >= 0 Then
        d(di) = CByte(c)
        di = di + 1
        If di = 4 Then
            abDecoded(iIndex) = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL6(d(2) And &H3) Or d(3)
            iIndex = iIndex + 1
            If d(3) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            If d(2) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            di = 0
        End If
    End If
Next I
DecodeStr64 = StrConv(abDecoded(), vbUnicode)
DecodeStr64 = Left(DecodeStr64, iIndex)
End Function

Private Sub EncodeQuantumB(b() As Byte)
Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
b0 = SHR2(b(0)) And &H3F
b1 = SHL4(b(0) And &H3) Or (SHR4(b(1)) And &HF)
b2 = SHL2(b(1) And &HF) Or (SHR6(b(2)) And &H3)
b3 = b(2) And &H3F
b(0) = aEncTab(b0)
b(1) = aEncTab(b1)
b(2) = aEncTab(b2)
b(3) = aEncTab(b3)
End Sub

Private Function MakeDecTab()
Dim t As Integer
Dim c As Integer
For c = 0 To 255
    aDecTab(c) = -1
Next
t = 0
For c = Asc("A") To Asc("Z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("a") To Asc("z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("0") To Asc("9")
    aDecTab(c) = t
    t = t + 1
Next
c = Asc("+")
aDecTab(c) = t
t = t + 1
c = Asc("/")
aDecTab(c) = t
t = t + 1
c = Asc("=")
aDecTab(c) = t
End Function

Private Function MakeEncTab()
Dim I As Integer
Dim c As Integer
I = 0
For c = Asc("A") To Asc("Z")
    aEncTab(I) = c
    I = I + 1
Next
For c = Asc("a") To Asc("z")
    aEncTab(I) = c
    I = I + 1
Next
For c = Asc("0") To Asc("9")
    aEncTab(I) = c
    I = I + 1
Next
c = Asc("+")
aEncTab(I) = c
I = I + 1
c = Asc("/")
aEncTab(I) = c
I = I + 1
End Function

Private Function SHL2(ByVal bytValue As Byte) As Byte
SHL2 = (bytValue * &H4) And &HFF
End Function

Private Function SHL4(ByVal bytValue As Byte) As Byte
SHL4 = (bytValue * &H10) And &HFF
End Function

Private Function SHL6(ByVal bytValue As Byte) As Byte
SHL6 = (bytValue * &H40) And &HFF
End Function

Private Function SHR2(ByVal bytValue As Byte) As Byte
SHR2 = bytValue \ &H4
End Function

Private Function SHR4(ByVal bytValue As Byte) As Byte
SHR4 = bytValue \ &H10
End Function

Private Function SHR6(ByVal bytValue As Byte) As Byte
SHR6 = bytValue \ &H40
End Function

' ------------------------------------------------------------
'    Progress Bar Picture sub (please adjust to your program code)
' ------------------------------------------------------------

Public Sub UpdateStatus(ByVal sngPercent As Single)
'With FormDemo.picProgress
'If sngPercent > 100 Then sngPercent = 100
'If sngPercent = 0 Then .Cls: Exit Sub
'.DrawMode = 13
'FormDemo.picProgress.Line (-10, -10)-(sngPercent, .Height + 75), .ForeColor, BF
'.Refresh
'End With
End Sub

Public Function CheckVersion(USB_Path As String) As Boolean
On Error GoTo erromod
Dim VerR_position, freefileInt As Integer
Dim version_Obj As VERSION_TYPE
    freefileInt = FreeFile
    If Dir(USB_Path & "\VERSION.DAT") <> "" Then
            freefileInt = FreeFile
            Open USB_Path & "\VERSION.DAT" For Binary Access Read As #freefileInt
            Get freefileInt, , version_Obj
            Close freefileInt
            VerR_position = InStrRev(TrimChr(version_Obj.Version), "R")
            If Left(TrimChr(version_Obj.Version), VerR_position - 1) = "PVT_GEN_12" Then
                 CheckVersion = True
            Else
                 CheckVersion = False
            End If
    End If
Exit Function
erromod:
    MsgBox "Error due to " & err.Description, vbInformation
End Function
