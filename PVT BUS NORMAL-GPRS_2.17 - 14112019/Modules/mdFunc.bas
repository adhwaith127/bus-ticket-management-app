Attribute VB_Name = "mdFunc"

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'Public shlShell As Shell32.Shell
'Public shlFolder As Shell32.Folder
'Public Const BIF_RETURNONLYFSDIRS = &H1

Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
 
Public Const LOCALE_SLONGDATE = &H20        '  long date format string
Public Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_SETTINGCHANGE = &H1A
Public Const WM_CLOSE = &H10


Public SYSDIR As String
Public WINPATH As String


Public PrjValidity As PROJECT_VALIDITY
Public Const NOOFDAYS_VALIDITY = 30

Public Const SRILANKA = 1
Public Const KERALA = 2

Public Const PROJECT = SRILANKA
'Public Const PROJECT = KERALA
'Public Const LANGAGE_ENABLED = 1   '''vaisakh 30.03.11


Public SeletedRoute() As String
Public SeletedRouteCount As Integer

Public adminflag As Boolean
Public Frmflag As Boolean
'Public Function TrimChr(ByVal Buf As String) As String
'    Dim tmp As Integer
'    Dim Buf1 As String
'    tmp = InStr(Buf, Chr$(0))
'    If tmp > 0 Then
'        TrimChr = Trim(Mid(Buf, 1, tmp - 1))
'    Else
'        TrimChr = Trim(Buf)
'    End If
'    'TrimChr = left(Buf, InStr(1, strString, Chr$(0)) - 1) 'Mid(Buf, 1, (InStr(Buf, Chr(0)) - 1))
'End Function
Public Function CreateCurrency() As Boolean
On Error GoTo err
Dim strSend As String * 16
Dim hndl1 As Integer
Dim hndl2 As Integer
Dim slen As Byte
    hndl1 = FreeFile()
    CreateCurrency = True
    Open App.Path & "\2.txt" For Input As #hndl1
    hndl2 = FreeFile()
    If Dir(App.Path & "\currency.dat") <> "" Then
        Kill App.Path & "\currency.dat"
    End If
    Open App.Path & "\currency.dat" For Binary Access Write As #hndl2
    Do While Not EOF(hndl1)
        Input #hndl1, strSend
        slen = Len(strSend)
        If slen < 16 Then
            For I = 1 To 15 - slen
                strSend = strSend & " "
            Next
        ElseIf slen > 16 Then
            strSend = Mid$(strSend, 1, 16)
        End If
        strSend = strSend & Chr(0)
        Put #hndl2, , strSend
    Loop
    Close #hndl1
    Close #hndl2
Exit Function
err:
MsgBox "Error in Currency Conversion" & vbCrLf & err.Number & " , " & err.Description
CreateCurrency = False
Exit Function
End Function
Public Function CheckTableExistsOrNot(strTableName As String) As Boolean
On Error GoTo err
    
    Dim strSql As String
    Dim DB As New ADODB.Connection
    Dim rs1 As ADODB.Recordset
    'CONNECTDB
    sDataBase = App.Path & "\PVT.MDB"
    CheckTableExistsOrNot = False
    
    If DB.State <> 0 Then DB.Close
        DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source= " & sDataBase & "; Jet OLEDB:Database Password =silbus"
    Set rs1 = New ADODB.Recordset
    
    strSql = "SELECT * FROM " & strTableName
    If rs1.State <> 0 Then rs1.Close
    
     rs1.Open strSql, DB, adOpenDynamic, adLockOptimistic
     rs1.Close
    CheckTableExistsOrNot = True
    Exit Function
err:
    If rs1.State <> 0 Then rs1.Close
    
    CheckTableExistsOrNot = False
End Function

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
Dim iNull As Integer
Dim lpIDList As Long
Dim lResult As Long
Dim sPath As String
Dim udtBI As BrowseInfo
On Error Resume Next
    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With
     lpIDList = SHBrowseForFolder(udtBI)
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If
     BrowseForFolder = sPath
End Function


Public Sub CenterForm(pobjForm As Form)
 On Error Resume Next
   With pobjForm
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub

Public Sub CreateErrorLog(ErrorMessage As String)
On Error Resume Next
Dim Handle As Integer
    Handle = FreeFile
    If Dir(App.Path & "\Error Log.txt") = "" Then
        Open App.Path & "\Error Log.txt" For Output As #Handle
    Else
        Open App.Path & "\Error Log.txt" For Append As #Handle
    End If
    ErrorMessage = "----------------------------------------------------------------------" & vbCrLf & _
                   Day(Date) & "/" & Month(Date) & "/" & Year(Date) & " - " & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & vbCrLf & _
                   ErrorMessage & vbCrLf & _
                   "----------------------------------------------------------------------" & vbCrLf
    Print #Handle, ErrorMessage
    Close #Handle
End Sub

Public Function LoadFooter() As String
Dim strBuffer As String * 250
Dim strpath As String
Dim intCheckSum As Integer
Dim intHandle As Integer
Dim iHdl As Integer
Dim strFooter As String
Dim strEncData As String
    
    strFooter = ""
    LoadFooter = ""
    iHdl = FreeFile()
    
    Open App.Path & "\patch.pth" For Input As #iHdl
        If EOF(iHdl) = True Then Exit Function
        Input #iHdl, strEncData
    Close #iHdl
    
    strFooter = PROBAdecodeString(DecodeStr64(strEncData), "ajsndd", True)
    
    
'    Dim i As Integer
'    For i = 0 To Len(strFooter) - 1
'        Debug.Print "asc: " & Asc(Mid(strFooter, i + 1, 1)) & " - " & Mid(strFooter, i + 1, 1)
'        If Asc(Mid(strFooter, i + 1, 1)) > 122 Then
'            Exit Function
'        End If
'    Next
'
    If InStr(strFooter, "SOFTLAND") = 0 Then
         strFooter = ""
         Exit Function
    End If
    LoadFooter = strFooter
End Function

