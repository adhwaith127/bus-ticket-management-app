Attribute VB_Name = "MDLHindi"
Dim sltxt As String * 24

Private Type PORTSETUP
    Port As Byte 'String * 1
    baud As String * 6
End Type

Private PSetup As PORTSETUP

Public Sub connect()
    If dbcon.State <> 1 Then
        dbcon.Provider = "Microsoft.jet.oledb.4.0"
        dbcon.Properties("Jet OLEDB:Database Password") = "siljvvnl"
        dbcon.ConnectionString = App.Path & "\JVVNL.mdb"
        dbcon.Open
    End If
End Sub

Public Sub sndPrintH(PStr As String)
  Dim I As Integer
  Dim Cvstr As String * 2
  Dim sBt As Byte
   For I = 1 To Len(PStr) Step 2
     Cvstr = Mid$(PStr, I, 2)
     sBt = StrToHex(Cvstr)
'     'SendByte sBt
     DWait (ToutDly)
   Next I
'     'SendByte &HD
     DWait (ToutDly)
End Sub

Public Function StrToHex(Cvstr As String) As Byte
  Dim msB As String * 1
  Dim lsB As String * 1
  msB = Mid$(Cvstr, 1, 1)
  lsB = Mid$(Cvstr, 2, 1)
    Select Case msB
        Case 0: StrToHex = 16 * 0
        Case 1: StrToHex = 16 * 1
        Case 2: StrToHex = 16 * 2
        Case 3: StrToHex = 16 * 3
        Case 4: StrToHex = 16 * 4
        Case 5: StrToHex = 16 * 5
        Case 6: StrToHex = 16 * 6
        Case 7: StrToHex = 16 * 7
        Case 8: StrToHex = 16 * 8
        Case 9: StrToHex = 16 * 9
        Case "A": StrToHex = 16 * 10
        Case "B": StrToHex = 16 * 11
        Case "C": StrToHex = 16 * 12
        Case "D": StrToHex = 16 * 13
        Case "E": StrToHex = 16 * 14
        Case "F": StrToHex = 16 * 15
    End Select
    Select Case lsB
        Case 0: StrToHex = StrToHex + 0
        Case 1: StrToHex = StrToHex + 1
        Case 2: StrToHex = StrToHex + 2
        Case 3: StrToHex = StrToHex + 3
        Case 4: StrToHex = StrToHex + 4
        Case 5: StrToHex = StrToHex + 5
        Case 6: StrToHex = StrToHex + 6
        Case 7: StrToHex = StrToHex + 7
        Case 8: StrToHex = StrToHex + 8
        Case 9: StrToHex = StrToHex + 9
        Case "A": StrToHex = StrToHex + 10
        Case "B": StrToHex = StrToHex + 11
        Case "C": StrToHex = StrToHex + 12
        Case "D": StrToHex = StrToHex + 13
        Case "E": StrToHex = StrToHex + 14
        Case "F": StrToHex = StrToHex + 15
    End Select
End Function

Public Sub sndPrintE(PStr As String)
  Dim I As Integer
  Dim Cvstr As String * 2
  Dim sBt As Byte
   For I = 1 To Len(PStr) Step 1
     Cvstr = Mid$(PStr, I, 1)
     sBt = Asc(Cvstr)
'     'SendByte sBt
     DWait (ToutDly)
   Next I
'   'SendByte &HD
   DWait (ToutDly)
End Sub

Public Function initconn() As Boolean
'Dim Filehdl As Integer
'''CCOUNT = 1
'Set SerialComm = H_Convert.MSComm1
'Filehdl = FreeFile()
'Open App.Path & "\TRANS.DAT" For Binary Access Read Write As #Filehdl
'    Get #Filehdl, , PSetup
'    Close #Filehdl
'    If H_Convert.MSComm1.PortOpen = True Then
'        H_Convert.MSComm1.PortOpen = False
'    End If
'    If Not InitPort(val(PSetup.Port), PSetup.baud) Then
'        If TransMsg <> "" Then
'           MsgBox TransMsg
'        End If
'           initconn = False
'       Exit Function
'    Else
'           initconn = True
'           H_Convert.TextP.Text = PSetup.Port
'    End If
'    'MsgBox "Connection is OK", vbExclamation, "KOT"
End Function

Public Sub DWait(dly As Integer)
  Dim Tin As Double
  Dim Tout As Double
    Tin = Timer
    Do While (1)
        Tout = Timer
        If Tout - Tin > dly Then Exit Do
    Loop
End Sub
