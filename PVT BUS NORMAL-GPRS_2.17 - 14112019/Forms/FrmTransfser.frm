VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "FrmTransfser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBoud 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   15
      Top             =   3885
      Width           =   1125
   End
   Begin VB.TextBox txtPort 
      Enabled         =   0   'False
      Height          =   315
      Left            =   165
      TabIndex        =   14
      Top             =   3870
      Width           =   810
   End
   Begin VB.CommandButton cmdDownloadRoute 
      Caption         =   "Download Route"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2955
      TabIndex        =   13
      Top             =   2595
      Width           =   1680
   End
   Begin VB.CommandButton cmdCrew 
      Caption         =   "&Download Crew Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2895
      TabIndex        =   12
      Top             =   2100
      Width           =   1680
   End
   Begin VB.CommandButton cmdSndSetup 
      Caption         =   "&Download Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2895
      TabIndex        =   11
      Top             =   1635
      Width           =   1680
   End
   Begin VB.FileListBox FileContainer 
      Height          =   3015
      Left            =   5385
      Pattern         =   "*.dat;*.lst"
      TabIndex        =   10
      Top             =   180
      Width           =   2235
   End
   Begin VB.CommandButton CmdDownload 
      Caption         =   "&PC to PMTC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2895
      TabIndex        =   2
      Top             =   1185
      Width           =   1680
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   9
      Top             =   4935
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   "Idle"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm SerialCom 
      Left            =   0
      Top             =   4335
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton CmdUpload 
      Caption         =   "P&MTC to PC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2895
      TabIndex        =   1
      Top             =   720
      Width           =   1680
   End
   Begin VB.ListBox LstPalmTec 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3210
      Left            =   165
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   180
      Width           =   2415
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   255
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   2670
      TabIndex        =   5
      Top             =   90
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Boud Rate"
      Height          =   195
      Left            =   1230
      TabIndex        =   17
      Top             =   3630
      Width           =   780
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   210
      Left            =   180
      TabIndex        =   16
      Top             =   3615
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   2730
      X2              =   5250
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      X1              =   2730
      X2              =   5295
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   2715
      X2              =   5235
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2715
      X2              =   5235
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Convert Utility"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   3300
      TabIndex        =   8
      Top             =   3420
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "and"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3885
      TabIndex        =   7
      Top             =   3195
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Transfer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   3375
      TabIndex        =   6
      Top             =   2925
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   2625
      Top             =   2685
      Width           =   765
   End
   Begin VB.Label LblMsg 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   4485
      Width           =   45
   End
End
Attribute VB_Name = "FrmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Type PORTSETUP
'    Port As Integer
'    baud As String * 6
'End Type
'Dim fso As New FileSystemObject
'Dim PSetup As PORTSETUP
'Dim FileRpt As Integer
'Dim i As Single
'Dim j As Integer
'Dim msg As String
'Dim SysDate As Date
'Dim DiffDate As Date
'Dim gd As DATETYPE
'Dim gt As TIMETYPE
'Dim Pathname As String
'Dim FileName As String
'
'
'Private Sub cmdDelete_Click() 'To delete Files
'On Error GoTo ErrLoc
'LblMsg = ""
'LblMsg.Caption = ""
'FileName = Left(LstPalmTec.List(LstPalmTec.ListIndex), 8)
'FileName = Trim(FileName)
'FileName = FileName & ".DAT"
'If FileName = "" Then
'    MsgBox "No File Selected", vbInformation, "BUS Data Transfer"
'    Exit Sub
'End If
'Trans (" 14 " & FileName)
'LblMsg.Caption = "Successfully Deleted File..."
'CmdRefresh_Click
'Exit Sub
'ErrLoc:
'    MsgBox err.Description, vbCritical, "BUS Data Transfer"
'End Sub
'
'Private Sub cmdCrew_Click()
'Dim FName As String
'
'        FName = App.Path & "\CREW.DAT"
'        If Dir(FName) <> "" Then
'            Disableall
'            If Trans(" 2 CREW.DAT") = True Then
'                LblMsg.Caption = "CREW Details  Downloaded "
'                MsgBox "CREW Details  Download successfully"
'                Enableall
'                Exit Sub
'            Else
'                MsgBox TransMsg
'                Exit Sub
'            End If
'        Else
'            MsgBox FName & vbCrLf & "File Not Found"
'            Exit Sub
'        End If
'
'End Sub
'
'Private Sub cmdDownload_Click()
'
'On Error GoTo ErrLoc
'Dim i As Integer
'    Disableall
'    CreateSchedule
'
'
'    FileContainer.Path = App.Path
'    FileContainer.Refresh
'
'    If Dir$(App.Path & "\*.LST") = "" Then
'         MsgBox "LST Files Not Found!", vbInformation
'         Enableall
'         Exit Sub
'    End If
'    If Dir$(App.Path & "\*.DAT") = "" Then
'         MsgBox "DAT Files Not Found!", vbInformation
'         Enableall
'         Exit Sub
'    End If
'    FileName = ""
'    'newly added by deej for downloading currency file
'    'added on 30.04.07
'    '-------------------------------
'    If CreateCurrency = True Then
'         FileName = "currency.dat"
'         LblMsg.Caption = FileName & " Downloading..."
'         If Trans(" 2 " & FileName) = False Then
'            MsgBox TransMsg, vbCritical
'            Enableall
'            Exit Sub
'         End If
'    Else
'        MsgBox "Error in creating Currency file"
'        Enableall
'        Exit Sub
'    End If
'    '------------------------
''***********DOWNLOADING TO PALMTEC**********************
''   For i = 0 To FileContainer.ListCount - 1   ' Loop through File list.
''      If FileContainer.List(i) <> "TRANS.DAT" And FileContainer.List(i) <> "DIR.LST" Then  'FileContainer.List(i) <> "DATE.DAT" And
''         FileName = FileContainer.List(i)
''         If Trans(" 2 " & FileName) = False Then
''            MsgBox TransMsg, vbCritical
''            Exit Sub
''         End If
''         LblMsg.Caption = FileName & " Downloading..."
''      End If
''   Next i
''*******************************************************
'
''***********DOWNLOADING TO PALMTEC**********************
''   For i = 0 To UBound(SeletedRoute) 'SeletedRouteCount - 1
''      If SeletedRoute(i) <> "" Then
'         FileName = "RTE.DAT" 'SeletedRoute(i) & ".Dat"
'         LblMsg.Caption = FileName & " Downloading..."
'         If Trans(" 2 " & FileName) = False Then
'            MsgBox TransMsg, vbCritical
'            Enableall
'            Exit Sub
'         End If
''      End If
''   Next i
'
'    If Dir(App.Path & "\Crew.dat") <> "" Then
'        LblMsg.Caption = "Crew.dat" & " Downloading..."
'        If Trans(" 2 " & "Crew.dat") = False Then
'           MsgBox TransMsg, vbCritical
'           Enableall
'           Exit Sub
'        End If
'    Else
'        MsgBox "Crew.Dat Not Found !", vbInformation
'    End If
'
'    If Dir(App.Path & "\RouteLst.LST") <> "" Then
'        LblMsg.Caption = "RouteLst.LST" & " Downloading..."
'        If Trans(" 2 " & "RouteLst.LST") = False Then
'           MsgBox TransMsg, vbCritical
'           Enableall
'           Exit Sub
'        End If
'    Else
'        MsgBox "RouteLst.LST Not Found !", vbInformation
'    End If
'
'    If Dir(App.Path & "\Stage.lst") <> "" Then
'        LblMsg.Caption = "Stage.lst" & " Downloading..."
'        If Trans(" 2 " & "Stage.lst") = False Then
'           MsgBox TransMsg, vbCritical
'           Enableall
'           Exit Sub
'        End If
'    Else
'        MsgBox "Satge.lst Not Found !", vbInformation
'    End If
'
'    If Dir(App.Path & "\Bus.dat") <> "" Then
'        LblMsg.Caption = "Bus.dat" & " Downloading..."
'        If Trans(" 2 " & "Bus.dat") = False Then
'           MsgBox TransMsg, vbCritical
'           Enableall
'           Exit Sub
'        End If
'    Else
'        MsgBox "Bus.dat Not Found !", vbInformation
'    End If
'
''*******************************************************
'
'
'
''sleep (800)
'delayTr 8
'FileName = ""
'StatusBar1.SimpleText = ""
'LblMsg.Caption = "Successfully Downloaded"
'MsgBox "Successfully Downloaded"
'Enableall
'Exit Sub
'
'ErrLoc:
'    MsgBox err.Description, vbCritical
'    CmdDownload.Enabled = True
'    CmdRefresh.Enabled = True
'    CmdUpload.Enabled = True
'    Enableall
'    Exit Sub
'
'End Sub
'
'Private Sub cmdDownloadRoute_Click()
'    FrmSelectRoute.Show vbModal
'End Sub
'
'Private Sub CmdRefresh_Click()
' On Error GoTo ERR1
'    CmdRefresh.Enabled = False
'
'    Dim S As String
'    Dim i, Filehdl As Integer
'    If Dir$(App.Path & "\Dir.lst") <> "" Then
'      Kill App.Path & "\Dir.lst"
'    End If
'    Disableall
'    Filehdl = FreeFile()
'    LblMsg.Caption = "Collecting File List in Palmtec"
'    'Trans ("3")
'    If Not Trans("3") Then 'UPloading Dir.Lst from palmtec
'        LblMsg.Caption = "No Device - Old File List"
'    Else
'        LblMsg.Caption = "Palmtec File List"
'    End If
'
'    If Dir$(TransPath & "\Dir.lst") <> "" Then
'        Open App.Path & "\Dir.lst" For Input As #Filehdl
'        LstPalmTec.Clear
'        Do While Not EOF(Filehdl)
'            Line Input #Filehdl, S
'            If UCase$(Trim$(Mid$(S, 10, 3))) = "DAT" And UCase$(Trim$(Mid$(S, 1, 3))) = "RPT" Then
'               S = Mid(S, 1, InStr(1, S, "-") - 1)
'               LstPalmTec.AddItem (Trim(Mid$(S, 1, 5)) & ".DAT")
'            End If
''            If UCase$(Trim$(Mid$(s, 10, 3))) = "DAT" And (UCase$(Trim$(Mid$(s, 1, 3))) = "RPT" Or UCase$(Trim$(Mid$(s, 1, 4))) = "TKTS") Then
''                If UCase$(Trim$(Mid$(s, 1, 3))) = "RPT" Then
''                 pmtcID = Mid(s, 4, 2)
''                End If
'                'By for Bye > 8 for uploading ticket files
'
''                Dim by As Integer
''                by = Mid(s, 14, InStr(1, s, "-") + 6)
''               If by > 8 Then
''                 LstPalmTec.AddItem (Trim(Mid$(s, 1, 6)) & ".DAT")
''               End If
' '           End If
'        Loop
'        Close #Filehdl
'    End If
'    CmdRefresh.Enabled = True
'    CmdUpload.Enabled = True
'    CmdDownload.Enabled = True
'    Enableall
'    Exit Sub
'ERR1:
'    MsgBox err.Description
'    Close #Filehdl
'    CmdRefresh.Enabled = True
'    Enableall
'End Sub
'
'Private Sub cmdSndSetup_Click()
'Dim FName As String
'    FName = App.Path & "\BUS.DAT"
'    If Dir(FName) <> "" Then
'    Disableall
'        If Trans(" 2 BUS.DAT") = True Then
'            LblMsg.Caption = "Setup  Download successfully"
'            MsgBox "Setup  Download successfully"
'            Enableall
'            Exit Sub
'        Else
'            MsgBox TransMsg
'            Exit Sub
'        End If
'    Else
'        MsgBox FName & vbCrLf & "File Not Found"
'        Enableall
'        Exit Sub
'    End If
'End Sub
'
'Private Sub CmdUpload_Click()
'On Error GoTo ErrLoc
'
'    Dim n As Integer
'    Dim num As Integer
'    Dim TktFilename As String
'
'    Dim TODATE As String
'    Dim temp As String
'    Dim tmpFileName As String
'    Dim FirstChk As Integer
'    Dim che As String
'    Dim checkfile As Integer
'    Dim DirRPT As Integer
'    Dim strDir As String
'    Dim l As Integer
'            l = 1
'    TODATE = Format(Date, "ddmmyy")
'    num = 1
'    FirstChk = 0
'    CmdUpload.Enabled = False
'    CmdDownload.Enabled = False
'    CmdRefresh.Enabled = False
'    On Error GoTo ErrLoc
'    LblMsg = ""
'    LblMsg.Caption = ""
'    tmpFileName = ""
'    FileName = ""
'    FileName = LstPalmTec.List(LstPalmTec.ListIndex)
'    pmtcID = Mid(FileName, 4, 2)
'
'    If Dir$(App.Path & "\" & FileName) <> "" Then
'          Kill App.Path & "\" & FileName
'    End If
'
'    If FileName = "" Then
'        MsgBox "No File Selected", vbInformation, "BUS Data Transfer"
'        Exit Sub
'    End If
'
'    Disableall
'    If Trans(" 1 " & FileName) = False Then
'        MsgBox TransMsg, vbCritical
'        Exit Sub
'    End If
'
'    If Trans(" 1 " & "Bus.DAT") = False Then 'Bus.DAT for writing header
'        MsgBox TransMsg, vbCritical
'        Exit Sub
'    End If
'
'    ChDir App.Path
'
'
'    'To upload the Files and converts
'     If RptConvert(FileName) = True Then
'                    LblMsg.Caption = FileName & "File converted..."
'     Else
'                    LblMsg.Caption = FileName & "Not converted..."
'     End If
'  '   DBRpt (FileName) ' inserting Rpt*.dat to Access 30/03/06
'     'RPT01.DAT
''     Sleep (500)
''     str1 = App.Path & "\BUSCONV " & FileName         'RPT01.DAT
''      Sleep (100)
''     ret = WinExec(str1, vbHide) 'BUSCONV.EXE is calling
''     Sleep (500)
''
''            Do While 1
''               str1 = FindWindow(vbNullString, "BUSCONV")
''               If str1 = "0" Then
''                  Exit Do
''               End If
''               DoEvents
''            Loop
''
'    OneLine = ""
'    FileDateCret = ""
'    'If fso.FileExists(App.Path & "\TICKET.TXT") = True Then
'    FileRpt = FreeFile()
'    'Open App.Path & "\TICKET.TXT" For Binary Access Read Write As #FileRpt
'    Open App.Path & "\TICKET.TXT" For Input As #FileRpt
'         'Do While (Not EOF(FileRpt))
'         While (Not EOF(FileRpt))
'            'Line Input #FileRpt, OneLine
'            Input #FileRpt, OneLine
'            DirRPT = FreeFile()
'            Open App.Path & "\Dir.LST" For Input As #DirRPT
'            Do While Not (EOF(DirRPT))
'            Input #DirRPT, strDir
'               If Trim(Mid(strDir, 1, 6)) = Trim(Mid(OneLine, 1, 6)) Then
'                 TktFilename = Trim(Mid(OneLine, 1, 10))
'
'            scduID = Trim(Mid(TktFilename, 5, 2))
'            'TODATE = Trim(Mid(OneLine, 14, 8))    '15/01/00
'            'temp = Mid(TODATE, 1, 2) & Mid(TODATE, 4, 2) & Mid(TODATE, 7, 2)
'            FileDateCret = TODATE 'temp
'            For i = 1 To LstPalmTec.ListCount
'            Next i
'            che = LstPalmTec.List(l)
'            If che <> TktFilename Then
'               checkfile = 1
'            End If
'    '        If tmpFileName <> FileDateCret Then
'    '           FirstChk = 0
'    '        End If
'
'           ' tmpFileName = FileDateCret
'            'Folder creation
'            If Dir$(App.Path & "\" & FileDateCret, vbDirectory) <> "" Then
'                If Dir$(App.Path & "\" & FileDateCret & "\" & pmtcID, vbDirectory) = "" Then
'                   MkDir App.Path & "\" & FileDateCret & "\" & pmtcID
'                   MkDir App.Path & "\" & FileDateCret & "\" & pmtcID & "\TripDetails"
'                   FirstChk = 1
'                End If
'            Else
'                MkDir App.Path & "\" & FileDateCret
'                If Dir$(App.Path & "\" & FileDateCret & "\" & pmtcID, vbDirectory) = "" Then
'                   MkDir App.Path & "\" & FileDateCret & "\" & pmtcID
'                   MkDir App.Path & "\" & FileDateCret & "\" & pmtcID & "\TripDetails"
'                   FirstChk = 1
'                End If
'            End If
'            FirstChk = 1
'            If checkfile = 1 Then  'inner if
'               If Trans(" 1 " & TktFilename) = False Then
'                  MsgBox TransMsg, vbCritical
'                  CmdDownload.Enabled = True
'                  CmdUpload.Enabled = True
'                  CmdRefresh.Enabled = True
'                  Close FileRpt
'                  Exit Sub
'               End If
'
'               LblMsg.Caption = TktFilename & "File Uploaded..."
'    '           Sleep (800)
'    '        If Mid(FileName, 1, 4) = "TKTS" Then
'              If TktsConvert(TktFilename) = True Then
'               LblMsg.Caption = TktFilename & " File converted..."
'              Else
'               LblMsg.Caption = TktFilename & "File Not converted..."
'              End If
''               Sleep (500)
''               str1 = App.Path & "\BUSCONV " & TktFilename       'TKTS01.DAT
''                Sleep (100)
''               ret = WinExec(str1, vbHide) 'BUSCONV.EXE is calling
''               Sleep (500)
''               Do While 1
''                  str1 = FindWindow(vbNullString, "BUSCONV")
''                  If str1 = "0" Then
''                     Exit Do
''                  End If
''                  DoEvents
''               Loop
'           '  End If
'    '      If Mid(FileName, 1, 4) = "TKTS" Then
'    '         scduID = Trim(Mid(FileName, 5, 2))
'    '         'pmtcid loaded in form
'
'              If ReportCnv(pmtcID, scduID) = True Then
'                LblMsg.Caption = FileName & "File Converted..."
'              Else
'                LblMsg.Caption = FileName & "File Not Converted..."
'              End If
''              Sleep (500)
''               str1 = App.Path & "\BUSCONV REPORT " & pmtcID & "-" & scduID  'REPORT 01-02
''                Sleep (100)
''               ret = WinExec(str1, vbHide)
''               LblMsg.Caption = FileName & "File Converted..."
''               Sleep (500)
''
''               Do While 1
''                  str1 = FindWindow(vbNullString, "BUSCONV")
''                  If str1 = "0" Then
''                     Exit Do
''                  End If
''                  DoEvents
''               Loop
'              'End If
'     '         DBTKTS (TktFilename) ' for updating TKTS table 30/03/06
'                If Dir(App.Path & "\TKTS" & scduID & ".CSV") <> "" Then
'                  If Dir(App.Path & "\" & FileDateCret & "\" & pmtcID & "\TKTS" & scduID & ".CSV") = "" Then
'                     fso.MoveFile App.Path & "\TKTS" & scduID & ".CSV", App.Path & "\" & FileDateCret & "\" & pmtcID & "\"
'                  Else
'                     fso.DeleteFile App.Path & "\" & FileDateCret & "\" & pmtcID & "\TKTS" & scduID & ".CSV"
'                     fso.MoveFile App.Path & "\TKTS" & scduID & ".CSV", App.Path & "\" & FileDateCret & "\" & pmtcID & "\"
'                  End If
'                End If
'
'                If Dir(App.Path & "\REPORT" & scduID & ".TXT") <> "" Then
'                  If Dir(App.Path & "\" & FileDateCret & "\" & pmtcID & "\REPORT" & scduID & ".TXT") = "" Then
'                     fso.MoveFile App.Path & "\REPORT" & scduID & ".TXT", App.Path & "\" & FileDateCret & "\" & pmtcID & "\"
'                  Else
'                     fso.DeleteFile App.Path & "\" & FileDateCret & "\" & pmtcID & "\REPORT" & scduID & ".TXT"
'                     fso.MoveFile App.Path & "\REPORT" & scduID & ".TXT", App.Path & "\" & FileDateCret & "\" & pmtcID & "\"
'                  End If
'                End If
'                If Dir(App.Path & "\Rpt" & pmtcID & ".CSV") <> "" Then
'                  If Dir(App.Path & "\" & FileDateCret & "\" & pmtcID & "\RPT" & pmtcID & ".CSV") = "" Then
'                     fso.MoveFile App.Path & "\Rpt" & pmtcID & ".CSV", App.Path & "\" & FileDateCret & "\" & pmtcID & "\"
'                  Else
'                     fso.DeleteFile App.Path & "\" & FileDateCret & "\" & pmtcID & "\Rpt" & pmtcID & ".CSV"
'                     fso.MoveFile App.Path & "\Rpt" & pmtcID & ".CSV", App.Path & "\" & FileDateCret & "\" & pmtcID & "\"
'                  End If
'                End If
'
'            End If 'inner if
'         '  If LstPalmTec.ListIndex > 2 Then
'         '  l = l + 1
'              End If
'            Loop
'            Close #DirRPT
'
'         Wend
'      Close #FileRpt
'    'End If
'    'Ends Here
'
'    If Dir(App.Path & "\TICKET.TXT") <> "" Then
'       Kill App.Path & "\TICKET.TXT"
'    End If
'    If Dir(App.Path & "\RPT*.DAT") <> "" Then
'       Kill App.Path & "\RPT*.DAT"
'    End If
''     If Dir(App.Path & "\Bus.DAT") <> "" Then
''       Kill App.Path & "\Bus.DAT"
''    End If
'    If Dir(App.Path & "\TKTS*.DAT") <> "" Then
'       Kill App.Path & "\TKTS*.DAT"
'    End If
'
'    LblMsg.Caption = "All Files are Uploaded"
'    MsgBox "All Files are Uploaded"
'    CmdDownload.Enabled = True
'    CmdUpload.Enabled = True
'    CmdRefresh.Enabled = True
'    Enableall
'    Exit Sub
'ErrLoc:
'    MsgBox err.Description, vbCritical, "BUS Data Transfer"
'    CmdUpload.Enabled = True
'    CmdDownload.Enabled = True
'    CmdRefresh.Enabled = True
'    Close #FileRpt
'    Close #DirRPT
'    Enableall
'    Exit Sub
'End Sub
'Private Sub Command1_Click()
''Dim OneLine As String
''OneLine = ""
''FileRpt = FreeFile()
''Open App.Path & "\TICKET.TXT" For Input As #FileRpt
''     'Do While (Not EOF(FileRpt))
''     While (Not EOF(FileRpt))
''         Input #FileRpt, OneLine
''     Wend
''     'Loop
''Close #FileRpt
'End Sub
'Sub Disableall()
'    CmdUpload.Enabled = False
'    CmdRefresh.Enabled = False
'    CmdDownload.Enabled = False
'    cmdSndSetup.Enabled = False
'    cmdCrew.Enabled = False
'    cmdDownloadRoute.Enabled = False
'
'End Sub
'Sub Enableall()
'    CmdUpload.Enabled = True
'    CmdRefresh.Enabled = True
'    CmdDownload.Enabled = True
'    cmdSndSetup.Enabled = True
'    cmdCrew.Enabled = True
'    cmdDownloadRoute.Enabled = True
'End Sub
'Private Sub Form_Activate()
'Dim Filehdl As Integer
'
'    CmdDownload.Enabled = True
'  '  CmdUpload.Enabled = True
'    cmdDownloadRoute.Left = 2895
'    cmdDownloadRoute.Top = 1185
'End Sub
'Private Sub Form_Load()
'Dim Filehdl As Integer
'
'  'CmdUpload.Enabled = False
'  Set fso = CreateObject("Scripting.FileSystemObject")
'  FileContainer.Path = App.Path
'  ArryCount = 0
'
'  If Dir(App.Path & "\TRANS.DAT") <> "" Then
'   Kill App.Path & "\TRANS.DAT"
'  End If
'    CONNECT_DB
'    Set res = DB.OpenRecordset("PORT", dbOpenDynaset)
'    If res.RecordCount = 0 Then
'        PortNo = 1
'    Else
'        PortNo = res!Port
'    End If
'    If Dir$(TransPath & "\TRANS.DAT") = "" Then
'        Filehdl = FreeFile()
'        Open App.Path & "\TRANS.DAT" For Binary Access Read Write As #Filehdl
'        PSetup.Port = PortNo
'        PSetup.baud = res!Boud
'        Put #Filehdl, 1, PSetup
'        Close #Filehdl
'    End If
'    txtPort = PortNo
'    txtBoud = res!Boud
'    res.Close
'
'Set SerialComm = SerialCom
'    'TransPath = App.Path & "\Data"
'    TransPath = App.Path
'    Filehdl = FreeFile()
'    Open App.Path & "\TRANS.DAT" For Binary Access Read Write As #Filehdl
'    Get #Filehdl, , PSetup
'    Close #Filehdl
'
'    If Not InitPort(PSetup.Port, PSetup.baud) Then
'        MsgBox TransMsg
'    End If
'
'
'    Exit Sub
'chk:
'    Select Case err.Number
'        Case 70
'            MsgBox "   Database already running    " & vbCrLf & "   Please close database  ", vbInformation, "Bus"
'        Case Else
'            MsgBox "Error No: " & err.Number & "     " & err.Description, vbInformation, "Bus"
'    End Select
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
' 'PC_to_PMTC_Cntr = 0
' If Dir(App.Path & "\TICKET.TXT") <> "" Then
'   Kill App.Path & "\TICKET.TXT"
'End If
' If Dir(App.Path & "\RPT*.DAT") <> "" Then
'   Kill App.Path & "\RPT*.DAT"
'End If
' If Dir(App.Path & "\TKTS*.DAT") <> "" Then
'   Kill App.Path & "\TKTS*.DAT"
'End If
' If Dir(App.Path & "\RPT*.CSV") <> "" Then
'   Kill App.Path & "\RPT*.CSV"
'End If
'If Dir(App.Path & "\TKTS*.CSV") <> "" Then
'   Kill App.Path & "\TKTS*.CSV"
'End If
'If Dir(App.Path & "\REPORT*.TXT") <> "" Then
'   Kill App.Path & "\REPORT*.TXT"
'End If
'
'
'    If SerialComm.PortOpen = True Then SerialComm.PortOpen = False
''    If Not InitPort(PSetup.Port, PSetup.baud) Then
''        MsgBox TransMsg, vbInformation, "Error"
''    End If
''Dim Filehdl As Integer
''  If Dir(App.Path & "\TRANS.DAT") <> "" Then
''   Kill App.Path & "\TRANS.DAT"
''  End If
''    Filehdl = FreeFile()
''    Open App.Path & "\TRANS.DAT" For Binary Access Read Write As #Filehdl
''    If OptPort(0).Value = True Then
''       PSetup.Port = "1"
''    Else
''        PSetup.Port = "2"
''    End If
''    PSetup.baud = "115200"
''    Put #Filehdl, 1, PSetup
''    Close #Filehdl
'End Sub
Private Sub cmdCrew_Click()

End Sub

