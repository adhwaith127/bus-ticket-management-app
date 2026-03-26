VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmSchedule 
   BackColor       =   &H00E0E0E0&
   Caption         =   "TRIP Report"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frSelectSchedule 
      BackColor       =   &H00E0E0E0&
      Height          =   3060
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   5715
      Begin VB.ComboBox cmbtripno 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cmbSchedule 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cmbPalmID 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optReportType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Schedule Wise"
         Height          =   300
         Index           =   0
         Left            =   -120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.OptionButton optReportType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Trip Wise"
         Height          =   300
         Index           =   1
         Left            =   1815
         TabIndex        =   4
         Top             =   225
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.ComboBox cmbrouteno 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   375
         Left            =   2850
         TabIndex        =   2
         Top             =   2160
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmSchedule.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSchsmryRpt 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   2160
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         TX              =   "&Export"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmSchedule.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16576
         Format          =   71696385
         CurrentDate     =   39536
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   330
         Left            =   4200
         TabIndex        =   15
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16576
         Format          =   71696385
         CurrentDate     =   39536
      End
      Begin VB.Label lblTripno 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "    Trip No    "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbltodate 
         BackStyle       =   0  'Transparent
         Caption         =   "  End Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2880
         TabIndex        =   14
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label lblfromdate 
         BackStyle       =   0  'Transparent
         Caption         =   "       Start Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -240
         TabIndex        =   12
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lblPalmId 
         BackStyle       =   0  'Transparent
         Caption         =   "       PalmtecID  "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -240
         TabIndex        =   10
         Top             =   960
         Width           =   1770
      End
      Begin VB.Label lblSchedule 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "        Schedule "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblRouteNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "            Route       "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   -600
         TabIndex        =   8
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRIP Summary Report"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1005
      TabIndex        =   11
      Top             =   -495
      Width           =   3720
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmbPalmID_Click()
If cmbPalmID.Text <> "All" And cmbPalmID.Text <> "ALL" Then
     sql = "SELECT DISTINCT SCHEDULE  FROM RPT WHERE PALMID='" & cmbPalmID.Text & "' and  DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
     Else
        sql = "SELECT DISTINCT SCHEDULE  FROM RPT WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
     End If
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
       If cmbPalmID.Text <> "All" And cmbPalmID.Text <> "ALL" Then
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbSchedule.Clear
        cmbSchedule.AddItem "ALL"
        Do While Not RES.EOF
            cmbSchedule.AddItem RES!SCHEDULE
            RES.MoveNext
        Loop
      Else
        cmbSchedule.Clear
        cmbSchedule.AddItem "ALL"
      End If
        If cmbSchedule.ListCount > 0 Then
            cmbSchedule.Text = cmbSchedule.List(0)
        End If
        'If RES. =  Then RES.Close
End Sub

Private Sub cmbSchedule_Click()
     If cmbSchedule.Text <> "All" And cmbSchedule.Text <> "ALL" Then
      sql = "SELECT DISTINCT TRIPNO  FROM RPT WHERE SCHEDULE=" & cmbSchedule.Text & " and PALMID='" & cmbPalmID.Text & "' and  DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
     Else
        sql = "SELECT DISTINCT TRIPNO  FROM RPT WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
     End If
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
       If cmbSchedule.Text <> "All" And cmbSchedule.Text <> "ALL" Then
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbtripno.Clear
        cmbtripno.AddItem "ALL"
        Do While Not RES.EOF
            cmbtripno.AddItem RES!TripNo
            RES.MoveNext
        Loop
      Else
        cmbtripno.Clear
        cmbtripno.AddItem "ALL"
      End If
        If cmbtripno.ListCount > 0 Then
            cmbtripno.Text = cmbtripno.List(0)
        End If
        RES.Close
End Sub

Private Sub cmdSchsmryRpt_Click()
On Error Resume Next
  Call SchsmryRptCnv
    
'   Cmd.Filter = "TXT (*.TXT)|*.TXT"
'    Cmd.ShowOpen
''    rtfText.Width = frmReport.Width - 300
''    rtfText.Height = frmReport.Height - 200
'    rchtxtbox.Font = "Lucida console"
'    rchtxtbox.Font.Size = "10"
'    rchtxtbox.Locked = True
'    rchtxtbox.LoadFile (Cmd.filename)
'    rchtxtfrm.Show vbModal
'    cmdPrint.Enabled = True
End Sub

Public Function SchsmryRpt() As Boolean

''Public Function CovertColln(Fint As String) As Boolean
' On Error GoTo errLn
     Dim FS As New FileSystemObject
     Dim TcketPath As String
'    Dim FHndl As Integer
     Dim fShndl As Integer
'    Dim tktK As PTicket
     Dim fBuff As String
     Dim FnameUp As String
'    Dim PFname As String
'    Dim pHandle As Integer
'    Dim gPass As PASSCONC
'    '''gPassCount added by syam
'    Dim gPassCount As Long
'    Dim iFull As Integer
'    Dim iHalf As Integer
'    Dim iPhy As Integer
'    Dim iLugg As Integer
'    Dim iSt As Integer
'    Dim lTotPassenger As Long
'    Dim fTotAmount As Single
'    Dim fTotLuggAmount As Single
'    Dim strYear As String
'    Dim SysD, SysT, PID As String
'    gPassCount = 0
        TSQL = "SELECT * FROM PCSETUP"
        Set RES = CNN.OpenRecordset(TSQL, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            TcketPath = RES!TICKET_PATH
            TransPath = RES!TRANSFER_PATH
        End If
        RES.Close
'
'        sql = "SELECT PALMTECID FROM SETTINGS"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        If RES.RecordCount > 0 Then
'            PalmId = RES!PalmtecID
'        End If
'        RES.Close
'
         sql = "SELECT * FROM RPT"

'        FnameUp = "TKTS" & Fint & ".DAT"
'        'If Dir(TransPath & "\" & FnameUp) <> "" Then
         SysD = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
'        'SysT = Replace(Time, ":", ".")
'        SysT = Format(Time, "hhmmAM/PM")
'        PID = Replace(PalmId, Chr(0), "")
        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
        TcketPath = TcketPath & "\" & SysD
            If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
            FnameUp = "SCHEDULESUMMARYRPT.TXT"
            '& Fint & "-" & SysT & PID
            If Dir(TcketPath & "\" & FnameUp) <> "" Then Kill TcketPath & "\" & FnameUp
            fShndl = FreeFile()
            Open TcketPath & "\" & FnameUp For Binary Access Write As #fShndl
                fBuff = ""
                fBuff = String(84, "_") & vbCrLf
                Put #fShndl, , fBuff
                fBuff = Format("DATE|", "@@@@@@@@@@@@") & vbCrLf
                fBuff = fBuff & Format("PALMID| ", "@@@@")
                fBuff = fBuff & Format("SCHEDULENO| ", "@@@@") & vbCrLf
'                fBuff = fBuff & Format("LG| ", "@@@@")
'                fBuff = fBuff & Format("PH| ", "@@@@")
'                fBuff = fBuff & Format("ST| ", "@@@@")
'                fBuff = fBuff & Format("PASS No|", "@@@@@@@@@")
'                'fBuff = fBuff & Format("FROM| ", "@@@@@")
'                'fBuff = fBuff & Format("TO| ", "@@@@@")
'                fBuff = fBuff & Format("AMOUNT|", "@@@@@@@")
'                fBuff = fBuff & Format("LUG_AMT| ", "@@@@@@@@")
'                'fBuff = fBuff & Format("TIME|", "@@@@@@")
'                'fBuff = fBuff & Format("DATE|", "@@@@@@@@@@@@") & vbCrLf
                 Put #fShndl, , fBuff
                 fBuff = String(84, "_") & "|" & vbCrLf
                 Put #fShndl, , fBuff
'
'
'        FnameUp = "TKTS" & Fint & ".DAT"
'        PFname = App.Path & "\PASS.PAS"
'        FHndl = FreeFile()
'            Open TransPath & "\" & FnameUp For Binary Access Read As #FHndl
'               Do While Not EOF(FHndl)
'                Dim str As String
'                Get #FHndl, , tktK
'                If tktK.TicketNo = -1 Then Exit Do
'                If EOF(FHndl) = True Then Exit Do
'                 With tktK
'                    fBuff = Format(.TicketNo, " 00000000") & "| "
'                    fBuff = fBuff & Format(.Full, "00") & "| "
'                    fBuff = fBuff & Format(.Half, "00") & "| "
'                    fBuff = fBuff & Format(.Lugg, "00") & "| "
'                    fBuff = fBuff & Format(.Phy, "00") & "| "
'                    fBuff = fBuff & Format(.st, "00") & "| "
'
'                    iFull = iFull + .Full
'                    iHalf = iHalf + .Half
'                    iLugg = iLugg + .Lugg
'                    iPhy = iPhy + .Phy
'                    iSt = iSt + .st
'
'                    If .Typ = 32 Then
'                        If Dir(PFname) <> "" Then
'                            pHandle = FreeFile()
'                            Open PFname For Binary Access Read As pHandle
'                            Do While Not EOF(pHandle)
'                                Get #pHandle, , gPass
'                                If .TicketNo = gPass.TicketNo Then Exit Do
'                            Loop
'                            Close #pHandle
'                            str = TrimChr(gPass.PassNo)
'                            'MsgBox str
'                        Else
'                            str = "  "
'                        End If

'                        If str <> "  " Then
'                        gPassCount = gPassCount + 1
'                        End If
'
'
'
'                        ''''''''''''''
'
'                        fBuff = fBuff & Format(str & "|", "@@@@@@@@@")
'                    Else
'                        fBuff = fBuff & String(7, " ") & "-|"
'                    End If
'                    fBuff = fBuff & Format(.From, " 000") & "|"
'                    fBuff = fBuff & Format(.To, " 000") & "|"
'                    str = Format(.Amount, "0.00")
'                    fBuff = fBuff & Format(str, "@@@@@@@") & "|"
'                    str = Format(.Luggage, "0.00")
'                    fBuff = fBuff & Format(str, "@@@@@@@") & "|"
'                    fBuff = fBuff & " " & Format(.Hr & ":" & .Minut, "HH:MM") & "|"
'                    strYear = ""  '05/01/2010
'                    strYear = DatePart("YYYY", Date) '05/01/2010
'                    fBuff = fBuff & " " & Format(.Dy & "/" & .Mn & "/" & strYear, "DD/MM/YYYY") & "|" & vbCrLf  '05/01/2010
'                    Put #fShndl, , fBuff
'
'
'                    fTotAmount = fTotAmount + .Amount
'                    fTotLuggAmount = fTotLuggAmount + .Luggage
'                    fBuff = ""
'                 End With
'               Loop
'                fBuff = String(84, "_") & "|" & vbCrLf '05/01/2010
'                Put #fShndl, , fBuff
'                lTotPassenger = iFull + iHalf + iPhy + iSt + gPassCount
'
'                fBuff = "TOTAL FULL       |" & Format(iFull, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                fBuff = "TOTAL HALF       |" & Format(iHalf, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                fBuff = "TOTAL PHY        |" & Format(iPhy, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                fBuff = "TOTAL ST         |" & Format(iSt, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                fBuff = "TOTAL LUGGAGE    |" & Format(iLugg, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                 ''''''''''''''''''''''syam
'
'
'               ' fBuff = "TOTAL PASS" & gPassCount
'                '  Put #fShndl, , fBuff
'
'                fBuff = "TOTAL PASS       |" & Format(gPassCount, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'               ''''''''''''''''''''''''''
'
'                fBuff = "TOTAL PASSENGER  |" & Format(lTotPassenger, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'                str = Format(fTotLuggAmount, "0.00")
'                fBuff = "TOTAL LUGGAGE    |" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'                str = Format(fTotAmount, "0.00")
'                fBuff = "TOTAL AMOUNT     |" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'              DBTKTS (FnameUp)
'
'
'            Close #FHndl
'         Close #fShndl
'         CovertColln = True
'        Exit Function
'      'End If
'errLn:
'  CovertColln = False
    
End Function
'End Function

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Load()
On Error Resume Next
    Me.Icon = frmMainform.Icon
    CONNECTDB
    sql = "SELECT DISTINCT PALMID FROM RPT"
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
   'optReportType(1).Value = True
   
    
        DTfrom.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
        DTTo.Value = DateValue(Format(Now, "DD/MM/YYYY"))
   If optReportType(0).Value = True Then
        'DTEnd.Visible = False
        'DTStart.Visible = False
        
        DTfrom.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
        DTTo.Value = DateValue(Format(Now, "DD/MM/YYYY"))
        'DTFrom.Top = 960
        'DTFrom.Left = 2000
        'DTTo.Top = 960
        'DTTo.Left = 4080

        'cmbrouteno.Top = 2520
        'cmbrouteno.Left = 2000
        'lblSDate.Visible = False
        'lblEDate.Visible = False
'        lblTripno.Visible = False
'        cmbtripno.Visible = False
        
'        cmbStartTkt.Top = 2280
'        cmbStartTkt.Left = 2000
'        cmbEndTkt.Top = 2280
'        cmbEndTkt.Left = 4440
'        DTSchDate.Top = 1200
'        DTSchDate.Left = 3500
'        DTSchDate.Day = Day(Now)
'        DTSchDate.Month = Month(Now)
'        DTSchDate.Year = Year(Now)
'        DTSchDate.Enabled = False
        
        
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbPalmID.Clear
        cmbPalmID.AddItem "ALL"
        Do While Not RES.EOF
            cmbPalmID.AddItem RES!PalmID
            RES.MoveNext
        Loop
        If cmbPalmID.ListCount > 0 Then
            cmbPalmID.Text = cmbPalmID.List(0)
        End If
      '  RES.Close
      
        sql = "SELECT DISTINCT SCHEDULE  FROM RPT " 'WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbSchedule.Clear
        cmbSchedule.AddItem "ALL"
        Do While Not RES.EOF
            cmbSchedule.AddItem RES!SCHEDULE
            RES.MoveNext
        Loop
        If cmbSchedule.ListCount > 0 Then
            cmbSchedule.Text = cmbSchedule.List(0)
        End If
        RES.Close
        
'        sql = "SELECT DISTINCT TRIPNO  FROM RPT WHERE PALMID='" & cmbPalmID.Text & "'"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        If RES.RecordCount > 0 Then RES.MoveFirst
'        cmbtripno.Clear
'        Do While Not RES.EOF
'            cmbtripno.AddItem RES!TripNo
'            RES.MoveNext
'        Loop
'        If cmbtripno.ListCount > 0 Then
'            cmbtripno.Text = cmbtripno.List(0)
'        End If
        
'        sql = "SELECT DISTINCT RouteCode  FROM RPT " ' WHERE PALMID='" & cmbPalmID.Text & "'"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        If RES.RecordCount > 0 Then RES.MoveFirst
'        cmbrouteno.Clear
'        cmbrouteno.AddItem "ALL"
'        Do While Not RES.EOF
'            cmbrouteno.AddItem RES!RouteCode
'            RES.MoveNext
'        Loop
'        If cmbrouteno.ListCount > 0 Then
'            cmbrouteno.Text = cmbrouteno.List(0)
'        End If

         sql = "SELECT DISTINCT TRIPNO  FROM RPT " ' WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbtripno.Clear
        cmbtripno.AddItem "ALL"
        Do While Not RES.EOF
            cmbtripno.AddItem RES!TripNo
            RES.MoveNext
        Loop
        If cmbtripno.ListCount > 0 Then
            cmbtripno.Text = cmbtripno.List(0)
        End If
'
    
    Else
        'lblfromdate.Visible = False
        'lbltodate.Visible = False
        'lblSdateOrID.Caption = "Start Date  :"
        'lblEndDateOrSch.Caption = " End Date :"
        'lblRouteNo.Caption = " Trip No :"
        'DTEnd.Visible = True
        'DTStart.Visible = True
        'cmbPalmID.Visible = False
        'cmbShedule.Visible = False
        
        'lblSDate.Visible = True
        'lblEDate.Visible = True
        lbltripno.Visible = True
        cmbtripno.Visible = True
       
        'DTStart.Top = 840
        'DTStart.Left = 2000
        'DTEnd.Top = 1320
        'DTEnd.Left = 2000
        'DTStart.Width = 1300
        'DTEnd.Width = 1300
        'cmbtripno.Top = 1800                        '''rnc
        'cmbtripno.Left = 2000
        'DTStart.Day = Day(Now)
        'DTStart.Month = Month(Now)
        'DTStart.Year = Year(Now)
       'DTEnd.Day = Day(Now)
        'DTEnd.Month = Month(Now)
        'DTEnd.Year = Year(Now)
        
        sql = "SELECT DISTINCT TRIPNO  FROM RPT " ' WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbtripno.Clear
        cmbtripno.AddItem "ALL"
        Do While Not RES.EOF
            cmbtripno.AddItem RES!TripNo
            RES.MoveNext
        Loop
        If cmbtripno.ListCount > 0 Then
            cmbtripno.Text = cmbtripno.List(0)
        End If
        End If
        
'        If optReportType(0).Value = True Then                       ''rnc
'            lbltripno.Visible = True
'            cmbtripno.Visible = True
'    End If
    RES.Close
End Sub

Private Sub optReportType_Click(Index As Integer)
On Error Resume Next

    If optReportType(0).Value = True Then
    
        DTfrom.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
        DTTo.Value = DateValue(Format(Now, "DD/MM/YYYY"))
        'lblSdateOrID.Caption = " Palmtec ID    :"
        'lblEndDateOrSch.Caption = "Schedule No :"
        'DTEnd.Visible = False
        'DTStart.Visible = False
        'cmbPalmID.Visible = True
        'cmbShedule.Visible = True
        'DTFrom.Top = 960
        'DTFrom.Left = 2000
        'DTTo.Top = 960
        'DTTo.Left = 4080
        'cmbPalmID.Top = 1560
        'cmbPalmID.Left = 2000
        'cmbShedule.Top = 2040
        'cmbShedule.Left = 2000
       'cmbrouteno.Top = 2520
        'cmbrouteno.Left = 2000
        
        'lblSDate.Visible = False
        'lblEDate.Visible = False
        lbltripno.Visible = False
        cmbtripno.Visible = False
        cmbPalmID.Clear
        
        sql = "SELECT DISTINCT PALMID FROM RPT"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        Do While Not RES.EOF
            cmbPalmID.AddItem RES!PalmID
            RES.MoveNext
        Loop
        If cmbPalmID.ListCount > 0 Then
            cmbPalmID.Text = cmbPalmID.List(0)
        End If
        
     '   RES.Close
        
        sql = "SELECT DISTINCT SCHEDULE FROM RPT" ' WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbSchedule.Clear
        Do While Not RES.EOF
            cmbSchedule.AddItem RES!SCHEDULE
            RES.MoveNext
        Loop
        If cmbSchedule.ListCount > 0 Then
            cmbSchedule.Text = cmbSchedule.List(0)
        End If
        RES.Close
        'sql = "SELECT DISTINCT TRIPNO FROM RPT WHERE PALMID='" & cmbPalmID.Text & "'"
        'Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        'If RES.RecordCount > 0 Then RES.MoveFirst
        'cmbtripno.Clear
        'Do While Not RES.EOF
         '   cmbtripno.AddItem RES!TripNo
          '  RES.MoveNext
         'Loop
        'If cmbtripno.ListCount > 0 Then
         '   cmbtripno.Text = cmbtripno.List(0)
        'End If
        sql = "SELECT DISTINCT RouteCode  FROM RPT " 'WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbrouteno.Clear
        Do While Not RES.EOF
        
            cmbrouteno.AddItem RES!RouteCode
            RES.MoveNext
        Loop
        If cmbrouteno.ListCount > 0 Then
            cmbrouteno.Text = cmbrouteno.List(0)
        End If


    Else
       'lblfromdate.Visible = False
       'DTFrom.Visible = False
        'lbltodate.Visible = False
        'DTTo.Visible = False
        'lblSdateOrID.Caption = "Start Date  :"
        'lblEndDateOrSch.Caption = " End Date :"
        'lblRouteNo.Caption = "   Trip No  :"
        'DTEnd.Visible = True
        'DTStart.Visible = True
        'cmbPalmID.Visible = False
        'cmbShedule.Visible = False
        'cmbrouteno.Visible = False
        'cmbtripno.Visible = True
        'lblSdateOrID.Top = 960
        'lblSdateOrID.Left = 500
        'DTStart.Top = 960
        'DTStart.Left = 2000
        'lblEndDateOrSch.Top = 1560
        'lblEndDateOrSch.Left = 500
        'DTEnd.Top = 1560
        'DTEnd.Left = 2000
        'DTStart.Width = 960
        'DTEnd.Width = 960
        'lblRouteNo.Top = 2100
        'lblRouteNo.Left = 500
        'cmbtripno.Top = 2100                        '''rnc
        'cmbtripno.Left = 2000
        'lblSDate.Visible = True
        'DTStart.Visible = True
        'lblEDate.Visible = True
        'DTEnd.Visible = True
        lbltripno.Visible = True
        cmbtripno.Visible = True
        'DTStart.Day = Day(Now)
        'DTStart.Month = Month(Now)
        'DTStart.Year = Year(Now)
        'DTEnd.Day = Day(Now)
        'DTEnd.Month = Month(Now)
        'DTEnd.Year = Year(Now)
        
        sql = "SELECT DISTINCT TripNo  FROM RPT" 'WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbtripno.Clear
        cmbtripno.AddItem "ALL"
        Do While Not RES.EOF
            cmbtripno.AddItem RES!TripNo
            RES.MoveNext
        Loop
        If cmbtripno.ListCount > 0 Then
            cmbtripno.Text = cmbtripno.List(0)
        End If
       ' End If
        
    End If
    If optReportType(0).Value = True Then
        
        'lblSDate.Visible = False
        'lblEDate.Visible = False
        lbltripno.Visible = False
        cmbtripno.Visible = False
    End If
    RES.Close
End Sub
Public Sub SchsmryRptCnv()
On Error GoTo erromod
Dim FS As New FileSystemObject
'Dim fShndl As Integer
Dim SysD, FnameUp As String
Dim qry, sql, sSQL, subsql, sqlroute As String
Dim HDR1, HDR2 As String
Dim pamt As Integer
Dim Total As Double
Dim expsql As String
Dim cond As String
Dim Stno, Etno As String
Dim checkflag As Boolean
Dim mystr As String
'open file
   optReportType(1).Value = True
pamt = 0
pamt = 0
TSQL = "SELECT * FROM PCSETUP"
        Set RES = CNN.OpenRecordset(TSQL, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            TcketPath = RES!TICKET_PATH
            TransPath = RES!TRANSFER_PATH
        End If
        RES.Close
        SysD = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
        TcketPath = TcketPath & "\" & SysD
        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
        If optReportType(0).Value = True Then
            FnameUp = "SCHEDULE SUMMARY REPORT"
        Else
            FnameUp = "TRIP WISE DETAILS"
        End If
        '& Fint & "-" & SysT & PID
        'If Dir(TcketPath & "\" & FnameUp) <> "" Then Kill TcketPath & "\" & FnameUp
'''        fShndl = FreeFile()
    If Dir(TcketPath & "\" & FnameUp & ".xlsx", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xlsx")
    If Dir(TcketPath & "\" & FnameUp & ".xls", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xls")
        Dim ExlObj As New excel.Application
        ExlObj.Workbooks.Add
       ' ExlObj.Visible = True
       

    'print HEADER FROM SETTINGS
        qry = "SELECT HEADER1,HEADER2 FROM SETTINGS"
        Set RES = CNN.OpenRecordset(qry, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            HDR1 = RES!HEADER1
            HDR2 = RES!HEADER2
        End If
        RES.Close
       Dim exclrow As Integer, exclcol As Integer
       ExlObj.ActiveSheet.Cells(1, 5).Value = HDR1
       ExlObj.ActiveSheet.Cells(1, 5).HorizontalAlignment = xlCenter
       ExlObj.ActiveSheet.Cells(1, 5).Font.Bold = True
       ExlObj.ActiveSheet.Cells(2, 5).Value = HDR2
       ExlObj.ActiveSheet.Cells(2, 5).HorizontalAlignment = xlCenter
       ExlObj.ActiveSheet.Cells(2, 5).Font.Bold = True
       If optReportType(0).Value = True Then
            ExlObj.ActiveSheet.Cells(3, 5).Value = "Schedule wise Summary Report"
       Else
            ExlObj.ActiveSheet.Cells(3, 5).Value = "TRIP WISE DETAILS"
       End If
       ExlObj.ActiveSheet.Cells(3, 5).Font.Bold = True
       ExlObj.ActiveSheet.Cells(3, 5).HorizontalAlignment = xlCenter
       'exclrow = 4
       'ExlObj.ActiveSheet.Range("C5:G5").Value = "___________________"
       ' ExlObj.Range("C" & exclrow & ":" & "G" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
       
'       If (optReportType(0).Value = False) Then
'        sql = "SELECT RT.* FROM RPT RT WHERE DATE BETWEEN DATEVALUE('" & DTFrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "  'SANGEETHA
'
'        Else
'

   
    If (optReportType(0).Value = False) Then
    
     sql = "SELECT StartDate,UpDownTrip,TripNo,SCHEDULE,Conductor,Routecode,Driver,Cleaner,PalmID, Sum(Fulls) AS fullcnt, Sum(fullcoll) AS fullcln," _
          & "Sum(Luggage) AS Luggcnt,Sum(AdjustColl) as Adju,Sum(Adjust) As Adjcnt,Sum(Half) AS halfcnt, Sum(Halfcoll) AS halfcln, Sum(LuggageColl) AS Luggagecln," _
          & " Sum(Phy) AS phycnt, Sum(RPT.St) AS Stcnt, Sum(Stcoll) AS Stcln,Sum(Phycoll) AS Phycl, Sum(pass) AS passcnt, Sum(TotalColl) AS totcol" _
          & ",sum(EXPENSE) as Expns,sum(ladies_count) as lad_count,sum(ladies_coll) as lad_col,sum(senior_count) as sc_count,sum(senior_coll) as sc_col FROM RPT   WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "  'SANGEETHA
    Else

     sql = "SELECT StartDate,SCHEDULE,Conductor,Routecode,Driver,Cleaner,PalmID, Sum(Fulls) AS fullcnt, Sum(fullcoll) AS fullcln," _
          & "Sum(Luggage) AS Luggcnt,Sum(AdjustColl) as Adju,Sum(Adjust) As Adjcnt, Sum(Half) AS halfcnt, Sum(Halfcoll) AS halfcln, Sum(LuggageColl) AS Luggagecln," _
          & " Sum(Phy) AS phycnt, Sum(RPT.St) AS Stcnt, Sum(RPT.Stcoll) AS Stcln,Sum(Phycoll) AS Phycl, Sum(pass) AS passcnt, Sum(TotalColl) AS totcol" _
          & ",sum(EXPENSE) as Expns,sum(ladies_count) as lad_count,sum(ladies_coll) as lad_col,sum(senior_count) as sc_count,sum(senior_coll) as sc_col FROM RPT  WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "  'SANGEETHA
   End If
   
     sqlroute = "SELECT  Routecode  FROM RPT  WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "  'SANGEETHA
        If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
            sql = sql & " AND PalmID='" & cmbPalmID & "'"
            sqlroute = sqlroute & " AND rt.PalmID='" & cmbPalmID & "'"
        End If
    
        If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
            sql = sql & " AND SCHEDULE=" & val(cmbSchedule) & ""
            sqlroute = sqlroute & " AND rt.SCHEDULE=" & val(cmbSchedule) & ""
        End If
    
        If cmbrouteno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
            sql = sql & " AND ROUTECODE='" & cmbrouteno & "'"
            sqlroute = sqlroute & " AND rt.ROUTECODE='" & cmbrouteno & "'"
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If cmbtripno.ListIndex <> -1 And cmbtripno.Text <> "ALL" Then
            sql = sql & " AND TRIPNO=" & val(cmbtripno) & " "
            sqlroute = sqlroute & " AND rt.TRIPNO=" & val(cmbtripno) & ""
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If optReportType(0).Value = True Then
            'sql = sql & " ORDER BY RT.DATE,RT.PALMID,RT.SCHEDULE,RT.STicket"
            sql = sql & " GROUP BY StartDate,PALMID,SCHEDULE,Conductor,Driver,Cleaner,Routecode"
            
        Else
            sql = sql & " GROUP BY StartDate,PALMID,SCHEDULE,Conductor,Driver,Cleaner,Routecode,TripNo,UpDownTrip"
           sql = sql & " ORDER BY StartDate,PALMID,SCHEDULE,TripNo"
           
        End If
   '     getvalueQuery(sqlroute
'        sql = "SELECT Date,SCHEDULE,Conductor,Routecode,Driver,Cleaner,PalmID, (Select Sum(Luggage) AS Luggcnt,Sum(Fulls) AS fullcnt, Sum(fullcoll) AS fullcln, Sum(Half) AS halfcnt, Sum(Halfcoll) AS halfcln, Sum(LuggageColl) AS Luggagecln," _
'          & " Sum(Phy) AS phycnt, Sum(St) AS Stcnt, Sum(Stcoll) AS Stcln,Sum(Phycoll) AS Phycl, Sum(Val(pass)) AS passcnt, Sum(TotalColl) AS totcol" _
'          & " FROM RPT as RPT1 where RPT1.Trip_Master_ID=RPT2.Trip_Master_ID GROUP BY RPT1.DATE,RPT1.PALMID,RPT1.SCHEDULE,RPT1.Routecode)" _
'          & "FROM RPT as RPT2 WHERE cdate(DATE) BETWEEN DATEVALUE('" & DTFrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "  'SANGEETHA
'        ConnectDatabase adoc, "pvt.mdb", "silbus"
'        Set RES = adoc.Execute(sql)
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        
exclrow = 4
Dim gtotal As Double
Dim totalexp As Double
totalexp = 0
checkflag = False
If Not RES.EOF Then
   Do While Not RES.EOF
    Total = 0
     checkflag = True
          ExlObj.Range("A" & exclrow & ":" & "I" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
          'ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
          exclrow = exclrow + 1
          ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
          ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Date       "
          ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Bold = True
          ExlObj.ActiveCell(exclrow, 2).Style.NumberFormat = "@"
          ExlObj.ActiveSheet.Cells(exclrow, 2).Value = CStr(RES!StartDate)
          ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
          exclrow = exclrow + 1
          ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "PalmId      "
          ExlObj.ActiveSheet.Cells(exclrow, 2).Value = RES!PalmID
          ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Schedule No  "
          ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!SCHEDULE
          ExlObj.ActiveCell(exclrow, 9).Style.NumberFormat = "@"
          exclrow = exclrow + 1
          sSQL = ""
          mystr = ""
'         ' If optReportType(0).Value = True Then
                sSQL = "select Count(tripno) as CTRIP,min(STicket) as stno ,max(ETicketNo) as etno from rpt WHERE DATEVALUE(StartDate) ='" & RES!StartDate & "' "
                mystr = "select count(refundsts) as Recount,sum(refundamt) as refund from tkts where cdate(date) between DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "')"
                If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
                    sSQL = sSQL & " AND PalmID='" & cmbPalmID & "'"
                    mystr = mystr & " AND PalmID='" & cmbPalmID & "'"

                Else
                    sSQL = sSQL & " AND PalmID='" & Trim(TrimChr(RES!PalmID)) & "'"
                    mystr = mystr & " AND PalmID='" & Trim(TrimChr(RES!PalmID)) & "'"
                 End If
                If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
                    sSQL = sSQL & " AND SCHEDULE=" & val(cmbSchedule) & " "
                    mystr = mystr & " AND SCHDULE=" & val(cmbSchedule) & " "
                Else
                    sSQL = sSQL & " AND SCHEDULE=" & RES!SCHEDULE & " "
                    mystr = mystr & " AND SCHDULE=" & RES!SCHEDULE & " "
                End If
                If cmbrouteno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
                    sSQL = sSQL & " AND ROUTECODE='" & cmbrouteno & "'"
                Else
                    sSQL = sSQL & " AND ROUTECODE='" & RES!RouteCode & "' "
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If cmbtripno.ListIndex <> -1 And cmbtripno.Text <> "ALL" Then
                    sSQL = sSQL & " AND TRIPNO=" & val(cmbtripno) & ""
                    mystr = mystr & " AND TRIPNO=" & val(cmbtripno) & ""
                Else
                    sSQL = sSQL & " AND TRIPNO= " & RES!TripNo & ""
                    mystr = mystr & " AND TRIPNO= " & RES!TripNo & ""
                End If
'          'End If '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
          If optReportType(0).Value = True Then
                sSQL = sSQL
          Else
                sSQL = sSQL
          End If
          mystr = mystr & " and refundsts=1"
         Set res1 = CNN.OpenRecordset(sSQL, dbOpenDynaset)
         Set res9 = CNN.OpenRecordset(mystr, dbOpenDynaset)
         If optReportType(0).Value = True Then
            ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "No Of Trips   "
            ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "  " & res1!CTRIP
         End If
'         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Route No      "
'         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!RouteCode
        ' exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Conductor   "
         ExlObj.ActiveSheet.Cells(exclrow, 2).Value = RES!Conductor
         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Driver      "
         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!Driver
         exclrow = exclrow + 1
         If optReportType(1).Value = True Then
             Dim sroutename As String, sroutesql As String
             sroutesql = "select rutname from ROUTE where rutcode='" & RES!RouteCode & "'"
             Set RESROUTE = CNN.OpenRecordset(sroutesql, dbOpenDynaset)
             ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Trip Number & Flag      "
             ExlObj.ActiveSheet.Cells(exclrow, 2).Value = RES!TripNo & "   " & IIf(RES!UpDownTrip = "U", "UP", "DOWN")
             ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Route No.            "
             ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!RouteCode
             Set RESROUTE = Nothing
             exclrow = exclrow + 1
         End If
         ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Start TktNo    "
         ExlObj.ActiveSheet.Cells(exclrow, 2).Value = res1!Stno
         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "End Tkt No     "
         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = res1!Etno
         'exclrow = exclrow + 1
         ExlObj.Range("A" & exclrow & ":" & "I" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous  'SANGEETHA
         'ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "      " & "Count"
         ExlObj.ActiveSheet.Cells(exclrow, 3).Value = "      " & "Amount"
         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "      " & "Count"
         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = "      " & "Amount"
         exclrow = exclrow + 1
         ExlObj.Range("B" & exclrow & ":" & "C" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
         ExlObj.Range("H" & exclrow & ":" & "I" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
         'ExlObj.ActiveSheet.Range("B" & exclrow & ":" & "C" & exclrow).Value = "_______________"
         'ExlObj.ActiveSheet.Range("H" & exclrow & ":" & "I" & exclrow).Value = "_______________"
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 1) = " Full       "
         ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!fullcnt
         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!fullcln), 2)
         'ExlObj.ActiveSheet.Cells(exclrow, 3).Align = 6
         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
         ExlObj.ActiveSheet.Cells(exclrow, 7) = " Half       "
         ExlObj.ActiveSheet.Cells(exclrow, 8) = RES!halfcnt
         ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 9) = Round(val(RES!halfcln), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlRight
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 1) = " ST       "
         ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!stcnt
         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!stcln), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
         ExlObj.ActiveSheet.Cells(exclrow, 7) = " PH       "
         ExlObj.ActiveSheet.Cells(exclrow, 8) = RES!phycnt
         ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 9) = Round(val(RES!Phycl), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlRight
         exclrow = exclrow + 1
         
         ExlObj.ActiveSheet.Cells(exclrow, 1) = " SC       "
         ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!sc_count
         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 3) = IIf(IsNull(RES!sc_col), 0, Round(RES!sc_col, 2))
         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
         
         ExlObj.ActiveSheet.Cells(exclrow, 7) = " Ladies   "
         ExlObj.ActiveSheet.Cells(exclrow, 8) = RES!lad_count
         ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 9) = Round(val(RES!lad_col), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlRight
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 1) = " REFUND       "
         ExlObj.ActiveSheet.Cells(exclrow, 2) = res9!Recount
         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 3) = IIf(IsNull(res9!refund), 0, Round(res9!refund, 2))
         refund = refund + IIf(IsNull(res9!refund), 0, Round(res9!refund, 2))
         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
        exclrow = exclrow + 1
        
        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Pass       "
        ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!passcnt
        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(0), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
        exclrow = exclrow + 1
        'new
        
        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Luggage       "
        ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!luggcnt
        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
        
        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!Luggagecln), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
        exclrow = exclrow + 1
        
        
        ''''''''''''''''''''''''
        
        
        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Adjust       "
        ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!Adjcnt
        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!Adju), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
        'exclrow = exclrow + 1
        ExlObj.Range("A" & exclrow & ":" & "I" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
        'ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
        exclrow = exclrow + 1
        
       ' exclrow = exclrow + 1
'        ExlObj.ActiveSheet.Cells(exclrow, 1) = RES!Expns
'        ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(RES2!EXPAmount), 2)
'        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
'        totalexp = totalexp + RES2!EXPAmount
       
        Total = Total + RES!fullcln + RES!halfcln + RES!stcln + RES!Phycl + RES!lad_col + RES!sc_col + RES!Luggagecln - RES!Adju
        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Total       "
        ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(Total), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Bold = True
        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
        'exclrow = exclrow + 1
        gtotal = gtotal + Total
        ExlObj.Range("A" & exclrow & ":" & "I" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
        exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 1) = " Expense       "
       ' ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!Expns
       ' ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 3) = Format(IIf(IsNull(RES!Expns), 0, (RES!Expns)), "#0.00")
        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
        'exclrow = exclrow + 1
      
        totalexp = totalexp + IIf(IsNull(RES!Expns), 0, (RES!Expns))
       ' exclrow = exclrow + 1
        RES.MoveNext
    Loop
End If
    
'ExlObj.Range("A" & exclrow & ":" & "I" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
''ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
exclrow = exclrow + 1
'expsql = "SELECT EXP_NAME, sum(ExpAmt) as EXPAmount FROM expmaster AS m, expense AS e ,rpt rt Where m.exp_code = val(e.ExpCode) and chr(rt.Trip_Master_ID) =chr(e.TripMasterReferenceId)  and e.DATE BETWEEN DATEVALUE('" & DTFrom.Value & "')AND DATEVALUE('" & DTTo.Value & "') " & cond & " group by EXP_NAME"
'expsql = " SELECT ExpName, sum(ExpAmt) AS EXPAmount FROM expense AS e, rpt AS rt WHERE cstr(rt.Trip_Master_ID) =cstr(e.TripMasterReferenceId)  and e.DATE BETWEEN DATEVALUE('" & DTFrom.Value & "')AND DATEVALUE('" & DTTo.Value & "') " & cond & "  GROUP BY ExpName"
''expsql = "select sum(EXPENSE) from RPT where TripNo=" &  & ""
'Set RES2 = CNN.OpenRecordset(expsql, dbOpenDynaset)
'If Not RES2.EOF Then
'    While Not RES2.EOF
'    exclrow = exclrow + 1
'        ExlObj.ActiveSheet.Cells(exclrow, 1) = RES2!expname
 '       ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(totalexp), 2)
  '      ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
       ' totalexp = totalexp + RES2!EXPAmount
        'RES2.MoveNext
        'exclrow = exclrow + 1
 '   Wend
'End If
ExlObj.Range("A" & exclrow & ":" & "I" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
'ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
exclrow = exclrow + 1
ExlObj.ActiveSheet.Cells(exclrow, 1) = "Total Expense  "
ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(totalexp), 2)
ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Bold = True
ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
exclrow = exclrow + 1
ExlObj.ActiveSheet.Cells(exclrow, 1) = "Amount To Remit  "
ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(gtotal - totalexp - refund), 2)
ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Bold = True
ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
'With Worksheets("Sheet1").Columns("A")
'    .ColumnWidth = (.ColumnWidth * 1.75) + 5
'End With
'With Worksheets("Sheet1").Columns("B")
'    .ColumnWidth = .ColumnWidth * 1.75
'End With
'With Worksheets("Sheet1").Columns("C")
'    .ColumnWidth = .ColumnWidth * 1.75
'End With
'With Worksheets("Sheet1").Columns("H")
'    .ColumnWidth = .ColumnWidth * 1.75
'End With
'With Worksheets("Sheet1").Columns("I")
'    .ColumnWidth = .ColumnWidth * 1.75
'End With

If checkflag = True Then
    ExlObj.ActiveSheet.Name = FnameUp
    ExlObj.ActiveWorkbook.SaveAs TcketPath & "\" & FnameUp
   ExlObj.ActiveWorkbook.Close False
    MsgBox "Report Exported Successfully", vbOKOnly, gblstrPrjTitle
Else
      Shell "taskkill /f /im ""EXCEL.exe"""
      MsgBox "No data for export", vbOKOnly, gblstrPrjTitle
     
End If
Exit Sub

erromod:
If err.Number = 429 Then
    MsgBox "To export data MS Office Excel should be installed.!", vbExclamation, PrjTitleMsg
ElseIf InStr(1, err.Description, "cannot find the file specified") > 0 Then
    MsgBox "To export data Please install MS Office properly.!", vbExclamation, gblstrPrjTitle

Else
    MsgBox "Error occured Due to " & err.Description, vbExclamation, gblstrPrjTitle
End If
End Sub



