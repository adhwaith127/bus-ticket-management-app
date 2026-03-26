VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmSchedule1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Schedule Report"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSchpID 
      BackColor       =   &H00E0E0E0&
      Height          =   3900
      Left            =   120
      TabIndex        =   0
      Top             =   675
      Width           =   5715
      Begin VB.ComboBox cmbtripno 
         Height          =   315
         Left            =   4320
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cmbSchedule 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbPalmID 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton optReportType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Schedule Wise"
         Height          =   300
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton optReportType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Trip Wise"
         Height          =   300
         Index           =   1
         Left            =   2415
         TabIndex        =   4
         Top             =   225
         Width           =   1590
      End
      Begin VB.ComboBox cmbrouteno 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   2520
         Width           =   855
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   450
         Left            =   3570
         TabIndex        =   2
         Top             =   3240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   794
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Copy (2) of frmSchedule.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSchsmryRpt 
         Height          =   450
         Left            =   2400
         TabIndex        =   3
         Top             =   3240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   794
         TX              =   "&OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Copy (2) of frmSchedule.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   345
         Left            =   1560
         TabIndex        =   13
         Top             =   960
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   609
         _Version        =   393216
         Format          =   20905985
         CurrentDate     =   39536
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   330
         Left            =   4320
         TabIndex        =   15
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Format          =   20905985
         CurrentDate     =   39536
      End
      Begin VB.Label lblTripno 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "    Trip No  :     "
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
         Left            =   2880
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbltodate 
         BackStyle       =   0  'Transparent
         Caption         =   "   To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2880
         TabIndex        =   14
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label lblfromdate 
         BackStyle       =   0  'Transparent
         Caption         =   "        From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -240
         TabIndex        =   12
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label lblPalmId 
         BackStyle       =   0  'Transparent
         Caption         =   "       PalmtecID  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -240
         TabIndex        =   10
         Top             =   1560
         Width           =   1770
      End
      Begin VB.Label lblSchedule 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "        Schedule :"
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
         Left            =   -120
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblRouteNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "            Route   :     "
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
         Left            =   -120
         TabIndex        =   8
         Top             =   2520
         Width           =   1695
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Schedule Summary Report"
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
      Left            =   765
      TabIndex        =   11
      Top             =   105
      Width           =   3720
   End
End
Attribute VB_Name = "frmSchedule1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
    CONNECTDB
    sql = "SELECT DISTINCT PALMID FROM RPT"
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
   
    If optReportType(0).Value = True Then
        'DTEnd.Visible = False
        'DTStart.Visible = False
        
        DTFrom.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
        DTTo.Value = DateValue(Format(Now, "DD/MM/YYYY"))
        'DTFrom.Top = 960
        'DTFrom.Left = 2000
        'DTTo.Top = 960
        'DTTo.Left = 4080

        'cmbrouteno.Top = 2520
        'cmbrouteno.Left = 2000
        'lblSDate.Visible = False
        'lblEDate.Visible = False
        lbltripno.Visible = False
        cmbtripno.Visible = False
        
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
            cmbPalmID.AddItem RES!PalmId
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
        
        sql = "SELECT DISTINCT RouteCode  FROM RPT " ' WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbrouteno.Clear
        cmbrouteno.AddItem "ALL"
        Do While Not RES.EOF
            cmbrouteno.AddItem RES!RouteCode
            RES.MoveNext
        Loop
        If cmbrouteno.ListCount > 0 Then
            cmbrouteno.Text = cmbrouteno.List(0)
        End If
        
    
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
    
        DTFrom.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
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
            cmbPalmID.AddItem RES!PalmId
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
On Error Resume Next
Dim FS As New FileSystemObject
Dim fShndl As Integer
Dim SysD, FnameUp As String
Dim qry, sql, sSQL, subsql As String
Dim HDR1, HDR2 As String
Dim pamt As Integer
Dim total As Double
Dim expsql As String
Dim cond As String
'open file
        
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
            FnameUp = "TRIP WISE SCHEDULE DETAILS"
        End If
        '& Fint & "-" & SysT & PID
        'If Dir(TcketPath & "\" & FnameUp) <> "" Then Kill TcketPath & "\" & FnameUp
'''        fShndl = FreeFile()
    If Dir(TcketPath & "\" & FnameUp & ".xlsx", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xlsx")
    If Dir(TcketPath & "\" & FnameUp & ".xls", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xls")
        Dim ExlObj As New excel.Application
        ExlObj.Workbooks.Add
        'ExlObj.Visible = True
       

    'print HEADER FROM SETTINGS
        qry = "SELECT HEADER1,HEADER2 FROM SETTINGS"
        Set RES = CNN.OpenRecordset(qry, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            HDR1 = RES!HEADER1
            HDR2 = RES!HEADER2
        End If
        RES.Close
        
       ExlObj.ActiveSheet.Cells(2, 5).Value = HDR1
       ExlObj.ActiveSheet.Cells(2, 5).Font.Bold = True
       ExlObj.ActiveSheet.Cells(3, 5).Value = HDR2
       ExlObj.ActiveSheet.Cells(3, 5).Font.Bold = True
       If optReportType(0).Value = True Then
            ExlObj.ActiveSheet.Cells(4, 5).Value = "Schedule wise Summary Report"
       Else
            ExlObj.ActiveSheet.Cells(4, 5).Value = "Trip wise Summary Report"
       End If
       ExlObj.ActiveSheet.Cells(4, 5).Font.Bold = True
       ExlObj.ActiveSheet.Range("C5:G5").Value = "___________________"

        sql = "SELECT RT.* FROM RPT RT WHERE DATE BETWEEN DATEVALUE('" & DTFrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "

        If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
            sql = sql & " AND PalmID='" & cmbPalmID & "'"
            cond = cond & " AND rt.PalmID='" & cmbPalmID & "'"
        End If
    
        If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
            sql = sql & " AND SCHEDULE=" & val(cmbSchedule) & ""
            cond = cond & " AND rt.SCHEDULE=" & val(cmbSchedule) & ""
        End If
    
        If cmbrouteno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
            sql = sql & " AND ROUTECODE='" & cmbrouteno & "'"
            cond = cond & " AND rt.ROUTECODE='" & cmbrouteno & "'"
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If cmbtripno.ListIndex <> -1 And cmbtripno.Text <> "ALL" Then
            sql = sql & " AND TRIPNO=" & val(cmbtripno) & " "
            cond = cond & " AND rt.TRIPNO=" & val(cmbtripno) & ""
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If optReportType(0).Value = True Then
            sql = sql & " ORDER BY RT.DATE,RT.PALMID,RT.SCHEDULE,RT.STicket"
        Else
            sql = sql & " ORDER BY RT.DATE,RT.PALMID,RT.SCHEDULE, RT.TRIPNO,RT.STicket"
        End If
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        Dim exclrow As Integer, exclcol As Integer
exclrow = 6
Dim gtotal As Double
If Not RES.EOF Then

    While Not RES.EOF
    total = 0
          ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
          exclrow = exclrow + 1
          ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Date       "
          ExlObj.ActiveSheet.Cells(exclrow, 2).Value = " " & RES!Date
          exclrow = exclrow + 1
          ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "PalmId      "
          ExlObj.ActiveSheet.Cells(exclrow, 2).Value = RES!PalmId
          ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Schedule No  "
          ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!SCHEDULE
          ExlObj.ActiveCell(exclrow, 9).Style.NumberFormat = "@"
          exclrow = exclrow + 1
          If optReportType(0).Value = True Then
                sSQL = "select count(tripno) as 0 from rpt rt WHERE DATE BETWEEN DATEVALUE('" & DTFrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') and SCHEDULE=" & RES!SCHEDULE
                If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
                    sSQL = sSQL & " AND PalmID='" & cmbPalmID & "'"
                End If
                If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
                sSQL = sSQL & " AND SCHEDULE=" & val(cmbSchedule) & " "
                End If
                If cmbrouteno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
                sSQL = sSQL & " AND ROUTECODE='" & cmbrouteno & "'"
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If cmbtripno.ListIndex <> -1 And cmbtripno.Text <> "ALL" Then
                sSQL = sSQL & " AND TRIPNO=" & val(cmbtripno) & ""
                End If
          End If '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
          If optReportType(0).Value = True Then
                sSQL = sSQL
          Else
                sSQL = sSQL
          End If
         Set res1 = CNN.OpenRecordset(sSQL, dbOpenDynaset)
         If optReportType(0).Value = True Then
            ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "No Of Trips   "
            ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "  " & res1!CTRIP
         End If
         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Route No      "
         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!RouteCode
         exclrow = exclrow + 1
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
             ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Route No. & Name            "
             ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!RouteCode & "   " & RESROUTE!rutname
             Set RESROUTE = Nothing
             exclrow = exclrow + 1
         End If
         ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Start TktNo    "
         ExlObj.ActiveSheet.Cells(exclrow, 2).Value = RES!STicket
         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "End Tkt No     "
         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!ETicketNo
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "      " & "Count"
         ExlObj.ActiveSheet.Cells(exclrow, 3).Value = "      " & "Amount"
         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "      " & "Count"
         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = "      " & "Amount"
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Range("B" & exclrow & ":" & "C" & exclrow).Value = "_______________"
         ExlObj.ActiveSheet.Range("H" & exclrow & ":" & "I" & exclrow).Value = "_______________"
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 1) = " Full       "
         ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!Full
         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!FullColl), 2)
         'ExlObj.ActiveSheet.Cells(exclrow, 3).Align = 6
         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
         ExlObj.ActiveSheet.Cells(exclrow, 7) = " Half       "
         ExlObj.ActiveSheet.Cells(exclrow, 8) = RES!Half
         ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 9) = Round(val(RES!HalfColl), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlRight
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 1) = " ST       "
         ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!st
         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!STColl), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
         ExlObj.ActiveSheet.Cells(exclrow, 7) = " PH       "
         ExlObj.ActiveSheet.Cells(exclrow, 8) = RES!Phy
         ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 9) = Round(val(RES!PhyColl), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlRight
         exclrow = exclrow + 1
         
         
'''         subsql = "select count(Pass) as cpass from rpt rt WHERE Pass <> '' "
'''         If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
'''                subsql = subsql & " AND PalmID='" & cmbPalmID & "'"
'''            End If
'''            If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
'''                subsql = subsql & " AND SCHEDULE= " & val(cmbSchedule) & ""
'''            End If
'''            If cmbrouteno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
'''                subsql = subsql & " AND ROUTECODE= '" & cmbrouteno & "'"
'''            End If
'''            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''            If cmbtripno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
'''                subsql = sSQL & " AND TRIPNO=" & val(cmbtripno) & ""
'''        End If
'''
''''        If optReportType(0).Value = True Then
''''                subsql = sSQL
''''          Else
''''                subsql = sSQL
''''          End If
'''        Set RES2 = CNN.OpenRecordset(subsql, dbOpenDynaset)
        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Pass       "
        ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!pass
        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(pamt), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
        exclrow = exclrow + 1
        'new
        
        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Luggage       "
        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!LuggageColl), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
        exclrow = exclrow + 1
        
        
        ''''''''''''''''''''''''
        
        
        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Adjust       "
        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!AdjustColl), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
        exclrow = exclrow + 1
        ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
        exclrow = exclrow + 1
        total = total + RES!FullColl + RES!HalfColl + RES!STColl + RES!PhyColl + RES!LuggageColl - RES!AdjustColl
        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Total       "
        ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
        ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(total), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Bold = True
        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
        exclrow = exclrow + 1
        gtotal = gtotal + total
        RES.MoveNext
    Wend
End If
    

ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
exclrow = exclrow + 1
Dim totalexp As Double
totalexp = 0
'expsql = "SELECT EXP_NAME, sum(ExpAmt) as EXPAmount FROM expmaster AS m, expense AS e ,rpt rt Where m.exp_code = val(e.ExpCode) and chr(rt.Trip_Master_ID) =chr(e.TripMasterReferenceId)  and e.DATE BETWEEN DATEVALUE('" & DTFrom.Value & "')AND DATEVALUE('" & DTTo.Value & "') " & cond & " group by EXP_NAME"
expsql = " SELECT ExpName, sum(ExpAmt) AS EXPAmount FROM expense AS e, rpt AS rt WHERE chr(rt.Trip_Master_ID) =chr(e.TripMasterReferenceId)  and e.DATE BETWEEN DATEVALUE('" & DTFrom.Value & "')AND DATEVALUE('" & DTTo.Value & "') " & cond & "  GROUP BY ExpName"
Set RES2 = CNN.OpenRecordset(expsql, dbOpenDynaset)
If Not RES2.EOF Then
    While Not RES2.EOF
        ExlObj.ActiveSheet.Cells(exclrow, 1) = RES2!expname
        ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(RES2!EXPAmount), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
        totalexp = totalexp + RES2!EXPAmount
        RES2.MoveNext
        exclrow = exclrow + 1
    Wend
End If
ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
exclrow = exclrow + 1
ExlObj.ActiveSheet.Cells(exclrow, 1) = "Total Expense  "
ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(totalexp), 2)
ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Bold = True
ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
exclrow = exclrow + 1
ExlObj.ActiveSheet.Cells(exclrow, 1) = "Amount To Remit  "
ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(gtotal - totalexp), 2)
ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Bold = True
ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
With Worksheets("Sheet1").Columns("A")
    .ColumnWidth = (.ColumnWidth * 1.75) + 5
End With
With Worksheets("Sheet1").Columns("B")
    .ColumnWidth = .ColumnWidth * 1.75
End With
With Worksheets("Sheet1").Columns("C")
    .ColumnWidth = .ColumnWidth * 1.75
End With
With Worksheets("Sheet1").Columns("H")
    .ColumnWidth = .ColumnWidth * 1.75
End With
With Worksheets("Sheet1").Columns("I")
    .ColumnWidth = .ColumnWidth * 1.75
End With
ExlObj.ActiveSheet.Name = FnameUp
ExlObj.ActiveWorkbook.SaveAs TcketPath & "\" & FnameUp
ExlObj.ActiveSheet.Close
ExlObj.Workbooks.Close
'ExlObj.ActiveWorkbook.Close False
MsgBox "Report Exported Successfully"
Exit Sub
End Sub

