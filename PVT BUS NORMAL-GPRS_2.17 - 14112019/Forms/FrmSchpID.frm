VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmSchpID 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Schedule Report"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSchpID 
      BackColor       =   &H00E0E0E0&
      Height          =   2460
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   5715
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
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   375
         Left            =   2850
         TabIndex        =   1
         Top             =   1800
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
         MICON           =   "FrmSchpID.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSchRpt 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1800
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
         MICON           =   "FrmSchpID.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   330
         Left            =   4080
         TabIndex        =   8
         Top             =   600
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
      Begin MSComCtl2.DTPicker DTfrom 
         Height          =   330
         Left            =   1365
         TabIndex        =   9
         Top             =   600
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
      Begin VB.Label lblfromdate 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
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
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1650
      End
      Begin VB.Label lbltodate 
         BackStyle       =   0  'Transparent
         Caption         =   "   End Date"
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
         Left            =   2805
         TabIndex        =   10
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label lblPalmId 
         BackStyle       =   0  'Transparent
         Caption         =   "PalmtecID  "
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
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Label lblSchedule 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule "
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
         Left            =   3000
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Schedule Report"
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
      TabIndex        =   7
      Top             =   -495
      Width           =   3720
   End
End
Attribute VB_Name = "frmSchpID"
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
        RES.Close
End Sub

Private Sub cmdSchRpt_Click()
On Error Resume Next
  Call SchRptCnv
    
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

Public Function SchRpt() As Boolean

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
  'Dim totalexp As Single
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

Private Sub cmdSchsmryRpt_Click()

End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub DTTo_Change()
   sql = "SELECT DISTINCT PALMID FROM RPT WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
   Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
  ' DTTo.Value = DateValue(Format(Now, "DD/MM/YYYY"))
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


End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Icon = frmMainform.Icon
    CONNECTDB
    sql = "SELECT DISTINCT PALMID FROM RPT "
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
   
    
        'DTEnd.Visible = False
        'DTStart.Visible = False
        
        DTfrom.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
        DTTo.Value = DateValue(Format(Now, "DD/MM/YYYY"))
     
        
        
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
      
        sql = "SELECT DISTINCT SCHEDULE  FROM RPT WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
        'WHERE PALMID='" & cmbPalmID.Text & "'"
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
'        'cmbrouteno.Clear
'        cmbrouteno.AddItem "ALL"
'        Do While Not RES.EOF
'            cmbrouteno.AddItem RES!RouteCode
'            RES.MoveNext
'        Loop
'        If cmbrouteno.ListCount > 0 Then
'            cmbrouteno.Text = cmbrouteno.List(0)
'        End If
        
    
   
        
'        sql = "SELECT DISTINCT TRIPNO  FROM RPT " ' WHERE PALMID='" & cmbPalmID.Text & "'"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        If RES.RecordCount > 0 Then RES.MoveFirst
'        cmbtripno.Clear
'        cmbtripno.AddItem "ALL"
'        Do While Not RES.EOF
'            cmbtripno.AddItem RES!TripNo
'            RES.MoveNext
'        Loop
'        If cmbtripno.ListCount > 0 Then
'            cmbtripno.Text = cmbtripno.List(0)
'        End If
      
'        If optReportType(0).Value = True Then                       ''rnc
'            lbltripno.Visible = True
'            cmbtripno.Visible = True
'    End If
    'RES.Close
End Sub


Public Sub SchRptCnv()
On Error GoTo erromod
Dim FS As New FileSystemObject
Dim fShndl As Integer
Dim SysD, FnameUp As String
Dim qry, sql, str, sSQL, Newstr, subsql As String
Dim HDR1, HDR2 As String
Dim pamt As Integer
Dim Total, ctotal As Double, totaladj As Double
Dim totalcnt As Double
Dim totcol As Double
Dim totalcol As Double
Dim fullcount As Double
Dim fullcoln As Double
Dim halfcount As Double
Dim halfcoln As Double
Dim stcount As Double
Dim stcoln As Double
Dim phycount As Double
Dim phycoln As Double
Dim Luggcount As Double
Dim Luggcoln As Double
Dim Adjcoln As Double
Dim passcount As Double
Dim Totc As Double
Dim checkflag As Boolean
Dim checkme As Boolean
Dim expsql As String
Dim expsql1 As String
Dim cond As String
Dim totalexp As Single
Dim mystr As String, lad_ct As Long, sc_ct As Long, lad_co As Double, sc_co As Double
'open file
checkflag = False
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
       
            FnameUp = "SCHEDULE REPORT"
       
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
        
       ExlObj.ActiveSheet.Cells(1, 5).Value = HDR1
       ExlObj.Range("A1:Q1").MergeCells = True
       ExlObj.ActiveSheet.Cells(1, 5).HorizontalAlignment = xlCenter
       ExlObj.Range("A1:Q1").MergeCells = True
       ExlObj.Range("1:5").Font.FontStyle = "Bold"
       ExlObj.ActiveSheet.Cells(2, 5).Value = HDR2
       ExlObj.ActiveSheet.Cells(2, 5).HorizontalAlignment = xlCenter
       ExlObj.Range("2:5").Font.FontStyle = "Bold"
      
            ExlObj.Range("A2:Q2").MergeCells = True
            ExlObj.Range("A3:Q3").MergeCells = True
            ExlObj.Range("A3:Q3").Value = "SCHEDULE REPORT"
            ExlObj.Range("A3:Q3").HorizontalAlignment = xlCenter
            ExlObj.Range("A3:Q3").Font.FontStyle = "Bold"
        str = ""
        sql = ""
        expsql = ""
        expsql1 = ""
        str = "SELECT PalmID,STicket,ETicketNo,SCHEDULE,StartDate,Conductor,Driver,Cleaner FROM RPT  WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
        'sql = "SELECT RT.* FROM RPT RT WHERE DATE BETWEEN DATEVALUE('" & DTFrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
        sql = "SELECT TripNO,RouteCode,UpDownTrip,STicket,ETicketNo,StartDate,Conductor,Driver,Cleaner,PalmID, Sum(Fulls) AS fullcnt, Sum(fullcoll) AS fullcln," _
          & "Sum(Luggage) AS Luggcnt, Sum(Half) AS halfcnt, Sum(AdjustColl) AS Adj,Sum(Halfcoll) AS halfcln, Sum(LuggageColl) AS Luggagecln," _
          & " Sum(Phy) AS phycnt, Sum(St) AS Stcnt, Sum(Stcoll) AS Stcln,Sum(Phycoll) AS Phycl, Sum(pass) AS passcnt, Sum(TotalColl) AS totcol" _
          & ",Sum(RPT.ladies_count) AS lad_cout, Sum(RPT.ladies_coll) AS lad_cole,Sum(RPT.senior_count) AS sc_count, Sum(RPT.senior_coll) AS sc_col " _
          & " FROM RPT  WHERE  DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
        
        
         If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
            str = str & " AND PalmID='" & cmbPalmID & "'"
'         Else
'             str = str & "AND PalmID='" & Trim(TrimChr(RES3!PalmId)) & "'"
         End If
             
         If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
            str = str & " AND SCHEDULE=" & val(cmbSchedule) & ""
'         Else
'            str = str & " AND SCHEDULE=" & RES3!SCHEDULE & " "
            
          End If
        str = str & " ORDER BY StartDate,PALMID,SCHEDULE"
        Set RES3 = CNN.OpenRecordset(str, dbOpenDynaset)
             
        mystr = "select sum(refundamt) as refund from tkts where datevalue(date) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
        expsql = "SELECT Sum(ExpAmt) as Expamt1,ExpName FROM EXPENSE  WHERE DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
        expsql1 = "SELECT Sum(ExpAmt) as EXP FROM EXPENSE  WHERE DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
        If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
            sql = sql & " AND PalmID='" & cmbPalmID & "'"
            expsql = expsql & " AND PalmID='" & cmbPalmID & "'"
            expsql1 = expsql1 & " AND PalmID='" & cmbPalmID & "'"
            'str = str & " AND PalmID='" & cmbPalmID & "'"
            cond = cond & " AND rt.PalmID='" & cmbPalmID & "'"
            mystr = mystr & " AND PalmID='" & cmbPalmID & "'"
       
        Else
        If RES3.EOF <> True Then
            sql = sql & " AND PalmID='" & Trim(TrimChr(RES3!PalmID)) & "'"
            expsql = expsql & " AND PalmID='" & Trim(TrimChr(RES3!PalmID)) & "'"
            expsql1 = expsql1 & " AND PalmID='" & Trim(TrimChr(RES3!PalmID)) & "'"
            mystr = mystr & " AND PalmID='" & Trim(TrimChr(RES3!PalmID)) & "'"
            'str = str & "AND PalmID='" & Trim(TrimChr(RES3!PalmId)) & "'"
            End If
        End If
    
        If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
            sql = sql & " AND SCHEDULE=" & val(cmbSchedule) & ""
            expsql = expsql & " AND SCHEDULENO=" & val(cmbSchedule) & ""
            expsql1 = expsql1 & " AND SCHEDULENO=" & val(cmbSchedule) & ""
            'str = str & " AND SCHEDULE=" & val(cmbSchedule) & ""
            cond = cond & " AND rt.SCHEDULE=" & val(cmbSchedule) & ""
            mystr = mystr & " AND SCHDULE=" & val(cmbSchedule) & ""
        Else
            If RES3.EOF <> True Then
            sql = sql & " AND SCHEDULE=" & RES3!SCHEDULE & " "
            expsql = expsql & " AND SCHEDULENO=" & RES3!SCHEDULE & " "
            expsql1 = expsql1 & " AND SCHEDULENO=" & RES3!SCHEDULE & " "
            mystr = mystr & " AND SCHDULE=" & RES3!SCHEDULE & " "
            End If
            'str = str & " AND SCHEDULE=" & RES3!SCHEDULE & " "
        End If
    
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
           
             sql = sql & " GROUP BY StartDate,PalmID,Conductor,Driver,Cleaner,TripNo,STicket,ETicketNo,UpDownTrip,RouteCode"
             'str = str & " ORDER BY DATE,PALMID,SCHEDULE"
            expsql = expsql & " GROUP BY ExpName,Date,PalmID"
            expsql1 = expsql1 & " GROUP BY Date,PalmID"
            mystr = mystr & " GROUP BY Date,PalmID,SCHDULE"
            
           ' sql = sql & " ORDER BY RT.DATE,RT.PALMID,RT.SCHEDULE,RT.STicket"
        
        
        Dim exclrow As Integer, exclcol As Integer
exclrow = 4
Dim gtotal As Double
checkme = False
If Not RES3.EOF Then
   Do While Not RES3.EOF
    If RES3.EOF = True Then Exit Do
       
        
        
        Dim Stno, Etno As String
        Dim I As Integer

       
          exclrow = exclrow + 1
                    
            checkme = True
        
        
        'Newstr=Select * From RPT where
        ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "Date       "
        ExlObj.ActiveSheet.Cells(exclrow, 3).Value = " " & RES3!StartDate
        ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "PalmId      "
        ExlObj.ActiveSheet.Cells(exclrow, 6).Value = " " & RES3!PalmID
        exclrow = exclrow + 1
          
        sSQL = "select Count(tripno) as [CTRIP] from rpt  WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
        If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
            sSQL = sSQL & "AND PalmID='" & cmbPalmID & "'"
        Else
            sSQL = sSQL & "AND PalmID='" & Trim(TrimChr(RES3!PalmID)) & "'"
        End If
        '''" & RES3!PalmId & "'"
        
        If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
            sSQL = sSQL & " AND SCHEDULE=" & val(cmbSchedule) & " "
         Else
            sSQL = sSQL & " AND SCHEDULE=" & RES3!SCHEDULE & " "
        End If
        
            'sSQL = sSQL & "GROUP BY TRIPNO"
            
     If checkflag = True Then
     
        sql = ""
        expsql = ""
        expsql = ""
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           sql = "SELECT TripNO,RouteCode,UpDownTrip,STicket,ETicketNo,StartDate,Conductor,Driver,Cleaner,PalmID, Sum(Fulls) AS fullcnt, Sum(fullcoll) AS fullcln," _
          & "Sum(Luggage) AS Luggcnt, Sum(Half) AS halfcnt, Sum(Halfcoll) AS halfcln, Sum(LuggageColl) AS Luggagecln," _
          & " Sum(Phy) AS phycnt,Sum(AdjustColl) AS Adj,Sum(RPT.St) AS Stcnt, Sum(RPT.Stcoll) AS Stcln,Sum(Phycoll) AS Phycl, Sum(pass) AS passcnt, Sum(TotalColl) AS totcol" _
          & ",Sum(RPT.ladies_count) AS lad_cout, Sum(RPT.ladies_coll) AS lad_cole,Sum(RPT.senior_count) AS sc_count, Sum(RPT.senior_coll) AS sc_col " _
          & " FROM RPT  WHERE  DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
          
           expsql = "SELECT Sum(ExpAmt) as Expamt1,ExpName FROM EXPENSE  WHERE DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
           expsql1 = "SELECT Sum(ExpAmt) As EXP FROM EXPENSE  WHERE DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
           mystr = "select sum(refundamt) as refund from tkts where datevalue(date) BETWEEN DATEVALUE('" & DTfrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
        DoEvents
        If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
            sql = sql & " AND PalmID='" & cmbPalmID & "'"
            expsql = expsql & " AND PalmID='" & cmbPalmID & "'"
            expsql1 = expsql1 & " AND PalmID='" & cmbPalmID & "'"
            mystr = mystr & " AND PalmID='" & cmbPalmID & "'"
           
       
        Else
            sql = sql & "AND PalmID='" & Trim(TrimChr(RES3!PalmID)) & "'"
            expsql = expsql & "AND PalmID='" & Trim(TrimChr(RES3!PalmID)) & "'"
            expsql1 = expsql1 & "AND PalmID='" & Trim(TrimChr(RES3!PalmID)) & "'"
            mystr = mystr & " AND PalmID='" & Trim(TrimChr(RES3!PalmID)) & "'"
           
        End If
    DoEvents
        If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
            sql = sql & " AND SCHEDULE=" & val(cmbSchedule) & ""
            expsql = expsql & " AND SCHEDULENO=" & val(cmbSchedule) & ""
            expsql1 = expsql1 & " AND SCHEDULENO=" & val(cmbSchedule) & ""
            mystr = mystr & " AND SCHDULE=" & val(cmbSchedule) & ""
        Else
            sql = sql & " AND SCHEDULE=" & RES3!SCHEDULE & " "
            expsql = expsql & " AND SCHEDULENO=" & RES3!SCHEDULE & " "
            expsql1 = expsql1 & " AND SCHEDULENO=" & RES3!SCHEDULE & " "
            mystr = mystr & " AND SCHDULE=" & RES3!SCHEDULE & " "
        End If
        
        DoEvents
            sql = sql & " GROUP BY StartDate,PalmID,Conductor,Driver,Cleaner,TripNo,STicket,ETicketNo,RouteCode,UpDownTrip"
             'str = str & " ORDER BY DATE,PALMID,SCHEDULE"
            expsql = expsql & " GROUP BY ExpName,Date,PalmID"
             mystr = mystr & " GROUP BY Date,PalmID,SCHDULE"
           ' expsql1 = expsql1 & " GROUP BY ExpAmt"
           ' sql = sql & " ORDER BY RT.DATE,RT.PALMID,RT.SCHEDULE,RT.STicket"
        
    checkflag = False
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       End If
        Set RES5 = CNN.OpenRecordset(sql, dbOpenDynaset)
        Rcount = RES5.RecordCount
        If RES5.EOF <> True Then
            Stno = ""
            
            For I = 1 To Rcount
                If I = 1 Then Stno = RES5!STicket
                If I = Rcount Then
                    Etno = RES5!ETicketNo
                    Exit For
                End If
                RES5.MoveNext
            Next I
         End If
         RES5.Close
            
        Set res1 = CNN.OpenRecordset(sSQL, dbOpenDynaset)
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "No Of Trips   "
        ExlObj.ActiveSheet.Cells(exclrow, 3).Value = "  " & res1!CTRIP
        ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "Schedule No  "
        ExlObj.ActiveSheet.Cells(exclrow, 6).Value = RES3!SCHEDULE
        exclrow = exclrow + 1
        
        ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "Conductor   "
        ExlObj.ActiveSheet.Cells(exclrow, 3).Value = RES3!Conductor
        ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "Driver      "
        ExlObj.ActiveSheet.Cells(exclrow, 6).Value = RES3!Driver
        exclrow = exclrow + 1
        ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "ST TKTNO"
        ExlObj.ActiveSheet.Cells(exclrow, 3).Value = Stno
        ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "ET TKTNO"
        ExlObj.ActiveSheet.Cells(exclrow, 6).Value = Etno
        exclrow = exclrow + 1
         
If Not RES.EOF Then
    Total = 0
    Dim sroutename As String, sroutesql As String
         exclrow = exclrow + 1
         ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "      " & "TRIP"
        ' ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "      " & "ROUTE"
         ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "      " & "ROUTE CODE"
         ExlObj.ActiveSheet.Cells(exclrow, 3).Value = "      " & "ST TKTNO"
         ExlObj.ActiveSheet.Cells(exclrow, 4).Value = "      " & "ET TKTNO"
         ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "      " & "FULL"
         ExlObj.ActiveSheet.Cells(exclrow, 6).Value = "      " & "AMT"
         ExlObj.ActiveSheet.Cells(exclrow, 7).Value = "      " & "HALF"
         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "     " & "AMT"
         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = "     " & "ST"
         ExlObj.ActiveSheet.Cells(exclrow, 10).Value = "      " & "AMT"
         
         ExlObj.ActiveSheet.Cells(exclrow, 11).Value = "     " & "Ladies"
         ExlObj.ActiveSheet.Cells(exclrow, 12).Value = "     " & "AMT"
         
         ExlObj.ActiveSheet.Cells(exclrow, 13).Value = "     " & "SC"
         ExlObj.ActiveSheet.Cells(exclrow, 14).Value = "     " & "AMT"
         
         ExlObj.ActiveSheet.Cells(exclrow, 15).Value = "      " & "PH"
         ExlObj.ActiveSheet.Cells(exclrow, 16).Value = "      " & "AMT"
         ExlObj.ActiveSheet.Cells(exclrow, 17).Value = "      " & "PASS"
         ExlObj.ActiveSheet.Cells(exclrow, 18).Value = "      " & "AMT"
         ExlObj.ActiveSheet.Cells(exclrow, 19).Value = "      " & "LUGGAGE"
         ExlObj.ActiveSheet.Cells(exclrow, 20).Value = "      " & "AMT"
         ExlObj.ActiveSheet.Cells(exclrow, 21).Value = "      " & "ADJUST"
         ExlObj.ActiveSheet.Cells(exclrow, 22).Value = "      " & "TOTAL"
         ExlObj.ActiveSheet.Cells(exclrow, 23).Value = "      " & "AMT"
         exclrow = exclrow + 1
    While Not RES.EOF
        totcol = 0
        totalcnt = 0
        ExlObj.ActiveSheet.Cells(exclrow, 1) = RES!TripNo & "   " & IIf(RES!UpDownTrip = "U", "UP", "DOWN")
        ' ExlObj.ActiveSheet.Cells(exclrow, 1) = RES!TripNo
         ExlObj.ActiveSheet.Cells(exclrow, 1).HorizontalAlignment = xlCenter
         
         ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!RouteCode
         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
         
'         ExlObj.ActiveSheet.Cells(exclrow, 3) = RES!STicket
'         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlCenter
         
         ExlObj.ActiveSheet.Cells(exclrow, 3) = RES!STicket
         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlCenter
         
         ExlObj.ActiveSheet.Cells(exclrow, 4) = RES!ETicketNo
         ExlObj.ActiveSheet.Cells(exclrow, 4).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 5) = RES!fullcnt
         totalcnt = totalcnt + RES!fullcnt
         fullcount = fullcount + RES!fullcnt
         ExlObj.ActiveSheet.Cells(exclrow, 5).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 6) = Round(val(RES!fullcln), 2)
         totalcol = totalcol + Round(val(RES!fullcln), 2)
           totcol = totcol + Round(val(RES!fullcln), 2)
         fullcoln = fullcoln + Round(val(RES!fullcln), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 6).HorizontalAlignment = xlRight
         ExlObj.ActiveSheet.Cells(exclrow, 7) = RES!halfcnt
         totalcnt = totalcnt + RES!halfcnt
         halfcount = halfcount + RES!halfcnt
         halfcoln = halfcoln + Round(val(RES!halfcln), 2)
         totalcol = totalcol + Round(val(RES!halfcln), 2)
         totcol = totcol + Round(val(RES!halfcln), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 7).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 8) = Round(val(RES!halfcln), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlRight
         
         ExlObj.ActiveSheet.Cells(exclrow, 9) = RES!stcnt
         totalcnt = totalcnt + RES!stcnt
         stcount = stcount + RES!stcnt
         stcoln = stcoln + Round(val(RES!stcln), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 10) = Round(val(RES!stcln), 2)
         totalcol = totalcol + Round(val(RES!stcln), 2)
         totcol = totcol + Round(val(RES!stcln), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 10).HorizontalAlignment = xlRight
         'lad
         ExlObj.ActiveSheet.Cells(exclrow, 11) = RES!lad_cout
         totalcnt = totalcnt + RES!lad_cout
         lad_ct = lad_ct + RES!lad_cout
         lad_co = lad_co + Round(val(RES!lad_cole), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 11).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 12) = Round(val(RES!lad_cole), 2)
         totalcol = totalcol + Round(val(RES!lad_cole), 2)
         totcol = totcol + Round(val(RES!lad_cole), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 12).HorizontalAlignment = xlRight
         'sc
         ExlObj.ActiveSheet.Cells(exclrow, 13) = RES!sc_count
         totalcnt = totalcnt + RES!sc_count
         sc_ct = sc_ct + RES!sc_count
         sc_co = sc_co + Round(val(RES!sc_col), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 13).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 14) = Round(val(RES!sc_col), 2)
         totalcol = totalcol + Round(val(RES!sc_col), 2)
         totcol = totcol + Round(val(RES!sc_col), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 14).HorizontalAlignment = xlRight
         ''''
         ExlObj.ActiveSheet.Cells(exclrow, 15) = RES!phycnt
         totalcnt = totalcnt + RES!phycnt
         phycount = phycount + RES!phycnt
         phycoln = phycoln + Round(val(RES!Phycl), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 15).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 16) = Round(val(RES!Phycl), 2)
         totalcol = totalcol + Round(val(RES!Phycl), 2)
         totcol = totcol + Round(val(RES!Phycl), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 16).HorizontalAlignment = xlRight
         ExlObj.ActiveSheet.Cells(exclrow, 17) = RES!passcnt
         totalcnt = totalcnt + IIf(IsNull(RES!passcnt), 0, RES!passcnt)
         passcount = passcount + IIf(IsNull(RES!passcnt), 0, RES!passcnt)
         ExlObj.ActiveSheet.Cells(exclrow, 17).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 18) = 0
         ExlObj.ActiveSheet.Cells(exclrow, 18).HorizontalAlignment = xlRight
         ExlObj.ActiveSheet.Cells(exclrow, 19) = RES!luggcnt
         ExlObj.ActiveSheet.Cells(exclrow, 19).HorizontalAlignment = xlCenter
       '  totalcnt = totalcnt + RES!Luggcnt
         
         Luggcount = Luggcount + RES!luggcnt
         ExlObj.ActiveSheet.Cells(exclrow, 20) = Round(val(RES!Luggagecln), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 20).HorizontalAlignment = xlRight
         Luggcoln = Luggcoln + Round(val(RES!Luggagecln), 2)
         totalcol = totalcol + Round(val(RES!Luggagecln), 2)
         totcol = totcol + Round(val(RES!Luggagecln), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 21) = Round(val(RES!Adj), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 21).HorizontalAlignment = xlRight
         Adjcoln = Adjcoln + Round(val(RES!Adj), 2)
    
         
         
         
         ExlObj.ActiveSheet.Cells(exclrow, 22) = totalcnt
         ExlObj.ActiveSheet.Cells(exclrow, 22).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 23) = totcol '(Round(val(RES!totcol), 2) + Round(val(RES!Adj), 2)) + val(totalexp)
         ExlObj.ActiveSheet.Cells(exclrow, 23).HorizontalAlignment = xlRight
         
         
         
       
         exclrow = exclrow + 1
        Total = Total + (totcol) ' total + (RES!totcol + Round(val(RES!Adj), 2))
        totaladj = totaladj + Round(val(RES!Adj), 2)
        
        ' total = total + (RES!totcol)
         ctotal = ctotal + totalcnt
         
         
        totalcnt = 0
         'totalcol = 0
         

'''        Set RES2 = CNN.OpenRecordset(expsql, dbOpenDynaset)
                 
        RES.MoveNext
        RES3.MoveNext
    Wend
End If
                
             exclrow = exclrow + 1
             ExlObj.ActiveSheet.Cells(exclrow, 1).MergeCells = True
             ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "TOTAL"
             ExlObj.ActiveSheet.Cells(exclrow, 1).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 1).Font.FontStyle = "Bold"
'             ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(total), 2)
'             ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
             ExlObj.ActiveSheet.Cells(exclrow, 5) = Round(val(fullcount), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 5).HorizontalAlignment = xlCenter
             ExlObj.ActiveSheet.Cells(exclrow, 6) = Round(val(fullcoln), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 6).HorizontalAlignment = xlRight
             
             ExlObj.ActiveSheet.Cells(exclrow, 7) = Round(val(halfcount), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 7).HorizontalAlignment = xlCenter
             ExlObj.ActiveSheet.Cells(exclrow, 8) = Round(val(halfcoln), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlRight
             
             ExlObj.ActiveSheet.Cells(exclrow, 9) = Round(val(stcount), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlCenter
             ExlObj.ActiveSheet.Cells(exclrow, 10) = Round(val(stcoln), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 10).HorizontalAlignment = xlRight
             
             ExlObj.ActiveSheet.Cells(exclrow, 11) = Round(val(lad_ct), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 11).HorizontalAlignment = xlCenter
             ExlObj.ActiveSheet.Cells(exclrow, 12) = Round(val(lad_co), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 12).HorizontalAlignment = xlRight
             
             ExlObj.ActiveSheet.Cells(exclrow, 13) = Round(val(sc_ct), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 13).HorizontalAlignment = xlCenter
             ExlObj.ActiveSheet.Cells(exclrow, 14) = Round(val(sc_co), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 14).HorizontalAlignment = xlRight
             
             ExlObj.ActiveSheet.Cells(exclrow, 15) = Round(val(phycount), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 15).HorizontalAlignment = xlCenter
             ExlObj.ActiveSheet.Cells(exclrow, 16) = Round(val(phycoln), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 16).HorizontalAlignment = xlRight
             ExlObj.ActiveSheet.Cells(exclrow, 17) = passcount
             ExlObj.ActiveSheet.Cells(exclrow, 17).HorizontalAlignment = xlCenter
             ExlObj.ActiveSheet.Cells(exclrow, 18) = 0
             ExlObj.ActiveSheet.Cells(exclrow, 18).HorizontalAlignment = xlRight
             ExlObj.ActiveSheet.Cells(exclrow, 19) = Luggcount
             ExlObj.ActiveSheet.Cells(exclrow, 19).HorizontalAlignment = xlCenter
             ExlObj.ActiveSheet.Cells(exclrow, 20) = Round(val(Luggcoln), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 20).HorizontalAlignment = xlRight
             ExlObj.ActiveSheet.Cells(exclrow, 21) = Round(val(Adjcoln), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 21).HorizontalAlignment = xlRight
             ExlObj.ActiveSheet.Cells(exclrow, 22) = ctotal
             ExlObj.ActiveSheet.Cells(exclrow, 22).HorizontalAlignment = xlCenter
             ExlObj.ActiveSheet.Cells(exclrow, 23) = Round(val(Total), 2)
             ExlObj.ActiveSheet.Cells(exclrow, 23).HorizontalAlignment = xlRight

    Set RES2 = CNN.OpenRecordset(expsql, dbOpenDynaset)
    Set RES14 = CNN.OpenRecordset(expsql1, dbOpenDynaset)
    Set res9 = CNN.OpenRecordset(mystr, dbOpenDynaset)
    'exclrow = exclrow + 1              EXP
    exclrow = exclrow + 1
If Not RES14.EOF Then
    While Not RES14.EOF
    
             ExlObj.ActiveSheet.Cells(exclrow, 1).MergeCells = True
             ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "EXPENSES:"
             ExlObj.ActiveSheet.Cells(exclrow, 1).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 1).Font.FontStyle = "Bold"
            If IsNull(RES14!exp) = True Then
             ExlObj.ActiveSheet.Cells(exclrow, 2) = 0
             
             Else
                 ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(RES14!exp), 2)
                 'ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlgnment = xlRight
            End If
             
             
        totalexp = Round(val(IIf(IsNull(RES14!exp), 0, (RES14!exp))), 2)
        
        ' total = total + totalexp
       RES14.MoveNext
    Wend
End If
 While Not RES2.EOF
    
             
             

    
             exclrow = exclrow + 1
             
    
         ExlObj.ActiveSheet.Cells(exclrow, 1) = RES2!expname
         ExlObj.ActiveSheet.Cells(exclrow, 1).HorizontalAlignment = xlCenter
         ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(RES2!ExpAmt1), 2)
         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
         
    
  
     RES2.MoveNext
            Wend
  
    
    exclrow = exclrow + 1
    exclrow = exclrow + 1
            ' ExlObj.Range("B" & exclrow & ":C" & exclrow).MergeCells = True
             ExlObj.ActiveSheet.Cells(exclrow, 2).MergeCells = True
             ExlObj.ActiveSheet.Cells(exclrow, 2).Value = " TOTAL COLLECTION IN SCHEDULE"
             ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 2).Font.FontStyle = "Bold"
             ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "Rs"
             ExlObj.ActiveSheet.Cells(exclrow, 5).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.FontStyle = "Bold"
             ExlObj.ActiveSheet.Cells(exclrow, 6) = Total
             ExlObj.ActiveSheet.Cells(exclrow, 6).HorizontalAlignment = xlRight
             exclrow = exclrow + 1
             
             
             'Format(IIf(IsNull(san!EXPAmount), 0, (san!EXPAmount)), "#0.00")
             ExlObj.ActiveSheet.Cells(exclrow, 2).MergeCells = True
             ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "TOTAL EXPENSE"
             ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 2).Font.FontStyle = "Bold"
             ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "Rs"
             ExlObj.ActiveSheet.Cells(exclrow, 5).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.FontStyle = "Bold"
             
             
             'ExlObj.ActiveSheet.Cells(exclrow, 6) = IIf(IsNull(totalexp), 0, totalexp)
             ExlObj.ActiveSheet.Cells(exclrow, 6) = Format(IIf(IsNumeric(totalexp), totalexp, 0), "#0.00")
             ExlObj.ActiveSheet.Cells(exclrow, 6).HorizontalAlignment = xlRight
             
             exclrow = exclrow + 1
             
             
             ExlObj.ActiveSheet.Cells(exclrow, 2).MergeCells = True
             ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "TOTAL REFUND"
             ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 2).Font.FontStyle = "Bold"
             ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "Rs"
             ExlObj.ActiveSheet.Cells(exclrow, 5).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.FontStyle = "Bold"
              ExlObj.ActiveSheet.Cells(exclrow, 6) = Format(IIf(IsNumeric(res9("refund")), res9("refund"), 0), "#0.00")
             ExlObj.ActiveSheet.Cells(exclrow, 6).HorizontalAlignment = xlRight
             
             exclrow = exclrow + 1
             
             ExlObj.ActiveSheet.Cells(exclrow, 2).MergeCells = True
             ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "NET AMT IN SCHEDULE"
             ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 2).Font.FontStyle = "Bold"
             ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "Rs"
             ExlObj.ActiveSheet.Cells(exclrow, 5).HorizontalAlignment = xlLeft
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.FontStyle = "Bold"
             ExlObj.ActiveSheet.Cells(exclrow, 6) = Total - totalexp - totaladj - res9("refund")
             ExlObj.ActiveSheet.Cells(exclrow, 6).HorizontalAlignment = xlRight

             ExlObj.Range("A" & exclrow & ":" & "Z" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
                 
                 
           '   exclrow = exclrow + 1
              exclrow = exclrow + 1
              
              
totalcnt = 0
fullcount = 0
totalcol = 0
fullcoln = 0
halfcount = 0
halfcoln = 0
stcount = 0
stcoln = 0
phycount = 0
phycoln = 0
passcount = 0
Luggcount = 0
Luggcoln = 0
totalexp = 0
Total = 0
Adjcoln = 0
ctotal = 0
        checkflag = True
       ' RES.MoveNext
       ' RES14.MoveNext
       ' RES2.MoveNext
       ' RES3.MoveNext
    Loop
End If
'Dim totalexp As Double


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
If checkme = True Then
    ExlObj.ActiveSheet.Name = FnameUp
    ExlObj.ActiveWorkbook.SaveAs TcketPath & "\" & FnameUp
 'ExlObj.ActiveWorkbook.SaveAs ExcelPath
    ExlObj.ActiveWorkbook.Close False
'ExlObj.ActiveWorkbook.Close False
    
    MsgBox "Report Exported Successfully", vbOKOnly, gblstrPrjTitle
Else
     Shell "taskkill /f /im ""EXCEL.exe"""
     MsgBox "No data for export", vbOKOnly, gblstrPrjTitle
     'ExlObj.ActiveWorkbook.Close
     
'     Set ExlObj.ActiveSheet = Nothing
'     ExlObj.ActiveWorkbook.Close (False)
'     Set ExlObj.ActiveWorkbook = Nothing
'     ExlObj.Quit
'     Set ExcelObj = Nothing
    
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


