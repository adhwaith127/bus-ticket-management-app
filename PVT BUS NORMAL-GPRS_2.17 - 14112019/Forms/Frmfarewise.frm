VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form Frmfarewise 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fare Wise Report"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSchpID 
      BackColor       =   &H00E0E0E0&
      Height          =   3060
      Left            =   0
      TabIndex        =   9
      Top             =   -120
      Width           =   5955
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
         TabIndex        =   4
         Top             =   1560
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
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
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
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   375
         Left            =   3450
         TabIndex        =   7
         Top             =   2400
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
         MICON           =   "Frmfarewise.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSchRpt 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   2400
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
         MICON           =   "Frmfarewise.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   330
         Left            =   3840
         TabIndex        =   1
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
         Format          =   96927745
         CurrentDate     =   39536
      End
      Begin MSComCtl2.DTPicker DTfrom 
         Height          =   330
         Left            =   1365
         TabIndex        =   0
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
         Format          =   96927745
         CurrentDate     =   39536
      End
      Begin CCRProgressBar6.ccrpProgressBar CCPB_Pbr 
         Height          =   255
         Left            =   120
         Top             =   2040
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   450
         FillColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin JeweledBut.JeweledButton cmdShow 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   2400
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         TX              =   "&Show"
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
         MICON           =   "Frmfarewise.frx":0038
         BC              =   12632256
         FC              =   0
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
         Left            =   1320
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
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
         Left            =   2760
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1770
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
         Left            =   2565
         TabIndex        =   11
         Top             =   600
         Width           =   1290
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
         TabIndex        =   10
         Top             =   600
         Width           =   1650
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fare Wise Report"
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
      Left            =   1440
      TabIndex        =   8
      Top             =   -480
      Width           =   3720
   End
End
Attribute VB_Name = "Frmfarewise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myrs As New ADODB.Recordset

Private Sub cmbPalmID_Change()
If cmbPalmID.Text <> "All" And cmbPalmID.Text <> "" Then
    FillComboPrivate cmbSchedule, "SELECT DISTINCT SCHNO FROM FARERPT where pid='" & cmbPalmID.Text & "' and DATEVALUE(SCH_STDATE) between  DATEVALUE('" & DTfrom.Value & " ') and  DATEVALUE('" & DTTo.Value & "')", False, "All"
End If
End Sub

Private Sub cmbPalmID_Click()
If cmbPalmID.Text <> "All" And cmbPalmID.Text <> "" Then
   FillComboPrivate cmbSchedule, "SELECT DISTINCT SCHNO FROM FARERPT where pid='" & cmbPalmID.Text & "' and  DATEVALUE(SCH_STDATE) between  DATEVALUE('" & DTfrom.Value & " ') and  DATEVALUE('" & DTTo.Value & "')", False, "All"
End If
End Sub



Private Sub cmbPalmID_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmbSchedule.SetFocus

End Sub

Private Sub cmbSchedule_Change()
If cmbSchedule.Text <> "All" And cmbSchedule.Text <> "" And cmbPalmID.Text <> "All" And cmbPalmID.Text <> "" Then
    FillComboPrivate cmbtripno, "SELECT TRPNO TripNo FROM FARERPT where SCHNO=" & cmbSchedule.Text & " and pid='" & cmbPalmID.Text & "' and DATEVALUE(SCH_STDATE) between  DATEVALUE('" & DTfrom.Value & " ') and  DATEVALUE('" & DTTo.Value & "')", False, "All"
ElseIf cmbSchedule.Text <> "All" And cmbSchedule.Text <> "" Then
    FillComboPrivate cmbtripno, "SELECT DISTINCT TRPNO FROM FARERPT where SCHNO=" & cmbSchedule.Text & " and DATEVALUE(SCH_STDATE) between  DATEVALUE('" & DTfrom.Value & " ') and  DATEVALUE('" & DTTo.Value & "')", False, "All"
End If
End Sub

Private Sub cmbSchedule_Click()
If cmbSchedule.Text <> "All" And cmbSchedule.Text <> "" And cmbPalmID.Text <> "All" And cmbPalmID.Text <> "" Then
    FillComboPrivate cmbtripno, "SELECT DISTINCT TRPNO  FROM FARERPT where SCHNO=" & cmbSchedule.Text & " and pid='" & cmbPalmID.Text & "' and DATEVALUE(SCH_STDATE) between  DATEVALUE('" & DTfrom.Value & " ') and  DATEVALUE('" & DTTo.Value & "')", False, "All"
ElseIf cmbSchedule.Text <> "All" And cmbSchedule.Text <> "" Then
    FillComboPrivate cmbtripno, "SELECT DI                                                                                                                                                         STINCT TRPNO FROM FARERPT where SCHNO=" & cmbSchedule.Text & " and DATEVALUE(SCH_STDATE) between  DATEVALUE('" & DTfrom.Value & " ') and  DATEVALUE('" & DTTo.Value & "')", False, "All"
End If
End Sub

Private Sub cmbSchedule_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmbtripno.SetFocus
End Sub

Private Sub cmbtripno_Change()
Dim rtname As String
If cmbtripno.Text <> "All" And cmbSchedule.Text <> "All" And cmbPalmID.Text <> "All" Then
    rtname = getvalueQuery("select RouteCode from rpt where TripNo=" & cmbtripno.Text & " and SCHEDULE=" & cmbSchedule.Text & " and PalmID='" & cmbPalmID.Text & "' and DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')")
    TxtRoute.Text = getvalueQuery(" select rutname from ROUTE where rutcode='" & TrimChr(rtname) & "'")
Else
    TxtRoute.Text = ""
End If
End Sub

Private Sub cmbtripno_Click()
Dim rtname As String
'If cmbtripno.Text <> "All" And cmbSchedule.Text <> "All" And cmbPalmID.Text <> "All" Then
'    rtname = getvalueQuery("select RouteCode from rpt where TripNo=" & cmbtripno.Text & " and SCHEDULE=" & cmbSchedule.Text & " and PalmID='" & cmbPalmID.Text & "' and DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')")
'    TxtRoute.Text = getvalueQuery(" select rutname from ROUTE where rutcode='" & TrimChr(rtname) & "'")
'Else
'    TxtRoute.Text = ""
'End If
End Sub

Private Sub cmbtripno_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdShow.SetFocus
End Sub

Private Sub cmdSchRpt_Click()
On Error GoTo lblErr
Dim exclrow, exclcol As Integer
Dim FS As New FileSystemObject
Dim SysD, FnameUp As String
Dim RouteCode, TcketPath, TransPath As String
Dim luggcnt, phcnt, stcnt, fullcnt, halfcnt, passcnt As Long
Dim TotAmt As Double
CCPB_Pbr.Value = 0
Frmfarewise.Enabled = False

        TSQL = "SELECT * FROM PCSETUP"
        Set myrs = New ADODB.Recordset
        myrs.Open TSQL, gbladoCon, 2, 2
        If myrs.EOF <> True Then
            TcketPath = myrs!TICKET_PATH
            TransPath = myrs!TRANSFER_PATH
        End If
        myrs.Close
        myrs.Open
        SysD = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
        TcketPath = TcketPath & "\" & SysD
        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
       
        FnameUp = "FARE WISE REPORT"
        
        If Dir(TcketPath & "\" & FnameUp & ".xlsx", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xlsx")
        If Dir(TcketPath & "\" & FnameUp & ".xls", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xls")
        Dim ExlObj As New excel.Application
        ExlObj.Workbooks.Add
        
        ExlObj.ActiveSheet.Cells(1, 3).Value = FnameUp
        'ExlObj.Range("A1:Q1").MergeCells = True
        ExlObj.ActiveSheet.Cells(1, 3).HorizontalAlignment = xlCenter
        'ExlObj.Range("A1:Q1").MergeCells = True
         ExlObj.ActiveSheet.Cells(1, 3).Font.Bold = True
      '  ExlObj.Range("1:2").Font.FontStyle = "Bold"
       
        TSQL = ""
        TSQL = "SELECT DISTINCT tb1.FARE, (select count(FARE) from FARERPT where FARE= tb1.FARE and pid=tb1.pid and schno=tb1.schno and TRPNO=tb1.TRPNO and  DATEVALUE(SCH_STDATE) =tb1.SCH_STDATE ) AS [count], tb1.FARE * (select count(FARE) from FARERPT where FARE= tb1.FARE and pid=tb1.pid and TRPNO=tb1.TRPNO and schno=tb1.schno and  DATEVALUE(SCH_STDATE) =tb1.SCH_STDATE ) as total,pid,SCHNO,TRPNO FROM FARERPT AS tb1 where DATEVALUE(SCH_STDATE) between  DATEVALUE('" & DTfrom.Value & " ') AND  DATEVALUE('" & DTTo.Value & "')"
        
        If cmbPalmID.Text <> "All" Then
            TSQL = TSQL & " and PID='" & cmbPalmID.Text & "'"
        End If
        If cmbSchedule.Text <> "All" Then
            TSQL = TSQL & " and SCHNO=" & cmbSchedule.Text & ""
        End If
        If cmbtripno.Text <> "All" Then
            TSQL = TSQL & " and TRPNO=" & cmbtripno.Text & ""
        End If
               
        TSQL = TSQL & " ORDER BY Fare,PID,SCHNO,TRPNO"
        
        Set myres = New ADODB.Recordset
        myres.Open TSQL, gbladoCon, 2, 2
        
        If myres.EOF <> True Then
             exclrow = 3
             
             ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "PALMTEC ID"
             ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "SCHEDULE NO"
             ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 3).Value = "TRIP NO"
             ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 4).Value = "FARE"
             ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "COUNT"
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 6).Value = "TOTAL AMT"
             ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Color = vbBlack
           
           

             exclrow = exclrow + 1
           
             
            Do While Not myres.EOF
            
               
                ExlObj.ActiveSheet.Cells(exclrow, 1).Value = myres("pid")
                ExlObj.ActiveSheet.Cells(exclrow, 2).Value = myres("SCHNO")
                ExlObj.ActiveSheet.Cells(exclrow, 3).Value = myres("TRPNO")
                ExlObj.ActiveSheet.Cells(exclrow, 4).Value = myres("FARE")
                ExlObj.ActiveSheet.Cells(exclrow, 5).Value = myres("count")
                ExlObj.ActiveSheet.Cells(exclrow, 6).Value = myres("total")
                TotAmt = TotAmt + myres("total")
                
                myres.MoveNext
                
                exclrow = exclrow + 1
             If Me.CCPB_Pbr.Value >= Me.CCPB_Pbr.Max - 1 Then
                Me.CCPB_Pbr.Value = 0
             Else
                Me.CCPB_Pbr.Value = Me.CCPB_Pbr.Value + 1
             End If
            Loop
             exclrow = exclrow + 1
            

            ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "TOTAL"
            ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Color = vbBlack
            'exclrow = exclrow + 1
            
            ExlObj.ActiveSheet.Cells(exclrow, 6).Value = Format(TotAmt, "0.00")
            ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Color = vbBlack
            
            ExlObj.ActiveSheet.Name = FnameUp
            ExlObj.ActiveWorkbook.SaveAs TcketPath & "\" & FnameUp
            ExlObj.ActiveWorkbook.Close False
            Me.CCPB_Pbr.Value = Me.CCPB_Pbr.Max
            MsgBox "Report Exported Successfully", vbOKOnly, gblstrPrjTitle
            
        Else
           Shell "taskkill /f /im ""EXCEL.exe"""
           MsgBox "No data for Export", vbOKOnly, gblstrPrjTitle
        End If
        Me.CCPB_Pbr.Value = 0
        Frmfarewise.Enabled = True
Exit Sub
lblErr:


If err.Number = 429 Then
    MsgBox "To export data MS Office Excel should be installed.!", vbExclamation, gblstrPrjTitle
ElseIf InStr(1, err.Description, "cannot find the file specified") > 0 Then
    MsgBox "To export data Please install MS Office properly.!", vbExclamation, gblstrPrjTitle
Else
    MsgBox err.Description & vbCrLf & err.Number, vbOKOnly
End If
'WriteToFile App.Path & "\ini.ini", err.Description
'WriteToFile App.Path & "\ini.ini", "cannot find the file specified"
'WriteToFile App.Path & "\ini.ini", InStr(1, err.Description, "cannot find the file specified")

Frmfarewise.Enabled = True
End Sub
Private Sub cmdshow_Click()
Dim rs As New ADODB.Recordset
On Error GoTo lblErr
Dim pass As Long

        TSQL = ""
        
        TSQL = "SELECT DISTINCT tb1.FARE, (select count(FARE) from FARERPT where FARE= tb1.FARE and pid=tb1.pid and schno=tb1.schno and TRPNO=tb1.TRPNO and  DATEVALUE(SCH_STDATE) =tb1.SCH_STDATE ) AS [count], tb1.FARE * (select count(FARE) from FARERPT where FARE= tb1.FARE and pid=tb1.pid and schno=tb1.schno and TRPNO=tb1.TRPNO and  DATEVALUE(SCH_STDATE) =tb1.SCH_STDATE ) as total,pid,SCHNO,TRPNO FROM FARERPT AS tb1 where DATEVALUE(SCH_STDATE) between  DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')"
        
        If cmbPalmID.Text <> "All" Then
            TSQL = TSQL & " and PID='" & cmbPalmID.Text & "'"
        End If
        
        If cmbSchedule.Text <> "All" Then
            TSQL = TSQL & " and SCHNO=" & cmbSchedule.Text & ""
        End If
        If cmbtripno.Text <> "All" Then
            TSQL = TSQL & " and TRPNO=" & cmbtripno.Text & ""
        End If
               
        TSQL = TSQL & " ORDER BY Fare,PID,SCHNO,TRPNO"
        
   If rs.State = adStateOpen Then rs.Close
    
 
    If rs.State <> adStateClosed Then rs.Close
    rs.Open TSQL, gbladoCon, adOpenStatic, adLockReadOnly
                                        
    If rs.RecordCount <> 0 Then
        InsptrRpt.Sections("Section4").Controls("rhDate").caption = Format(Now, "dd/mm/yyyy  hh:mm:ss")
        Set InsptrRpt.Icon = frmMainform.Icon
        Set InsptrRpt.DataSource = rs
        InsptrRpt.Show vbModal
    Else
         MsgBox "No Data to Print...", vbOKOnly
        'Unload Me
    End If
    
    rs.Close
    
Exit Sub
lblErr:
MsgBox "Error due to :" & err.Description, vbExclamation, gblstrPrjTitle
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub DTfrom_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then DTTo.SetFocus
End Sub

Private Sub DTTo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then cmbPalmID.SetFocus
End Sub

Private Sub Form_Load()
    Me.Icon = frmMainform.Icon
    Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
            
        DTfrom.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
        DTTo.Value = DateValue(Format(Now, "DD/MM/YYYY"))
        FillComboPrivate cmbPalmID, "SELECT DISTINCT pid FROM FARERPT ", False, "All"
        FillComboPrivate cmbSchedule, "SELECT DISTINCT SCHNO FROM FARERPT ", False, "All"
        FillComboPrivate cmbtripno, "SELECT DISTINCT TRPNO FROM FARERPT ", False, "All"
        
End Sub
Public Sub FillComboPrivate(objComboBox As ComboBox, strSql As String, Optional strFieldForItemData As String, Optional firstTxt As String)

On Error GoTo lblErr
  Dim oRS As New ADODB.Recordset
  Set oRS = gbladoCon.Execute(strSql)
  With objComboBox
    .Clear
    'Cname_cmb.Clear
    
    If firstTxt <> "" Then
        .AddItem firstTxt
       ' Cname_cmb.AddItem firstTxt
       
    Else
        .AddItem ""
        'Cname_cmb.AddItem ""
    End If
    
    If strFieldForItemData = False Then
      Do While Not oRS.EOF      '(without ItemData)
        .AddItem oRS.Fields(0).Value
         'Cname_cmb.AddItem oRS.Fields(0).value
        oRS.MoveNext
      Loop
    Else
      Do While Not oRS.EOF      '(with ItemData)
        .AddItem oRS.Fields(1).Value
      '  Cname_cmb.AddItem oRS.Fields(2).value
        .ItemData(.NewIndex) = oRS.Fields(0).Value
       ' Cname_cmb.ItemData(.NewIndex) = oRS.Fields(0).value
        oRS.MoveNext
      Loop
    End If
    If .ListCount > 0 Then .ListIndex = 0
    'If Cname_cmb.ListCount > 0 Then Cname_cmb.ListIndex = 0
  End With
  oRS.Close
  Set oRS = Nothing
Exit Sub
lblErr:
MsgBox err.Description, vbOKOnly
End Sub

Private Sub JeweledButton1_Click()

End Sub
