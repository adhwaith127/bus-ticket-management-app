VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form FrmConsole 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stage Wise Consolidated Report"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSchpID 
      BackColor       =   &H00E0E0E0&
      Height          =   3060
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6435
      Begin VB.TextBox TxtRoute 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1560
         Width           =   2055
      End
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
         Left            =   1320
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
         Left            =   3570
         TabIndex        =   8
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
         MICON           =   "FrmConsole.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSchRpt 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
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
         MICON           =   "FrmConsole.frx":001C
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
         Format          =   35061761
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
         Format          =   35061761
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
         Left            =   1200
         TabIndex        =   14
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
         MICON           =   "FrmConsole.frx":0038
         BC              =   12632256
         FC              =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Route"
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
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
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
         Left            =   -120
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   600
         Width           =   1650
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stage Wise Consolidated Report"
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
      TabIndex        =   6
      Top             =   -480
      Width           =   3720
   End
End
Attribute VB_Name = "FrmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myrs As New ADODB.Recordset

Private Sub cmbPalmID_Change()
If cmbPalmID.Text <> "All" And cmbPalmID.Text <> "" Then
    FillComboPrivate cmbSchedule, "SELECT DISTINCT Schdule FROM TKTS where PalmId='" & cmbPalmID.Text & "' and DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')", False, "All"
End If
End Sub

Private Sub cmbPalmID_Click()
If cmbPalmID.Text <> "All" And cmbPalmID.Text <> "" Then
    FillComboPrivate cmbSchedule, "SELECT DISTINCT Schdule FROM TKTS where PalmId='" & cmbPalmID.Text & "' and DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')", False, "All"
End If
End Sub
Private Sub cmbPalmID_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmbSchedule.SetFocus
End Sub

Private Sub cmbSchedule_Change()
If cmbSchedule.Text <> "All" And cmbSchedule.Text <> "" And cmbPalmID.Text <> "All" And cmbPalmID.Text <> "" Then
    FillComboPrivate cmbtripno, "SELECT DISTINCT TripNo FROM TKTS where Schdule=" & cmbSchedule.Text & " and PalmId='" & cmbPalmID.Text & "' and DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')", False, "All"
ElseIf cmbSchedule.Text <> "All" And cmbSchedule.Text <> "" Then
    FillComboPrivate cmbtripno, "SELECT DISTINCT TripNo FROM TKTS where Schdule=" & cmbSchedule.Text & " and DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')", False, "All"
End If
End Sub

Private Sub cmbSchedule_Click()
If cmbSchedule.Text <> "All" And cmbSchedule.Text <> "" And cmbPalmID.Text <> "All" And cmbPalmID.Text <> "" Then
    FillComboPrivate cmbtripno, "SELECT DISTINCT TripNo FROM TKTS where Schdule=" & cmbSchedule.Text & " and PalmId='" & cmbPalmID.Text & "' and DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')", False, "All"
ElseIf cmbSchedule.Text <> "All" And cmbSchedule.Text <> "" Then
    FillComboPrivate cmbtripno, "SELECT DISTINCT TripNo FROM TKTS where Schdule=" & cmbSchedule.Text & " and DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')", False, "All"
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
If cmbtripno.Text <> "All" And cmbSchedule.Text <> "All" And cmbPalmID.Text <> "All" Then
    rtname = getvalueQuery("select RouteCode from rpt where TripNo=" & cmbtripno.Text & " and SCHEDULE=" & cmbSchedule.Text & " and PalmID='" & cmbPalmID.Text & "' and DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')")
    TxtRoute.Text = getvalueQuery(" select rutname from ROUTE where rutcode='" & TrimChr(rtname) & "'")
Else
    TxtRoute.Text = ""
End If
End Sub

Private Sub cmbtripno_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdSchRpt.SetFocus
End Sub

Private Sub cmdSchRpt_Click()
On Error GoTo lblErr
Dim exclrow, exclcol As Integer
Dim FS As New FileSystemObject
Dim SysD, FnameUp As String
Dim RouteCode, TcketPath, TransPath As String
Dim luggcnt, phcnt, stcnt, fullcnt, halfcnt, passcnt As Long
Dim lad_cnt As Long, sc_cnt As Long
Dim TotAmt, refund As Double
    CCPB_Pbr.Value = 0
    FrmConsole.Enabled = False
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
    FnameUp = "STAGE WISE CONSOLIDATED REPORT"
    If Dir(TcketPath & "\" & FnameUp & ".xlsx", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xlsx")
    If Dir(TcketPath & "\" & FnameUp & ".xls", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xls")
    Dim ExlObj As New excel.Application
    ExlObj.Workbooks.Add
        
    ExlObj.ActiveSheet.Cells(1, 5).Value = FnameUp
    'ExlObj.Range("A1:Q1").MergeCells = True
    ExlObj.ActiveSheet.Cells(1, 5).HorizontalAlignment = xlCenter
    'ExlObj.Range("A1:Q1").MergeCells = True
    ExlObj.Range("1:3").Font.FontStyle = "Bold"
       
    TSQL = ""
    TSQL = "select Sum([Full]) as ful,sum(refundamt) as refund,Sum(Half) as Haf,Sum(LuggAmont) AS lugamt,Sum(Phy) as ph,Sum(st) as stu,Sum(Amount) as amt,FromStage,ToStage,PalmId,Schdule,TripNo,sum(ladies_count) as lad,sum(senior_count) as sc from TKTS where DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')"
    If cmbPalmID.Text <> "All" Then
        TSQL = TSQL & " and PalmId='" & cmbPalmID.Text & "'"
    End If
    If cmbSchedule.Text <> "All" Then
        TSQL = TSQL & " and Schdule=" & cmbSchedule.Text & ""
    End If
    If cmbtripno.Text <> "All" Then
        TSQL = TSQL & " and TripNo=" & cmbtripno.Text & ""
    End If
        
        'TSQL = TSQL & " ORDER BY TicketNo,Date,PalmId,Schdule,TripNo"
        TSQL = TSQL & " GROUP BY FromStage,ToStage,PalmId,Schdule,TripNo"
        Set myres = New ADODB.Recordset
        myres.Open TSQL, gbladoCon, 2, 2
        'Myrs.Open TSQL, gbladoCon, 2, 2
        If myres.EOF <> True Then
        
             exclrow = 5
'             ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "TICKET No"
'             ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
'             ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "START STAGE"
             ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "END STAGE"
             ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 3).Value = "FL"
             ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 4).Value = "HF"
             ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "ST"
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 6).Value = "PH"
             ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Color = vbBlack
             
             ExlObj.ActiveSheet.Cells(exclrow, 7).Value = "SC"
             ExlObj.ActiveSheet.Cells(exclrow, 7).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 7).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Ladies"
             ExlObj.ActiveSheet.Cells(exclrow, 8).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 8).Font.Color = vbBlack
             
             ExlObj.ActiveSheet.Cells(exclrow, 9).Value = "LUG"
             ExlObj.ActiveSheet.Cells(exclrow, 9).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 9).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 10).Value = "REFUND AMOUNT"
             ExlObj.ActiveSheet.Cells(exclrow, 10).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 10).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 11).Value = "TOTAL AMOUNT"
             ExlObj.ActiveSheet.Cells(exclrow, 11).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 11).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 12).Value = "PALMID"
             ExlObj.ActiveSheet.Cells(exclrow, 12).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 12).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 13).Value = "SCHEDULE"
             ExlObj.ActiveSheet.Cells(exclrow, 13).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 13).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 14).Value = "TRIP"
             ExlObj.ActiveSheet.Cells(exclrow, 14).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 14).Font.Color = vbBlack
             exclrow = exclrow + 1
             
            Do While Not myres.EOF
            '''''SANGEETHA 09-05-14
'                routecode = ""
'                routecode = getvalueQuery("select RouteCode from rpt where TripNo=" & Myres("TripNo") & " and SCHEDULE=" & Myres("Schdule") & " and PalmID='" & TrimChr(Myres("PalmId")) & "' ")
                'routecode = getvalueQuery("select RouteCode from rpt where TripNo=" & Myres("TripNo") & " and SCHEDULE=" & Myres("Schdule") & " and PalmID='" & TrimChr(Myres("PalmId")) & "' and DATEVALUE('" & Myres("Date") & "') BETWEEN DATEVALUE(StartDate) AND DATEVALUE(EndDate)")
            ''''''''''''''''''''''''''''
                'ExlObj.ActiveSheet.Cells(exclrow, 1).Value = Myres("TicketNo") 'FromStage
                ExlObj.ActiveSheet.Cells(exclrow, 1).Value = myres("FromStage") - 1
                ExlObj.ActiveSheet.Cells(exclrow, 2).Value = myres("ToStage") - 1
'                ExlObj.ActiveSheet.Cells(exclrow, 1).Value = getvalueQuery("SELECT TOP 1 StageName From Stage WHERE route='" & routecode & "' and  StageName  not in (SELECT TOP " & Myres("FromStage") & " StageName  from STAGE where  route='" & routecode & "' ORDER BY id)ORDER BY id")
'                ExlObj.ActiveSheet.Cells(exclrow, 2).Value = getvalueQuery("SELECT TOP 1 StageName From Stage WHERE route='" & routecode & "' and  StageName  not in (SELECT TOP " & Myres("ToStage") & " StageName  from STAGE where  route='" & routecode & "' ORDER BY id)ORDER BY id")
                ExlObj.ActiveSheet.Cells(exclrow, 3).Value = myres("ful")
                fullcnt = fullcnt + myres("Ful")
                ExlObj.ActiveSheet.Cells(exclrow, 4).Value = myres("Haf")
                halfcnt = halfcnt + myres("Haf")
                ExlObj.ActiveSheet.Cells(exclrow, 5).Value = myres("stu")
                stcnt = stcnt + myres("stu")
                ExlObj.ActiveSheet.Cells(exclrow, 6).Value = myres("Ph")
                phcnt = phcnt + myres("Ph")
                
                ExlObj.ActiveSheet.Cells(exclrow, 7).Value = myres("sc")
                sc_cnt = sc_cnt + myres("sc")
                ExlObj.ActiveSheet.Cells(exclrow, 8).Value = myres("lad")
                lad_cnt = lad_cnt + myres("lad")
                
                ExlObj.ActiveSheet.Cells(exclrow, 9).Value = myres("LugAmt")
                luggcnt = luggcnt + CLng(myres("LugAmt"))
'                ExlObj.ActiveSheet.Cells(exclrow, 9).Value = Myres("PassNo")
'                If TrimChr(Myres("PassNo")) <> "" And TrimChr(Myres("PassNo")) <> " " Then
'                    passcnt = passcnt + 1
'                End If
                ExlObj.ActiveSheet.Cells(exclrow, 10).Value = myres("refund")
                refund = refund + myres("refund")
                ExlObj.ActiveSheet.Cells(exclrow, 11).Value = myres("Amt")
                TotAmt = TotAmt + myres("Amt")
'                ExlObj.ActiveSheet.Cells(exclrow, 11).Value = Myres("Time")
'                ExlObj.ActiveSheet.Cells(exclrow, 12).Value = Myres("Date")
                ExlObj.ActiveSheet.Cells(exclrow, 12).Value = myres("PalmId")
                ExlObj.ActiveSheet.Cells(exclrow, 13).Value = myres("Schdule")
                ExlObj.ActiveSheet.Cells(exclrow, 14).Value = myres("TripNo")
                myres.MoveNext
                exclrow = exclrow + 1
                 If Me.CCPB_Pbr.Value >= Me.CCPB_Pbr.Max - 1 Then
                Me.CCPB_Pbr.Value = 0
             Else
                Me.CCPB_Pbr.Value = Me.CCPB_Pbr.Value + 1
             End If
            Loop
             exclrow = exclrow + 1
             ExlObj.ActiveSheet.Cells(exclrow, 3).Value = "FL"
             ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 4).Value = "HF"
             ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 5).Value = "ST"
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Color = vbBlack
             ExlObj.ActiveSheet.Cells(exclrow, 6).Value = "PH"
             ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Color = vbBlack
             
             ExlObj.ActiveSheet.Cells(exclrow, 7).Value = "SC"
             ExlObj.ActiveSheet.Cells(exclrow, 7).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 7).Font.Color = vbBlack
             
             ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Ladies"
             ExlObj.ActiveSheet.Cells(exclrow, 8).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 8).Font.Color = vbBlack
             
             ExlObj.ActiveSheet.Cells(exclrow, 9).Value = "LUG"
             ExlObj.ActiveSheet.Cells(exclrow, 9).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 9).Font.Color = vbBlack


             ExlObj.ActiveSheet.Cells(exclrow, 10).Value = "REFUND AMOUNT"
             ExlObj.ActiveSheet.Cells(exclrow, 10).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 10).Font.Color = vbBlack
             
             ExlObj.ActiveSheet.Cells(exclrow, 11).Value = "TOTAL AMOUNT"
             ExlObj.ActiveSheet.Cells(exclrow, 11).Font.Bold = True
             ExlObj.ActiveSheet.Cells(exclrow, 11).Font.Color = vbBlack
            exclrow = exclrow + 1
            ExlObj.ActiveSheet.Cells(exclrow, 3).Value = fullcnt
            ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 3).Font.Color = vbBlack
            ExlObj.ActiveSheet.Cells(exclrow, 4).Value = halfcnt
            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Color = vbBlack
            ExlObj.ActiveSheet.Cells(exclrow, 5).Value = stcnt
            ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Color = vbBlack
            ExlObj.ActiveSheet.Cells(exclrow, 6).Value = phcnt
            ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Color = vbBlack
            
            ExlObj.ActiveSheet.Cells(exclrow, 7).Value = sc_cnt
            ExlObj.ActiveSheet.Cells(exclrow, 7).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 7).Font.Color = vbBlack
            
            ExlObj.ActiveSheet.Cells(exclrow, 8).Value = lad_cnt
            ExlObj.ActiveSheet.Cells(exclrow, 8).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 8).Font.Color = vbBlack
            
            ExlObj.ActiveSheet.Cells(exclrow, 9).Value = luggcnt
            ExlObj.ActiveSheet.Cells(exclrow, 9).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 9).Font.Color = vbBlack
            ExlObj.ActiveSheet.Cells(exclrow, 10).Value = Format(refund, "0.00")
            ExlObj.ActiveSheet.Cells(exclrow, 10).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 10).Font.Color = vbBlack
           ' exclrow = exclrow + 1
            
            ExlObj.ActiveSheet.Cells(exclrow, 11).Value = Format(TotAmt, "0.00")
            ExlObj.ActiveSheet.Cells(exclrow, 11).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 11).Font.Color = vbBlack
            exclrow = exclrow + 1
             exclrow = exclrow + 1
            ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "TOTAL PASS"
            ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Color = vbBlack
            
            TSQL = ""
            TSQL = "SELECT count(passNo) AS pascnt From TKTS where DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "') And PassNo <> """""
            
            If cmbPalmID.Text <> "All" Then
            TSQL = TSQL & " and PalmId='" & cmbPalmID.Text & "'"
            End If
            If cmbSchedule.Text <> "All" Then
                TSQL = TSQL & " and Schdule=" & cmbSchedule.Text & ""
            End If
            If cmbtripno.Text <> "All" Then
                TSQL = TSQL & " and TripNo=" & cmbtripno.Text & ""
            End If
            Set myres = New ADODB.Recordset
            myres.Open TSQL, gbladoCon, 2, 2
        
            If myres.EOF <> True Then
                ExlObj.ActiveSheet.Cells(exclrow, 4).Value = myres!pascnt
            Else
                ExlObj.ActiveSheet.Cells(exclrow, 4).Value = 0
            End If
            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Color = vbBlack
            
            
            exclrow = exclrow + 1
            ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "TOTAL LUGG"
            ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Color = vbBlack
            ExlObj.ActiveSheet.Cells(exclrow, 4).Value = luggcnt & "U"
            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Color = vbBlack
            ExlObj.ActiveSheet.Cells(exclrow, 4).HorizontalAlignment = xlRight
            
            
            exclrow = exclrow + 1
            ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "TOTAL AMOUNT"
            ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Color = vbBlack
            ExlObj.ActiveSheet.Cells(exclrow, 4).Value = TotAmt - refund
            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Bold = True
            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Color = vbBlack
            
            
            
            
'            ExlObj.ActiveSheet.Cells(exclrow, 4).Value = fullcnt
'            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Bold = True
'            ExlObj.ActiveSheet.Cells(exclrow, 4).Font.Color = vbBlack
'            ExlObj.ActiveSheet.Cells(exclrow, 5).Value = halfcnt
'            ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Bold = True
'            ExlObj.ActiveSheet.Cells(exclrow, 5).Font.Color = vbBlack
'            ExlObj.ActiveSheet.Cells(exclrow, 6).Value = luggcnt
'            ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Bold = True
'            ExlObj.ActiveSheet.Cells(exclrow, 6).Font.Color = vbBlack
'            ExlObj.ActiveSheet.Cells(exclrow, 7).Value = phcnt
'            ExlObj.ActiveSheet.Cells(exclrow, 7).Font.Bold = True
'            ExlObj.ActiveSheet.Cells(exclrow, 7).Font.Color = vbBlack
'            ExlObj.ActiveSheet.Cells(exclrow, 8).Value = stcnt
'            ExlObj.ActiveSheet.Cells(exclrow, 8).Font.Bold = True
'            ExlObj.ActiveSheet.Cells(exclrow, 8).Font.Color = vbBlack
'            ExlObj.ActiveSheet.Cells(exclrow, 10).Value = Format(TotAmt, "0.00")
'            ExlObj.ActiveSheet.Cells(exclrow, 10).Font.Bold = True
'            ExlObj.ActiveSheet.Cells(exclrow, 10).Font.Color = vbBlack
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
        FrmConsole.Enabled = True
Exit Sub
lblErr:
If err.Number = 429 Then
    MsgBox "To export data MS Office Excel should be installed.!", vbExclamation, gblstrPrjTitle
ElseIf InStr(1, err.Description, "cannot find the file specified") > 0 Then
    MsgBox "To export data Please install MS Office properly.!", vbExclamation, gblstrPrjTitle
Else
    MsgBox err.Description, vbOKOnly, gblstrPrjTitle
End If
FrmConsole.Enabled = True
End Sub
Private Sub cmdshow_Click()
On Error GoTo lblErr
Dim pass As Long
        TSQL = ""
        TSQL = "select Sum([Full]) as ful,sum(refundamt) as refund ,Sum(Half) as Haf,Sum(LuggAmont) AS lugamt,Sum(Phy) as ph,Sum(st) as stu,Sum(Amount) as amt,FromStage-1 as FromStage ,ToStage-1 as ToStage, PalmId,Schdule,TripNo,sum(ladies_count) as lad,sum(senior_count) as sc from TKTS where DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "')"
        If cmbPalmID.Text <> "All" Then
            TSQL = TSQL & " and PalmId='" & cmbPalmID.Text & "'"
        End If
        If cmbSchedule.Text <> "All" Then
            TSQL = TSQL & " and Schdule=" & cmbSchedule.Text & ""
        End If
        If cmbtripno.Text <> "All" Then
            TSQL = TSQL & " and TripNo=" & cmbtripno.Text & ""
        End If
               
        TSQL = TSQL & " GROUP BY FromStage,ToStage,PalmId,Schdule,TripNo"
        Set myres = New ADODB.Recordset
        myres.Open TSQL, gbladoCon, 2, 2
        passTSQL = ""
            passTSQL = "SELECT count(passNo) AS pascnt From TKTS where DATEVALUE(Date) BETWEEN DATEVALUE('" & DTfrom.Value & " ') AND DATEVALUE('" & DTTo.Value & "') And PassNo <> """""
            
            If cmbPalmID.Text <> "All" Then
            passTSQL = passTSQL & " and PalmId='" & cmbPalmID.Text & "'"
            End If
            If cmbSchedule.Text <> "All" Then
                passTSQL = passTSQL & " and Schdule=" & cmbSchedule.Text & ""
            End If
            If cmbtripno.Text <> "All" Then
                passTSQL = passTSQL & " and TripNo=" & cmbtripno.Text & ""
            End If
            Set Mypassres = New ADODB.Recordset
            Mypassres.Open passTSQL, gbladoCon, 2, 2
            If Mypassres.EOF <> True Then
                pass = Mypassres!pascnt
            End If
            
        If myres.EOF <> True Then
            Set stageRpt.Icon = frmMainform.Icon
            Set stageRpt.DataSource = myres
            With stageRpt
            .Sections("Section4").Controls("rhDate").caption = Format(Now, "dd/mm/yyyy  hh:mm:ss")
            For Each Section In .Sections
                If Section.Name = "GroupHeader" Then
                    For Each Control In Section.Controls
                        If TypeOf Control Is RptTextBox Or TypeOf Control Is RptFunction Then
                            Control.DataMember = ""
                        End If
                    Next Control
                End If
                If Section.Name = "Detail" Or Section.Name = "GroupFooter" Then
                    For Each Control In Section.Controls
                        If TypeOf Control Is RptTextBox Or TypeOf Control Is RptFunction Then
                            Control.DataMember = ""
                        End If
                    Next Control
                End If
            Next Section
            
            .Sections("Detail").Controls("FromStage").DataField = "FromStage"
            .Sections("Detail").Controls("ToStage").DataField = "ToStage"
            .Sections("Detail").Controls("ful").DataField = "ful"
            .Sections("Detail").Controls("Haf").DataField = "Haf"
            .Sections("Detail").Controls("stu").DataField = "stu"
            .Sections("Detail").Controls("ph").DataField = "ph"
            .Sections("Detail").Controls("lugamt").DataField = "lugamt"
            .Sections("Detail").Controls("amt").DataField = "amt"
            .Sections("Detail").Controls("PalmId").DataField = "PalmId"
            .Sections("Detail").Controls("Schdule").DataField = "Schdule"
            .Sections("Detail").Controls("TripNo").DataField = "TripNo"
            stageRpt.Sections("Detail5").Controls("label16").caption = pass
             
            .Refresh
            .Show vbModal
        End With
    Else
        MsgBox "No Data to Print...", vbOKOnly
        Exit Sub
    End If
             
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
        FillComboPrivate cmbPalmID, "SELECT DISTINCT PalmId FROM TKTS ", False, "All"
        FillComboPrivate cmbSchedule, "SELECT DISTINCT Schdule FROM TKTS ", False, "All"
        FillComboPrivate cmbtripno, "SELECT DISTINCT TripNo FROM TKTS ", False, "All"
        
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

