VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmExpenseRpt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expense Report"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frSelectSchedule 
      BackColor       =   &H00E0E0E0&
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5115
      Begin VB.ComboBox cmbShedule 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cmbPalmID 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optReportType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Schedule Wise"
         Height          =   300
         Index           =   0
         Left            =   285
         TabIndex        =   8
         Top             =   225
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton optReportType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Date Wise"
         Height          =   300
         Index           =   1
         Left            =   2415
         TabIndex        =   7
         Top             =   225
         Width           =   1590
      End
      Begin VB.CheckBox chkEnableDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Date"
         Height          =   240
         Left            =   3500
         TabIndex        =   4
         Top             =   800
         Width           =   1245
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   450
         Left            =   3570
         TabIndex        =   1
         Top             =   2235
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
         MICON           =   "FrmOdmtr.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdOdmtr 
         Height          =   450
         Left            =   2475
         TabIndex        =   2
         Top             =   2235
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
         MICON           =   "FrmOdmtr.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker DTSchDate 
         Height          =   300
         Left            =   690
         TabIndex        =   3
         Top             =   1770
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   20709377
         CurrentDate     =   39536
      End
      Begin MSComCtl2.DTPicker DTEnd 
         Height          =   330
         Left            =   3300
         TabIndex        =   5
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   39536
      End
      Begin MSComCtl2.DTPicker DTStart 
         Height          =   345
         Left            =   3285
         TabIndex        =   6
         Top             =   900
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   609
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   39536
      End
      Begin VB.Label lblSdateOrID 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter PalmtecID:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label lblEndDateOrSch 
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
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Odometer Report"
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
      Left            =   1200
      TabIndex        =   13
      Top             =   240
      Width           =   2880
   End
End
Attribute VB_Name = "frmExpenseRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PRES As DAO.Recordset
Dim PALM_ID As String
Dim Conn As New ADODB.Connection
Dim RES As DAO.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String

Private Sub cmdOdmtr_Click()
On Error Resume Next
    Set rs = New ADODB.Recordset
    Set Conn = New ADODB.Connection
    If optReportType(0).Value = True Then
        If cmbPalmID = "" Or cmbShedule = "" Then
            MsgBox "Select PalmtecID and Schedule", vbOKOnly, "PVT"
            Exit Sub
        End If
        If chkEnableDate.Value = 0 Then
            sql = "select * from ODOMETER OD where OD.PalmID='" & cmbPalmID & "' and OD.SCHEDULENO=" & cmbShedule & " order by OD.SDATE, OD.TRIPNO"
        ElseIf chkEnableDate.Value = 1 Then
            sql = "select * from ODOMETER OD where OD.PalmID='" & cmbPalmID & "' and OD.SCHEDULENO=" & cmbShedule & " AND OD.SDATE=DateValue ('" & DTSchDate.Value & "') ORDER BY OD.SDATE, OD.TRIPNO"
        End If
    Else
    '''******************
        sql = "SELECT * FROM ODOMETER OD WHERE OD.SDATE BETWEEN DateValue ('" & DTStart.Value & "') AND DateValue ('" & DTEnd.Value & "') AND OD.EDATE BETWEEN DateValue ('" & DTStart.Value & "') AND DateValue ('" & DTEnd.Value & "')ORDER BY OD.SDATE, OD.PALMID, OD.SCHEDULENO, OD.TRIPNO"
    End If
    '''************
    If rs.State = adStateOpen Then rs.Close
    
    Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
    Conn.ConnectionString = "Data Source=" & App.Path & "\Pvt.mdb"
    Conn.Properties("Jet OLEDB:Database Password") = "silbus"
    Conn.Open
    
    If rs.State <> adStateClosed Then rs.Close
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly
    If rs.RecordCount <> 0 Then
    
        Set OdmtrRpt.DataSource = rs
        OdmtrRpt.Show vbModal
    Else
        MsgBox "No Record Found.Please upload data from Palmtech"
        Unload Me
    End If
    rs.Close
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()                         ''''''''''''RNC
On Error Resume Next
    CONNECTDB
    sql = "SELECT DISTINCT PALMID FROM ODOMETER"
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
    lblSdateOrID.Left = 300
    lblEndDateOrSch.Left = 300
    If optReportType(0).Value = True Then
        lblSdateOrID.Caption = "Palmtec ID    :"
        lblEndDateOrSch.Caption = "Schedule No :"
        DTEnd.Visible = False
        DTStart.Visible = False
        cmbPalmID.Visible = True
        cmbShedule.Visible = True
        cmbPalmID.Top = 840
        cmbPalmID.Left = 2000
        cmbShedule.Top = 1320
        cmbShedule.Left = 2000
        DTSchDate.Top = 1200
        DTSchDate.Left = 3500
        DTSchDate.Day = Day(Now)
        DTSchDate.Month = Month(Now)
        DTSchDate.Year = Year(Now)
        DTSchDate.Enabled = False
        
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbPalmID.Clear
        Do While Not RES.EOF
            cmbPalmID.AddItem RES!PalmId
            RES.MoveNext
        Loop
        If cmbPalmID.ListCount > 0 Then
            cmbPalmID.Text = cmbPalmID.List(0)
        End If
        
        sql = "SELECT DISTINCT SCHEDULENO FROM ODOMETER WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbShedule.Clear
        Do While Not RES.EOF
            cmbShedule.AddItem RES!ScheduleNo
            RES.MoveNext
        Loop
        If cmbShedule.ListCount > 0 Then
            cmbShedule.Text = cmbShedule.List(0)
        End If
        
        Else
        lblSdateOrID.Caption = "Start Date  :"
        lblEndDateOrSch.Caption = " End Date :"
        DTEnd.Visible = True
        DTStart.Visible = True
        cmbPalmID.Visible = False
        cmbShedule.Visible = False
        DTStart.Top = 840
        DTStart.Left = 2000
        DTEnd.Top = 1320
        DTEnd.Left = 2000
        DTStart.Width = 1300
        DTEnd.Width = 1300
        DTStart.Day = Day(Now)
        DTStart.Month = Month(Now)
        DTStart.Year = Year(Now)
        DTEnd.Day = Day(Now)
        DTEnd.Month = Month(Now)
        DTEnd.Year = Year(Now)
    End If
    RES.Close
End Sub

Private Sub optReportType_Click(Index As Integer)
On Error Resume Next

    If optReportType(0).Value = True Then
        lblSdateOrID.Caption = "Palmtec ID    :"
        lblEndDateOrSch.Caption = "Schedule No :"
        DTEnd.Visible = False
        DTStart.Visible = False
        cmbPalmID.Visible = True
        cmbShedule.Visible = True
        chkEnableDate.Visible = True
        DTSchDate.Visible = True
        chkEnableDate.Value = 0
        DTSchDate.Enabled = False
        cmbPalmID.Top = 840
        cmbPalmID.Left = 2000
        cmbShedule.Top = 1320
        cmbShedule.Left = 2000
        cmbPalmID.Clear
        sql = "SELECT DISTINCT PALMID FROM ODOMETER"
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
        
        sql = "SELECT DISTINCT SCHEDULENO FROM ODOMETER WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbShedule.Clear
        Do While Not RES.EOF
            cmbShedule.AddItem RES!ScheduleNo
            RES.MoveNext
        Loop
        If cmbShedule.ListCount > 0 Then
            cmbShedule.Text = cmbShedule.List(0)
        End If
    Else
        lblSdateOrID.Caption = "Start Date :"
        lblEndDateOrSch.Caption = "End Date  :"
        DTEnd.Visible = True
        DTStart.Visible = True
        cmbPalmID.Visible = False
        cmbShedule.Visible = False
        chkEnableDate.Visible = False
        DTSchDate.Visible = False
        DTStart.Top = 840
        DTStart.Left = 2000
        DTEnd.Top = 1320
        DTEnd.Left = 2000
        DTStart.Width = 1300
        DTEnd.Width = 1300
        DTStart.Day = Day(Now)
        DTStart.Month = Month(Now)
        DTStart.Year = Year(Now)
        DTEnd.Day = Day(Now)
        DTEnd.Month = Month(Now)
        DTEnd.Year = Year(Now)
        
    End If
'    RES.Close

End Sub

Private Sub chkEnableDate_Click()
On Error Resume Next

    If chkEnableDate.Value = 1 Then
        DTSchDate.Enabled = True
        sql = "select DISTINCT PALMID from ODOMETER where SDATE=DateValue('" & DTSchDate.Value & "')"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        cmbPalmID.Clear
        cmbShedule.Clear
        If RES.RecordCount > 0 Then
            RES.MoveFirst
            Do While Not RES.EOF
                cmbPalmID.AddItem RES!PalmId
                RES.MoveNext
            Loop
            If cmbPalmID.ListCount > 0 Then
                cmbPalmID.Text = cmbPalmID.List(0)
            End If
         '   RES.Close
            sql = "select DISTINCT SCHEDULENO from ODOMETER where PALMID='" & cmbPalmID.Text & "'"
            Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
            RES.MoveFirst
            cmbShedule.Clear
            Do While Not RES.EOF
                cmbShedule.AddItem RES!ScheduleNo
                RES.MoveNext
            Loop
            If cmbShedule.ListCount > 0 Then
                cmbShedule.Text = cmbShedule.List(0)
            End If
       
        Else
        
        MsgBox "Please Select a Valid Trip Date"
        '''''''''''''
        
        End If
        
    Else
        DTSchDate.Enabled = False
        sql = "SELECT DISTINCT PALMID FROM ODOMETER"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        cmbPalmID.Clear
        If RES.RecordCount > 0 Then RES.MoveFirst
        Do While Not RES.EOF
            cmbPalmID.AddItem RES!PalmId
            RES.MoveNext
        Loop
        If cmbPalmID.ListCount > 0 Then
            cmbPalmID.Text = cmbPalmID.List(0)
        End If
        
'        RES.Close
        
        sql = "SELECT DISTINCT SCHEDULENO FROM ODOMETER WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbShedule.Clear
        Do While Not RES.EOF
            cmbShedule.AddItem RES!ScheduleNo
            RES.MoveNext
        Loop
        If cmbShedule.ListCount > 0 Then
            cmbShedule.Text = cmbShedule.List(0)
        End If
    
    End If
End Sub
