VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmExpenseRpt 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expense Report"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkEnableDate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enable Date"
      Height          =   240
      Left            =   0
      TabIndex        =   12
      Top             =   -360
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.OptionButton optReportType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Date Wise"
      Height          =   300
      Index           =   1
      Left            =   2130
      TabIndex        =   11
      Top             =   -360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.OptionButton optReportType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Schedule Wise"
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   -240
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Frame frSelectSchedule 
      BackColor       =   &H00E0E0E0&
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin VB.ComboBox cmbShedule 
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
         TabIndex        =   6
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   375
         Left            =   2970
         TabIndex        =   1
         Top             =   1635
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         TX              =   "&Exit"
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
         MICON           =   "FrmExpenseRpt.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdExpense 
         Height          =   375
         Left            =   1875
         TabIndex        =   2
         Top             =   1635
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         TX              =   "&OK"
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
         MICON           =   "FrmExpenseRpt.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker DTEnd 
         Height          =   330
         Left            =   4080
         TabIndex        =   3
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
         Format          =   232718337
         CurrentDate     =   39536
      End
      Begin MSComCtl2.DTPicker DTStart 
         Height          =   345
         Left            =   1560
         TabIndex        =   4
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
         Format          =   232718337
         CurrentDate     =   39536
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "End  Date"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   360
         Width           =   1770
      End
      Begin VB.Label Label1 
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
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1770
      End
      Begin VB.Label lblSdateOrID 
         BackStyle       =   0  'Transparent
         Caption         =   "PalmtecID"
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
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   1770
      End
      Begin VB.Label lblEndDateOrSch 
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
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
   End
   Begin MSComCtl2.DTPicker DTSchDate 
      Height          =   300
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   232718337
      CurrentDate     =   39536
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Expense Report"
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
      TabIndex        =   9
      Top             =   -480
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
Dim flagdate As Boolean
Dim sql As String

Private Sub LoadSchedule(cmbSchedule As ComboBox, PalmID As String)
If PalmID <> "All" Then
    sql = "select DISTINCT SCHEDULENO from EXPENSE where PALMID='" & PalmID & "'"
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
    RES.MoveFirst
    cmbSchedule.Clear
    cmbSchedule.AddItem "All"
    Do While Not RES.EOF
    cmbSchedule.AddItem RES!ScheduleNo
    RES.MoveNext
    Loop
    If cmbSchedule.ListCount > 0 Then
    cmbSchedule.Text = cmbSchedule.List(0)
    End If
End If
End Sub

Private Sub cmbPalmID_Change()
    If Len(cmbPalmID) > 0 Then
        LoadSchedule cmbShedule, cmbPalmID.Text
    End If
End Sub

Private Sub cmbPalmID_Click()
    If Len(cmbPalmID) > 0 Then
        LoadSchedule cmbShedule, cmbPalmID.Text
    End If
End Sub


Private Sub cmdExpense_Click()
On Error GoTo ErrorMod
Dim mystr As String
    Set rs = New ADODB.Recordset
    Set Conn = New ADODB.Connection
    flagdate = True
    If flagdate = True Then
    chkEnableDate.Value = 1
    flagdate = False
    End If
    
    If optReportType(0).Value = True Then
        If cmbPalmID = "" Or cmbShedule = "" Then
            MsgBox "Select PalmtecID and Schedule", vbOKOnly, "PVT"
            Exit Sub
        End If
        If chkEnableDate.Value = 0 Then
            sql = "select * from EXPENSE EX where EX.PalmID='" & cmbPalmID & "' and EX.SCHEDULENO=" & cmbShedule & " order by EX.DATE"
        ElseIf chkEnableDate.Value = 1 Then
        If cmbPalmID.Text <> "All" Then
                mystr = mystr & " and EX.PalmID='" & cmbPalmID & "'"
            End If
            If cmbShedule.Text <> "All" Then
                If mystr <> "" Then
                     mystr = mystr & " and EX.SCHEDULENO=" & cmbShedule & ""
                Else
                    mystr = mystr & "EX.SCHEDULENO=" & cmbShedule & ""
                End If
            End If
          If mystr = "" Then
            mystr = mystr & "AND " & 1
          End If
          'sql = "select * from EXPENSE EX where EX.PalmID='" & cmbPalmID & "' and EX.SCHEDULENO=" & cmbShedule & " AND EX.DATE=DateValue ('" & DTSchDate.Value & "') ORDER BY EX.DATE"
            sql = "select * from EXPENSE EX where DateValue(EX.DATE) BETWEEN DateValue ('" & DTStart.Value & "') AND DateValue ('" & DTEnd.Value & "')  " & mystr & " ORDER BY EX.DATE"
        End If
'            sql = "select * from EXPENSE EX where EX.PalmID='" & cmbPalmID & "' and EX.SCHEDULENO=" & cmbShedule & " AND EX.DATE=DateValue ('" & DTSchDate.Value & "') ORDER BY EX.DATE"
'        End If
    Else
        sql = "SELECT * FROM EXPENSE EX WHERE EX.DATE BETWEEN DateValue ('" & DTStart.Value & "') AND DateValue ('" & DTEnd.Value & "') ORDER BY EX.DATE, EX.PALMID, EX.SCHEDULENO"
    End If
    
    If rs.State = adStateOpen Then rs.Close
    
    Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
    Conn.ConnectionString = "Data Source=" & App.Path & "\Pvt.mdb"
    Conn.Properties("Jet OLEDB:Database Password") = "silbus"
    Conn.Open
    
    If rs.State <> adStateClosed Then rs.Close
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly
    If rs.RecordCount <> 0 Then
        Set ExpenseRpt.Icon = frmMainform.Icon
        Set ExpenseRpt.DataSource = rs
        ExpenseRpt.Show vbModal
    Else
        MsgBox "No Record Found.Please upload data from Palmtec"
        Unload Me
    End If
    rs.Close
Exit Sub

ErrorMod:
If err.Number = 429 Then
    MsgBox "To export data MS Office Excel should be installed.!", vbExclamation, gblstrPrjTitle
Else
    MsgBox "Error occured Due to " & err.Description, vbExclamation, gblstrPrjTitle
End If
End Sub

Private Sub DTSchDate_Change()
On Error Resume Next
    sql = "select DISTINCT PALMID from EXPENSE where DATE=DateValue ('" & DTSchDate & "')"
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
    cmbPalmID.Clear
    cmbShedule.Clear
    If RES.RecordCount > 0 Then
        RES.MoveFirst
        Do While Not RES.EOF
            cmbPalmID.AddItem RES!PalmID
            RES.MoveNext
        Loop
        If cmbPalmID.ListCount > 0 Then
            cmbPalmID.Text = cmbPalmID.List(0)
        End If
        RES.Close
        sql = "select DISTINCT ScheduleNo from EXPENSE where PalmId='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        RES.MoveFirst
        Do While Not RES.EOF
            cmbShedule.AddItem RES!ScheduleNo
            RES.MoveNext
        Loop
'        If cmbShedule.ListCount > 0 Then
'            cmbShedule.Text = cmbShedule.List(0)
'        End If
RES.Close
    End If
    
    chkEnableDate.Value = False
    Call chkEnableDate_Click
    flagdate = True
    
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()                         ''''''''''''RNC
On Error Resume Next
    Me.Icon = frmMainform.Icon
    CONNECTDB
    sql = "SELECT DISTINCT PALMID FROM EXPENSE"
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
  '  lblSdateOrID.Left = 300
    'lblEndDateOrSch.Left = 300
    If optReportType(0).Value = True Then
       ' lblSdateOrID.caption = "Palmtec ID    :"
        'lblEndDateOrSch.caption = "Schedule No :"
       ' DTEnd.Visible = False
       ' DTStart.Visible = False
        DTEnd.Visible = True
        DTStart.Visible = True
        cmbPalmID.Visible = True
        cmbShedule.Visible = True
        DTStart.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
        DTEnd.Value = DateValue(Format(Now, "DD/MM/YYYY"))
'        DTSchDate.Day = Day(Now)
'        DTSchDate.Month = Month(Now)
'        DTSchDate.Year = Year(Now)
'        DTSchDate.Enabled = False
         
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbPalmID.Clear
        cmbPalmID.AddItem "All"
        Do While Not RES.EOF
            cmbPalmID.AddItem RES!PalmID
            RES.MoveNext
        Loop
        If cmbPalmID.ListCount > 0 Then
            cmbPalmID.Text = cmbPalmID.List(0)
        End If
        
        sql = "SELECT DISTINCT SCHEDULENO FROM EXPENSE WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbShedule.Clear
        cmbShedule.AddItem "All"
        Do While Not RES.EOF
            cmbShedule.AddItem RES!ScheduleNo
            RES.MoveNext
        Loop
        If cmbShedule.ListCount > 0 Then
            cmbShedule.Text = cmbShedule.List(0)
        End If
        
        Else
       ' lblSdateOrID.caption = "Start Date  :"
       ' lblEndDateOrSch.caption = " End Date :"
        DTEnd.Visible = True
        DTStart.Visible = True
        cmbPalmID.Visible = False
        cmbShedule.Visible = False
'        DTStart.Top = 840
'        DTStart.Left = 2000
'        DTEnd.Top = 1320
'        DTEnd.Left = 2000
'        DTStart.Width = 1300
'        DTEnd.Width = 1300
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
        lblSdateOrID.caption = "Palmtec ID    :"
        lblEndDateOrSch.caption = "Schedule No :"
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
        sql = "SELECT DISTINCT PALMID FROM EXPENSE"
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
        
        sql = "SELECT DISTINCT SCHEDULENO FROM EXPENSE WHERE PALMID='" & cmbPalmID.Text & "'"
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
        lblSdateOrID.caption = "Start Date :"
        lblEndDateOrSch.caption = "End Date  :"
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
    'if res.s RES.Close

End Sub

Private Sub chkEnableDate_Click()
On Error Resume Next

  If chkEnableDate.Value = 1 Then
        DTSchDate.Enabled = True
'        sql = "select DISTINCT PALMID from EXPENSE where DATE=DateValue('" & DTSchDate.Value & "')"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        cmbPalmID.Clear
'        cmbShedule.Clear
'        If RES.RecordCount > 0 Then
'            RES.MoveFirst
'            Do While Not RES.EOF
'                cmbPalmID.AddItem RES!PalmId
'                RES.MoveNext
'            Loop
'            If cmbPalmID.ListCount > 0 Then
'                cmbPalmID.Text = cmbPalmID.List(0)
'            End If
'         '   RES.Close
'            sql = "select DISTINCT SCHEDULENO from EXPENSE where PALMID='" & cmbPalmID.Text & "'"
'            Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'            RES.MoveFirst
'            cmbShedule.Clear
'            Do While Not RES.EOF
'                cmbShedule.AddItem RES!ScheduleNo
'                RES.MoveNext
'            Loop
'            If cmbShedule.ListCount > 0 Then
'                cmbShedule.Text = cmbShedule.List(0)
'            End If
'
'        Else
'
'        MsgBox "Please Select a Valid Trip Date"
'        '''''''''''''
'
'        End If
'
'    Else
'        DTSchDate.Enabled = False
'        sql = "SELECT DISTINCT PALMID FROM EXPENSE"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        cmbPalmID.Clear
'        If RES.RecordCount > 0 Then RES.MoveFirst
'        Do While Not RES.EOF
'            cmbPalmID.AddItem RES!PalmId
'            RES.MoveNext
'        Loop
'        If cmbPalmID.ListCount > 0 Then
'            cmbPalmID.Text = cmbPalmID.List(0)
'        End If
'
''        RES.Close
'
'        sql = "SELECT DISTINCT SCHEDULENO FROM EXPENSE WHERE PALMID='" & cmbPalmID.Text & "'"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        If RES.RecordCount > 0 Then RES.MoveFirst
'        cmbShedule.Clear
'        Do While Not RES.EOF
'            cmbShedule.AddItem RES!ScheduleNo
'            RES.MoveNext
'        Loop
'        If cmbShedule.ListCount > 0 Then
'            cmbShedule.Text = cmbShedule.List(0)
'        End If
'
    End If
End Sub
