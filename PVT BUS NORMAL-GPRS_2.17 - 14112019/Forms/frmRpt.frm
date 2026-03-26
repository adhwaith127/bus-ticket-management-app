VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmRpt1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6435
   Icon            =   "frmRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frSelectSchedule 
      BackColor       =   &H00E0E0E0&
      Height          =   2835
      Left            =   0
      TabIndex        =   6
      Top             =   -15
      Width           =   6435
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
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   375
         Left            =   3210
         TabIndex        =   10
         Top             =   1995
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
         MICON           =   "frmRpt.frx":0CCA
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdRpt1 
         Height          =   375
         Left            =   2115
         TabIndex        =   9
         Top             =   1995
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
         MICON           =   "frmRpt.frx":0CE6
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker DTEnd 
         Height          =   330
         Left            =   4320
         TabIndex        =   4
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
         Format          =   72155137
         CurrentDate     =   39536
      End
      Begin MSComCtl2.DTPicker DTStart 
         Height          =   345
         Left            =   1560
         TabIndex        =   3
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
         Format          =   72155137
         CurrentDate     =   39536
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
         TabIndex        =   0
         Top             =   840
         Width           =   1335
      End
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
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   1335
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
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   1770
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
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   1770
      End
      Begin VB.Label lbltripno 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "          Trip No    "
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
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
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
         Left            =   2760
         TabIndex        =   8
         Top             =   840
         Width           =   1695
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
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   1770
      End
   End
   Begin VB.CheckBox chkEnableDate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enable Date"
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.OptionButton optReportType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Date Wise"
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.OptionButton optReportType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Schedule Wise"
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker DTSchDate 
      Height          =   300
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   72155137
      CurrentDate     =   39536
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Report"
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
      Left            =   1965
      TabIndex        =   5
      Top             =   -465
      Width           =   1800
   End
End
Attribute VB_Name = "frmRpt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PRES As DAO.Recordset
Dim PALM_ID As String
Dim Conn As ADODB.Connection
Dim RES As DAO.Recordset
Dim rs As ADODB.Recordset
Dim flagdate1 As Boolean
Dim sql As String

Private Sub chkEnableDate_Click()
On Error Resume Next
    If chkEnableDate.Value = 1 Then
        DTSchDate.Enabled = True
'        sql = "select DISTINCT PALMID from RPT where DATE=DateValue ('" & DTSchDate.Value & "')"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        cmbPalmID.Clear
'        cmbShedule.Clear
'        If RES.RecordCount > 0 Then
'            RES.MoveFirst
'            Do While Not RES.EOF
'                cmbPalmID.AddItem RES!PalmID
'                RES.MoveNext
'            Loop
'            If cmbPalmID.ListCount > 0 Then
'                cmbPalmID.Text = cmbPalmID.List(0)
'            End If
'         '   RES.Close
'            sql = "select DISTINCT SCHEDULE from RPT where PALMID='" & cmbPalmID.Text & "'"
'            Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'            RES.MoveFirst
'            cmbShedule.Clear
'            Do While Not RES.EOF
'                cmbShedule.AddItem RES!SCHEDULE
'                RES.MoveNext
'            Loop
'            If cmbShedule.ListCount > 0 Then
'                cmbShedule.Text = cmbShedule.List(0)
'            End If
'        '''''''''syam added
'        Else
'
'        MsgBox "Please Select a Valid Trip Date"
'        '''''''''''''
'
'        End If
'
'    Else
'        DTSchDate.Enabled = False
'        sql = "SELECT DISTINCT PALMID FROM RPT"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        cmbPalmID.Clear
'        If RES.RecordCount > 0 Then RES.MoveFirst
'        Do While Not RES.EOF
'            cmbPalmID.AddItem RES!PalmID
'            RES.MoveNext
'        Loop
'        If cmbPalmID.ListCount > 0 Then
'            cmbPalmID.Text = cmbPalmID.List(0)
'        End If
'
''        RES.Close
'
'        sql = "SELECT DISTINCT SCHEDULE FROM RPT WHERE PALMID='" & cmbPalmID.Text & "'"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        If RES.RecordCount > 0 Then RES.MoveFirst
'        cmbShedule.Clear
'        Do While Not RES.EOF
'            cmbShedule.AddItem RES!SCHEDULE
'            RES.MoveNext
'        Loop
'        If cmbShedule.ListCount > 0 Then
'            cmbShedule.Text = cmbShedule.List(0)
'        End If
'
    End If
End Sub

Private Sub cmbPalmID_Change()
On Error Resume Next

    cmbShedule.Enabled = True
    cmbShedule.Clear
    If cmbPalmID <> "" Then
        If cmbPalmID.Text = "All" Then
            sql = "select distinct SCHEDULE from RPT where DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTStart.Value & " ') AND DATEVALUE('" & DTEnd.Value & "') "
        Else
            sql = "select distinct SCHEDULE from RPT where PalmID='" & cmbPalmID.Text & "' AND DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTStart.Value & " ') AND DATEVALUE('" & DTEnd.Value & "') "
        End If
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            cmbShedule.Clear
             cmbShedule.AddItem "All"
            RES.MoveFirst
            Do While Not RES.EOF
               
                cmbShedule.AddItem RES.Fields(0)
                RES.MoveNext
            Loop
        End If
    End If
    RES.Close
    If cmbShedule.ListCount > 0 Then
        cmbShedule.Text = cmbShedule.List(0)
    End If

End Sub

Private Sub cmbPalmID_Click()
On Error Resume Next

    cmbShedule.Enabled = True
    cmbShedule.Clear
    If cmbPalmID <> "" Then
        'If cmbPalmID.Text = "All" Then
          '  sql = "select distinct SCHEDULE from RPT where DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTStart.Value & " ') AND DATEVALUE('" & DTEnd.Value & "') "
        'Else
            sql = "select distinct SCHEDULE from RPT where PalmID='" & cmbPalmID.Text & "' AND DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTStart.Value & " ') AND DATEVALUE('" & DTEnd.Value & "') "
       ' End If
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If cmbPalmID.Text <> "All" Then
        If RES.RecordCount > 0 Then
            RES.MoveFirst
            cmbShedule.Clear
            cmbShedule.AddItem "All"
            Do While Not RES.EOF
                 
                cmbShedule.AddItem RES.Fields(0)
                RES.MoveNext
            Loop
        Else
            cmbShedule.Clear
            cmbShedule.AddItem "All"
        End If
        
        Else
            cmbShedule.Clear
            cmbShedule.AddItem "All"
        End If
    End If
    RES.Close
    If cmbShedule.ListCount > 0 Then
        cmbShedule.Text = cmbShedule.List(0)
    End If
    
'    If Len(cmbPalmID) > 0 Then
'        LoadSchedule1 cmbShedule, cmbtripno, cmbPalmID.Text
'    End If
    
    
End Sub

Private Sub cmbShedule_Change()
    cmbtripno.Enabled = True
    cmbtripno.Clear
    If cmbPalmID <> "" Then
        If cmbShedule <> "" Then
        sql = "SELECT DISTINCT TRIPNO  FROM RPT WHERE PALMID='" & cmbPalmID.Text & "' and SCHEDULE='" & cmbShedule.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'         Set RES = New ADODB.Recordset
'        RES.Open sql, CON, 2, 2
        If RES.EOF <> True Then
            RES.MoveFirst
            cmbtripno.AddItem "All"
            Do While Not RES.EOF
                cmbtripno.AddItem RES.Fields(0)
                RES.MoveNext
            Loop
        End If
      End If
    End If
    RES.Close
    If cmbtripno.ListCount > 0 Then
        cmbtripno.Text = cmbtripno.List(0)
    End If

End Sub

Private Sub cmbShedule_Click()
cmbtripno.Enabled = True
    cmbtripno.Clear
    If cmbPalmID <> "" Then
       ' If cmbShedule <> "" Then
        If cmbShedule <> "" And cmbShedule <> "All" Then
         cmbtripno.Clear
        sql = "SELECT DISTINCT TRIPNO  FROM RPT WHERE PALMID='" & cmbPalmID.Text & "' and SCHEDULE=" & cmbShedule.Text & ""
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        ' Set RES = New ADODB.Recordset
      ' RES.Open sql, CNN, adOpenStatic, adLockReadOnly
        If RES.EOF <> True Then
            RES.MoveFirst
             cmbtripno.AddItem "All"
            Do While Not RES.EOF
                cmbtripno.AddItem RES.Fields(0)
                RES.MoveNext
            Loop
        End If
      Else
        cmbtripno.AddItem "All"
      End If
    End If
    'RES.Close
    If cmbtripno.ListCount > 0 Then
        cmbtripno.Text = cmbtripno.List(0)
    End If
End Sub

Private Sub CmdRpt1_Click()
On Error Resume Next
 Dim mystr As String
    
    Set rs = New ADODB.Recordset
    Set Conn = New ADODB.Connection
    flagdate1 = 1
   
        If cmbPalmID = "" Or cmbShedule = "" Or cmbtripno = "" Then
            MsgBox "Select PalmtecID , Schedule And Trip", vbOKOnly, "PVT"
            Exit Sub
        End If

        
          If cmbPalmID.Text <> "All" Then
                mystr = mystr & " RT.PalmID='" & cmbPalmID & "'"
            End If
            If cmbShedule.Text <> "All" Then
                If mystr <> "" Then
                     mystr = mystr & " and RT.SCHEDULE=" & cmbShedule & ""
                Else
                    mystr = mystr & "RT.SCHEDULE=" & cmbShedule & ""
                End If
            End If
            If cmbtripno.Text <> "All" Then
                If mystr <> "" Then
                     mystr = mystr & " AND RT.TRIPNO=" & cmbtripno & ""
                Else
                    mystr = mystr & "RT.TRIPNO=" & cmbtripno & ""
                End If
            End If
           ' Else
           If mystr = "" Then
            mystr = 1
           End If
           
           sql = "select RT.*,(RT.TotalColl)AS TotAdjust,(select sum(refundamt) from tkts where  SCHDULE=RT.SCHEDULE and  RT.PALMID=PALMID and TRIPNO=RT.TRIPNO and val(tripmasterid)=val(RT.trip_master_id)) as Refund ,(RT.TotalColl-(IIf((AdjustColl) Is Null, 0, (AdjustColl)))-(IIf((Expense) Is Null, 0, (Expense)))-(IIf((Refund) Is Null, 0, (Refund)))) as Nettot, IIf((Expense) Is Null, 0, (Expense)) as EXPENSE1 from RPT RT where  " & mystr & " AND DateValue(RT.StartDate) BETWEEN  DateValue ('" & DTStart.Value & "') AND DateValue ('" & DTEnd.Value & "') ORDER BY RT.DATE, RT.PALMID, RT.SCHEDULE, RT.TRIPNO"
           
  
       
   
    If rs.State <> adStateClosed Then rs.Close
    
    Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
    Conn.ConnectionString = "Data Source=" & App.Path & "\Pvt.mdb"
    Conn.Properties("Jet OLEDB:Database Password") = "silbus"
    Conn.Open
    
    If rs.State <> adStateClosed Then rs.Close
   
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly
    
    ''''''''syam added
    If rs.RecordCount <> 0 Then
    
    Set RPT.Icon = frmMainform.Icon
    Set RPT.DataSource = rs
    RPT.Show vbModal
    Else
    MsgBox "No Data to Print...", vbOKOnly
    Unload Me
    End If
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTSchDate_Change()
On Error Resume Next
    sql = "select DISTINCT PALMID from RPT where DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTStart.Value & " ')AND DATEVALUE('" & DTEnd.Value & "') "
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
    cmbPalmID.Clear
    cmbShedule.Clear
    cmbtripno.Clear
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
         cmbShedule.Clear
        sql = "select DISTINCT SCHEDULE from RPT where PalmId='" & cmbPalmID.Text & "' and DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTStart.Value & " ')AND DATEVALUE('" & DTEnd.Value & "') "
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        RES.MoveFirst
        Do While Not RES.EOF
            cmbShedule.AddItem RES!SCHEDULE
            RES.MoveNext
        Loop
        If cmbShedule.ListCount > 0 Then
            cmbShedule.Text = cmbShedule.List(0)
        End If
        RES.Close
        'cmbtripno.Clear
        sql = "SELECT DISTINCT TRIPNO  FROM RPT WHERE PALMID='" & cmbPalmID.Text & "' and DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTStart.Value & " ')AND DATEVALUE('" & DTEnd.Value & "') "
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbtripno.Clear
        Do While Not RES.EOF
            cmbtripno.AddItem RES!TripNo
            RES.MoveNext
        Loop
        If cmbtripno.ListCount > 0 Then
            cmbtripno.Text = cmbtripno.List(0)
        End If
        
    
    End If
    chkEnableDate.Value = False
    Call chkEnableDate_Click
    flagdate1 = True
    
End Sub


'Private Sub DTSchDate_Click()
'    chkEnableDate.Value = False
'    chkEnableDate_Click
'    chkEnableDate.Value = True
'End Sub

Private Sub DTSchDate_LostFocus()
    chkEnableDate.Value = False
    chkEnableDate_Click
   ' chkEnableDate.Value = True
End Sub



Private Sub Form_Load()
On Error Resume Next
    Me.Icon = frmMainform.Icon
    CONNECTDB
    sql = "SELECT DISTINCT PALMID FROM RPT"
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
   ' lblSdateOrID.Left = 300
   ' lblEndDateOrSch.Left = 300
  '  If optReportType(0).Value = True Then
   If 1 Then
       ' lblSdateOrID.caption = "Palmtec ID    :"
       ' lblEndDateOrSch.caption = "Schedule No :"
       ' DTEnd.Visible = False
        'DTStart.Visible = False
        DTEnd.Visible = True
        DTStart.Visible = True
        cmbPalmID.Visible = True
        cmbShedule.Visible = True
        DTEnd.Visible = True
        DTEnd.Visible = True
        cmbtripno.Visible = True
        
      
'        DTSchDate.Day = Day(Now)
'        DTSchDate.Month = Month(Now)
'        DTSchDate.Year = Year(Now)
'        DTSchDate.Enabled = False
        DTStart.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
        DTEnd.Value = DateValue(Format(Now, "DD/MM/YYYY"))
        
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
      '  RES.Close
        
        sql = "SELECT DISTINCT SCHEDULE  FROM RPT WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbShedule.Clear
        cmbShedule.AddItem "All"
        Do While Not RES.EOF
            cmbShedule.AddItem RES!SCHEDULE
            RES.MoveNext
        Loop
        If cmbShedule.ListCount > 0 Then
            cmbShedule.Text = cmbShedule.List(0)
        End If
        'If RES.State = 1 Then RES.Close
        
        sql = "SELECT DISTINCT TRIPNO  FROM RPT WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbtripno.Clear
        cmbtripno.AddItem "All"
        Do While Not RES.EOF
            cmbtripno.AddItem RES!TripNo
            RES.MoveNext
        Loop
        If cmbtripno.ListCount > 0 Then
            cmbtripno.Text = cmbtripno.List(0)
        End If
    
    Else
        lblSdateOrID.caption = "Start Date  :"
        lblEndDateOrSch.caption = " End Date :"
        DTEnd.Visible = True
        DTStart.Visible = True
        cmbPalmID.Visible = False
        cmbShedule.Visible = False
        cmbtripno.Visible = False
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
        
        sql = "SELECT DISTINCT SCHEDULE FROM RPT WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbShedule.Clear
        Do While Not RES.EOF
            cmbShedule.AddItem RES!SCHEDULE
            RES.MoveNext
        Loop
        If cmbShedule.ListCount > 0 Then
            cmbShedule.Text = cmbShedule.List(0)
        End If
        RES.Close
        sql = "SELECT DISTINCT TRIPNO FROM RPT WHERE PALMID='" & cmbPalmID.Text & "'"
        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbtripno.Clear
        Do While Not RES.EOF
            cmbtripno.AddItem RES!TripNo
            RES.MoveNext
        Loop
        If cmbtripno.ListCount > 0 Then
            cmbtripno.Text = cmbtripno.List(0)
        End If
    Else
        lblSdateOrID.caption = "Start Date :"
        lblEndDateOrSch.caption = "End Date  :"
        DTEnd.Visible = True
        DTStart.Visible = True
        cmbPalmID.Visible = False
        cmbShedule.Visible = False
        
        If optReportType(1).Value = True Then
            lbltripno.Visible = False
            cmbtripno.Visible = False
        End If
        
'        If optReportType(0).Value = True Then                       ''rnc
'            lbltripno.Visible = True
'            cmbtripno.Visible = True
'        End If

        

        
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
    If optReportType(0).Value = True Then                       ''rnc
            lbltripno.Visible = True
            cmbtripno.Visible = True
    End If
    RES.Close

End Sub
Private Sub LoadSchedule1(cmbSchedule As ComboBox, cmbtripno As ComboBox, PalmID As String)
    sql = "select DISTINCT SCHEDULE,TRIPNO from EXPENSE where PALMID='" & PalmID & "'"
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
    RES.MoveFirst
    cmbSchedule.Clear
    Do While Not RES.EOF
    cmbSchedule.AddItem RES!SCHEDULE
    RES.MoveNext
    Loop
    If cmbSchedule.ListCount > 0 Then
    cmbSchedule.Text = cmbSchedule.List(0)
    End If
    If RES.RecordCount > 0 Then RES.MoveFirst
        cmbtripno.Clear
        Do While Not RES.EOF
            cmbtripno.AddItem RES!TripNo
            RES.MoveNext
        Loop
        If cmbtripno.ListCount > 0 Then
            cmbtripno.Text = cmbtripno.List(0)
        End If
End Sub

