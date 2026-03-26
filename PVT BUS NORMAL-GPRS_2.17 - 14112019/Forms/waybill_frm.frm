VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form waybill_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Waybill Allocation"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   8625
      TabIndex        =   7
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox conductor_name_cbo 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1530
         Width           =   2775
      End
      Begin VB.TextBox Service_No_txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         IMEMode         =   3  'DISABLE
         Left            =   5160
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3120
         Width           =   3255
      End
      Begin VB.ListBox Service_No_lst 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   810
         Left            =   5040
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ComboBox bus_no_cbo 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3810
         Width           =   2775
      End
      Begin VB.TextBox schedule_Trip_txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   3
         Top             =   3210
         Width           =   2775
      End
      Begin VB.TextBox schedule_KM_txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   2
         Top             =   2610
         Width           =   2775
      End
      Begin VB.ComboBox driver_name_cbo 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2070
         Width           =   2775
      End
      Begin VB.TextBox manual_way_no_txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox waybill_no_txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin JeweledBut.JeweledButton Clear_cmd 
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Top             =   4440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         TX              =   "Clea&r"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "waybill_frm.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton save_cmd 
         Height          =   315
         Left            =   4440
         TabIndex        =   5
         Top             =   4440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         TX              =   "&Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "waybill_frm.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Delete_cmd 
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         TX              =   "&Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "waybill_frm.frx":0038
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker duty_date_dtp 
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   900
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   -2147483644
         Format          =   36962305
         CurrentDate     =   42142
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Duty Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Conductor Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Service No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   15
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Bus Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Trip"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule KM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Waybill no"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Waybill No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "waybill_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bus_no_cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Service_No_txt.SetFocus
End Sub

Private Sub Clear_cmd_Click()
    clearfield
End Sub

Private Sub conductor_name_cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then driver_name_cbo.SetFocus
End Sub

Private Sub driver_name_cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then schedule_KM_txt.SetFocus
End Sub

Private Sub duty_date_dtp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then conductor_name_cbo.SetFocus
End Sub

Private Sub Form_Load()
    Me.Icon = frmMainform.Icon
    Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
    Fillconductor
    FillDriver
    FillBus
    clearfield
End Sub

Private Sub Fillconductor()
Dim RES As New ADODB.Recordset
    TSQL = ""
    TSQL = "SELECT EMPLOYEEID,EMPLOYEENAME,ET.EMPLOYEETYPEID FROM CREW CR inner join EMPLOYEETYPE ET on CR.EMPLOYEETYPEID=ET.EmployeeTypeId where ET.TypeName='Conductor'"
    Set RES = gbladoCon.Execute(TSQL)
    conductor_name_cbo.Clear
    conductor_name_cbo.AddItem "Select"
    conductor_name_cbo.ItemData(conductor_name_cbo.NewIndex) = 0
    If Not RES.EOF Then
        RES.MoveFirst
        
        Do While Not RES.EOF
            conductor_name_cbo.AddItem (RES("EMPLOYEENAME"))
            conductor_name_cbo.ItemData(conductor_name_cbo.NewIndex) = RES("EMPLOYEEID")
            RES.MoveNext
        Loop
    End If
End Sub
Private Sub FillDriver()
Dim RES As ADODB.Recordset
    TSQL = ""
    TSQL = "SELECT EMPLOYEEID,EMPLOYEENAME,ET.EMPLOYEETYPEID FROM CREW CR inner join EMPLOYEETYPE ET on CR.EMPLOYEETYPEID=ET.EmployeeTypeId where ET.TypeName='Driver'"
    Set RES = gbladoCon.Execute(TSQL)
    driver_name_cbo.Clear
    driver_name_cbo.AddItem "Select"
    driver_name_cbo.ItemData(driver_name_cbo.NewIndex) = 0
    If Not RES.EOF Then
        RES.MoveFirst
       
        Do While Not RES.EOF
            driver_name_cbo.AddItem (RES("EMPLOYEENAME"))
            driver_name_cbo.ItemData(driver_name_cbo.NewIndex) = RES("EMPLOYEEID")
            RES.MoveNext
            Loop
    End If
End Sub
Private Sub FillBus()
Dim RES As ADODB.Recordset
    TSQL = ""
    TSQL = "SELECT BUSID,BUSNO,BUSTYPE FROM VEHICLETYPE"
    Set RES = gbladoCon.Execute(TSQL)
    bus_no_cbo.Clear
    bus_no_cbo.AddItem "Select"
    bus_no_cbo.ItemData(bus_no_cbo.NewIndex) = 0
    If Not RES.EOF Then
        RES.MoveFirst
       
        Do While Not RES.EOF
            bus_no_cbo.AddItem (RES("BUSNO"))
            bus_no_cbo.ItemData(bus_no_cbo.NewIndex) = RES("BUSID")
            RES.MoveNext
            Loop
    End If
End Sub
Private Function clearfield()
On Error Resume Next
    waybill_no_txt = ""
    manual_way_no_txt = ""
    duty_date_dtp = Now()
    conductor_name_cbo.ListIndex = 0
    driver_name_cbo.ListIndex = 0
    schedule_KM_txt = ""
    schedule_Trip_txt = ""
    bus_no_cbo.ListIndex = 0
    Service_No_Rtxt = ""
    manual_way_no_txt = ""
    Service_No_txt.Text = ""
End Function

Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> 1 And WindowState <> 2 Then
        Me.Left = (Screen.Width / 2) - (Me.Width / 2)
        Me.Top = ((Screen.Height - 1200) / 2) - (Me.Height / 2)
    End If
End Sub

Private Sub manual_way_no_txt_KeyPress(KeyAscii As Integer)
    ValidationMode = StrictAlphaNumeric
    ValidateKeyPress manual_way_no_txt, KeyAscii
End Sub

Private Sub save_cmd_Click()
On Error GoTo er1
    If Trim(manual_way_no_txt) = "" Then
        MsgBox "Please enter a waybill number.", vbExclamation, gblstrPrjTitle
        manual_way_no_txt.SetFocus
        Exit Sub
    End If
    If conductor_name_cbo.ListIndex < 1 Then
        MsgBox "Please select a conductor.", vbExclamation, gblstrPrjTitle
        conductor_name_cbo.SetFocus
        Exit Sub
    End If
    If driver_name_cbo.ListIndex < 1 Then
        MsgBox "Please select a driver.", vbExclamation, gblstrPrjTitle
        driver_name_cbo.SetFocus
        Exit Sub
    End If
    If Trim(schedule_KM_txt) = "" Then
        MsgBox "Please enter schedule KM.", vbExclamation, gblstrPrjTitle
        schedule_KM_txt.SetFocus
        Exit Sub
    End If
    If Trim(schedule_Trip_txt) = "" Then
        MsgBox "Please enter a schedule Trip.", vbExclamation, gblstrPrjTitle
        schedule_Trip_txt.SetFocus
        Exit Sub
    End If
    If bus_no_cbo.ListIndex < 1 Then
        MsgBox "Please select a bus.", vbExclamation, gblstrPrjTitle
        bus_no_cbo.SetFocus
        Exit Sub
    End If
    
    If val(getvalueQuery("select count(waybill_ID) from Waybill_tab where waybill_number='" & manual_way_no_txt & "'")) > 0 Then
        MsgBox "Waybill number already exists.Please choose another one.", vbExclamation, gblstrPrjTitle
        manual_way_no_txt.SetFocus
        Exit Sub
    End If
    sql = "insert into Waybill_tab(waybill_number,Duty_date,Conductor,Driver,Schedule_KM,Schedule_Trip,Bus_no,Service_No) " _
        & "values('" & Trim(manual_way_no_txt) & "','" & DateValue(duty_date_dtp.Value) & "','" & Trim(conductor_name_cbo.Text) & "','" _
        & Trim(driver_name_cbo.Text) & "','" & Trim(schedule_KM_txt) & "','" & Trim(schedule_Trip_txt) & "','" & Trim(bus_no_cbo.Text) & "','" & Trim(Service_No_txt.Text) & "')"
    gbladoCon.Execute (sql)
    MsgBox "Waybill details saved..", vbInformation, gblstrPrjTitle
    clearfield
Exit Sub
er1:
    MsgBox "Error in waybill saving.." & vbCrLf & err.Number & err.Description, vbExclamation, gblstrPrjTitle
End Sub

Private Sub schedule_KM_txt_KeyPress(KeyAscii As Integer)
    ValidationMode = FloatingPointValue
    ValidateKeyPress schedule_KM_txt, KeyAscii
End Sub

Private Sub schedule_Trip_txt_KeyPress(KeyAscii As Integer)
    ValidationMode = AlphaNumeric
    ValidateKeyPress schedule_Trip_txt, KeyAscii
End Sub

Private Sub Service_No_txt_KeyPress(KeyAscii As Integer)
    ValidationMode = AlphaNumeric
    ValidateKeyPress Service_No_txt, KeyAscii
End Sub
