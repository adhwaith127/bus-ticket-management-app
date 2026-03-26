VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form Crew 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crew "
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   0
      TabIndex        =   6
      Top             =   -105
      Width           =   7110
      Begin VB.TextBox txtpswd 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1680
         Width           =   2580
      End
      Begin JeweledBut.JeweledButton ccmdclear 
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   2640
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         TX              =   "Clear"
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
         MICON           =   "Crew.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin VB.CheckBox chkView 
         BackColor       =   &H00E0E0E0&
         Caption         =   "View All Details"
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   2280
         Width           =   3975
      End
      Begin VB.ComboBox cmbTypeId 
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
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtEmployeeId 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   1
         Top             =   720
         Width           =   2580
      End
      Begin VB.TextBox txtEmployeeName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1200
         Width           =   2580
      End
      Begin JeweledBut.JeweledButton cmdCancel 
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   2640
         Width           =   1200
         _ExtentX        =   2117
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
         MICON           =   "Crew.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSave 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   2640
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         TX              =   "&Save"
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
         MICON           =   "Crew.frx":0038
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdDel 
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   2640
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         TX              =   "&Delete"
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
         MICON           =   "Crew.frx":0054
         BC              =   12632256
         FC              =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp Pswd"
         Height          =   495
         Left            =   1200
         TabIndex        =   15
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp Id"
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp Type"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp Name"
         Height          =   495
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   1995
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flGrd 
      Height          =   3255
      Left            =   0
      TabIndex        =   13
      Top             =   3120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   15846563
      BackColorBkg    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
      FormatString    =   "Emp Type    |<Emp Id     |<Emp Name                                  |<Emp Pswd      "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Crew Details"
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
      Height          =   510
      Left            =   1050
      TabIndex        =   12
      Top             =   -480
      Width           =   5490
   End
End
Attribute VB_Name = "Crew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentEmpId As Double
Dim blnEdit As Boolean
Private Sub ccmdclear_Click()
    txtEmployeeId.Text = ""
    txtEmployeeName.Text = ""
    txtpswd.Text = ""
    cmdDel.Enabled = False
    blnEdit = False
End Sub

Private Sub chkView_Click()
On Error GoTo lblErr
    FillEmployee
Exit Sub
lblErr:
End Sub

Private Sub cmbTypeId_Click()
On Error GoTo lblErr
    If blnEdit = True Then Exit Sub
    FillEmployee
Exit Sub
lblErr:
End Sub

Private Sub cmbTypeId_KeyPress(KeyAscii As Integer)
On Error GoTo lblErr
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    End If
Exit Sub
lblErr:
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
On Error GoTo CatchError
    If flGrd.TextMatrix(flGrd.row, 0) = "" Or flGrd.row = 0 Then Exit Sub
    CONNECTDB
    Dim sqlstr As String
    If (MsgBox("Do you want to remove Crew  " & txtEmployeeName & " from Crew details ", vbQuestion + vbYesNo)) = vbYes Then
        sqlstr = "DELETE * FROM crew WHERE [EMPLOYEEID]=" & currentEmpId
        CNN.Execute (sqlstr)
        CNN.Close
        MsgBox "Crew " & txtEmployeeName & " successfully removed from Crew List", vbInformation, App.ProductName
        FillEmployee
    End If
    cmbTypeId.ListIndex = 0
    txtEmployeeId.Text = ""
    txtEmployeeName.Text = ""
    txtpswd.Text = ""
    cmdDel.Enabled = False
    blnEdit = False
    flGrd.row = 0
Exit Sub
CatchError:
    MsgBox "Error : " & err.Number & vbTab & err.Description & vbCrLf & _
        "Crew " & txtEmployeeName & "  is not removed", vbExclamation, App.ProductName
End Sub
Private Sub cmdSave_Click()
On Error GoTo err
Dim strSql, sql1 As String
Dim recedit As ADODB.Recordset
    Set RSDT = New ADODB.Recordset
    CONNECTDB
    Dim EmployeeTypeId As String
    EmployeeTypeId = getemployeetypeid(Trim(cmbTypeId.Text))
    If Trim(txtEmployeeId.Text) = "" Then
        MsgBox "Please enter the Employee Id", vbInformation, App.ProductName
        txtEmployeeId.SetFocus
        Exit Sub
    End If
    If Trim(txtEmployeeName.Text) = "" Then
        MsgBox "Please enter the Employee Name", vbInformation, App.ProductName
        txtEmployeeName.SetFocus
        Exit Sub
    End If
    If Trim(txtpswd.Text) = "" Then
        MsgBox "Please enter the Employee Password", vbInformation, App.ProductName
        txtpswd.SetFocus
        Exit Sub
    End If
    strSql = "SELECT count(*) as cnt FROM CREW "
    Set res1 = CNN.OpenRecordset(strSql, dbOpenDynaset)
    rec = res1!cnt
    Set res1 = Nothing
    If rec = 128 Then
        MsgBox "Crew Entry Limit Reached!", vbInformation
        cmbTypeId.ListIndex = 0
        txtEmployeeId.Text = ""
        txtEmployeeName.Text = ""
        txtpswd.Text = ""
        Exit Sub
    End If
    If blnEdit = False Then
        strSql = "SELECT * FROM CREW "
        Set res1 = CNN.OpenRecordset(strSql, dbOpenDynaset)
        If res1.RecordCount > 0 Then
            res1.MoveLast
        End If
        res1.AddNew
        res1!EmployeeTypeId = EmployeeTypeId
        res1!EMPLOYEEID = txtEmployeeId
        res1!EMPLOYEENAME = txtEmployeeName
        res1!PSWD = txtpswd
        res1.Update
        res1.Close
        txtEmployeeId = ""
        txtEmployeeName = ""
        txtpswd.Text = ""
        cmbTypeId.SetFocus
        MsgBox "Details added successfully", vbInformation, App.ProductName
        CNN.Close
        FillEmployee
        Set res1 = Nothing
        cmdDel.Enabled = False
        flGrd.row = 0
        Exit Sub
    Else
        RSql = "SELECT EMPLOYEEID FROM crew WHERE EMPLOYEEID = " & txtEmployeeId & " and EMPLOYEEID <> " & currentEmpId
        Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            MsgBox "Employee Id already exist." & vbCrLf & "Please give another Id", vbInformation, gblstrPrjTitle
            txtEmployeeId.SetFocus
            Exit Sub
        End If
        Set res1 = CNN.OpenRecordset("SELECT * FROM crew WHERE [EMPLOYEEID]=" & currentEmpId, dbOpenDynaset)
        If res1.RecordCount > 0 Then
            res1.Edit
            res1!EmployeeTypeId = EmployeeTypeId
            res1!EMPLOYEEID = txtEmployeeId
            res1!EMPLOYEENAME = txtEmployeeName
            res1!PSWD = txtpswd
            res1.Update
            res1.Close
            cmbTypeId.ListIndex = 0
            txtEmployeeId.Text = ""
            txtEmployeeName.Text = ""
            txtpswd.Text = ""
            blnEdit = False
            MsgBox "Crew details updated successfully", vbInformation, App.ProductName
            CNN.Close
            FillEmployee
            cmdDel.Enabled = False
            Set res1 = Nothing
            flGrd.row = 0
            Exit Sub
        Else
            MsgBox "Current record not found! Unable to update", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    Dim Ctrl As Control
    If Trim(txtEmployeeId) = "" And Trim(txtEmployeeName) = "" Then
        MsgBox "No fields should be empty", vbInformation, App.ProductName
        Exit Sub
    End If
Exit Sub
err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Sub
Public Function getemployeetypeid(EmpType As String) As Double
On Error GoTo lblErr
Dim rsEmployee As ADODB.Recordset
Dim intCount As Integer
    Set rsEmployee = New ADODB.Recordset
    rsEmployee.Open "SELECT EmployeeTypeId from EMPLOYEETYPE where TypeName='" & EmpType & "'", gbladoCon, adOpenDynamic, adLockOptimistic
    If rsEmployee.State = adStateOpen Then
        getemployeetypeid = rsEmployee!EmployeeTypeId
    End If
    If rsEmployee.State = adStateOpen Then rsEmployee.Close
Exit Function
lblErr:
End Function

Private Sub flGrd_DblClick()
Dim intCont As Integer
On Error GoTo lblErr
    cmdCancel.Enabled = True
    Me.MousePointer = vbHourglass
    If flGrd.row = 0 Or flGrd.TextMatrix(flGrd.row, 0) = "" Then Exit Sub
    blnEdit = True
    cmbTypeId.Text = flGrd.TextMatrix(flGrd.row, 0)
    txtEmployeeId.Text = flGrd.TextMatrix(flGrd.row, 1)
    txtEmployeeName = flGrd.TextMatrix(flGrd.row, 2)
    txtpswd = flGrd.TextMatrix(flGrd.row, 3)
    currentEmpId = flGrd.TextMatrix(flGrd.row, 1)
    cmbTypeId.SetFocus
    Me.MousePointer = vbNormal
    cmdDel.Enabled = True
Exit Sub
lblErr:
End Sub

Private Sub Form_Load()
On Error GoTo err
    Me.Icon = frmMainform.Icon
    Call FillEmployeeType
    CONNECTDB
    Set RSDT = New ADODB.Recordset
    If CheckTableExistsOrNot("CREW") = False Then
        If RSDT.State <> 0 Then RSDT.Close
        sSQL = "CREATE TABLE CREW " & _
            "(EMPLOYEEID NUMBER , " & _
            "EMPLOYEEID NUMBER , " & _
            "EMPLOYEENAME VARCHAR(64), " & _
            "EMPLOYEETYPEID VARCHAR(8)," & _
            "PSWD VARCHAR(8)," & _
            "BUS_NO VARCHAR(16))"
        CNN.Execute (sSQL)
    End If
    cmbTypeId.Text = cmbTypeId.List(0)
    Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
    FillEmployee
    cmdDel.Enabled = False
    blnEdit = False
Exit Sub
err:
    Call ErrorHandle("Form_Load", err.Number, err.Description)
End Sub
Private Sub FillEmployee()
Dim rsEmployee As DAO.Recordset
Dim intCount As Integer
On Error GoTo CatchError
Dim cond As String
Dim myyy As String
    If cmbTypeId.Text <> "" And chkView.Value = 0 Then
        cond = cond & " and TypeName='" & Trim(cmbTypeId.Text) & "'"
    End If
    CONNECTDB
    flGrd.Rows = 1
    flGrd.FormatString = "<Emp Type    |<Emp Id     |<Emp Name                                  |<Emp Pswd      "
    Set rsEmployee = CNN.OpenRecordset("SELECT TypeName,EMPLOYEEID,EMPLOYEENAME,PSWD FROM crew c,EMPLOYEETYPE e where e.EmployeeTypeId=c.EMPLOYEETYPEID" & cond, adOpenDynamic)
    intCount = 1
    If rsEmployee.RecordCount > 0 Then
        While Not rsEmployee.EOF
            If intCount >= flGrd.Rows Then flGrd.Rows = flGrd.Rows + 1
            flGrd.TextMatrix(intCount, 0) = rsEmployee!TypeName
            flGrd.TextMatrix(intCount, 1) = rsEmployee!EMPLOYEEID
            flGrd.TextMatrix(intCount, 2) = rsEmployee!EMPLOYEENAME
            flGrd.TextMatrix(intCount, 3) = IIf(IsNull(rsEmployee!PSWD), "", rsEmployee!PSWD)
            intCount = intCount + 1
            rsEmployee.MoveNext
        Wend
    End If
    Set rsEmployee = Nothing
    CNN.Close
Exit Sub
CatchError:
    MsgBox "Err:" & err.Number & vbCrLf & err.Description
End Sub
Private Sub FillEmployeeType()
Dim ETYPE As DAO.Database
Dim ETYPEREC As DAO.Recordset
    Set ETYPE = DAO.OpenDatabase(App.Path & "\pvt.mdb")
    TSQL = "SELECT * FROM EMPLOYEETYPE"
    Set ETYPEREC = ETYPE.OpenRecordset(TSQL, dbOpenDynaset)
    If ETYPEREC.RecordCount > 0 Then
        ETYPEREC.MoveFirst
        Do While Not ETYPEREC.EOF
            cmbTypeId.AddItem (ETYPEREC!TypeName)
            ETYPEREC.MoveNext
        Loop
    End If
End Sub
Public Function CheckTableExistsOrNot(strTableName As String) As Boolean
On Error GoTo err
Dim strSql As String
Dim DB As New ADODB.Connection
Dim rs1 As ADODB.Recordset
    CONNECTDB
    sDataBase = App.Path & "\PVT.MDB"
    CheckTableExistsOrNot = False
    If DB.State <> 0 Then DB.Close
        DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source= " & sDataBase & "; Jet OLEDB:Database Password =silbus"
    Set rs1 = New ADODB.Recordset
    strSql = "SELECT * FROM " & strTableName
    If rs1.State <> 0 Then rs1.Close
    rs1.Open strSql, DB, adOpenDynamic, adLockOptimistic
    rs1.Close
    CheckTableExistsOrNot = True
Exit Function
err:
    If rs1.State <> 0 Then rs1.Close
    CheckTableExistsOrNot = False
End Function
Public Function ErrorHandle(strError As String, ErrNumber As Integer, ErrDescription As String)
Dim sErrorString As String
    sErrorString = strError & "_Error!" & vbCrLf & "Error No :" & ErrNumber & vbCrLf & ErrDescription
    Call CreateErrorLog(sErrorString)
    MsgBox strError & "_Error!" & vbCrLf & "Error No :" & ErrNumber & vbCrLf & ErrDescription, vbInformation, gblstrPrjTitle
End Function
Private Sub txtEmployeeId_KeyPress(KeyAscii As Integer)
On Error GoTo lblErr
    If TextBoxValidityNumeric(KeyAscii) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    End If
Exit Sub
lblErr:
End Sub
Public Function EmpIDExists(EmpId As Double) As Boolean
On Error GoTo err
    EmpIDExists = False
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    RSql = "SELECT EMPLOYEEID FROM crew WHERE EMPLOYEEID = " & EmpId
    Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
    If RES.RecordCount > 0 Then EmpIDExists = True
Exit Function
err:
    Select Case err.Number
    Case Else
        MsgBox "Error No : " & err.Number & vbCrLf & err.Description, vbInformation, "Route"
        Exit Function
    End Select
End Function

Private Sub txtEmployeeId_LostFocus()
On Error GoTo lblErr
    If blnEdit = False Then
        If txtEmployeeId.Text <> "" Then
            If EmpIDExists(val(txtEmployeeId.Text)) = True Then
                MsgBox "Employee Id already exist." & vbCrLf & "Please give another Id", vbInformation, gblstrPrjTitle
                txtEmployeeId = ""
                txtEmployeeId.SetFocus
            End If
        End If
    End If
Exit Sub
lblErr:
End Sub
Private Sub txtEmployeeName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtpswd_KeyPress(KeyAscii As Integer)
    If TextBoxValidity(KeyAscii) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    End If
End Sub
