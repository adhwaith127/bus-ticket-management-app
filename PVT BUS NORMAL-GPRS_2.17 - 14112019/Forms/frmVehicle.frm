VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmVehicle 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vehicle Details"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5085
   Icon            =   "frmVehicle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin JeweledBut.JeweledButton cmdDel 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2160
      Width           =   1080
      _ExtentX        =   1905
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
      MICON           =   "frmVehicle.frx":0CCA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdCancel 
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   1080
      _ExtentX        =   1905
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
      MICON           =   "frmVehicle.frx":0CE6
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdSave 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1080
      _ExtentX        =   1905
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
      MICON           =   "frmVehicle.frx":0D02
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2805
      Left            =   -60
      TabIndex        =   6
      Top             =   0
      Width           =   5070
      Begin JeweledBut.JeweledButton cmdClear 
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   2160
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         TX              =   "Clear"
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
         MICON           =   "frmVehicle.frx":0D1E
         BC              =   12632256
         FC              =   0
      End
      Begin VB.TextBox txtBusNo 
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
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   0
         Top             =   600
         Width           =   2580
      End
      Begin VB.ComboBox cmbBusType 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtBusNo1 
         Height          =   330
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Bus Type Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Bus No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1755
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flGrd 
      Height          =   3015
      Left            =   0
      TabIndex        =   11
      Top             =   2760
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   15846563
      BackColorBkg    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
      FormatString    =   "ID |BUS No.             | BUS TYPE                           "
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
      Caption         =   "Vehicle Details"
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
      Left            =   1440
      TabIndex        =   10
      Top             =   -480
      Width           =   2610
   End
End
Attribute VB_Name = "frmVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cHandle As Integer
Dim currentVehId As String
Dim blnEdit As Boolean
Public Function EditCrewToDatabase(cdtr_code As String, cdtr_name As String, Dr_code As String, Dr_Name As String, cln_code As String, cln_name As String, bus_no As String) As Boolean
On Error GoTo err
CONNECTDB
If CheckCrewTableExistsOrNot("CREWDET") = False Then
    strSql = "SELECT DISTINCT * FROM CREWDET "
    RSDT.CursorLocation = adUseClient
    RSDT.Open strSql, DbZb, adOpenStatic, adLockOptimistic
    If RSDT.RecordCount > 0 Then
        While RSDT.EOF <> True
            RSDT!Id = Id
            RSDT!CDTR_ID = cdtr_code
            RSDT!cdtr_name = cdtr_name
            RSDT!DR_ID = Dr_code
            RSDT!Dr_Name = Dr_Name
            RSDT!CLNR_ID = cln_code
            RSDT!CLNR_NAME = cln_name
            RSDT!bus_no = bus_no
            RSDT.Update
            RSDT.MoveNext
            RSDT.Close
        Wend
    End If
End If
err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Function
Public Function GetTotalCrewRecord() As Long
On Error GoTo err
CONNECTDB
If CheckCrewTableExistsOrNot("CREWDET") = False Then
sql = "SELECT count(ID) as  CREWCOUNT FROM CREWDET "
    RSDT.CursorLocation = adUseClient
    RSDT.Open sql, DbZb, adOpenStatic, adLockOptimistic
    If Not RSDT.EOF Then
        GetTotalCrewRecord = RSDT("CREWCOUNT")
    End If
End If
err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Function
Public Function DeleteCrewDetails(Id As Integer) As Boolean
On Error GoTo err
CONNECTDB
If CheckCrewTableExistsOrNot("CREWDET") = False Then
    sql = "DELETE FROM CREWDET WHERE ID=" & Id
    DbZb.Execute (sql)
    MsgBox "Details deleted successfully from database", vbInformation, gblstrPrjTitle
        sql = " UPDATE CREWDET SET ID=id-1 WHERE ID>" & Id
        DbZb.Execute (sql)
        DeleteCrewDetails = True
End If
Exit Function
err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Function
Public Function CheckCrewTableExistsOrNot(table_name As String) As Boolean
On Error GoTo err
CONNECTDB
Set RsZb = New ADODB.Recordset
If CheckTableExistsOrNot("CREWDET") = False Then
    If RsZb.State <> 0 Then RsZb.Close
        sSQL = "CREATE TABLE CREWDET " & _
            "(ID NUMBER, " & _
            "CDTR_ID VARCHAR(8), " & _
            "CDTR_NAME VARCHAR(64), " & _
            "DR_ID VARCHAR(8), " & _
            "DR_NAME VARCHAR(64), " & _
            "CLNR_ID VARCHAR(8), " & _
            "CLNR_NAME VARCHAR(64), " & _
            "BUS_NO VARCHAR(16)) "
        DbZb.Execute (sSQL)
    End If
err:
   If RsZb.State <> 0 Then RsZb.Close
    CheckCrewTableExistsOrNot = False
End Function
Public Function AddCrewToDatabase(cdtr_code As String, cdtr_name As String, Dr_code As String, Dr_Name As String, cln_code As String, cln_name As String, bus_no As String) As Boolean
On Error GoTo err
Set RSDT = New ADODB.Recordset
CONNECTDB
If CheckCrewTableExistsOrNot("CREWDET") = False Then
    Id = 0
    strSql = "SELECT DISTINCT * FROM CREWDET "
    RSDT.CursorLocation = adUseClient
    RSDT.Open strSql, DbZb, adOpenStatic, adLockOptimistic
    If RSDT.RecordCount > 0 Then
        While RSDT.EOF <> True
            Id = Id + 1
            RSDT.AddNew
            RSDT!Id = Id
            RSDT!CDTR_ID = cdtr_code
            RSDT!cdtr_name = cdtr_name
            RSDT!DR_ID = Dr_code
            RSDT!Dr_Name = Dr_Name
            RSDT!CLNR_ID = cln_code
            RSDT!CLNR_NAME = cln_name
            RSDT!bus_no = bus_no
            RSDT.Update
            RSDT.MoveNext
            RSDT.Close
        Wend
    Else
       MsgBox "Record not found", vbInformation, gblstrPrjTitle
    End If
    If GetTotalCrewRecord() > 250 Then
        MsgBox "Record limit exceeds maximum! Record addition failed", vbInformation
    End If
End If
err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Function

Public Function CheckTableExistsOrNot(strTableName As String) As Boolean
On Error GoTo err
Dim strSql As String
Dim DB As New ADODB.Connection
Dim rs As ADODB.Recordset
sDataBase = App.Path & "\PVT.MDB"
CheckTableExistsOrNot = False
If DB.State <> 0 Then DB.Close
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source= " & sDataBase & "; Jet OLEDB:Database Password =silbus"
Set rs = New ADODB.Recordset
strSql = "SELECT * FROM " & strTableName
If rs.State <> 0 Then rs.Close
rs.Open strSql, DB, adOpenDynamic, adLockOptimistic
rs.Close
CheckTableExistsOrNot = True
Exit Function
err:
    If rs.State <> 0 Then rs.Close
    CheckTableExistsOrNot = False
End Function
Private Sub cmbBustype_KeyPress(KeyAscii As Integer)
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

Private Sub cmdClear_Click()
cmdDel.Enabled = False
txtBusNo.Text = ""
'cmbBusType.Text = ""
End Sub

Private Sub cmdDel_Click()
On Error GoTo CatchError
  If flGrd.TextMatrix(flGrd.row, 0) = "" Or flGrd.row = 0 Then Exit Sub
  CONNECTDB
  Dim sqlstr As String
  If (MsgBox("Do you want to remove Bus No " & txtBusNo & " from Vehicle details ?", vbQuestion + vbYesNo)) = vbYes Then
        sqlstr = "DELETE * FROM vehicletype WHERE [BUSNO]='" & currentVehId & "'"
        CNN.Execute (sqlstr)
        CNN.Close
        MsgBox "Vehicle " & txtBusNo & " successfully removed from vehicle list", vbInformation, App.ProductName
        FillVehicleView
        cmdDel.Enabled = False
   End If
        txtBusNo.Text = ""
        cmbBusType.ListIndex = 0
        blnEdit = False
  Exit Sub
CatchError:
    MsgBox "Error : " & err.Number & vbTab & err.Description & vbCrLf & _
           "Crew " & txtEmployeeName & "  is not removed", vbExclamation, App.ProductName
End Sub
Private Sub cmdSave_Click()
On Error GoTo err
Dim sql, sql1 As String
Dim Ctrl As Control
Set RSDT = New ADODB.Recordset
    CONNECTDB
     strSql = "SELECT count(*) as cnt FROM VEHICLETYPE "
     Set res1 = CNN.OpenRecordset(strSql, dbOpenDynaset)
     rec = res1!cnt
     Set res1 = Nothing
     If rec = 256 Then
         MsgBox "Vehicle entry limit reached!", vbInformation, gblstrPrjTitle
         txtBusNo.Text = ""
         cmbBusType.ListIndex = 0
         Exit Sub
     End If
    If Trim(txtBusNo) = "" Then
        MsgBox "Empty field ! enter busno ", vbInformation, App.ProductName
        Exit Sub
    End If
    If BusNoExists(txtBusNo.Text) = True Then
        MsgBox "Bus No. already exists" & vbCrLf & "Please give another Bus No.", vbInformation, gblstrPrjTitle
        txtBusNo.Text = ""
        txtBusNo.SetFocus
        Exit Sub
    End If
    If blnEdit = False Then
            sql = "SELECT * FROM VEHICLETYPE "
            Set res1 = CNN.OpenRecordset(sql, dbOpenDynaset)
            If res1.RecordCount > 0 Then
                res1.MoveLast
            End If
            res1.AddNew
            res1!Busno = txtBusNo
            res1!BusType = cmbBusType
            sql1 = "SELECT * FROM BUSTYPE"
            Set RES4 = CNN.OpenRecordset(sql1, adOpenDynamic)
            If RES4.RecordCount > 0 Then
                Do While Not RES4.EOF
                    If (res1!BusType = RES4!Name) Then
                        res1!BUSID = RES4!Id
                        Exit Do
                    End If
                    RES4.MoveNext
                Loop
            End If
            RES4.Close
            res1.Update
            res1.Close
            Set res1 = Nothing
            'txtBusId = ""
            txtBusNo = ""
            txtBusNo.SetFocus
            cmbBusType.Text = cmbBusType.List(0)
            MsgBox "Details added successfully ", vbInformation, App.ProductName
            CNN.Close
            Set res1 = Nothing
            flGrd.row = 0
            FillVehicleView
            cmdDel.Enabled = False
            Exit Sub
    Else
            Set res1 = CNN.OpenRecordset("SELECT * FROM vehicletype WHERE [BUSNO]='" & currentVehId & "'", dbOpenDynaset)
            If res1.RecordCount > 0 Then
                res1.Edit
                res1!Busno = txtBusNo
                res1!BusType = cmbBusType
                res1.Update
                res1.Close
                txtBusNo.Text = ""
                cmbBusType.ListIndex = 0
                blnEdit = False
                CNN.Close
                MsgBox "Vehicle details updated successfully", vbInformation, App.ProductName
                flGrd.row = 0
                FillVehicleView
                cmdDel.Enabled = False
                Exit Sub
            Else
                MsgBox "Current record not found! Unable to update", vbInformation, App.ProductName
                Exit Sub
            End If
       Exit Sub
    End If
err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Sub

Private Sub flGrd_Click()
Dim intCont As Integer
On Error GoTo lblErr
    cmdCancel.Enabled = True
    Me.MousePointer = vbNormal
    If flGrd.row = 0 Or flGrd.TextMatrix(flGrd.row, 0) = "" Then Exit Sub
    txtBusNo.Text = flGrd.TextMatrix(flGrd.row, 1)
    cmbBusType.Text = flGrd.TextMatrix(flGrd.row, 2)
    currentVehId = flGrd.TextMatrix(flGrd.row, 1)
    blnEdit = True
    txtBusNo.SetFocus
    cmdDel.Enabled = True
    Me.MousePointer = vbNormal
Exit Sub
lblErr:
End Sub

Private Sub Form_Load()                             ''on 12/09/11
On Error Resume Next
    Me.Icon = frmMainform.Icon
    Call FillBusTypeName
    CONNECTDB
    Set RSDT = New ADODB.Recordset
        If CheckTableExistsOrNot("VEHICLETYPE") = False Then
            If RSDT.State <> 0 Then RSDT.Close
                sSQL = "CREATE TABLE VEHICLETYPE " & _
                        "(BUSID VARCHAR(8) , " & _
                        "BUSNO VARCHAR(16) , " & _
                        "BUSTYPE VARCHAR(16))"
                    CNN.Execute (sSQL)
            End If
            cmbBusType.Text = cmbBusType.List(0)
            If RSDT.State = adStateOpen Then RSDT.Close
            Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
            blnEdit = False
            FillVehicleView
            cmdDel.Enabled = False
Exit Sub
End Sub

Private Sub FillVehicleView()
Dim intCount As Integer
On Error GoTo CatchError
    CONNECTDB
    flGrd.Clear
    flGrd.Rows = 1
    flGrd.FormatString = "^ID |^Bus No.               |<Bus Type Name                     "
    Dim rsvehilce As DAO.Recordset
     Set rsvehilce = CNN.OpenRecordset("SELECT * FROM vehicletype", adOpenDynamic)
    intCount = 1
    If rsvehilce.RecordCount > 0 Then
        While Not rsvehilce.EOF
            If intCount >= flGrd.Rows Then flGrd.Rows = flGrd.Rows + 1
            flGrd.TextMatrix(intCount, 0) = rsvehilce!BUSID
            flGrd.TextMatrix(intCount, 1) = rsvehilce!Busno
            flGrd.TextMatrix(intCount, 2) = rsvehilce!BusType
            intCount = intCount + 1
            rsvehilce.MoveNext
        Wend
    End If
     Set rsvehilce = Nothing
     flGrd.ColWidth(0) = 0
    Exit Sub
CatchError:
    MsgBox err.Number & vbCrLf & err.Description
End Sub
Private Sub FillBusTypeName()
Dim VCLTYPE As DAO.Database
Dim VCLTYPEREC As DAO.Recordset
    Set VCLTYPE = DAO.OpenDatabase(App.Path & "\pvt.mdb")
 
    TSQL = "SELECT * FROM BUSTYPE"
    Set VCLTYPEREC = VCLTYPE.OpenRecordset(TSQL, dbOpenDynaset)
    If VCLTYPEREC.RecordCount > 0 Then
        VCLTYPEREC.MoveFirst
        'cmbBusType.Clear
        Do While Not VCLTYPEREC.EOF
            cmbBusType.AddItem (VCLTYPEREC!Name)
            ''MsgBox BTYPEREC!Name
            VCLTYPEREC.MoveNext
            Loop
    End If
End Sub

Private Sub txtBus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtConductor <> "" Then txtCleaner.SetFocus
End Sub

Private Sub txtCleaner_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCleaner <> "" Then cmdSave.SetFocus
End Sub

Private Sub txtConductor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtConductor <> "" Then txtBus.SetFocus
End Sub

Private Sub txtDriver_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtDriver <> "" Then txtConductor.SetFocus
End Sub

Private Sub txtBusNo_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 And txtBusNo <> "" Then
        If blnEdit = False Then
            If BusNoExists(txtBusNo.Text) = True Then
                MsgBox "Bus no. already exists" & vbCrLf & "Please give another Bus no.", vbInformation, gblstrPrjTitle
                txtBusNo.Text = ""
                txtBusNo.SetFocus
                Exit Sub
            End If
        End If
        SendKeys ("{TAB}")
    End If
End Sub
Public Function BusNoExists(Busno As String) As Boolean
On Error GoTo err
    BusNoExists = False
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    RSql = "SELECT BUSNO FROM VEHICLETYPE WHERE BUSNO = '" & Busno & "'"
    Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
    If RES.RecordCount > 0 Then BusNoExists = True
    Exit Function
err:
    Select Case err.Number
        Case Else
            MsgBox "Error No : " & err.Number & vbCrLf & err.Description, vbInformation, "Route"
            Exit Function
    End Select
End Function
