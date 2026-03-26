VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmVehicle 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   5085
   ControlBox      =   0   'False
   Icon            =   "frmCrew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin JeweledBut.JeweledButton cmdCancel 
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   2955
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      MICON           =   "frmCrew.frx":0CCA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdSave 
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2955
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      TX              =   "&Save"
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
      MICON           =   "frmCrew.frx":0CE6
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2685
      Left            =   180
      TabIndex        =   2
      Top             =   870
      Width           =   4710
      Begin VB.ComboBox cmbBusType 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtBusNo 
         Height          =   330
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1440
         Width           =   2580
      End
      Begin VB.TextBox txtBusId 
         Height          =   330
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   0
         Top             =   300
         Width           =   2580
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "BUS TYPE NAME"
         Height          =   255
         Left            =   225
         TabIndex        =   7
         Top             =   870
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "BUS NO"
         Height          =   255
         Left            =   225
         TabIndex        =   6
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "BUS  ID"
         Height          =   285
         Left            =   225
         TabIndex        =   5
         Top             =   300
         Width           =   1755
      End
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
      Left            =   1230
      TabIndex        =   8
      Top             =   255
      Width           =   2610
   End
End
Attribute VB_Name = "frmVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cHandle As Integer
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
    'Set RSDT = New ADODB.Recordset
    CONNECTDB
        If CheckCrewTableExistsOrNot("CREWDET") = False Then
            sql = "DELETE FROM CREWDET WHERE ID=" & Id
            DbZb.Execute (sql)
            MsgBox "Details deleted successfully from database", vbInformation, gblstrPrjTitle
'            RSDT.CursorLocation = adUseClient
'            RSDT.Open strSQL, DbZb, adOpenStatic, adLockOptimistic
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
                MsgBox "Record Not Found", vbInformation
             End If
             If GetTotalCrewRecord() > 250 Then
                MsgBox "Record Limit exceeds Maximum! Record Addition Failed", vbInformation
                
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
    
    CheckTableExistsOrNot = False
    
    If DB.State <> 0 Then DB.Close
        DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source= " & sDataBase & "; Jet OLEDB:Database Password = softland"
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



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim Ctrl As Control
    If Trim(txtDriver.Text) = "" And Trim(txtConductor.Text) = "" And Trim(txtBus.Text) = "" And Trim(txtCleaner.Text) = "" Then
        MsgBox "No Data to save! Please enter any details.", vbInformation, App.ProductName
        Exit Sub
    End If
    
'    For Each Ctrl In Me.Controls
'        If TypeOf Ctrl Is TextBox Then
'            If Ctrl.Text = "" Then
'                MsgBox "Some field missing!", vbInformation, "Crew Details"
'                Exit Sub
'            End If
'        End If
'    Next
    
    Crew.Conductor = Trim(txtConductor) & Chr(0)
    Crew.Driver = Trim(txtDriver) & Chr(0)
    Crew.Cleaner = Trim(txtCleaner) & Chr(0)
    Crew.BusNo = Trim(txtBus) & Chr(0)
    If Dir(App.Path & "\CREW.DAT") <> "" Then Kill App.Path & "\CREW.DAT"
    cHandle = FreeFile()
    Open App.Path & "\CREW.DAT" For Binary Access Write As #cHandle
        Put #cHandle, , Crew
    Close #cHandle
    CREW_FLAG = True
    MsgBox "Crew Details Successfully saved", vbInformation, "CREW"
    Unload Me
End Sub

Private Sub Form_Load()
    CREW_FLAG = False
    cHandle = FreeFile()
    If Dir(App.Path & "\CREW.DAT") <> "" Then
        Open App.Path & "\CREW.DAT" For Binary Access Read As #cHandle
            Get #cHandle, , Crew
        Close #cHandle
        txtConductor = Crew.Conductor
        txtDriver = Crew.Driver
        txtCleaner = Crew.Cleaner
        txtBus = Crew.BusNo
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
