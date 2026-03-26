VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmCrewDetails 
   BorderStyle     =   0  'None
   Caption         =   "Crew Details"
   ClientHeight    =   10410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15210
   Icon            =   "frmCrewDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin JeweledBut.JeweledButton cmdAdd 
      Height          =   495
      Left            =   10320
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      TX              =   "&Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmCrewDetails.frx":6852
      BC              =   12632256
      FC              =   0
   End
   Begin VB.ComboBox cmbBusType 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   7
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txtGrdVal 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      MaxLength       =   15
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtBusNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      MaxLength       =   15
      TabIndex        =   6
      Top             =   2520
      Width           =   1395
   End
   Begin VB.TextBox txtCleaner_Name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      MaxLength       =   63
      TabIndex        =   5
      ToolTipText     =   "Maximum 22 charactors can be printed and 15 charactors can be displayed"
      Top             =   3060
      Width           =   2955
   End
   Begin VB.TextBox txtCleaner_Code 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      MaxLength       =   7
      TabIndex        =   4
      Top             =   2580
      Width           =   1035
   End
   Begin VB.TextBox txtCond_Name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      MaxLength       =   63
      TabIndex        =   3
      ToolTipText     =   "Maximum 22 charactors can be printed and 15 charactors can be displayed"
      Top             =   1860
      Width           =   2955
   End
   Begin VB.TextBox txtCond_Code 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1380
      Width           =   1035
   End
   Begin VB.TextBox txtDrvr_Name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      MaxLength       =   22
      TabIndex        =   1
      ToolTipText     =   "Maximum 22 charactors can be printed and 15 charactors can be displayed"
      Top             =   1860
      Width           =   2955
   End
   Begin VB.TextBox txtDrvr_Code 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1380
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid flGrd 
      Height          =   5775
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   10186
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      BackColor       =   16777215
      BackColorFixed  =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton cmdSave 
      Height          =   495
      Left            =   11400
      TabIndex        =   9
      Top             =   9720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      TX              =   "&Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmCrewDetails.frx":686E
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdDelete 
      Height          =   495
      Left            =   12720
      TabIndex        =   10
      Top             =   9720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      TX              =   "&Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmCrewDetails.frx":688A
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdExit 
      Height          =   495
      Left            =   13920
      TabIndex        =   11
      Top             =   9720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmCrewDetails.frx":68A6
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Bus Type"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CREW DETAILS"
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
      Height          =   495
      Left            =   3960
      TabIndex        =   21
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bus No"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   19
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cleaner Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cleaner Code"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Conductor Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   1920
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Conductor Code"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Driver Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Driver Code"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmCrewDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Id, newid As Integer
Dim sql, strSql As String
Dim Ctrl As Control
Dim flclick As Boolean
Dim cdtr_code As String
Dim cdtr_name As String
Dim Dr_code As String
Dim Dr_Name As String
Dim cln_code As String
Dim cln_name As String
Dim bus_no As String
Dim BusType As String
Public rowselected As Integer


Public Function EditCrewToDatabase(cdtr_code As String, cdtr_name As String, Dr_code As String, Dr_Name As String, cln_code As String, cln_name As String, bus_no As String) As Boolean
On Error GoTo err
    CONNECTDB
         newid = flGrd.TextMatrix(rowselected, 0)
         
           strSql = "SELECT * FROM CREWDET WHERE ID=" & newid
            Set res1 = CNN.OpenRecordset(strSql, dbOpenDynaset)
            If res1.RecordCount > 0 Then
                        res1.Edit
                        res1!Id = flGrd.TextMatrix(newid, 0)
                        res1!DR_ID = flGrd.TextMatrix(newid, 1)
                        res1!Dr_Name = flGrd.TextMatrix(newid, 2)
                        res1!CDTR_ID = flGrd.TextMatrix(newid, 3)
                        res1!cdtr_name = flGrd.TextMatrix(newid, 4)
                        res1!CLNR_ID = flGrd.TextMatrix(newid, 5)
                        res1!CLNR_NAME = flGrd.TextMatrix(newid, 6)
                        res1!bus_no = flGrd.TextMatrix(newid, 7)
                        res1!BusTypeName = flGrd.TextMatrix(newid, 8)
                        res1.Update
                        res1.Close
                       EditCrewToDatabase = True
             End If
    Exit Function
err:
   MsgBox "Error due to " & err.Description, vbCritical, gblstrPrjTitle
End Function
Public Function GetTotalCrewRecord() As Long
On Error GoTo err
Dim cnt As Integer
   CONNECTDB

        sql = "SELECT count(ID) as CREWCOUNT FROM CREWDET "
           Set res1 = CNN.OpenRecordset(sql, adOpenDynamic)
          
            cnt = res1!CREWCOUNT
            GetTotalCrewRecord = res1!CREWCOUNT
            res1.Close
            
    Exit Function
err:
   MsgBox "Error due to " & err.Description, vbCritical, gblstrPrjTitle
End Function
Public Function DeleteCrewDetails() As Boolean
 Dim I As Integer
On Error GoTo err
    
    CONNECTDB
        newid = flGrd.TextMatrix(rowselected, 0)
        Msg = MsgBox("Are you sure to delete ?", vbYesNo)
        
            If (Msg = vbNo) Then
                Exit Function
            End If
            
            strSql = "DELETE * FROM CREWDET WHERE ID=" & newid
            CNN.Execute (strSql)

            sql = "UPDATE CREWDET SET ID=ID - 1 WHERE ID>" & newid
            CNN.Execute (sql)
    
            sql = "SELECT * FROM CREWDET"
            Set res1 = CNN.OpenRecordset(sql, adOpenDynamic)
            ClearGrid
            flGrd.TextMatrix(0, 0) = "ID"
            flGrd.TextMatrix(0, 1) = "Driver Code"
            flGrd.TextMatrix(0, 2) = "Driver Name"
            flGrd.TextMatrix(0, 3) = "Conductor Code"
            flGrd.TextMatrix(0, 4) = "Conductor Name"
            flGrd.TextMatrix(0, 5) = "Cleaner Code"
            flGrd.TextMatrix(0, 6) = "Cleaner Name"
            flGrd.TextMatrix(0, 7) = "Bus No"
            flGrd.TextMatrix(0, 8) = "Bus Type"
            If res1.RecordCount > 0 Then
            With flGrd
                .Rows = 1
                .Cols = 9
                While res1.EOF <> True
                 .Rows = .Rows + 1
                 .TextMatrix(.Rows - 1, 0) = res1!Id                'ID
                 .TextMatrix(.Rows - 1, 1) = res1!DR_ID             'Driver Code
                 .TextMatrix(.Rows - 1, 2) = res1!Dr_Name           'Driver Name
                 .TextMatrix(.Rows - 1, 3) = res1!CDTR_ID           'Conductor Code
                 .TextMatrix(.Rows - 1, 4) = res1!cdtr_name         'Conductor Name
                 .TextMatrix(.Rows - 1, 5) = res1!CLNR_ID           'Cleaner Code
                 .TextMatrix(.Rows - 1, 6) = res1!CLNR_NAME         'Cleaner Name
                 .TextMatrix(.Rows - 1, 7) = res1!bus_no            'Bus No
                 .TextMatrix(.Rows - 1, 8) = res1!BusTypeName       'Bus Type
                 res1.MoveNext
                Wend
            End With
            End If
            res1.Close
            DeleteCrewDetails = True
        
    Exit Function
err:
   MsgBox "Error due to " & err.Description, vbCritical, gblstrPrjTitle

End Function
Public Function CheckCrewTableExistsOrNot(table_name As String) As Boolean
On Error GoTo err
    CONNECTDB
    Set RSDT = New ADODB.Recordset
        If CheckTableExistsOrNot("CREWDET") = False Then
            If RSDT.State <> 0 Then RSDT.Close
                sSQL = "CREATE TABLE CREWDET " & _
                        "(ID NUMBER, " & _
                        "CDTR_ID VARCHAR(8), " & _
                        "CDTR_NAME VARCHAR(64), " & _
                        "DR_ID VARCHAR(8), " & _
                        "DR_NAME VARCHAR(64), " & _
                        "CLNR_ID VARCHAR(8), " & _
                        "CLNR_NAME VARCHAR(64), " & _
                        "BUS_NO VARCHAR(16), " & _
                        "BUSTYPENAME VARCHAR(64)," & _
                        "BUSTYPEID VARCHAR(8)) "
                DbZb.Execute (sSQL)
            End If
        Exit Function

err:
   MsgBox "Error due to " & err.Description, vbCritical, gblstrPrjTitle
    
End Function

Public Function AddCrewToDatabase(cdtr_code As String, cdtr_name As String, Dr_code As String, Dr_Name As String, cln_code As String, cln_name As String, bus_no As String, BusType As String) As Boolean
On Error GoTo err
Dim strSql, sql1 As String
Dim rec As Integer
    Set RSDT = New ADODB.Recordset
    CONNECTDB
        
            Id = 0
            strSql = "SELECT * FROM CREWDET "
            Set res1 = CNN.OpenRecordset(strSql, dbOpenDynaset)
        
            rec = res1.RecordCount
            If res1.RecordCount > 0 Then
                res1.MoveLast
                Id = res1!Id
            Else
                Id = 0
            End If
                    
            res1.AddNew
                
                Id = Id + 1
                    res1!Id = Id
                    res1!CDTR_ID = cdtr_code
                    res1!cdtr_name = cdtr_name
                    res1!DR_ID = Dr_code
                    res1!Dr_Name = Dr_Name
                    res1!CLNR_ID = cln_code
                    res1!CLNR_NAME = cln_name
                    res1!bus_no = bus_no
                    res1!BusTypeName = BusType
                    
                    sql1 = "SELECT * FROM BUSTYPE"
                     Set RES4 = CNN.OpenRecordset(sql1, adOpenDynamic)
                     If RES4.RecordCount > 0 Then
                        Do While Not RES4.EOF
                            If (res1!BusTypeName = RES4!Name) Then
                                res1!BustypeID = RES4!Id
                                'res4.MoveNext
                                Exit Do
                            End If
                            RES4.MoveNext
                        Loop
                    End If
                    RES4.Close
                    res1.Update
                    res1.Close
                    
                    txtDrvr_Code = ""
                    txtDrvr_Name = ""
                    txtCond_Code = ""
                    txtCond_Name = ""
                    txtCleaner_Code = ""
                    txtCleaner_Name = ""
                    txtBusNo = ""
                    txtDrvr_Code.SetFocus
                    AddCrewToDatabase = True
            
             If rec > 249 Then
                MsgBox "Record Limit exceeds Maximum! Record Addition Failed", vbInformation
                Exit Function
             End If
             
        Exit Function

err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Function



Private Sub cmdAdd_Click()
Dim sql1 As String
On Error GoTo err
    cdtr_code = txtCond_Code
    cdtr_name = txtCond_Name
    Dr_code = txtDrvr_Code
    Dr_Name = txtDrvr_Name
    cln_code = txtCleaner_Code
    cln_name = txtCleaner_Name
    bus_no = txtBusNo
    BusType = cmbBusType
    
'    If (cdtr_code = "") Or (cdtr_name = "") Or (Dr_code = "") Or (Dr_Name = "") Or (cln_code = "") Or (cln_name = "") Or (bus_no = "") Then
'        MsgBox "No fields should be empty", vbInformation
'        Exit Sub
'    End If
    
    If AddCrewToDatabase(cdtr_code, cdtr_name, Dr_code, Dr_Name, cln_code, cln_name, bus_no, BusType) = False Then Exit Sub

    sql = "SELECT * FROM CREWDET"
    Set res1 = CNN.OpenRecordset(sql, adOpenDynamic)
    ClearGrid
    flGrd.TextMatrix(0, 0) = "ID"
    flGrd.TextMatrix(0, 1) = "Driver Code"
    flGrd.TextMatrix(0, 2) = "Driver Name"
    flGrd.TextMatrix(0, 3) = "Conductor Code"
    flGrd.TextMatrix(0, 4) = "Conductor Name"
    flGrd.TextMatrix(0, 5) = "Cleaner Code"
    flGrd.TextMatrix(0, 6) = "Cleaner Name"
    flGrd.TextMatrix(0, 7) = "Bus No"
    flGrd.TextMatrix(0, 8) = "Bus Type"
    With flGrd
        .Rows = 1
        .Cols = 9
        res1.MoveFirst
        While res1.EOF <> True
         .Rows = .Rows + 1

         .TextMatrix(.Rows - 1, 0) = res1!Id                'ID
         .TextMatrix(.Rows - 1, 1) = res1!DR_ID             'Driver Code
         .TextMatrix(.Rows - 1, 2) = res1!Dr_Name           'Driver Name
         .TextMatrix(.Rows - 1, 3) = res1!CDTR_ID           'Conductor Code
         .TextMatrix(.Rows - 1, 4) = res1!cdtr_name         'Conductor Name
         .TextMatrix(.Rows - 1, 5) = res1!CLNR_ID           'Cleaner Code
         .TextMatrix(.Rows - 1, 6) = res1!CLNR_NAME         'Cleaner Name
         .TextMatrix(.Rows - 1, 7) = res1!bus_no            'Bus No
         .TextMatrix(.Rows - 1, 8) = res1!BusTypeName       'Bus Type
         
'         sql1 = "SELECT * FROM BUSTYPE"
'         Set res4 = CNN.OpenRecordset(sql1, adOpenDynamic)
'         If res4.RecordCount > 0 Then
'            Do While Not res4.EOF
'                If (res1!BusTypeName = res4!Name) Then
'                    res1!BUSTYPEID = res4!Id
'                End If
'                res4.MoveNext
'            Loop
'        End If
'         res4.Close
         
        res1.MoveNext
        Wend
    End With
    
    Exit Sub
err:
    Call ErrorHandle("cmdAdd_Click", err.Number, err.Description)
End Sub

Private Sub cmdDelete_Click()
On Error GoTo err
If DeleteCrewDetails() = False Then Exit Sub
         
Exit Sub
err:
    Call ErrorHandle("cmdDelete_Click", err.Number, err.Description)
End Sub

Private Sub cmdExit_Click()
On Error GoTo err
    Unload Me
    
    Exit Sub
err:
    Call ErrorHandle("cmdExit_Click", err.Number, err.Description)
End Sub

Private Sub cmdSave_Click()
On Error GoTo err

  If EditCrewToDatabase(cdtr_code, cdtr_name, Dr_code, Dr_Name, cln_code, cln_name, bus_no) = False Then Exit Sub
    Exit Sub
err:
    Call ErrorHandle("cmdSave_Click", err.Number, err.Description)
End Sub

Private Sub flGrd_Click()
On Error GoTo err
    With flGrd
        If .row = 0 Or .Col = 0 Then Exit Sub
        txtGrdVal.Left = .CellLeft + .Left
        txtGrdVal.Top = .CellTop + .Top
        txtGrdVal.Height = .CellHeight
        txtGrdVal.Width = .CellWidth
        txtGrdVal.FontBold = True
        txtGrdVal.ForeColor = vbRed
        txtGrdVal.Text = .TextMatrix(.row, .Col)
        rowselected = .row
    End With
    txtGrdVal.Visible = True
    txtGrdVal.SetFocus
    Exit Sub
err:
    Call ErrorHandle("flGrd_Click", err.Number, err.Description)
End Sub

Private Sub Form_Load()
On Error GoTo err
Call FillBusTypeName

    CONNECTDB
    Set RSDT = New ADODB.Recordset
        If CheckTableExistsOrNot("CREWDET") = False Then
            If RSDT.State <> 0 Then RSDT.Close
                sSQL = "CREATE TABLE CREWDET " & _
                        "(ID NUMBER, " & _
                        "DR_ID VARCHAR(8), " & _
                        "DR_NAME VARCHAR(64), " & _
                        "CDTR_ID VARCHAR(8), " & _
                        "CDTR_NAME VARCHAR(64), " & _
                        "CLNR_ID VARCHAR(8), " & _
                        "CLNR_NAME VARCHAR(64), " & _
                        "BUS_NO VARCHAR(16), " & _
                        "BUSTYPENAME VARCHAR(64)," & _
                        "BUSTYPEID VARCHAR(8)) "
                       
                    CNN.Execute (sSQL)
            End If
        
        For Each Ctrl In Me
            If TypeOf Ctrl Is TextBox Then
                Ctrl = ""
            End If
        Next
        
        flGrd.Clear
        flGrd.Rows = 1
    
'    For Each Ctrl In Me
'        If TypeOf Ctrl Is TextBox Then
'            If Ctrl = "" Then
'                If Ctrl.Name <> "txtCleaner_Name" And Ctrl.Name <> "txtCleaner_Code" And Ctrl.Name <> "txtGrdVal" Then
'                    MsgBox "Invalid data" & vbCrLf & "Please check the fields", vbExclamation
'                    Exit Sub
'                End If
'            End If
'        End If
'    Next
'    With flGrd
'       .TextMatrix(1, 0) = "ID"
'       .TextMatrix(1, 1) = "Driver Code"
'       .TextMatrix(1, 2) = "Driver Name"
'    End With
    
            
    sql = "SELECT * FROM CREWDET"
    Set res1 = CNN.OpenRecordset(sql, adOpenDynamic)
    ClearGrid
    flGrd.ColWidth(0) = 1000
    flGrd.ColWidth(1) = 1500
    flGrd.ColWidth(2) = 2450
    flGrd.ColWidth(3) = 1800
    flGrd.ColWidth(4) = 2450
    flGrd.ColWidth(5) = 1525
    flGrd.ColWidth(6) = 2450
    flGrd.ColWidth(7) = 1500
    flGrd.ColWidth(8) = 2450
    
    flGrd.TextMatrix(0, 0) = "ID"
    flGrd.TextMatrix(0, 1) = "Driver Code"
    flGrd.TextMatrix(0, 2) = "Driver Name"
    flGrd.TextMatrix(0, 3) = "Conductor Code"
    flGrd.TextMatrix(0, 4) = "Conductor Name"
    flGrd.TextMatrix(0, 5) = "Cleaner Code"
    flGrd.TextMatrix(0, 6) = "Cleaner Name"
    flGrd.TextMatrix(0, 7) = "Bus No"
    flGrd.TextMatrix(0, 8) = "Bus Type"
    If res1.RecordCount > 0 Then
    With flGrd
        .Rows = 1
        .Cols = 9
        While res1.EOF <> True
         .Rows = .Rows + 1
         .TextMatrix(.Rows - 1, 0) = res1!Id                'ID
         .TextMatrix(.Rows - 1, 1) = res1!DR_ID             'Driver Code
         .TextMatrix(.Rows - 1, 2) = res1!Dr_Name           'Driver Name
         .TextMatrix(.Rows - 1, 3) = res1!CDTR_ID           'Conductor Code
         .TextMatrix(.Rows - 1, 4) = res1!cdtr_name         'Conductor Name
         .TextMatrix(.Rows - 1, 5) = res1!CLNR_ID           'Cleaner Code
         .TextMatrix(.Rows - 1, 6) = res1!CLNR_NAME         'Cleaner Name
         .TextMatrix(.Rows - 1, 7) = res1!bus_no            'Bus No
         .TextMatrix(.Rows - 1, 8) = res1!BusTypeName       'Bus Type
         res1.MoveNext
        Wend
    End With
    End If
    cmbBusType.Text = cmbBusType.List(0)
    Exit Sub
err:
    Call ErrorHandle("Form_Load", err.Number, err.Description)
End Sub


Private Sub txtBusNo_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 And txtBusNo <> "" Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub txtCleaner_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 And txtCleaner_Code <> "" Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub txtCleaner_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCleaner_Name <> "" Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub txtCond_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 And txtCond_Code <> "" Then
        SendKeys ("{TAB}")
    End If
End Sub


Private Sub txtCond_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCond_Name <> "" Then
        SendKeys ("{TAB}")
    End If
End Sub



Private Sub txtDrvr_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 And txtDrvr_Code <> "" Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub txtDrvr_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtDrvr_Name <> "" Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub txtGrdVal_Change()
On Error GoTo err
    If txtGrdVal <> "" Then
        flGrd.TextMatrix(flGrd.row, flGrd.Col) = txtGrdVal
        flGrd.CellFontBold = True
        flGrd.CellFontSize = 12
    End If
    Exit Sub
err:
    Call ErrorHandle("txtVal_Change", err.Number, err.Description)
End Sub

Private Sub txtGrdVal_GotFocus()
    txtGrdVal.FontSize = 12
    txtGrdVal.SelStart = 0
    txtGrdVal.SelLength = Len(txtGrdVal)
End Sub

Private Sub txtGrdVal_KeyPress(KeyAscii As Integer)
On Error GoTo err
    
    If KeyAscii = Asc(vbTab) Or KeyAscii = 13 Or KeyAscii = 27 Or KeyAscii = Asc(vbBack) Then
        If KeyAscii = 13 Or KeyAscii = Asc(vbTab) Then
            
            With flGrd
                If .row = .Rows - 1 And .Col = .Cols - 1 Then
                    txtGrdVal.Visible = False
                    cmdSave.SetFocus
                    Exit Sub
                End If
                If .row < .Rows - 1 And .Col = .Cols - 1 Then
                    .row = .row + 1
                    .Col = 1
                    .CellFontBold = True
                    .CellFontSize = 12
                    .TextMatrix(.row, .Col) = .TextMatrix(.row - 1, .Col + 1)
                    .Col = 2
                    txtGrdVal = ""
                Else
                    If .Col < .Cols - 1 Then .Col = .Col + 1
                End If

                txtGrdVal.Left = .CellLeft + .Left
                txtGrdVal.Top = .CellTop + .Top
                txtGrdVal.Height = .CellHeight
                txtGrdVal.Width = .CellWidth
    
                txtGrdVal = .TextMatrix(flGrd.row, flGrd.Col)
                txtGrdVal.SetFocus
                txtGrdVal.FontSize = 12
                txtGrdVal.SelStart = 0
                txtGrdVal.SelLength = Len(txtGrdVal)
                
            End With
        ElseIf KeyAscii = 27 Then
            txtGrdVal.Visible = False
            flGrd.SetFocus
        End If
        Exit Sub
    End If
    Exit Sub
err:
    Call ErrorHandle("txtGrdVal_KeyPress", err.Number, err.Description)
End Sub

Public Sub CreateErrorLog(ErrorMessage As String)
On Error Resume Next
Dim Handle As Integer
    Handle = FreeFile
    If Dir(App.Path & "\Error Log.txt") = "" Then
        Open App.Path & "\Error Log.txt" For Output As #Handle
    Else
        Open App.Path & "\Error Log.txt" For Append As #Handle
    End If
    ErrorMessage = "----------------------------------------------------------------------" & vbCrLf & _
                   Day(Date) & "/" & Month(Date) & "/" & Year(Date) & " - " & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & vbCrLf & _
                   ErrorMessage & vbCrLf & _
                   "----------------------------------------------------------------------" & vbCrLf
    Print #Handle, ErrorMessage
    Close #Handle
End Sub


Public Function ErrorHandle(strError As String, ErrNumber As Integer, ErrDescription As String)
    Dim sErrorString As String
    sErrorString = strError & "_Error!" & vbCrLf & "Error No :" & ErrNumber & vbCrLf & ErrDescription
    Call CreateErrorLog(sErrorString)
    MsgBox strError & "_Error!" & vbCrLf & "Error No :" & ErrNumber & vbCrLf & ErrDescription, vbInformation, gblstrPrjTitle
End Function

Public Function ClearGrid()
flGrd.Clear
End Function
Private Sub FillBusTypeName()
Dim BTYPE As DAO.Database
Dim BTYPEREC As DAO.Recordset
    Set BTYPE = DAO.OpenDatabase(App.Path & "\pvt.mdb")
 
    TSQL = "SELECT * FROM BUSTYPE"
    Set BTYPEREC = BTYPE.OpenRecordset(TSQL, dbOpenDynaset)
    If BTYPEREC.RecordCount > 0 Then
        BTYPEREC.MoveFirst
        'cmbBusType.Clear
        Do While Not BTYPEREC.EOF
            cmbBusType.AddItem (BTYPEREC!Name)
            ''MsgBox BTYPEREC!Name
            BTYPEREC.MoveNext
            Loop
    End If
End Sub
