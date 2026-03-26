VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmExpense 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expense"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   5895
      Begin JeweledBut.JeweledButton JeweledButton1 
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   1080
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
         MICON           =   "frmExpense.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin VB.TextBox txtExpcode 
         Height          =   375
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtExpname 
         Height          =   375
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin JeweledBut.JeweledButton cmdAdd 
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   480
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         TX              =   "&Add"
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
         MICON           =   "frmExpense.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "EXPENSE "
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
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   -600
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Exp.Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Exp.Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin JeweledBut.JeweledButton cmdSave 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   5640
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      TX              =   "&Update"
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
      MICON           =   "frmExpense.frx":0038
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdDelete 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   5640
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
      MICON           =   "frmExpense.frx":0054
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdCancel 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   5640
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
      MICON           =   "frmExpense.frx":0070
      BC              =   12632256
      FC              =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flGrd 
      Height          =   3375
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   15846563
      BackColorBkg    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
      FormatString    =   "SL.NO           |   EXP CODE     | EXP NAME                "
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
End
Attribute VB_Name = "frmExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rowselected As Integer
Dim Id, newid As Integer
Dim sql, strSql As String
Dim Ctrl As Control
Dim flclick As Boolean
Dim EXP_CODE As String
Dim exp_name As String
Dim flgupdate As Double
Dim oldexpval, oldexpvalname
Public Function ClearGrid()
    flGrd.Clear
End Function
Public Function DeleteExpenseDetails() As Boolean
Dim I As Integer
On Error GoTo err
    CONNECTDB
        'newid = flGrd.TextMatrix(rowselected, 0)
        Msg = MsgBox("Are you sure to delete ?", vbYesNo)
            If (Msg = vbNo) Then
                Exit Function
            End If
            strSql = "DELETE * FROM EXPMASTER WHERE ID=" & newid
            CNN.Execute (strSql)

            sql = "UPDATE EXPMASTER SET ID=ID - 1 WHERE ID>" & newid
            CNN.Execute (sql)
    
            sql = "SELECT * FROM EXPMASTER "
            Set res1 = CNN.OpenRecordset(sql, adOpenDynamic)
            ClearGrid
            flGrd.TextMatrix(0, 0) = "ID"
            flGrd.TextMatrix(0, 1) = "Expense Code"
            flGrd.TextMatrix(0, 2) = "Name"
            
            If res1.RecordCount > 0 Then
            With flGrd
                .Rows = 2
                .Cols = 3
                
                .TextMatrix(.Rows - 1, 0) = "0"
                .TextMatrix(.Rows - 1, 1) = "1"        'Expense Code
                .TextMatrix(.Rows - 1, 2) = "Diesel Entry"
                
                While res1.EOF <> True
                 .Rows = .Rows + 1
                 .TextMatrix(.Rows - 1, 0) = res1!Id                'ID
                 .TextMatrix(.Rows - 1, 1) = res1!EXP_CODE             'Expense Code
                 .TextMatrix(.Rows - 1, 2) = res1!exp_name            ' Name
                 
                 res1.MoveNext
                Wend
            End With
            Else
               
                 With flGrd
                    .Rows = 2
                    .Cols = 3 '3
                    .TextMatrix(.Rows - 1, 0) = "0"
                    .TextMatrix(.Rows - 1, 1) = "1"        'Expense Code
                    .TextMatrix(.Rows - 1, 2) = "Diesel Entry"
                End With
            End If
            res1.Close
            DeleteExpenseDetails = True
            txtGrdVal.Visible = False
            cmdDelete.Enabled = False
            cmdSave.Enabled = False
    Exit Function
err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"

End Function
Private Sub cmdDelete_Click()
On Error GoTo err
If DeleteExpenseDetails() = False Then Exit Sub
cmdAdd.Enabled = True
txtExpcode.Text = ""
txtExpname.Text = ""
Exit Sub
err:
    Call ErrorHandle("cmdDelete_Click", err.Number, err.Description)
End Sub


Public Function AddExpenseToDatabase(EXP_CODE As String, exp_name As String) As Boolean
On Error GoTo err
Dim strSql, sql1 As String
Dim rec As Integer
    Set RSDT = New ADODB.Recordset
    CONNECTDB
        
            Id = 0
            strSql = "SELECT * FROM EXPMASTER "
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
                    res1!EXP_CODE = EXP_CODE
                    res1!exp_name = exp_name
                    
'                    sql1 = "SELECT PalmtecID FROM SETTINGS"
'                     Set res4 = CNN.OpenRecordset(sql1, adOpenDynamic)
'                    res1!PalmId = res4!PalmtecID
'                    res4.Close
                    
                    res1.Update
                    res1.Close
                    
                    txtExpcode = ""
                    txtExpname = ""
                    txtExpcode.SetFocus
                    AddExpenseToDatabase = True
            
             If rec > 249 Then
                MsgBox "Record Limit exceeds Maximum! Record Addition Failed", vbInformation
                Exit Function
             End If
             
        Exit Function

err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Function

Private Sub cmdAdd_Click()
On Error GoTo err
Dim rec As Integer
    EXP_CODE = txtExpcode
    exp_name = txtExpname
    'txtExpcode.SetFocus
    If EXP_CODE = "" Or exp_name = "" Then
        MsgBox "No fields should be empty", vbInformation, gblstrPrjTitle
        Exit Sub
    End If
    If EXP_CODE = 0 Then
        MsgBox "EXP_CODE should be greater than Zero", vbInformation
        Exit Sub
    End If
     strSql = "SELECT count(*) as cnt FROM EXPMASTER"
     Set res1 = CNN.OpenRecordset(strSql, dbOpenDynaset)
     rec = res1!cnt
     Set res1 = Nothing
     If rec = 24 Then
         MsgBox "Expense Entry Limit Reached!", vbInformation
         txtExpcode.Text = ""
         txtExpname.Text = ""
         Exit Sub
     End If
    If AddExpenseToDatabase(EXP_CODE, exp_name) = False Then Exit Sub

    sql = "SELECT * FROM EXPMASTER"
    Set res1 = CNN.OpenRecordset(sql, adOpenDynamic)
    ClearGrid
    flGrd.TextMatrix(0, 0) = "SL.No"
    flGrd.TextMatrix(0, 1) = "Expense Code"
    flGrd.TextMatrix(0, 2) = "Name"
    
    With flGrd
        .Rows = 2
        .Cols = 3
        res1.MoveFirst
        
        .TextMatrix(.Rows - 1, 0) = "0"
        .TextMatrix(.Rows - 1, 1) = "1"        'Expense Code
        .TextMatrix(.Rows - 1, 2) = "Diesel Entry"
        While res1.EOF <> True
         .Rows = .Rows + 1

         .TextMatrix(.Rows - 1, 0) = res1!Id                'ID
         .TextMatrix(.Rows - 1, 1) = res1!EXP_CODE             'Expense Code
         .TextMatrix(.Rows - 1, 2) = res1!exp_name           'Name
        
        res1.MoveNext
        Wend
    End With
    
    Exit Sub
err:
    Call ErrorHandle("cmdAdd_Click", err.Number, err.Description)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Function GetTotalExpenseRecord() As Long
On Error GoTo err
Dim cnt As Integer
   CONNECTDB

        sql = "SELECT count(ID) as EXPENSECOUNT FROM EXPMASTER "
           Set res1 = CNN.OpenRecordset(sql, adOpenDynamic)
          
            cnt = res1!EXPENSECOUNT
            GetTotalExpenseRecord = res1!EXPENSECOUNT
            res1.Close
            
    Exit Function
err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Function

Public Function EditExpenseToDatabase(EXP_CODE As String, exp_name As String) As Boolean
On Error GoTo err
    CONNECTDB
         'newid = flGrd.TextMatrix(rowselected, 0)
         
           strSql = "SELECT * FROM EXPMASTER WHERE ID=" & newid
            Set res1 = CNN.OpenRecordset(strSql, dbOpenDynaset)
            If res1.RecordCount > 0 Then
                        res1.Edit
                        'res1!Id = flGrd.TextMatrix(newid, 0)
                        res1!EXP_CODE = txtExpcode
                        res1!exp_name = txtExpname
                        res1.Update
                        res1.Close
                       EditExpenseToDatabase = True
             End If
              flGrd.Clear
        flGrd.Rows = 1
    
    sql = "SELECT * FROM EXPMASTER"
    Set res1 = CNN.OpenRecordset(sql, adOpenDynamic)
    ClearGrid
    'flGrd.ColWidth(0) = 1000
    flGrd.ColWidth(0) = 0
    flGrd.ColWidth(1) = 1600
    flGrd.ColWidth(2) = 3500
    
    
    flGrd.TextMatrix(0, 0) = "SL.No"
    flGrd.TextMatrix(0, 1) = "ExpenseCode"
    flGrd.TextMatrix(0, 2) = " Name"
    
    If res1.RecordCount > 0 Then
    With flGrd
        .Rows = 2
        .Cols = 3
         .TextMatrix(.Rows - 1, 0) = "0"
         .TextMatrix(.Rows - 1, 1) = "1"        'Expense Code
         .TextMatrix(.Rows - 1, 2) = "Diesel Entry"
        While res1.EOF <> True
         .Rows = .Rows + 1
         .TextMatrix(.Rows - 1, 0) = res1!Id                'ID
         .TextMatrix(.Rows - 1, 1) = res1!EXP_CODE          'Expense Code
         .TextMatrix(.Rows - 1, 2) = res1!exp_name          'Name
        
         res1.MoveNext
        Wend
    End With
    End If
    res1.Close
    Exit Function
err:
   MsgBox "Error due to " & err.Description, vbCritical, "BusTicket"
End Function

Private Sub cmdSave_Click()
On Error GoTo err
   If txtExpcode.Text = "" Or txtExpname.Text = "" Then
        MsgBox "No fields should be empty", vbInformation
        Exit Sub
   End If
   If val(txtExpcode.Text) = 0 Then
        MsgBox "EXP_CODE should be greater than Zero", vbInformation
        Exit Sub
    End If
   If ExpIDExists(val(txtExpcode.Text)) = True Or val(txtExpcode.Text) = 1 Then
        MsgBox "Expense CODE already Exists" & vbCrLf & "Please give another Code", vbInformation, gblstrPrjTitle
        txtExpcode = ""
        txtExpcode.SetFocus
        Exit Sub
    End If
  If EditExpenseToDatabase(EXP_CODE, exp_name) = False Then Exit Sub
  cmdSave.Enabled = False
  txtGrdVal.Visible = False
  cmdDelete.Enabled = False
  cmdAdd.Enabled = True
  txtExpcode.Text = ""
  txtExpname.Text = ""
Exit Sub
err:
    Call ErrorHandle("cmdSave_Click", err.Number, err.Description)
End Sub

Private Sub flGrd_DblClick()
On Error GoTo lblErr
If flGrd.TextMatrix(flGrd.row, 1) <> "1" Then
    newid = val(flGrd.TextMatrix(flGrd.row, 0))
    txtExpcode.Text = flGrd.TextMatrix(flGrd.row, 1)
    txtExpname.Text = flGrd.TextMatrix(flGrd.row, 2)
    cmdSave.Enabled = True
    cmdDelete.Enabled = True
    cmdAdd.Enabled = False
Else
    txtExpcode.Text = ""
    txtExpname.Text = ""
    newid = 0
    cmdDelete.Enabled = False
    cmdSave.Enabled = False
End If
Exit Sub
lblErr:
End Sub

Private Sub Form_Load()
On Error GoTo err
    Me.Icon = frmMainform.Icon
    CONNECTDB
    Set RSDT = New ADODB.Recordset
        If CheckTableExistsOrNot("EXPMASTER") = False Then
            If RSDT.State <> 0 Then RSDT.Close
                sSQL = "CREATE TABLE EXPMASTER" & _
                        "(ID NUMBER, " & _
                        "EXP_CODE VARCHAR(8), " & _
                        "EXP_NAME VARCHAR(64)) "
                       
                    CNN.Execute (sSQL)
            End If
        
        For Each Ctrl In Me
            If TypeOf Ctrl Is TextBox Then
                Ctrl = ""
            End If
        Next
        
        flGrd.Clear
        flGrd.Rows = 1
    
    sql = "SELECT * FROM EXPMASTER"
    Set res1 = CNN.OpenRecordset(sql, adOpenDynamic)
    ClearGrid
    'flGrd.ColWidth(0) = 1000
    flGrd.ColWidth(0) = 0
    flGrd.ColWidth(1) = 1600
    flGrd.ColWidth(2) = 3500
    
    
    flGrd.TextMatrix(0, 0) = "SL.No  "
    flGrd.TextMatrix(0, 1) = "Expense Code"
    flGrd.TextMatrix(0, 2) = " Name"
    
    If res1.RecordCount > 0 Then
    With flGrd
        .Rows = 2
        .Cols = 3 '3
         .TextMatrix(.Rows - 1, 0) = "0"
         .TextMatrix(.Rows - 1, 1) = "1"        'Expense Code
         .TextMatrix(.Rows - 1, 2) = "Diesel Entry"
       
        While res1.EOF <> True
         .Rows = .Rows + 1
         .TextMatrix(.Rows - 1, 0) = res1!Id                'ID
         .TextMatrix(.Rows - 1, 1) = res1!EXP_CODE            'Expense Code
         .TextMatrix(.Rows - 1, 2) = res1!exp_name           ' Name
        
         res1.MoveNext
        Wend
    End With
    Else
         With flGrd
             .Rows = 2
             .Cols = 3 '3
             .TextMatrix(.Rows - 1, 0) = "0"
             .TextMatrix(.Rows - 1, 1) = "1"        'Expense Code
             .TextMatrix(.Rows - 1, 2) = "Diesel Entry"
           End With
    End If
    res1.Close
    'txtExpcode.SetFocus
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    Exit Sub
err:
    Call ErrorHandle("Form_Load", err.Number, err.Description)
End Sub

Private Sub flGrd_Click()
'On Error GoTo err
'     If flGrd.Col = 1 Then
'        oldexpval = flGrd.TextMatrix(flGrd.row, 1)
'    End If
'    If flGrd.Col = 2 Then
'        oldexpvalname = flGrd.TextMatrix(flGrd.row, 2)
'    End If
'    With flGrd
'        If .row = 0 Or .Col = 0 Then Exit Sub
'        txtGrdVal.Left = .CellLeft + .Left
'        txtGrdVal.Top = .CellTop + .Top
'        txtGrdVal.Height = .CellHeight
'        txtGrdVal.Width = .CellWidth
'        txtGrdVal.FontBold = True
'        txtGrdVal.ForeColor = vbRed
'        txtGrdVal.Text = .TextMatrix(.row, .Col)
'        rowselected = .row
'    End With
'    txtGrdVal.Visible = True
'    If flGrd.Col = 1 Then
'        txtGrdVal.Text = oldexpval
'    End If
'    If flGrd.Col = 2 Then
'        txtGrdVal.Text = oldexpvalname
'    End If
'    txtGrdVal.SetFocus
'    cmdSave.Enabled = True
'    cmdDelete.Enabled = True
'    Exit Sub
'err:
'    Call ErrorHandle("flGrd_Click", err.Number, err.Description)

lblErr:
End Sub

Private Sub JeweledButton1_Click()
    txtExpcode.Text = ""
    txtExpname.Text = ""
    newid = 0
    EXP_CODE = 0
    exp_name = 0
    cmdDelete.Enabled = False
    cmdSave.Enabled = False
    cmdAdd.Enabled = True
End Sub

Private Sub txtExpcode_LostFocus()
On Error GoTo lblErr
If cmdDelete.Enabled = False Then
If val(txtExpcode.Text) > 0 Then
    If ExpIDExists(val(txtExpcode.Text)) = True Or val(txtExpcode.Text) = 1 Then
        MsgBox "Expense CODE already Exists" & vbCrLf & "Please give another Code", vbInformation, "Expense"
        txtExpcode = ""
        txtExpcode.SetFocus
        Exit Sub
    End If
End If
End If
Exit Sub
lblErr:
End Sub
Public Function ExpIDExists(EXP_CODE As Double) As Boolean
On Error GoTo err
Dim cond As String
If newid > 0 Then
  cond = " and ID<>" & newid
End If
    ExpIDExists = False
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    RSql = "SELECT EXP_CODE FROM  EXPMASTER WHERE EXP_CODE = " & EXP_CODE & cond
    Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
    If RES.RecordCount > 0 Then ExpIDExists = True
    Exit Function
err:
    Select Case err.Number
        Case Else
            MsgBox "Error No : " & err.Number & vbCrLf & err.Description, vbInformation, "Route"
            Exit Function
    End Select
End Function
Private Sub txtGrdVal_Change()
'On Error GoTo err
'
'    If flGrd.Col = 1 And Len(txtGrdVal.Text) > 5 Then
'      txtGrdVal.Text = oldexpval
'      SendKeys ("{END}")
'    End If
'
'    If flGrd.Col = 2 And Len(txtGrdVal.Text) > 14 Then
'      txtGrdVal.Text = oldexpvalname
'      SendKeys ("{END}")
'    End If
'
'    If txtGrdVal <> "" Then
'        flGrd.TextMatrix(flGrd.row, flGrd.Col) = txtGrdVal
'        flGrd.CellFontBold = True
'        flGrd.CellFontSize = 12
'    End If
'    If flGrd.Col = 1 Then
'        oldexpval = Trim(txtGrdVal.Text)
'    End If
'    If flGrd.Col = 2 Then
'        oldexpvalname = Trim(txtGrdVal.Text)
'    End If
'    Exit Sub
'err:
'    Call ErrorHandle("txtVal_Change", err.Number, err.Description)
End Sub


Private Sub txtGrdVal_KeyPress(KeyAscii As Integer)
'On Error GoTo err
'    If KeyAscii = 13 Then
'        'If flGrd.row = 1 Then
'           If Trim(flGrd.TextMatrix(flGrd.row, 1)) = 1 Then
'                MsgBox "Expense CODE already Exists" & vbCrLf & "Please give another Code", vbInformation, gblstrPrjTitle
'               txtGrdVal.SetFocus
'               Exit Sub
'           End If
'        'End If
'     End If
'     If flGrd.Col = 1 Then
'        If TextBoxValidityNumeric(KeyAscii) > 0 Then
'            KeyAscii = 0
'            Exit Sub
'        End If
'    End If
'    If KeyAscii = Asc(vbTab) Or KeyAscii = 13 Or KeyAscii = 27 Or KeyAscii = Asc(vbBack) Then
'        If KeyAscii = 13 Or KeyAscii = Asc(vbTab) Then
'
'            With flGrd
'                If .row = .Rows - 1 And .Col = .Cols - 1 Then
'                    txtGrdVal.Visible = False
'                    cmdSave.SetFocus
'                    Exit Sub
'                End If
'                If .row < .Rows - 1 And .Col = .Cols - 1 Then
'                    .row = .row + 1
'                    .Col = 1
'                    .CellFontBold = True
'                    .CellFontSize = 12
'                    .TextMatrix(.row, .Col) = .TextMatrix(.row - 1, .Col + 1)
'                    .Col = 2
'                    txtGrdVal = ""
'                Else
'                    If .Col < .Cols - 1 Then .Col = .Col + 1
'                End If
'
'                txtGrdVal.Left = .CellLeft + .Left
'                txtGrdVal.Top = .CellTop + .Top
'                txtGrdVal.Height = .CellHeight
'                txtGrdVal.Width = .CellWidth
'
'                txtGrdVal = .TextMatrix(flGrd.row, flGrd.Col)
'                txtGrdVal.SetFocus
'                txtGrdVal.FontSize = 12
'                txtGrdVal.SelStart = 0
'                txtGrdVal.SelLength = Len(txtGrdVal)
'
'            End With
'        ElseIf KeyAscii = 27 Then
'            txtGrdVal.Visible = False
'            flGrd.SetFocus
'        End If
'        Exit Sub
'    End If
'    Exit Sub
'err:
'    Call ErrorHandle("txtGrdVal_KeyPress", err.Number, err.Description)
End Sub

Private Sub txtGrdVal_GotFocus()
'    txtGrdVal.FontSize = 12
'    txtGrdVal.SelStart = 0
'    txtGrdVal.SelLength = Len(txtGrdVal)
End Sub

Private Sub txtExpcode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtExpcode <> "" Then
        SendKeys ("{TAB}")
    End If
    If TextBoxValidityonlyNumeric(KeyAscii) > 0 Then
            KeyAscii = 0
            Exit Sub
    End If
End Sub

Private Sub txtExpname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtExpname <> "" Then
        SendKeys ("{TAB}")
    End If
End Sub

Public Function CheckTableExistsOrNot(strTableName As String) As Boolean
On Error GoTo err
    
    Dim strSql As String
    Dim DB As New ADODB.Connection
    Dim rs1 As ADODB.Recordset
    'CONNECTDB
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

Private Sub txtGrdVal_LostFocus()
On Error GoTo lblErr
If flGrd.row = 2 Then
    'If Trim(txtGrdVal.Text) = 1 Then
     If Trim(flGrd.TextMatrix(flGrd.row, 1)) = 1 Then
         MsgBox "Expense CODE already Exists" & vbCrLf & "Please give another Code", vbInformation, gblstrPrjTitle
         txtGrdVal.SetFocus
         Exit Sub
    End If
End If
Exit Sub
lblErr:
End Sub
