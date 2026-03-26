VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form Frm_importexport 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Excel Import"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   Icon            =   "Frm_importexport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog SelectMdb_CD 
      Left            =   1080
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox path_Txt 
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
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   5175
   End
   Begin JeweledBut.JeweledButton file_brows_cmd 
      Height          =   315
      Left            =   6360
      TabIndex        =   1
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      TX              =   "...."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Frm_importexport.frx":11F84
      BC              =   12632256
      FC              =   0
   End
   Begin CCRProgressBar6.ccrpProgressBar CCPB_Pbr 
      Height          =   135
      Left            =   0
      Top             =   2040
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   238
      BackColor       =   14737632
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
   Begin JeweledBut.JeweledButton Import_cmd 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      TX              =   "&Import"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Frm_importexport.frx":11FA0
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label lbl 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Excel File"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   795
      Width           =   885
   End
End
Attribute VB_Name = "Frm_importexport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExlObj As New excel.Application
Public xl As New excel.Application
Public xlwbook As excel.Workbook
Public xlsheet As excel.Worksheet
Public ExcelCon As New ADODB.Connection
Private Type master_table
    AccountNo As String
    accountname As String
    phonnumber As String
    Address As String
    duedate As Date
    OpeningDate As Date
    lasttransactdate As Date
    AccType As String
    multiple As Double
    LoanAmt As Double
    agntcaode As String
    agntName As String
    amtpayed As Double
    loanbaldue As Double
    Period  As Double
    Lintrest As Double
    penalty As Double
    Total As Double
    DoorNo As String
    StreetName As String
End Type
Private Sub export_btn_Click()
On Error GoTo err:
    CCPB_Pbr.Value = 0
    lbl.caption = ""
    Me.Enabled = False
    sql = "select receiptno,collectamt,Acno,AcType,collectdt,chequeno,ModePay,ag_code,amtpaid,interestpaid,penalty,Cancel,Palmtec_ID from transact"
    Set rs = CON.Execute(sql)
    While Not rs.EOF
        ExlObj.Workbooks.Add
        ExlObj.Visible = True
        lbl.caption = "Exporting is in progress.Please wait..."
        ex1 = 2
        ex2 = 1
        For k = 1 To rs.Fields.Count
            ExlObj.ActiveSheet.Cells(ex1 - 1, ex2).Value = IIf(IsNull(rs.Fields(k - 1).Name), "", rs.Fields(k - 1).Name)
            ExlObj.ActiveSheet.Cells(ex1 - 1, ex2).NumberFormat = "@"
            ExlObj.ActiveSheet.Cells(ex1 - 1, ex2).HorizontalAlignment = excel.xlCenter
            ExlObj.ActiveSheet.Cells(ex1 - 1, ex2).Font.Bold = True
            ex2 = ex2 + 1
        Next k
        ex2 = 1
        Do While Not rs.EOF
            For k = 1 To rs.Fields.Count
                ExlObj.ActiveSheet.Cells(ex1, ex2).Value = IIf(IsNull(rs.Fields(k - 1)), "", rs.Fields(k - 1))
                ExlObj.ActiveSheet.Cells(ex1, ex2).NumberFormat = "@"
                ExlObj.ActiveSheet.Cells(ex1, ex2).HorizontalAlignment = excel.xlCenter
                ex2 = ex2 + 1
            Next k
            ex1 = ex1 + 1
            ex2 = 1
            If Me.CCPB_Pbr.Value >= Me.CCPB_Pbr.Max - 1 Then
                Me.CCPB_Pbr.Value = 0
            Else
                Me.CCPB_Pbr.Value = Me.CCPB_Pbr.Value + 1
            End If
            rs.MoveNext
            Me.CCPB_Pbr.Value = Me.CCPB_Pbr.Max
        Loop
        Me.CCPB_Pbr.Value = Me.CCPB_Pbr.Max
        MsgBox "Exporting Completed successfully", vbInformation, PrjTitleMsg
        Me.CCPB_Pbr.Value = 0
        lbl.caption = ""
        Me.Enabled = True
    Wend
Exit Sub
err:
    MsgBox "Error in Excel Function for Transaction Report", vbCritical, "Sil_Paycollect_Ver.7 "
    Me.CCPB_Pbr.Value = 0
    lbl.caption = ""
    Me.Enabled = True
End Sub
Private Sub file_brows_cmd_Click()
On Error Resume Next
    SelectMdb_CD.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx|"
    SelectMdb_CD.InitDir = ExcelInputPath 'App.Path '"E:"
    SelectMdb_CD.ShowOpen
    If Len(SelectMdb_CD.filename) > 0 Then
        path_Txt = SelectMdb_CD.filename
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMainform.Icon
    Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim check_me As Boolean
    If val(getvalueQuery("select count(*) FROM ROUTE WHERE rutcode='" & Rutecode_new & "'")) > 0 Then
        If val(getvalueQuery("select count(*) FROM FARE WHERE route='" & Rutecode_new & "'")) > 0 Then
            If val(getvalueQuery("select count(*) FROM STAGE WHERE route='" & Rutecode_new & "'")) > 0 Then
                check_me = True
            End If
        End If
    End If
    If check_me = False Then
        sql = "DELETE * FROM ROUTE WHERE rutcode='" & Rutecode_new & "'"
        gbladoCon.Execute (sql)
        
        sql = "DELETE * FROM FARE WHERE route='" & Rutecode_new & "'"
        gbladoCon.Execute (sql)
        
        sql = "DELETE * FROM STAGE WHERE route='" & Rutecode_new & "'"
        gbladoCon.Execute (sql)
    End If
End Sub
Private Sub Import_cmd_Click()
Dim rcnt As Long
Dim I, j, k, clnj, m As Long
Dim colcnt, totfldcnt As Long
Dim START As Single
Dim ENDKM As Single
Dim SQLQuery As String
Dim mastertable As master_table
On Error GoTo lblErr
Dim FARE As Single
Dim stgID As Long
Dim iStartLoc As Integer
    CCPB_Pbr.Value = 0
    Me.lbl.caption = ""
    If path_Txt = "" Then
        MsgBox "Please select the path", vbInformation, gblstrPrjTitle
        path_Txt.SetFocus
        Exit Sub
    End If
    Me.Enabled = False
    Me.lbl.caption = "Importing is in progress.Please wait..."
    Me.lbl.caption = "Importing is in progress.Please wait..."
    Import_cmd.Enabled = False
    If Dir(path_Txt) <> "" Then
        Set xlwbook = xl.Workbooks.Open(path_Txt)
        Set xlsheet = xl.ActiveSheet
        stgID = val(getvalueQuery("select max(id) from STAGE"))
        If val(stgID) > 0 Then stgID = stgID + 1 '' Asish consulted with Subodh
        rcnt = NOSTGS - 1
        colcnt = NOSTGS - 1
        totfldcnt = rcnt + 1
        iStartLoc = 1
        If xl.ActiveSheet.UsedRange.Cells(iStartLoc, 1).Value = "" Then iStartLoc = iStartLoc + 1
        If InStr(Replace(xl.ActiveSheet.UsedRange.Cells(iStartLoc, 1).Value, " ", ""), Replace(Rutecode_new, " ", "")) = 0 Then
            MsgBox "Route Mismatch!", vbExclamation, gblstrPrjTitle
            sql = "DELETE * FROM ROUTE WHERE rutcode='" & Rutecode_new & "'"
            gbladoCon.Execute (sql)
            Import_cmd.Enabled = True
            Unload Me 'subodh
            Exit Sub
        End If
        I = iStartLoc + 1
        j = 0
        For k = 0 To rcnt
            DoEvents
            sql = ""
            sql = "insert into STAGE (StageName,STG_LOCAL_LANGUAGE,Distance,route,id)values('" & xl.ActiveSheet.UsedRange.Cells(I, 1).Value & "','20-20-20-20',0,'" & Rutecode_new & "'," & stgID & ")"
            gbladoCon.Execute (sql)
            stgID = stgID + 1
            j = j + 1
            I = I + 1
            If Me.CCPB_Pbr.Value >= Me.CCPB_Pbr.Max - 1 Then
                Me.CCPB_Pbr.Value = 0
            Else
                Me.CCPB_Pbr.Value = Me.CCPB_Pbr.Value + 1
            End If
        Next k
        k = 1
        j = 1
        m = 0
        clnj = 1
        For j = 2 To colcnt + 2
            For I = 3 + m To totfldcnt + 1
                If xl.ActiveSheet.UsedRange.Cells(I, j).Value < val(min_fare) Then
                    MsgBox "fare is smaller than Minimum fare" & vbCrLf & "Please check the file!", vbExclamation, gblstrPrjTitle
                    sql = "DELETE * FROM FARE WHERE route='" & Rutecode_new & "'"
                    gbladoCon.Execute (sql)
                    sql = "DELETE * FROM STAGE WHERE route='" & Rutecode_new & "'"
                    gbladoCon.Execute (sql)
                    Me.CCPB_Pbr.Value = 0
                    xl.Quit
                    Set xl = Nothing
                    Set xlwbook = Nothing
                    Shell "taskkill /f /im ""EXCEL.exe"""
                    lbl.caption = ""
                    Me.Enabled = True
                    Import_cmd.Enabled = True
                    Exit Sub
                End If
                sql = ""
                sql = "insert into fare (row,COL,FARE,route)values( " & j - 1 & "," & k & ",'" & xl.ActiveSheet.UsedRange.Cells(I, j).Value & "','" & Rutecode_new & "')"
                gbladoCon.Execute (sql)
                Debug.Print sql
                k = k + 1
            Next I
            k = clnj + 1
            clnj = clnj + 1
            m = m + 1
            If Me.CCPB_Pbr.Value >= Me.CCPB_Pbr.Max - 1 Then
                Me.CCPB_Pbr.Value = 0
            Else
                Me.CCPB_Pbr.Value = Me.CCPB_Pbr.Value + 1
                DoEvents
            End If
        Next j
    End If
    If I > 0 And j > 0 Then
        MsgBox "Data imported successfully.", vbOKOnly, gblstrPrjTitle
        xl.ActiveWorkbook.Close False
        xl.Quit
        Set xl = Nothing
        Set xlwbook = Nothing
        Shell "taskkill /f /im ""EXCEL.exe"""
    End If
    Me.CCPB_Pbr.Value = Me.CCPB_Pbr.Max
    Me.CCPB_Pbr.Value = 0
    lbl.caption = ""
    Me.Enabled = True
    Import_cmd.Enabled = True
    Unload Me
Exit Sub
lblErr:
    Import_cmd.Enabled = True
    If err.Number = 429 Then
        MsgBox "To import data MS Office Excel should be installed.!", vbExclamation, gblstrPrjTitle
    Else
        MsgBox "Error due to " & err.Description & " ,Err No :" & err.Number, vbExclamation, gblstrPrjTitle
    End If
    MsgBox "Going to delete the route!", vbInformation, gblstrPrjTitle
    sql = "DELETE * FROM ROUTE WHERE rutcode='" & Rutecode_new & "'"
    gbladoCon.Execute (sql)
    sql = "DELETE * FROM FARE WHERE route='" & Rutecode_new & "'"
    gbladoCon.Execute (sql)
    sql = "DELETE * FROM STAGE WHERE route='" & Rutecode_new & "'"
    gbladoCon.Execute (sql)
    Me.CCPB_Pbr.Value = 0
    Shell "taskkill /f /im ""EXCEL.exe"""
    lbl.caption = ""
    Me.Enabled = True
End Sub
Public Function ConnectExcel(ExcelFilePath As String) As Boolean
On Error GoTo ErrorMod:
    If ExcelCon.State = 1 Then
        ExcelCon.Close
        ExcelCon.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & ExcelFilePath & ";DefaultDir=c:\mypath;"
        ExcelCon.Open
        ConnectExcel = True
    Else
        ExcelCon.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & ExcelFilePath & ";DefaultDir=c:\mypath;"
        ExcelCon.Open
        ConnectExcel = True
    End If
Exit Function
ErrorMod:
    ConnectExcel = False
    MsgBox "Error due to " & err.Description, vbExclamation, gblstrPrjTitle
End Function
