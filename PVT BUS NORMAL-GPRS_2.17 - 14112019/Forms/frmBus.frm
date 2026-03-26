VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmBus 
   BackColor       =   &H00E0E0E0&
   Caption         =   "BUS WISE REPORT"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frSelectSchedule 
      BackColor       =   &H00E0E0E0&
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   5715
      Begin VB.ComboBox cmbBusNo 
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
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   375
         Left            =   2850
         TabIndex        =   1
         Top             =   1920
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         TX              =   "E&xit"
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
         MICON           =   "frmBus.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdBusNoPalm 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1920
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         TX              =   "&Export"
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
         MICON           =   "frmBus.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   345
         Left            =   1200
         TabIndex        =   7
         Top             =   480
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
         Format          =   71630849
         CurrentDate     =   40939
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   330
         Left            =   3840
         TabIndex        =   9
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   71630849
         CurrentDate     =   40939
      End
      Begin VB.Label lbltodate 
         BackStyle       =   0  'Transparent
         Caption         =   "   End Date"
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
         Left            =   2640
         TabIndex        =   8
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label lblfromdate 
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
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1650
      End
      Begin VB.Label lblbusno 
         BackStyle       =   0  'Transparent
         Caption         =   "       Bus No "
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
         Left            =   1200
         TabIndex        =   4
         Top             =   1200
         Width           =   1770
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BUS WISE  REPORT"
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
      Left            =   165
      TabIndex        =   5
      Top             =   -495
      Width           =   5640
   End
End
Attribute VB_Name = "frmBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As ADODB.Connection
Dim san As New ADODB.Recordset 'SANGEETHA
Dim rsbustype As New ADODB.Recordset
Dim SysD As String
Private Sub cmdSchsmryRpt_Click()
On Error Resume Next
 ' Call SchsmryRptCnv
    
'   Cmd.Filter = "TXT (*.TXT)|*.TXT"
'    Cmd.ShowOpen
''    rtfText.Width = frmReport.Width - 300
''    rtfText.Height = frmReport.Height - 200
'    rchtxtbox.Font = "Lucida console"
'    rchtxtbox.Font.Size = "10"
'    rchtxtbox.Locked = True
'    rchtxtbox.LoadFile (Cmd.filename)
'    rchtxtfrm.Show vbModal
'    cmdPrint.Enabled = True
End Sub
Public Sub BussmryRptCnv()
On Error GoTo erromod
Dim FS As New FileSystemObject
Dim fShndl As Integer
Dim SysD, FnameUp, str, sql2 As String
Dim qry, sql, sSQL, subsql As String
Dim HDR1, HDR2 As String
Dim pamt As Integer
Dim Total As Double
Dim expsql, luggsql As String
Dim cond As String
Dim TcketPath As String
Dim fullcounttot As Single
Dim Stcounttot As Single
Dim Stcoltot As Single
Dim halfcounttot As Single
Dim Luggcounttot As Single
Dim Pascounttot As Single
Dim Phcounttot As Single
Dim fullcoltot As Single
Dim halfcoltot As Single
Dim phcoltot As Single
Dim Luggcoltot As Single
Dim Expcoltot As Single
Dim Totclntot As Single
Dim Netcoltot As Single
Dim Adjcoltot As Single
Dim Bus As String
Dim BusdATE As String
Dim flag1 As Boolean, sc_cont As Long, sc_col As Double
Dim check As Boolean, lad_cont As Long, lad_col As Double
Dim EXLRange As excel.Range
    pamt = 0
    pamt = 0
    TSQL = "SELECT * FROM PCSETUP"
    Set RES = CNN.OpenRecordset(TSQL, dbOpenDynaset)
    If RES.RecordCount > 0 Then
        TcketPath = RES!TICKET_PATH
        TransPath = RES!TRANSFER_PATH
    End If
    RES.Close
    SysD = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
    If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
    TcketPath = TcketPath & "\" & SysD
    If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
    FnameUp = "BUS WISE PALMTEC SUMMARY REPORT"
    If Dir(TcketPath & "\" & FnameUp & ".xlsx", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xlsx")
    If Dir(TcketPath & "\" & FnameUp & ".xls", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xls")
    Dim ExlObj As New excel.Application
    ExlObj.Workbooks.Add
    HDR1 = "BUS WISE PALMTEC SUMMARY REPORT"
    ExlObj.Range("H1:M1").MergeCells = True
    ExlObj.Range("H1:M1").Value = HDR1
    ExlObj.Range("H1:M1").Font.FontStyle = "Bold"
    ExlObj.Range("H1:M1").Font.Color = vbRed
    ExlObj.Range("H1:M1").Font.Size = 8
    sql = "SELECT RPT.StartDate, RPT.BusNo, RPT.Conductor, RPT.Driver, RPT.Cleaner, RPT.PalmID, Sum(RPT.Fulls) AS fullcnt, Sum(RPT.fullcoll) AS fullcln," _
        & "Sum(RPT.Luggage) AS Luggcnt, Sum(RPT.Half) AS halfcnt,Sum(RPT.Expense) AS Expamt,Sum(RPT.AdjustColl) AS Adj, Sum(RPT.Halfcoll) AS halfcln,Sum(RPT.St) AS Stcnt, Sum(RPT.Stcoll) AS Stcln, Sum(RPT.LuggageColl) AS Luggagecln," _
        & " Sum(RPT.Phy) AS phycnt, Sum(RPT.Phycoll) AS Phycl, Sum(Val(pass)) AS passcnt, Sum(RPT.TotalColl) AS totcol" _
        & ",Sum(RPT.ladies_count) AS lad_cout, Sum(RPT.ladies_coll) AS lad_cole,Sum(RPT.senior_count) AS sc_count, Sum(RPT.senior_coll) AS sc_col " _
        & " FROM rpt WHERE DATEVALUE(StartDate) BETWEEN DATEVALUE('" & DTfrom.Value & "') AND DATEVALUE('" & DTTo.Value & "') "
    If cmbBusNo.ListIndex <> -1 And cmbBusNo.Text <> "ALL" Then sql = sql & "AND rpt.BusNo='" & cmbBusNo & "'"
    sql = sql & " GROUP BY rpt.StartDate,rpt.BUSNO,rpt.PalmID,rpt.Conductor,rpt.Driver,rpt.Cleaner"
    Dim exclrow As Integer, exclcol As Integer
    exclrow = 2
    Dim gtotal As Double
    ExlObj.Range("A" & exclrow & ":" & "AA" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
    exclrow = exclrow + 1
    ExlObj.Range("A3:B3").MergeCells = True
    ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "DATE"
    ExlObj.Range("F3:G3").MergeCells = True
    ExlObj.Range("F3:G3").Value = DTfrom.Value
    exclrow = exclrow + 1
    ExlObj.Range("A4:B4").MergeCells = True
    ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "BUS NO"
    ExlObj.Range("F4:G4").MergeCells = True
    ExlObj.Range("F4:G4").Value = cmbBusNo.Text
    exclrow = exclrow + 1
    'ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_________________________________________________________________________________________________________________________________"
    ExlObj.Range("A" & exclrow & ":" & "AA" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
    exclrow = exclrow + 1
    
    ExlObj.Range("A6:A7").MergeCells = True
    ExlObj.Range("A6:A7").Value = "DATE "
    ExlObj.Range("B6:B7").MergeCells = True
    ExlObj.Range("B6:B7").Value = "BUS NO"
    ExlObj.Range("C6:C7").MergeCells = True
    ExlObj.Range("C6:C7").Value = "CONDUCTOR"
    ExlObj.Range("D6:D7").MergeCells = True
    ExlObj.Range("D6:D7").Value = "DRIVER"
    ExlObj.Range("E6:E7").MergeCells = True
    ExlObj.Range("E6:E7").Value = "CLEANER"
    ExlObj.Range("F6:F7").MergeCells = True
    ExlObj.Range("F6:F7").Value = "PALM ID"
    ExlObj.Range("G6:H6").MergeCells = True
    ExlObj.Range("G6:H6").Value = "FULL"
    ExlObj.Range("G6:H6").HorizontalAlignment = xlCenter
    ExlObj.Range("I6:J6").MergeCells = True
    ExlObj.Range("I6:J6").Value = "HALF"
    ExlObj.Range("I6:J6").HorizontalAlignment = xlCenter
    ExlObj.Range("K6:L6").MergeCells = True
    ExlObj.Range("K6:L6").Value = "LUGGAGE"
    ExlObj.Range("K6:L6").HorizontalAlignment = xlCenter
    ExlObj.Range("M6:N6").MergeCells = True
    ExlObj.Range("M6:N6").Value = "PH"
    ExlObj.Range("M6:N6").HorizontalAlignment = xlCenter
    
    ExlObj.Range("O6:P6").MergeCells = True
    ExlObj.Range("O6:P6").Value = "PASS"
    ExlObj.Range("O6:P6").HorizontalAlignment = xlCenter
    
    ExlObj.Range("Q6:R6").MergeCells = True
    ExlObj.Range("Q6:R6").Value = "ST"
    ExlObj.Range("Q6:R6").HorizontalAlignment = xlCenter
    
    ExlObj.Range("S6:T6").MergeCells = True
    ExlObj.Range("S6:T6").Value = "Ladies"
    ExlObj.Range("S6:T6").HorizontalAlignment = xlCenter
    
    ExlObj.Range("U6:V6").MergeCells = True
    ExlObj.Range("U6:V6").Value = "SC"
    ExlObj.Range("U6:V6").HorizontalAlignment = xlCenter
    
    ExlObj.Range("W6:W7").MergeCells = True
    ExlObj.Range("W6:W7").Value = "TOT COLL"
    ExlObj.Range("X6:X7").MergeCells = True
    ExlObj.Range("X6:X7").Value = "TOT ADJUST"
    ExlObj.Range("Y6:Y7").MergeCells = True
    ExlObj.Range("Y6:Y7").Value = "TOT OTHERCOLL"
    ExlObj.Range("Z6:Z7").MergeCells = True
    ExlObj.Range("Z6:Z7").Value = "TOT EXPENSE"
    ExlObj.Range("AA6:AA7").MergeCells = True
    ExlObj.Range("AA6:AA7").Value = "NET AMT"
    'ExlObj.Range("T4:T5").MergeCells = True
            
    ExlObj.Range("G7:G7").Value = "COUNT"
    ExlObj.Range("H7:H7").Value = "AMOUNT"
    ExlObj.Range("I7:I7").Value = "COUNT"
    ExlObj.Range("J7:J7").Value = "AMOUNT"
    ExlObj.Range("K7:K7").Value = "COUNT"
    ExlObj.Range("L7:L7").Value = "AMOUNT"
    ExlObj.Range("M7:M7").Value = "COUNT"
    ExlObj.Range("N7:N7").Value = "AMOUNT"
    ExlObj.Range("O7:O7").Value = "COUNT"
    ExlObj.Range("P7:P7").Value = "AMOUNT"
    ExlObj.Range("Q7:Q7").Value = "COUNT"
    ExlObj.Range("R7:R7").Value = "AMOUNT"
    
    ExlObj.Range("S7:S7").Value = "COUNT"
    ExlObj.Range("T7:T7").Value = "AMOUNT"
    ExlObj.Range("U7:U7").Value = "COUNT"
    ExlObj.Range("V7:V7").Value = "AMOUNT"
    Total = 0
    If DatabaseADOB_Connection() = False Then MsgBox "cannot connect db"
    If san.State = 1 Then san.Close
    san.Open sql, adoc, adOpenDynamic, adLockOptimistic
    If IsNull(san!Busno) Or san!Busno = "" Then
        MsgBox "Bus details not found!", vbExclamation, gblstrPrjTitle
        Shell "taskkill /f /im ""EXCEL.exe"""
        Exit Sub
    End If
    If Not san.EOF Then  'before
        Bus = san!Busno
        BusdATE = san!StartDate
        exclrow = 7
        exclrow = exclrow + 1
        ExlObj.ActiveSheet.Cells(exclrow, 2).Value = san!Busno  'SANGEETHA
        ExlObj.ActiveCell(exclrow, 1).Style.NumberFormat = "@"
        ExlObj.ActiveSheet.Cells(exclrow, 1).Value = CStr(san!StartDate)
        exclrow = 7
        check = False
        Do While Not san.EOF
            If san.EOF = True Then Exit Do
            exclrow = exclrow + 1
            If (Bus = san!Busno) Then
                flag1 = False
                Bus = san!Busno
                If (BusdATE <> san!StartDate) Then
                    ExlObj.ActiveSheet.Cells(exclrow, 1).Value = CStr(san!StartDate)
                    BusdATE = san!StartDate
                End If
                ExlObj.ActiveSheet.Cells(exclrow, 3).Value = san!Conductor
                ExlObj.ActiveSheet.Cells(exclrow, 4).Value = san!Driver
                ExlObj.ActiveSheet.Cells(exclrow, 5).Value = san!Cleaner
                ExlObj.ActiveSheet.Cells(exclrow, 6).Value = san!PalmID
                check = True
                ExlObj.ActiveSheet.Cells(exclrow, 7) = san!fullcnt
                fullcounttot = fullcounttot + san!fullcnt
                ExlObj.ActiveSheet.Cells(exclrow, 7).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 8) = Round(val(san!fullcln), 2)
                fullcoltot = fullcoltot + Round(val(san!fullcln), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 9) = san!halfcnt
                halfcounttot = halfcounttot + san!halfcnt
                ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 10) = Round(val(san!halfcln), 2)
                halfcoltot = halfcoltot + Round(val(san!halfcln), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 10).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 11) = san!luggcnt
                Luggcounttot = Luggcounttot + san!luggcnt
                ExlObj.ActiveSheet.Cells(exclrow, 11).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 12) = Round(val(san!Luggagecln), 2)
                Luggcoltot = Luggcoltot + Round(val(san!Luggagecln), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 12).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 13) = san!phycnt
                Phcounttot = Phcounttot + san!phycnt
                ExlObj.ActiveSheet.Cells(exclrow, 13).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 14) = Round(val(san!Phycl), 2)
                phcoltot = phcoltot + Round(val(san!Phycl), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 14).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 15) = san!passcnt
                Pascounttot = Pascounttot + san!passcnt
                ExlObj.ActiveSheet.Cells(exclrow, 15).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 16) = Round(0, 2)
                ExlObj.ActiveSheet.Cells(exclrow, 16).HorizontalAlignment = xlRight
         
                ExlObj.ActiveSheet.Cells(exclrow, 17) = san!stcnt
                Stcounttot = Stcounttot + san!stcnt
                ExlObj.ActiveSheet.Cells(exclrow, 17).HorizontalAlignment = xlCenter
                
                ExlObj.ActiveSheet.Cells(exclrow, 18) = Round(val(san!stcln), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 18).HorizontalAlignment = xlRight
                Stcoltot = Stcoltot + Round(val(san!stcln), 2)
                
                ExlObj.ActiveSheet.Cells(exclrow, 19) = san!lad_cout
                lad_cont = lad_cont + san!lad_cout
                ExlObj.ActiveSheet.Cells(exclrow, 19).HorizontalAlignment = xlCenter
                
                ExlObj.ActiveSheet.Cells(exclrow, 20) = Round(val(san!lad_cole), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 20).HorizontalAlignment = xlRight
                lad_col = lad_col + Round(val(san!lad_cole), 2)
                '''
                ExlObj.ActiveSheet.Cells(exclrow, 21) = san!sc_count
                sc_cont = sc_cont + san!sc_count
                ExlObj.ActiveSheet.Cells(exclrow, 21).HorizontalAlignment = xlCenter
                
                ExlObj.ActiveSheet.Cells(exclrow, 22) = Round(val(san!sc_col), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 22).HorizontalAlignment = xlRight
                sc_col = sc_col + Round(val(san!sc_col), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 23) = Round(val(san!totcol), 2)
                Totclntot = Totclntot + Round(val(san!totcol), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 23).HorizontalAlignment = xlRight
                
                ExlObj.ActiveSheet.Cells(exclrow, 24) = Round(val(san!Adj), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 24).HorizontalAlignment = xlRight
                
                ExlObj.ActiveSheet.Cells(exclrow, 25) = Round(0, 2)
                ExlObj.ActiveSheet.Cells(exclrow, 25).HorizontalAlignment = xlRight
                'ExlObj.Visible = True
                Adjcoltot = Adjcoltot + Round(val(san!Adj), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 26) = IIf(IsNull(san!ExpAmt), 0, (san!ExpAmt))
                Expcoltot = Expcoltot + IIf(IsNull(san!ExpAmt), 0, (san!ExpAmt))
                ExlObj.ActiveSheet.Cells(exclrow, 26).HorizontalAlignment = xlRight
                
                
                ExlObj.ActiveSheet.Cells(exclrow, 27) = (Round(val(san!totcol), 2) - IIf(IsNull(san!ExpAmt), 0, (san!ExpAmt)) - Round(val(san!Adj), 2))
                ExlObj.ActiveSheet.Cells(exclrow, 27).HorizontalAlignment = xlRight
                Netcoltot = Netcoltot + (Round(val(san!totcol), 2) - IIf(IsNull(san!ExpAmt), 0, (san!ExpAmt)) - Round(val(san!Adj), 2))
            Else
                Bus = san!Busno
                ExlObj.ActiveSheet.Cells(exclrow, 1) = "Total "
                ExlObj.ActiveSheet.Cells(exclrow, 7) = fullcounttot
                ExlObj.ActiveSheet.Cells(exclrow, 7).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 8) = fullcoltot
                ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 9) = halfcounttot
                ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 10) = Round(val(halfcoltot), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 10).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 11) = Luggcounttot
                ExlObj.ActiveSheet.Cells(exclrow, 11).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 12) = Round(val(Luggcoltot), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 12).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 13) = Phcounttot
                ExlObj.ActiveSheet.Cells(exclrow, 13).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 14) = Round(val(phcoltot), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 14).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 15) = Pascounttot
                ExlObj.ActiveSheet.Cells(exclrow, 15).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 16) = Round(0, 2)
                ExlObj.ActiveSheet.Cells(exclrow, 16).HorizontalAlignment = xlRight
                
                ExlObj.ActiveSheet.Cells(exclrow, 17) = Round(val(Stcounttot), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 17).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 18) = Round(val(Stcoltot), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 18).HorizontalAlignment = xlRight
                
                ExlObj.ActiveSheet.Cells(exclrow, 19) = Round(val(lad_cont), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 19).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 20) = Round(val(lad_col), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 20).HorizontalAlignment = xlRight
            
                ExlObj.ActiveSheet.Cells(exclrow, 21) = Round(val(sc_cont), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 21).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 22) = Round(val(sc_col), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 22).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 23) = Round(val(Totclntot), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 23).HorizontalAlignment = xlRight
                
                ExlObj.ActiveSheet.Cells(exclrow, 24) = Adjcoltot
                ExlObj.ActiveSheet.Cells(exclrow, 24).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 25) = Round(0, 2)
                ExlObj.ActiveSheet.Cells(exclrow, 25).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 26) = Round(val(Expcoltot), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 26).HorizontalAlignment = xlRight
                
                ExlObj.ActiveSheet.Cells(exclrow, 27) = Round(val(Netcoltot), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 27).HorizontalAlignment = xlRight
                BusdATE = ""
                fullcoltot = 0
                fullcounttot = 0
                halfcoltot = 0
                halfcounttot = 0
                phcoltot = 0
                Stcoltot = 0
                Stcounttot = 0
                Phcounttot = 0
                Luggcoltot = 0
                Luggcounttot = 0
                Expcoltot = 0
                Totclntot = 0
                Netcoltot = 0
                Pascounttot = 0
                Adjcoltot = 0
                flag1 = True
                ExlObj.Range("A" & exclrow & ":" & "AA" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
                exclrow = exclrow + 1
                If (BusdATE <> san!Date) Then
                    ExlObj.ActiveCell(exclrow, 1).Style.NumberFormat = "@"
                    ExlObj.ActiveSheet.Cells(exclrow, 1).Value = CStr(san!StartDate)
                    BusdATE = san!Date
                End If
                ExlObj.ActiveSheet.Cells(exclrow, 2).Value = san!Busno
                Bus = san!Busno
                ExlObj.ActiveSheet.Cells(exclrow, 3).Value = san!Conductor
                ExlObj.ActiveSheet.Cells(exclrow, 4).Value = san!Driver
                ExlObj.ActiveSheet.Cells(exclrow, 5).Value = san!Cleaner
                ExlObj.ActiveSheet.Cells(exclrow, 6).Value = san!PalmID
         
                ExlObj.ActiveSheet.Cells(exclrow, 7) = san!fullcnt
                fullcounttot = fullcounttot + san!fullcnt
                ExlObj.ActiveSheet.Cells(exclrow, 7).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 8) = Round(val(san!fullcln), 2)
                fullcoltot = fullcoltot + Round(val(san!fullcln), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 9) = san!halfcnt
                halfcounttot = halfcounttot + san!halfcnt
                ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 10) = Round(val(san!halfcln), 2)
                halfcoltot = halfcoltot + Round(val(san!halfcln), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 10).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 11) = san!luggcnt
                Luggcounttot = Luggcounttot + san!luggcnt
                ExlObj.ActiveSheet.Cells(exclrow, 11).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 12) = Round(val(san!Luggagecln), 2)
                Luggcoltot = Luggcoltot + Round(val(san!Luggagecln), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 12).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 13) = san!phycnt
                Phcounttot = Phcounttot + san!phycnt
                ExlObj.ActiveSheet.Cells(exclrow, 13).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 14) = Round(val(san!Phycl), 2)
                phcoltot = phcoltot + Round(val(san!Phycl), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 14).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 15) = san!passcnt
                Pascounttot = Pascounttot + san!passcnt
                ExlObj.ActiveSheet.Cells(exclrow, 15).HorizontalAlignment = xlCenter
                ExlObj.ActiveSheet.Cells(exclrow, 16) = Round(0, 2)
                ExlObj.ActiveSheet.Cells(exclrow, 16).HorizontalAlignment = xlRight
                
                ExlObj.ActiveSheet.Cells(exclrow, 17) = san!stcnt
                Stcounttot = Stcounttot + san!stcnt
                Stcoltot = Stcoltot + Round(val(san!stcln), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 17).HorizontalAlignment = xlCenter
                
                ExlObj.ActiveSheet.Cells(exclrow, 18) = Round(val(Stcoltot), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 18).HorizontalAlignment = xlRight
                
                
                ExlObj.ActiveSheet.Cells(exclrow, 19) = san!lad_cout
                lad_cont = lad_cont + san!lad_cout
                ExlObj.ActiveSheet.Cells(exclrow, 19).HorizontalAlignment = xlCenter
                
                ExlObj.ActiveSheet.Cells(exclrow, 20) = Round(val(san!lad_cole), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 20).HorizontalAlignment = xlRight
                lad_col = lad_col + Round(val(san!lad_cole), 2)
                '''
                ExlObj.ActiveSheet.Cells(exclrow, 21) = san!sc_count
                sc_cont = sc_cont + san!sc_count
                ExlObj.ActiveSheet.Cells(exclrow, 21).HorizontalAlignment = xlCenter
                
                ExlObj.ActiveSheet.Cells(exclrow, 22) = Round(val(san!sc_col), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 22).HorizontalAlignment = xlRight
                sc_col = sc_col + Round(val(san!sc_col), 2)
                
                
                
                ExlObj.ActiveSheet.Cells(exclrow, 23) = Round(val(san!totcol), 2)
                Totclntot = Totclntot + Round(val(san!totcol), 2)
         
                ExlObj.ActiveSheet.Cells(exclrow, 24) = Round(val(san!Adj), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 23).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 24).HorizontalAlignment = xlRight
                ExlObj.ActiveSheet.Cells(exclrow, 25) = Round(0, 2)
                ExlObj.ActiveSheet.Cells(exclrow, 25).HorizontalAlignment = xlRight
                Adjcoltot = Adjcoltot + Round(val(san!Adj), 2)
                ExlObj.ActiveSheet.Cells(exclrow, 26) = IIf(IsNull(san!ExpAmt), 0, (san!ExpAmt))
                Expcoltot = Expcoltot + IIf(IsNull(san!ExpAmt), 0, (san!ExpAmt))
                ExlObj.ActiveSheet.Cells(exclrow, 26).HorizontalAlignment = xlRight
                
                ExlObj.ActiveSheet.Cells(exclrow, 27) = (Round(val(san!totcol), 2) - IIf(IsNull(san!ExpAmt), 0, (san!ExpAmt)) - Round(val(san!Adj), 2))
                ExlObj.ActiveSheet.Cells(exclrow, 27).HorizontalAlignment = xlRight
                Netcoltot = Netcoltot + (Round(val(san!totcol), 2) - IIf(IsNull(san!ExpAmt), 0, (san!ExpAmt)) - Round(val(san!Adj), 2))
                flag1 = False
            End If
            san.MoveNext
        Loop
    End If
    If (flag1 = False) Then
        exclrow = exclrow + 1
        ExlObj.ActiveSheet.Cells(exclrow, 1) = "Total "
        ExlObj.ActiveSheet.Cells(exclrow, 7) = fullcounttot
        ExlObj.ActiveSheet.Cells(exclrow, 7).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 8) = fullcoltot
        ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlRight
        ExlObj.ActiveSheet.Cells(exclrow, 9) = halfcounttot
        ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 10) = Round(val(halfcoltot), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 10).HorizontalAlignment = xlRight
        ExlObj.ActiveSheet.Cells(exclrow, 11) = Luggcounttot
        ExlObj.ActiveSheet.Cells(exclrow, 11).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 12) = Round(val(Luggcoltot), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 12).HorizontalAlignment = xlRight
        ExlObj.ActiveSheet.Cells(exclrow, 13) = Phcounttot
        ExlObj.ActiveSheet.Cells(exclrow, 13).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 14) = Round(val(phcoltot), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 14).HorizontalAlignment = xlRight
        ExlObj.ActiveSheet.Cells(exclrow, 15) = Pascounttot
        ExlObj.ActiveSheet.Cells(exclrow, 15).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 16) = Round(0, 2)
        ExlObj.ActiveSheet.Cells(exclrow, 16).HorizontalAlignment = xlRight
        ExlObj.ActiveSheet.Cells(exclrow, 17) = Round(val(Stcounttot), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 17).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 18) = Round(val(Stcoltot), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 18).HorizontalAlignment = xlRight
        
        ExlObj.ActiveSheet.Cells(exclrow, 19) = Round(val(lad_cont), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 19).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 20) = Round(val(lad_col), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 20).HorizontalAlignment = xlRight
        
        ExlObj.ActiveSheet.Cells(exclrow, 21) = Round(val(sc_cont), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 21).HorizontalAlignment = xlCenter
        ExlObj.ActiveSheet.Cells(exclrow, 22) = Round(val(sc_col), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 22).HorizontalAlignment = xlRight
        
        ExlObj.ActiveSheet.Cells(exclrow, 23) = Round(val(Totclntot), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 23).HorizontalAlignment = xlRight
        ExlObj.ActiveSheet.Cells(exclrow, 24) = Adjcoltot
        ExlObj.ActiveSheet.Cells(exclrow, 24).HorizontalAlignment = xlRight
        ExlObj.ActiveSheet.Cells(exclrow, 25) = Round(0, 2)
        ExlObj.ActiveSheet.Cells(exclrow, 25).HorizontalAlignment = xlRight
        ExlObj.ActiveSheet.Cells(exclrow, 26) = Round(val(Expcoltot), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 26).HorizontalAlignment = xlRight

        ExlObj.ActiveSheet.Cells(exclrow, 27) = Round(val(Netcoltot), 2)
        ExlObj.ActiveSheet.Cells(exclrow, 27).HorizontalAlignment = xlRight
        fullcoltot = 0
        fullcounttot = 0
        halfcoltot = 0
        halfcounttot = 0
        phcoltot = 0
        Phcounttot = 0
        Luggcoltot = 0
        Luggcounttot = 0
        Expcoltot = 0
        Totclntot = 0
        Netcoltot = 0
        Pascounttot = 0
        Stcoltot = 0
        Stcounttot = 0
        Adjcoltot = 0
        ExlObj.Range("A" & exclrow & ":" & "AA" & exclrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
        exclrow = exclrow + 1
        flag1 = True
    End If
    Dim totalexp, refund As Double

    If check = True Then
        ExlObj.ActiveSheet.Name = FnameUp
        ExlObj.ActiveWorkbook.SaveAs TcketPath & "\" & FnameUp
        ExlObj.ActiveWorkbook.Close False
        MsgBox "Report Exported Successfully"
    Else
        Shell "taskkill /f /im ""EXCEL.exe"""
        MsgBox "No data for Export", vbOKOnly, gblstrPrjTitle
    End If
Exit Sub
erromod:
If err.Number = 429 Then
    MsgBox "To export data MS Office Excel should be installed.!", vbExclamation, gblstrPrjTitle
ElseIf err.Number = 91 Then
    MsgBox "To export data MS Office Excel Properties not supported.!" & vbCrLf & "Please check whether  MS office licence expired.", vbExclamation, gblstrPrjTitle
ElseIf InStr(1, err.Description, "cannot find the file specified") > 0 Then
    MsgBox "To export data Please install MS Office properly.!", vbExclamation, gblstrPrjTitle
Else
    MsgBox "Error occured Due to " & err.Description, vbExclamation, gblstrPrjTitle
End If
End Sub

Function ConnectNewDatabase(ByRef ConnectionObject As ADODB.Connection, _
                         ByVal Database As String, _
                         Optional ByVal DatabasePassword As String) As Boolean
On Error Resume Next
    If ConnectionObject.State = adStateOpen Then ConnectionObject.Close
        ConnectionObject.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database & ";Persist Security Info=False;Jet OLEDB:Database Password=" & DatabasePassword
    'ConnectionObject.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database & ";Persist Security Info=False;Jet OLEDB:Database Password=" & DatabasePassword
    ConnectionObject.Open
    If ConnectionObject.State = adStateOpen Then ConnectNewDatabase = True
    Exit Function
CatchError:
End Function


Public Function SchsmryRpt() As Boolean

''Public Function CovertColln(Fint As String) As Boolean
' On Error GoTo errLn
     Dim FS As New FileSystemObject
     Dim TcketPath As String
'    Dim FHndl As Integer
     Dim fShndl As Integer
'    Dim tktK As PTicket
     Dim fBuff As String
     Dim FnameUp As String
'    Dim PFname As String
'    Dim pHandle As Integer
'    Dim gPass As PASSCONC
'    '''gPassCount added by syam
'    Dim gPassCount As Long
'    Dim iFull As Integer
'    Dim iHalf As Integer
'    Dim iPhy As Integer
'    Dim iLugg As Integer
'    Dim iSt As Integer
'    Dim lTotPassenger As Long
'    Dim fTotAmount As Single
'    Dim fTotLuggAmount As Single
'    Dim strYear As String
'    Dim SysD, SysT, PID As String
'    gPassCount = 0
        TSQL = "SELECT * FROM PCSETUP"
        Set RES = CNN.OpenRecordset(TSQL, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            TcketPath = RES!TICKET_PATH
            TransPath = RES!TRANSFER_PATH
        End If
        RES.Close
'
'        sql = "SELECT PALMTECID FROM SETTINGS"
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        If RES.RecordCount > 0 Then
'            PalmId = RES!PalmtecID
'        End If
'        RES.Close
'
         sql = "SELECT * FROM RPT"

'        FnameUp = "TKTS" & Fint & ".DAT"
'        'If Dir(TransPath & "\" & FnameUp) <> "" Then
         SysD = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
'        'SysT = Replace(Time, ":", ".")
'        SysT = Format(Time, "hhmmAM/PM")
'        PID = Replace(PalmId, Chr(0), "")
        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
        TcketPath = TcketPath & "\" & SysD
            If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
            FnameUp = "SCHEDULESUMMARYRPT.TXT"
            '& Fint & "-" & SysT & PID
            If Dir(TcketPath & "\" & FnameUp) <> "" Then Kill TcketPath & "\" & FnameUp
            fShndl = FreeFile()
            Open TcketPath & "\" & FnameUp For Binary Access Write As #fShndl
                fBuff = ""
                fBuff = String(84, "_") & vbCrLf
                Put #fShndl, , fBuff
                fBuff = Format("DATE|", "@@@@@@@@@@@@") & vbCrLf
                fBuff = fBuff & Format("PALMID| ", "@@@@")
                fBuff = fBuff & Format("SCHEDULENO| ", "@@@@") & vbCrLf
'                fBuff = fBuff & Format("LG| ", "@@@@")
'                fBuff = fBuff & Format("PH| ", "@@@@")
'                fBuff = fBuff & Format("ST| ", "@@@@")
'                fBuff = fBuff & Format("PASS No|", "@@@@@@@@@")
'                'fBuff = fBuff & Format("FROM| ", "@@@@@")
'                'fBuff = fBuff & Format("TO| ", "@@@@@")
'                fBuff = fBuff & Format("AMOUNT|", "@@@@@@@")
'                fBuff = fBuff & Format("LUG_AMT| ", "@@@@@@@@")
'                'fBuff = fBuff & Format("TIME|", "@@@@@@")
'                'fBuff = fBuff & Format("DATE|", "@@@@@@@@@@@@") & vbCrLf
                 Put #fShndl, , fBuff
                 fBuff = String(84, "_") & "|" & vbCrLf
                 Put #fShndl, , fBuff
'
'
'        FnameUp = "TKTS" & Fint & ".DAT"
'        PFname = App.Path & "\PASS.PAS"
'        FHndl = FreeFile()
'            Open TransPath & "\" & FnameUp For Binary Access Read As #FHndl
'               Do While Not EOF(FHndl)
'                Dim str As String
'                Get #FHndl, , tktK
'                If tktK.TicketNo = -1 Then Exit Do
'                If EOF(FHndl) = True Then Exit Do
'                 With tktK
'                    fBuff = Format(.TicketNo, " 00000000") & "| "
'                    fBuff = fBuff & Format(.Full, "00") & "| "
'                    fBuff = fBuff & Format(.Half, "00") & "| "
'                    fBuff = fBuff & Format(.Lugg, "00") & "| "
'                    fBuff = fBuff & Format(.Phy, "00") & "| "
'                    fBuff = fBuff & Format(.st, "00") & "| "
'
'                    iFull = iFull + .Full
'                    iHalf = iHalf + .Half
'                    iLugg = iLugg + .Lugg
'                    iPhy = iPhy + .Phy
'                    iSt = iSt + .st
'
'                    If .Typ = 32 Then
'                        If Dir(PFname) <> "" Then
'                            pHandle = FreeFile()
'                            Open PFname For Binary Access Read As pHandle
'                            Do While Not EOF(pHandle)
'                                Get #pHandle, , gPass
'                                If .TicketNo = gPass.TicketNo Then Exit Do
'                            Loop
'                            Close #pHandle
'                            str = TrimChr(gPass.PassNo)
'                            'MsgBox str
'                        Else
'                            str = "  "
'                        End If

'                        If str <> "  " Then
'                        gPassCount = gPassCount + 1
'                        End If
'
'
'
'                        ''''''''''''''
'
'                        fBuff = fBuff & Format(str & "|", "@@@@@@@@@")
'                    Else
'                        fBuff = fBuff & String(7, " ") & "-|"
'                    End If
'                    fBuff = fBuff & Format(.From, " 000") & "|"
'                    fBuff = fBuff & Format(.To, " 000") & "|"
'                    str = Format(.Amount, "0.00")
'                    fBuff = fBuff & Format(str, "@@@@@@@") & "|"
'                    str = Format(.Luggage, "0.00")
'                    fBuff = fBuff & Format(str, "@@@@@@@") & "|"
'                    fBuff = fBuff & " " & Format(.Hr & ":" & .Minut, "HH:MM") & "|"
'                    strYear = ""  '05/01/2010
'                    strYear = DatePart("YYYY", Date) '05/01/2010
'                    fBuff = fBuff & " " & Format(.Dy & "/" & .Mn & "/" & strYear, "DD/MM/YYYY") & "|" & vbCrLf  '05/01/2010
'                    Put #fShndl, , fBuff
'
'
'                    fTotAmount = fTotAmount + .Amount
'                    fTotLuggAmount = fTotLuggAmount + .Luggage
'                    fBuff = ""
'                 End With
'               Loop
'                fBuff = String(84, "_") & "|" & vbCrLf '05/01/2010
'                Put #fShndl, , fBuff
'                lTotPassenger = iFull + iHalf + iPhy + iSt + gPassCount
'
'                fBuff = "TOTAL FULL       |" & Format(iFull, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                fBuff = "TOTAL HALF       |" & Format(iHalf, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                fBuff = "TOTAL PHY        |" & Format(iPhy, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                fBuff = "TOTAL ST         |" & Format(iSt, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                fBuff = "TOTAL LUGGAGE    |" & Format(iLugg, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'
'                 ''''''''''''''''''''''syam
'
'
'               ' fBuff = "TOTAL PASS" & gPassCount
'                '  Put #fShndl, , fBuff
'
'                fBuff = "TOTAL PASS       |" & Format(gPassCount, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'               ''''''''''''''''''''''''''
'
'                fBuff = "TOTAL PASSENGER  |" & Format(lTotPassenger, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'                str = Format(fTotLuggAmount, "0.00")
'                fBuff = "TOTAL LUGGAGE    |" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'                str = Format(fTotAmount, "0.00")
'                fBuff = "TOTAL AMOUNT     |" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'
'              DBTKTS (FnameUp)
'
'
'            Close #FHndl
'         Close #fShndl
'         CovertColln = True
'        Exit Function
'      'End If
'errLn:
'  CovertColln = False
    
End Function
'End Function

Private Sub cmdBusNoPalm_Click()
     Call BussmryRptCnv
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Load()
On Error Resume Next
   ' CONNECTDB
    
     Me.Icon = frmMainform.Icon
    sql = "SELECT DISTINCT BusNo FROM RPT"
    Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
   
        
        DTfrom.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
        DTTo.Value = DateValue(Format(Now, "DD/MM/YYYY"))
       
        
        If RES.RecordCount > 0 Then RES.MoveFirst
        cmbBusNo.Clear
        cmbBusNo.AddItem "ALL"
        Do While Not RES.EOF
        If RES!Busno <> Chr(0) Then
            cmbBusNo.AddItem IIf(IsNull(RES!Busno), "NILL", RES!Busno)
        End If
            RES.MoveNext
        Loop
        If cmbBusNo.ListCount > 0 Then
            cmbBusNo.Text = cmbBusNo.List(0)
        End If

    RES.Close
End Sub

Private Sub lblPalmId_Click()

End Sub

''Private Sub optReportType_Click(Index As Integer)
''On Error Resume Next
''
''    If optReportType(0).Value = True Then
''
''        DTFrom.Value = DateValue(Format(DateAdd("d", -1, Now), "DD/MM/YYYY"))
''        DTTo.Value = DateValue(Format(Now, "DD/MM/YYYY"))
''        'lblSdateOrID.Caption = " Palmtec ID    :"
''        'lblEndDateOrSch.Caption = "Schedule No :"
''        'DTEnd.Visible = False
''        'DTStart.Visible = False
''        'cmbPalmID.Visible = True
''        'cmbShedule.Visible = True
''        'DTFrom.Top = 960
''        'DTFrom.Left = 2000
''        'DTTo.Top = 960
''        'DTTo.Left = 4080
''        'cmbPalmID.Top = 1560
''        'cmbPalmID.Left = 2000
''        'cmbShedule.Top = 2040
''        'cmbShedule.Left = 2000
''       'cmbrouteno.Top = 2520
''        'cmbrouteno.Left = 2000
''
''        'lblSDate.Visible = False
''        'lblEDate.Visible = False
''        lblTripno.Visible = False
''        cmbtripno.Visible = False
''        cmbPalmID.Clear
''
''        sql = "SELECT DISTINCT PALMID FROM RPT"
''        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
''        If RES.RecordCount > 0 Then RES.MoveFirst
''        Do While Not RES.EOF
''            cmbPalmID.AddItem RES!PalmId
''            RES.MoveNext
''        Loop
''        If cmbPalmID.ListCount > 0 Then
''            cmbPalmID.Text = cmbPalmID.List(0)
''        End If
''
''     '   RES.Close
''
''        sql = "SELECT DISTINCT SCHEDULE FROM RPT" ' WHERE PALMID='" & cmbPalmID.Text & "'"
''        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
''        If RES.RecordCount > 0 Then RES.MoveFirst
''        cmbSchedule.Clear
''        Do While Not RES.EOF
''            cmbSchedule.AddItem RES!SCHEDULE
''            RES.MoveNext
''        Loop
''        If cmbSchedule.ListCount > 0 Then
''            cmbSchedule.Text = cmbSchedule.List(0)
''        End If
''        RES.Close
''        'sql = "SELECT DISTINCT TRIPNO FROM RPT WHERE PALMID='" & cmbPalmID.Text & "'"
''        'Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
''        'If RES.RecordCount > 0 Then RES.MoveFirst
''        'cmbtripno.Clear
''        'Do While Not RES.EOF
''         '   cmbtripno.AddItem RES!TripNo
''          '  RES.MoveNext
''         'Loop
''        'If cmbtripno.ListCount > 0 Then
''         '   cmbtripno.Text = cmbtripno.List(0)
''        'End If
''        sql = "SELECT DISTINCT RouteCode  FROM RPT " 'WHERE PALMID='" & cmbPalmID.Text & "'"
''        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
''        If RES.RecordCount > 0 Then RES.MoveFirst
''        cmbrouteno.Clear
''        Do While Not RES.EOF
''
''            cmbrouteno.AddItem RES!RouteCode
''            RES.MoveNext
''        Loop
''        If cmbrouteno.ListCount > 0 Then
''            cmbrouteno.Text = cmbrouteno.List(0)
''        End If
''
''
''    Else
''       'lblfromdate.Visible = False
''       'DTFrom.Visible = False
''        'lbltodate.Visible = False
''        'DTTo.Visible = False
''        'lblSdateOrID.Caption = "Start Date  :"
''        'lblEndDateOrSch.Caption = " End Date :"
''        'lblRouteNo.Caption = "   Trip No  :"
''        'DTEnd.Visible = True
''        'DTStart.Visible = True
''        'cmbPalmID.Visible = False
''        'cmbShedule.Visible = False
''        'cmbrouteno.Visible = False
''        'cmbtripno.Visible = True
''        'lblSdateOrID.Top = 960
''        'lblSdateOrID.Left = 500
''        'DTStart.Top = 960
''        'DTStart.Left = 2000
''        'lblEndDateOrSch.Top = 1560
''        'lblEndDateOrSch.Left = 500
''        'DTEnd.Top = 1560
''        'DTEnd.Left = 2000
''        'DTStart.Width = 960
''        'DTEnd.Width = 960
''        'lblRouteNo.Top = 2100
''        'lblRouteNo.Left = 500
''        'cmbtripno.Top = 2100                        '''rnc
''        'cmbtripno.Left = 2000
''        'lblSDate.Visible = True
''        'DTStart.Visible = True
''        'lblEDate.Visible = True
''        'DTEnd.Visible = True
''        lblTripno.Visible = True
''        cmbtripno.Visible = True
''        'DTStart.Day = Day(Now)
''        'DTStart.Month = Month(Now)
''        'DTStart.Year = Year(Now)
''        'DTEnd.Day = Day(Now)
''        'DTEnd.Month = Month(Now)
''        'DTEnd.Year = Year(Now)
''
''        sql = "SELECT DISTINCT TripNo  FROM RPT" 'WHERE PALMID='" & cmbPalmID.Text & "'"
''        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
''        If RES.RecordCount > 0 Then RES.MoveFirst
''        cmbtripno.Clear
''        cmbtripno.AddItem "ALL"
''        Do While Not RES.EOF
''            cmbtripno.AddItem RES!TripNo
''            RES.MoveNext
''        Loop
''        If cmbtripno.ListCount > 0 Then
''            cmbtripno.Text = cmbtripno.List(0)
''        End If
''       ' End If
''
''    End If
''    If optReportType(0).Value = True Then
''
''        'lblSDate.Visible = False
''        'lblEDate.Visible = False
''        lblTripno.Visible = False
''        cmbtripno.Visible = False
''    End If
''    RES.Close
''End Sub
'Public Sub SchsmryRptCnv()
'On Error Resume Next
'Dim FS As New FileSystemObject
'Dim fShndl As Integer
'Dim SysD, FnameUp As String
'Dim qry, sql, sSQL, subsql As String
'Dim HDR1, HDR2 As String
'Dim pamt As Integer
'Dim total As Double
'Dim expsql As String
'Dim cond As String
'Dim TcketPath As String
''open file
'
'pamt = 0
'pamt = 0
'TSQL = "SELECT * FROM PCSETUP"
'        Set RES = CNN.OpenRecordset(TSQL, dbOpenDynaset)
'        If RES.RecordCount > 0 Then
'            TcketPath = RES!TICKET_PATH
'            TransPath = RES!TRANSFER_PATH
'        End If
'        RES.Close
'        SysD = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
'        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
'        TcketPath = TcketPath & "\" & SysD
'        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
''        If optReportType(0).Value = True Then
''            FnameUp = "SCHEDULE SUMMARY REPORT"
''        Else
''            FnameUp = "TRIP WISE SCHEDULE DETAILS"
''        End If
'        '& Fint & "-" & SysT & PID
'        'If Dir(TcketPath & "\" & FnameUp) <> "" Then Kill TcketPath & "\" & FnameUp
''''        fShndl = FreeFile()
'    If Dir(TcketPath & "\" & FnameUp & ".xlsx", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xlsx")
'    If Dir(TcketPath & "\" & FnameUp & ".xls", vbNormal) <> "" Then Call Kill(TcketPath & "\" & FnameUp & ".xls")
'        Dim ExlObj As New excel.Application
'        ExlObj.Workbooks.Add
'        'ExlObj.Visible = True
'
'
'    'print HEADER FROM SETTINGS
'        qry = "SELECT HEADER1,HEADER2 FROM SETTINGS"
'        Set RES = CNN.OpenRecordset(qry, dbOpenDynaset)
'        If RES.RecordCount > 0 Then
'            HDR1 = RES!HEADER1
'            HDR2 = RES!HEADER2
'        End If
'        RES.Close
'
'       ExlObj.ActiveSheet.Cells(2, 5).Value = HDR1
'       ExlObj.ActiveSheet.Cells(2, 5).Font.Bold = True
'       ExlObj.ActiveSheet.Cells(3, 5).Value = HDR2
'       ExlObj.ActiveSheet.Cells(3, 5).Font.Bold = True
'       If optReportType(0).Value = True Then
'            ExlObj.ActiveSheet.Cells(4, 5).Value = "Schedule wise Summary Report"
'       Else
'            ExlObj.ActiveSheet.Cells(4, 5).Value = "Trip wise Summary Report"
'       End If
'       ExlObj.ActiveSheet.Cells(4, 5).Font.Bold = True
'       ExlObj.ActiveSheet.Range("C5:G5").Value = "___________________"
'
'        sql = "SELECT RT.* FROM RPT RT WHERE DATE BETWEEN DATEVALUE('" & DTFrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') "
'
'        If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
'            sql = sql & " AND PalmID='" & cmbPalmID & "'"
'            cond = cond & " AND rt.PalmID='" & cmbPalmID & "'"
'        End If
'
'        If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
'            sql = sql & " AND SCHEDULE=" & val(cmbSchedule) & ""
'            cond = cond & " AND rt.SCHEDULE=" & val(cmbSchedule) & ""
'        End If
'
'        If cmbrouteno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
'            sql = sql & " AND ROUTECODE='" & cmbrouteno & "'"
'            cond = cond & " AND rt.ROUTECODE='" & cmbrouteno & "'"
'        End If
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If cmbtripno.ListIndex <> -1 And cmbtripno.Text <> "ALL" Then
'            sql = sql & " AND TRIPNO=" & val(cmbtripno) & " "
'            cond = cond & " AND rt.TRIPNO=" & val(cmbtripno) & ""
'        End If
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If optReportType(0).Value = True Then
'            sql = sql & " ORDER BY RT.DATE,RT.PALMID,RT.SCHEDULE,RT.STicket"
'        Else
'            sql = sql & " ORDER BY RT.DATE,RT.PALMID,RT.SCHEDULE, RT.TRIPNO,RT.STicket"
'        End If
'        Set RES = CNN.OpenRecordset(sql, dbOpenDynaset)
'        Dim exclrow As Integer, exclcol As Integer
'exclrow = 6
'Dim gtotal As Double
'If Not RES.EOF Then
'
'    While Not RES.EOF
'    total = 0
'          ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
'          exclrow = exclrow + 1
'          ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Date       "
'          ExlObj.ActiveSheet.Cells(exclrow, 2).Value = " " & RES!Date
'          exclrow = exclrow + 1
'          ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "PalmId      "
'          ExlObj.ActiveSheet.Cells(exclrow, 2).Value = RES!PalmId
'          ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Schedule No  "
'          ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!SCHEDULE
'          ExlObj.ActiveCell(exclrow, 9).Style.NumberFormat = "@"
'          exclrow = exclrow + 1
'          If optReportType(0).Value = True Then
'                sSQL = "select count(tripno) as 0 from rpt rt WHERE DATE BETWEEN DATEVALUE('" & DTFrom.Value & " ')AND DATEVALUE('" & DTTo.Value & "') and SCHEDULE=" & RES!SCHEDULE
'                If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
'                    sSQL = sSQL & " AND PalmID='" & cmbPalmID & "'"
'                End If
'                If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
'                sSQL = sSQL & " AND SCHEDULE=" & val(cmbSchedule) & " "
'                End If
'                If cmbrouteno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
'                sSQL = sSQL & " AND ROUTECODE='" & cmbrouteno & "'"
'                End If
'                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                If cmbtripno.ListIndex <> -1 And cmbtripno.Text <> "ALL" Then
'                sSQL = sSQL & " AND TRIPNO=" & val(cmbtripno) & ""
'                End If
'          End If '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'          If optReportType(0).Value = True Then
'                sSQL = sSQL
'          Else
'                sSQL = sSQL
'          End If
'         Set res1 = CNN.OpenRecordset(sSQL, dbOpenDynaset)
'         If optReportType(0).Value = True Then
'            ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "No Of Trips   "
'            ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "  " & res1!CTRIP
'         End If
'         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Route No      "
'         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!RouteCode
'         exclrow = exclrow + 1
'         ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Conductor   "
'         ExlObj.ActiveSheet.Cells(exclrow, 2).Value = RES!Conductor
'         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Driver      "
'         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!Driver
'         exclrow = exclrow + 1
'         If optReportType(1).Value = True Then
'             Dim sroutename As String, sroutesql As String
'             sroutesql = "select rutname from ROUTE where rutcode='" & RES!RouteCode & "'"
'             Set RESROUTE = CNN.OpenRecordset(sroutesql, dbOpenDynaset)
'             ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Trip Number & Flag      "
'             ExlObj.ActiveSheet.Cells(exclrow, 2).Value = RES!TripNo & "   " & IIf(RES!UpDownTrip = "U", "UP", "DOWN")
'             ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "Route No. & Name            "
'             ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!RouteCode & "   " & RESROUTE!rutname
'             Set RESROUTE = Nothing
'             exclrow = exclrow + 1
'         End If
'         ExlObj.ActiveSheet.Cells(exclrow, 1).Value = "Start TktNo    "
'         ExlObj.ActiveSheet.Cells(exclrow, 2).Value = RES!STicket
'         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "End Tkt No     "
'         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = RES!ETicketNo
'         exclrow = exclrow + 1
'         ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
'         exclrow = exclrow + 1
'         ExlObj.ActiveSheet.Cells(exclrow, 2).Value = "      " & "Count"
'         ExlObj.ActiveSheet.Cells(exclrow, 3).Value = "      " & "Amount"
'         ExlObj.ActiveSheet.Cells(exclrow, 8).Value = "      " & "Count"
'         ExlObj.ActiveSheet.Cells(exclrow, 9).Value = "      " & "Amount"
'         exclrow = exclrow + 1
'         ExlObj.ActiveSheet.Range("B" & exclrow & ":" & "C" & exclrow).Value = "_______________"
'         ExlObj.ActiveSheet.Range("H" & exclrow & ":" & "I" & exclrow).Value = "_______________"
'         exclrow = exclrow + 1
'         ExlObj.ActiveSheet.Cells(exclrow, 1) = " Full       "
'         ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!Full
'         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
'         ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!FullColl), 2)
'         'ExlObj.ActiveSheet.Cells(exclrow, 3).Align = 6
'         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
'         ExlObj.ActiveSheet.Cells(exclrow, 7) = " Half       "
'         ExlObj.ActiveSheet.Cells(exclrow, 8) = RES!Half
'         ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlCenter
'         ExlObj.ActiveSheet.Cells(exclrow, 9) = Round(val(RES!HalfColl), 2)
'         ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlRight
'         exclrow = exclrow + 1
'         ExlObj.ActiveSheet.Cells(exclrow, 1) = " ST       "
'         ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!st
'         ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
'         ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!STColl), 2)
'         ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
'         ExlObj.ActiveSheet.Cells(exclrow, 7) = " PH       "
'         ExlObj.ActiveSheet.Cells(exclrow, 8) = RES!Phy
'         ExlObj.ActiveSheet.Cells(exclrow, 8).HorizontalAlignment = xlCenter
'         ExlObj.ActiveSheet.Cells(exclrow, 9) = Round(val(RES!PhyColl), 2)
'         ExlObj.ActiveSheet.Cells(exclrow, 9).HorizontalAlignment = xlRight
'         exclrow = exclrow + 1
'
'
''''         subsql = "select count(Pass) as cpass from rpt rt WHERE Pass <> '' "
''''         If cmbPalmID.ListIndex <> -1 And cmbPalmID.Text <> "ALL" Then
''''                subsql = subsql & " AND PalmID='" & cmbPalmID & "'"
''''            End If
''''            If cmbSchedule.ListIndex <> -1 And cmbSchedule.Text <> "ALL" Then
''''                subsql = subsql & " AND SCHEDULE= " & val(cmbSchedule) & ""
''''            End If
''''            If cmbrouteno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
''''                subsql = subsql & " AND ROUTECODE= '" & cmbrouteno & "'"
''''            End If
''''            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''            If cmbtripno.ListIndex <> -1 And cmbrouteno.Text <> "ALL" Then
''''                subsql = sSQL & " AND TRIPNO=" & val(cmbtripno) & ""
''''        End If
''''
'''''        If optReportType(0).Value = True Then
'''''                subsql = sSQL
'''''          Else
'''''                subsql = sSQL
'''''          End If
''''        Set RES2 = CNN.OpenRecordset(subsql, dbOpenDynaset)
'        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Pass       "
'        ExlObj.ActiveSheet.Cells(exclrow, 2) = RES!pass
'        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlCenter
'        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(pamt), 2)
'        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
'        exclrow = exclrow + 1
'        'new
'
'        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Luggage       "
'        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!LuggageColl), 2)
'        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
'        exclrow = exclrow + 1
'
'
'        ''''''''''''''''''''''''
'
'
'        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Adjust       "
'        ExlObj.ActiveSheet.Cells(exclrow, 3) = Round(val(RES!AdjustColl), 2)
'        ExlObj.ActiveSheet.Cells(exclrow, 3).HorizontalAlignment = xlRight
'        exclrow = exclrow + 1
'        ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
'        exclrow = exclrow + 1
'        total = total + RES!FullColl + RES!HalfColl + RES!STColl + RES!PhyColl + RES!LuggageColl - RES!AdjustColl
'        ExlObj.ActiveSheet.Cells(exclrow, 1) = " Total       "
'        ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
'        ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(total), 2)
'        ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Bold = True
'        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
'        exclrow = exclrow + 1
'        gtotal = gtotal + total
'        RES.MoveNext
'    Wend
'End If
'
'
'ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
'exclrow = exclrow + 1
'Dim totalexp As Double
'totalexp = 0
''expsql = "SELECT EXP_NAME, sum(ExpAmt) as EXPAmount FROM expmaster AS m, expense AS e ,rpt rt Where m.exp_code = val(e.ExpCode) and chr(rt.Trip_Master_ID) =chr(e.TripMasterReferenceId)  and e.DATE BETWEEN DATEVALUE('" & DTFrom.Value & "')AND DATEVALUE('" & DTTo.Value & "') " & cond & " group by EXP_NAME"
'expsql = " SELECT ExpName, sum(ExpAmt) AS EXPAmount FROM expense AS e, rpt AS rt WHERE chr(rt.Trip_Master_ID) =chr(e.TripMasterReferenceId)  and e.DATE BETWEEN DATEVALUE('" & DTFrom.Value & "')AND DATEVALUE('" & DTTo.Value & "') " & cond & "  GROUP BY ExpName"
'Set RES2 = CNN.OpenRecordset(expsql, dbOpenDynaset)
'If Not RES2.EOF Then
'    While Not RES2.EOF
'        ExlObj.ActiveSheet.Cells(exclrow, 1) = RES2!expname
'        ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(RES2!EXPAmount), 2)
'        ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
'        totalexp = totalexp + RES2!EXPAmount
'        RES2.MoveNext
'        exclrow = exclrow + 1
'    Wend
'End If
'ExlObj.ActiveSheet.Range("A" & exclrow & ":" & "H" & exclrow).Value = "_______________________________"
'exclrow = exclrow + 1
'ExlObj.ActiveSheet.Cells(exclrow, 1) = "Total Expense  "
'ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
'ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(totalexp), 2)
'ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Bold = True
'ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
'exclrow = exclrow + 1
'ExlObj.ActiveSheet.Cells(exclrow, 1) = "Amount To Remit  "
'ExlObj.ActiveSheet.Cells(exclrow, 1).Font.Bold = True
'ExlObj.ActiveSheet.Cells(exclrow, 2) = Round(val(gtotal - totalexp), 2)
'ExlObj.ActiveSheet.Cells(exclrow, 2).Font.Bold = True
'ExlObj.ActiveSheet.Cells(exclrow, 2).HorizontalAlignment = xlRight
'With Worksheets("Sheet1").Columns("A")
'    .ColumnWidth = (.ColumnWidth * 1.75) + 5
'End With
'With Worksheets("Sheet1").Columns("B")
'    .ColumnWidth = .ColumnWidth * 1.75
'End With
'With Worksheets("Sheet1").Columns("C")
'    .ColumnWidth = .ColumnWidth * 1.75
'End With
'With Worksheets("Sheet1").Columns("H")
'    .ColumnWidth = .ColumnWidth * 1.75
'End With
'With Worksheets("Sheet1").Columns("I")
'    .ColumnWidth = .ColumnWidth * 1.75
'End With
'ExlObj.ActiveSheet.Name = FnameUp
'ExlObj.ActiveWorkbook.SaveAs TcketPath & "\" & FnameUp
'ExlObj.ActiveSheet.Close
'ExlObj.Workbooks.Close
''ExlObj.ActiveWorkbook.Close False
'MsgBox "Report Exported Successfully"
'Exit Sub
'End Sub
'
Private Sub frSelectSchedule_DragDrop(Source As Control, x As Single, y As Single)

End Sub


