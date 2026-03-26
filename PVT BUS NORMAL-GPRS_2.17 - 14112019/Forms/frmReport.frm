VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6600
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   12555
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   5190
         Left            =   270
         TabIndex        =   2
         Top             =   900
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   9155
         _Version        =   393217
         BackColor       =   -2147483628
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmReport.frx":57E2
      End
      Begin JeweledBut.JeweledButton Command1 
         Height          =   375
         Left            =   4740
         TabIndex        =   8
         Top             =   405
         Width           =   1395
         _ExtentX        =   2461
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
         MICON           =   "frmReport.frx":5864
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdPrint 
         Height          =   375
         Left            =   3255
         TabIndex        =   7
         Top             =   405
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         TX              =   "&Print"
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
         MICON           =   "frmReport.frx":5880
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton CmdRept 
         Height          =   375
         Left            =   1770
         TabIndex        =   6
         Top             =   405
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         TX              =   "&Report"
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
         MICON           =   "frmReport.frx":589C
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdDisplay 
         Height          =   375
         Left            =   285
         TabIndex        =   5
         Top             =   405
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         TX              =   "&Text Report"
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
         MICON           =   "frmReport.frx":58B8
         BC              =   12632256
         FC              =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   585
         Left            =   960
         TabIndex        =   1
         Top             =   390
         Width           =   6705
      End
      Begin MSComDlg.CommonDialog Cmd 
         Left            =   0
         Top             =   1860
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reports"
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
      Left            =   6645
      TabIndex        =   3
      Top             =   -525
      Width           =   2145
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataBasePath As String
Dim DBPWD As String
Dim rsc As New ADODB.Recordset
Dim Conn As ADODB.Connection
Dim sql As String
Dim Fso As New FileSystemObject
Public DateTo As String
Dim rs As DAO.Recordset
Dim cn As DAO.Database
Dim Filename1 As String
Dim MaxFlg As Boolean


Private Sub cmdDisplay_Click()
On Error Resume Next
    Cmd.Filter = "TXT (*.TXT)|*.TXT"
    Cmd.ShowOpen
'    rtfText.Width = frmReport.Width - 300
'    rtfText.Height = frmReport.Height - 200
    rtfText.Font = "Lucida console"
    rtfText.Font.Size = "10"
    rtfText.Locked = True
    rtfText.LoadFile (Cmd.filename)
    cmdPrint.Enabled = True
    
    'fname = Cmd.FileName
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo err
Cmd.PrinterDefault = True
Cmd.CancelError = True
    On Error Resume Next
'   ' If ActiveForm Is Nothing Then Exit Sub
    With Cmd
        .DialogTitle = "Print"
        .CancelError = True
        '.Orientation = cdlLandscape
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If err <> MSComDlg.cdlCancel Then
            rtfText.SelPrint .hdc
        End If
    End With
    
    Exit Sub
err:
    MsgBox err.Number & ", " & err.Description, vbInformation, "BUS"
    Exit Sub
End Sub

Private Sub CmdRept_Click()
frmRpt1.Show vbModal
End Sub



Function CONNECTDB()
Set Conn = New ADODB.Connection
Set rsc = New ADODB.Recordset
Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
Conn.ConnectionString = "Data Source=" & App.Path & "\Pvt.mdb"
Conn.Properties("Jet OLEDB:Database Password") = "silbus"
Conn.Open
End Function

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If MaxFlg = True Then
       MaxFlg = False
       frmReport.WindowState = 0
    Else
       MaxFlg = True
       frmReport.WindowState = 2
    End If
End Sub
Private Sub Form_Load()
    Me.Icon = frmMainform.Icon
   Set cn = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
   MaxFlg = False
End Sub

Private Sub Form_Resize()
    Frame1.Width = frmReport.Width - 100
    'Frame1.Height = frmReport.Height - 2000
    rtfText.Width = Frame1.Width - 100
    rtfText.Height = Frame1.Height - 1500
    Command2.Left = frmReport.Width - 100
End Sub
