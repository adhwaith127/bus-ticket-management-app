VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmPortSetup 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Port Setup"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4545
   Icon            =   "frmPortSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1350
      Left            =   -90
      TabIndex        =   0
      Top             =   15
      Width           =   4620
      Begin JeweledBut.JeweledButton Command1 
         Height          =   375
         Left            =   3200
         TabIndex        =   7
         Top             =   780
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
         MICON           =   "frmPortSetup.frx":0CCA
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSave 
         Height          =   375
         Left            =   3200
         TabIndex        =   6
         Top             =   285
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
         MICON           =   "frmPortSetup.frx":0CE6
         BC              =   12632256
         FC              =   0
      End
      Begin VB.ComboBox cmbBoud 
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
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   735
         Width           =   1170
      End
      Begin VB.TextBox txtPort 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1335
         MaxLength       =   2
         TabIndex        =   1
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Boud Rate"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   315
         TabIndex        =   3
         Top             =   765
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   2
         Top             =   315
         Width           =   360
      End
   End
   Begin MSCommLib.MSComm SerialCom 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Port Setup"
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
      Left            =   1305
      TabIndex        =   5
      Top             =   -540
      Width           =   2325
   End
End
Attribute VB_Name = "frmPortSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim Port As Byte
Dim Boud As String
Private Sub cmbBoud_Click()
    Boud = cmbBoud.Text
End Sub

Private Sub cmbBoud_KeyPress(KeyAscii As Integer)
    cmdSave.SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo err
Dim Ctrl As Control
    
    For Each Ctrl In Me.Controls
        If TypeOf Ctrl Is TextBox Then
            If Ctrl.Text = "" Then
                MsgBox "Some field missing!", vbInformation, "Port Details"
                Exit Sub
            End If
        End If
    Next
    Set SerialComm = SerialCom
    If InitPort(val(txtPort.Text), cmbBoud.List(cmbBoud.ListIndex)) = False Then
        MsgBox "Port open error! Invalid Port number or Port already in use. Close any application using same port", vbExclamation, App.ProductName
        txtPort.SetFocus
        Call SendKeys("{HOME}+{END}")
        Exit Sub
    End If
    Set RES = DB.OpenRecordset("PORT", dbOpenDynaset)
    Port = val(txtPort)
    Boud = cmbBoud.Text
    If RES.RecordCount > 0 Then
        RES.Edit
        RES!Port = Port
        RES!Boud = Boud
        RES.Update
        MsgBox vbCrLf & "Port Settings saved Successfully.." & vbCrLf, vbInformation, App.ProductName
    Else
        RES.AddNew
        RES!Port = Port
        RES!Boud = Boud
        RES.Update
        MsgBox vbCrLf & "Port Settings saved Successfully.." & vbCrLf, vbInformation, App.ProductName
    End If
    RES.Close
    Unload Me
    Exit Sub
err:
    MsgBox err.Number & "  - " & err.Description, vbInformation, "Error"
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err
    Me.Icon = frmMainform.Icon
    CONNECT_DB
    LoadCombo
    Set RES = DB.OpenRecordset("PORT", dbOpenDynaset)
    If RES.RecordCount > 0 Then
        txtPort = RES!Port
        cmbBoud.Text = RES!Boud
        RES.Close
    End If
    
    Exit Sub
err:
    MsgBox err.Number & "  - " & err.Description, vbInformation, "Error"
End Sub

Private Sub txtPort_GotFocus()
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort)
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If TextBoxVal(KeyAscii) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 And txtPort <> "" Then cmbBoud.SetFocus
End Sub
Public Function TextBoxVal(Key As Integer) As Integer
Dim st As String
    Key = Asc(UCase(Chr(Key)))
    st = "ABCDEFGHIJKLMNOPQRSTUVWXYZ`~!@#$%^&*()_-+=|/?<>;',. {}[]:\"
    TextBoxVal = InStr(st, Chr(Key))
End Function

Public Sub LoadCombo()
    cmbBoud.Clear
    cmbBoud.AddItem "110"
    cmbBoud.AddItem "300"
    cmbBoud.AddItem "1200"
    cmbBoud.AddItem "2400"
    cmbBoud.AddItem "4800"
    cmbBoud.AddItem "9600"
    cmbBoud.AddItem "19200"
    cmbBoud.AddItem "38400"
    cmbBoud.AddItem "57600"
    cmbBoud.AddItem "115200"
    cmbBoud.AddItem "230400"
    cmbBoud.AddItem "460800"
    cmbBoud.AddItem "921600"
    cmbBoud.Text = "115200"
End Sub

