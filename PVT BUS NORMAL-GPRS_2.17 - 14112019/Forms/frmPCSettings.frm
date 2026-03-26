VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmPCSettings 
   Caption         =   "Path Settings"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   Icon            =   "frmPCSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton cmdCancel 
      Height          =   375
      Left            =   4635
      TabIndex        =   7
      Top             =   1605
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
      MICON           =   "frmPCSettings.frx":57E2
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdSave 
      Height          =   375
      Left            =   3150
      TabIndex        =   6
      Top             =   1600
      Width           =   1400
      _ExtentX        =   2461
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
      MICON           =   "frmPCSettings.frx":57FE
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdBTransfer 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   1000
      Width           =   1000
      _ExtentX        =   1773
      _ExtentY        =   661
      TX              =   "B&rowse"
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
      MICON           =   "frmPCSettings.frx":581A
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdBCollection 
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   250
      Width           =   1000
      _ExtentX        =   1773
      _ExtentY        =   661
      TX              =   "&Browse"
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
      MICON           =   "frmPCSettings.frx":5836
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox txtTransferPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1350
      TabIndex        =   3
      Top             =   1000
      Width           =   3500
   End
   Begin VB.TextBox txtCollectionPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1350
      TabIndex        =   2
      Top             =   250
      Width           =   3500
   End
   Begin VB.Label Label2 
      Caption         =   "Transfer Path"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   1
      Top             =   1005
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Collection Path"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   255
      Width           =   1140
   End
End
Attribute VB_Name = "frmPCSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FS As New FileSystemObject
Private Sub cmdBCollection_Click()
    
On Error Resume Next
    Load frmSelectFolder
    frmSelectFolder.Drive.Drive = Left$(txtCollectionPath, 2)
    frmSelectFolder.Dir.Path = txtCollectionPath
    frmSelectFolder.Show vbModal
    If Not FS.FolderExists(frmSelectFolder.Dir.Path) Then
        MsgBox "Folder Doesn't exists!"
        Unload frmSelectFolder
        Exit Sub
    End If
    txtCollectionPath = frmSelectFolder.Dir.Path
    Unload frmSelectFolder
End Sub

Private Sub cmdBTransfer_Click()
On Error Resume Next
    Load frmSelectFolder
    frmSelectFolder.Drive.Drive = Left$(txtTransferPath, 2)
    frmSelectFolder.Dir.Path = txtTransferPath
    frmSelectFolder.Show vbModal
    If Not FS.FolderExists(frmSelectFolder.Dir.Path) Then
        MsgBox "Folder Doesn't exists!"
        Unload frmSelectFolder
        Exit Sub
    End If
    txtTransferPath = frmSelectFolder.Dir.Path
    Unload frmSelectFolder
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    If txtTransferPath = "" Then
    
        MsgBox "Invalid Transfer Path!" & vbCrLf & "Please check Path", vbInformation
        txtTransferPath.SetFocus
        Exit Sub
        
    End If
    If txtCollectionPath = "" Then
    
        MsgBox "Invalid Collection Path!" & vbCrLf & "Please check Path", vbInformation
            txtCollectionPath.SetFocus
        Exit Sub
        
    End If
    
    TSQL = "SELECT * FROM PCSETUP"
    Set RES = CNN.OpenRecordset(TSQL, dbOpenDynaset)
    txtTransferPath = CreateFolder(txtTransferPath.Text)
    txtCollectionPath = CreateFolder(txtCollectionPath.Text)
    If RES.RecordCount > 0 Then
        RES.Edit
        RES!TRANSFER_PATH = txtTransferPath
        RES!TICKET_PATH = txtCollectionPath
        RES.Update
    Else
        RES.AddNew
        RES!TRANSFER_PATH = txtTransferPath
        RES!TICKET_PATH = txtCollectionPath
        RES.Update
    End If
    RES.Close
    MsgBox "Settings Saved Successfully"
End Sub

Private Sub Form_Load()
    Me.Icon = frmMainform.Icon
    TSQL = "SELECT * FROM PCSETUP"
    CONNECTDB
    Set RES = CNN.OpenRecordset(TSQL, dbOpenDynaset)
    If RES.RecordCount > 0 Then
        txtTransferPath = IIf(IsNull(RES!TRANSFER_PATH), App.Path, RES!TRANSFER_PATH)
        txtCollectionPath = IIf(IsNull(RES!TICKET_PATH), App.Path, RES!TICKET_PATH)
    End If
    RES.Close
End Sub

