VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmSelectFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Folder"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   Icon            =   "frmSelectFolder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin JeweledBut.JeweledButton cmdDir 
      Height          =   405
      Left            =   200
      TabIndex        =   4
      Top             =   3180
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   714
      TX              =   "Create &Dir"
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
      MICON           =   "frmSelectFolder.frx":57E2
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdCancel 
      Height          =   405
      Left            =   3615
      TabIndex        =   3
      Top             =   3180
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      TX              =   "&Exit"
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
      MICON           =   "frmSelectFolder.frx":57FE
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdOK 
      Height          =   405
      Left            =   2535
      TabIndex        =   2
      Top             =   3180
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      TX              =   "&OK"
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
      MICON           =   "frmSelectFolder.frx":581A
      BC              =   12632256
      FC              =   0
   End
   Begin VB.DirListBox Dir 
      Height          =   2340
      Left            =   210
      TabIndex        =   1
      Top             =   720
      Width           =   4425
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   300
      Width           =   2490
   End
End
Attribute VB_Name = "frmSelectFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDir_Click()
On Error GoTo err
Dim str As String
    str = InputBox("Enter Directory Name", "New Directory")
    If str = "" Then Exit Sub
    If Right$(Dir.Path, 1) <> "\" Then str = "\" & str
    MkDir Dir.Path & str
    Dir.Refresh
    Exit Sub
err:
    MsgBox "Error:" & err.Number & vbCrLf & err.Description
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Drive_Change()
    Dir.Path = Drive.Drive
End Sub

Private Sub Form_Load()
Me.Icon = frmMainform.Icon
End Sub
