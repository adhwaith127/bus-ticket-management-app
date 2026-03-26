VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmAddUser 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User"
   ClientHeight    =   2475
   ClientLeft      =   4185
   ClientTop       =   3345
   ClientWidth     =   4200
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B94A55&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   4605
      ScaleHeight     =   3330
      ScaleWidth      =   600
      TabIndex        =   8
      Top             =   -570
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CheckBox admin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Admin"
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
      Left            =   450
      TabIndex        =   7
      Top             =   1710
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2475
      Left            =   -90
      TabIndex        =   0
      Top             =   -30
      Width           =   4410
      Begin JeweledBut.JeweledButton CmdCancel 
         Height          =   375
         Left            =   2835
         TabIndex        =   11
         Top             =   1815
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
         MICON           =   "frmAddUser.frx":0CCA
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton CmdOK 
         Height          =   375
         Left            =   1530
         TabIndex        =   10
         Top             =   1815
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         TX              =   "&OK"
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
         MICON           =   "frmAddUser.frx":0CE6
         BC              =   12632256
         FC              =   0
      End
      Begin VB.TextBox txtnewusrnme 
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
         Left            =   2565
         MaxLength       =   10
         TabIndex        =   1
         Top             =   285
         Width           =   1365
      End
      Begin VB.TextBox txtnewpwd 
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
         IMEMode         =   3  'DISABLE
         Left            =   2565
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   750
         Width           =   1365
      End
      Begin VB.TextBox txtconewpwd 
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
         IMEMode         =   3  'DISABLE
         Left            =   2565
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1215
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Password "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   465
         TabIndex        =   6
         Top             =   765
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   465
         TabIndex        =   5
         Top             =   1178
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   435
         TabIndex        =   4
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add User"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   885
      TabIndex        =   9
      Top             =   -585
      Width           =   2835
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CON As DAO.Database
Public recs As DAO.Recordset
Public strSql As String
Public UsrName As String
Public Pwd As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If txtnewusrnme.Text <> "" And txtnewpwd.Text <> "" And txtconewpwd.Text <> "" Then
        If Trim(txtnewpwd.Text) = Trim(txtconewpwd.Text) Then
            strSql = "select UserName from LOGINTABLE where UserName ='" & Trim(txtnewusrnme) & "' or SuperUser='" & Trim(txtnewusrnme) & "'"
            Set RES = CON.OpenRecordset(strSql, dbOpenDynaset)
            If RES.RecordCount > 0 Then
                MsgBox vbCrLf & "UserName already exist" & vbCrLf & "Please choose different Username" & vbCrLf & vbCrLf, vbInformation
                txtnewusrnme.Text = ""
                txtnewpwd.Text = ""
                txtconewpwd.Text = ""
                txtnewusrnme.SetFocus
                Exit Sub
            End If
            strSql = "select * from LOGINTABLE "
            Set RES = CON.OpenRecordset(strSql, dbOpenDynaset)
            If RES.RecordCount > 0 Then RES.MoveLast
            If admin.Value = 1 Then
                RES.AddNew
                RES!Username = Trim(txtnewusrnme)
                RES!SUPERUSER = Trim(txtnewusrnme)
                RES!SuperPassword = Trim(txtnewpwd)
                RES.Update
                MsgBox "ADMIN User Added Successfully", vbInformation
                RES.Close
                Unload Me
            Else
                RES.AddNew
                RES!Username = Trim(txtnewusrnme)
                RES!PassWord = Trim(txtnewpwd)
                RES.Update
                MsgBox "New User Added Successfully", vbInformation
                RES.Close
                Unload Me
            End If
        Else
            MsgBox "PassWords does not match, Please re enter", vbCritical
            txtconewpwd.Text = ""
            txtconewpwd.SetFocus
        End If
    Else
        MsgBox "Enter All the Fields", vbInformation
    End If
End Sub
Private Sub Form_Load()
    Me.Icon = frmMainform.Icon
    Set CON = OpenDatabase(App.Path & "\PVT", _
    dbDriverComplete, False, ";UID=;PWD=silbus")
    If LoginStatus = 0 Then admin.Enabled = False Else: admin.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CON.Close
End Sub
Private Sub txtnewusrnme_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtnewusrnme <> "" Then txtnewpwd.SetFocus
End Sub
Private Sub txtnewpwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtnewpwd <> "" Then txtconewpwd.SetFocus
End Sub
Private Sub txtconewpwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtconewpwd <> "" Then cmdOK.SetFocus
End Sub
