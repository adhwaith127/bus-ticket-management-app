VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmRemoveUser 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete User"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4500
   Icon            =   "frmRemoveUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   4470
      Begin JeweledBut.JeweledButton cmdExit 
         Height          =   405
         Left            =   2325
         TabIndex        =   9
         Top             =   1785
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   714
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
         MICON           =   "frmRemoveUser.frx":0CCA
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdRemove 
         Height          =   405
         Left            =   1050
         TabIndex        =   8
         Top             =   1785
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   714
         TX              =   "&Remove"
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
         MICON           =   "frmRemoveUser.frx":0CE6
         BC              =   12632256
         FC              =   0
      End
      Begin VB.TextBox txtUser 
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
         Left            =   2385
         MaxLength       =   10
         TabIndex        =   3
         Top             =   300
         Width           =   1600
      End
      Begin VB.TextBox txtPWD 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2385
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   810
         Width           =   1600
      End
      Begin VB.TextBox txtCPWD 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2385
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1275
         Width           =   1600
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   315
         TabIndex        =   6
         Top             =   405
         Width           =   1740
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   5
         Top             =   870
         Width           =   1785
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   4
         Top             =   1335
         Width           =   1800
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete User"
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
      Left            =   1140
      TabIndex        =   7
      Top             =   -465
      Width           =   2325
   End
End
Attribute VB_Name = "frmRemoveUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
Dim sql As String
Dim rs As DAO.Recordset
    
    If txtpwd <> txtCPWD Then
        MsgBox "Password Doesn't match!", vbInformation, gblstrPrjTitle
        Exit Sub
    End If
    
    If GlobUsrName = txtUser Then
        MsgBox "Logined user cannot delete!", vbInformation, gblstrPrjTitle
        Exit Sub
    End If
    sql = "SELECT UserName ,SuperUser,SuperPassWord FROM LOGINTABLE WHERE  SuperUser= '" & txtUser & "' and SuperPassWord ='" & txtpwd & "'"
    Set rs = DB.OpenRecordset(sql, dbOpenDynaset)
    If rs.RecordCount > 0 Then
        If (txtUser.Text = rs!SUPERUSER) And (txtpwd.Text = rs!SuperPassword) And IsNull(rs!Username) = True Then
            MsgBox txtUser.Text & " is a Super Logined User" & vbCrLf & _
                " Can't Remove this Username and Password", vbInformation, gblstrPrjTitle
                txtUser = ""
                txtpwd = ""
                txtCPWD = ""
                txtUser.SetFocus
                rs.Close
            Exit Sub
        End If
    End If
    rs.Close
    
    sql = "SELECT UserName ,PassWord FROM LOGINTABLE WHERE USERNAME = '" & txtUser & "' and PassWord ='" & txtpwd & "'"
    Set rs = DB.OpenRecordset(sql, dbOpenDynaset)
    If rs.RecordCount > 0 Then
        sql = "DELETE UserName ,PassWord FROM LOGINTABLE WHERE USERNAME = '" & txtUser & "' and PassWord ='" & txtpwd & "'"
        DB.Execute (sql)
        MsgBox "User removed Successfully", vbInformation, gblstrPrjTitle
        txtUser = ""
        txtpwd = ""
        txtCPWD = ""
        txtUser.SetFocus
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    sql = "SELECT UserName ,SuperUser,SuperPassWord FROM LOGINTABLE WHERE USERNAME = '" & txtUser & "' and SuperUser= '" & txtUser & "' and SuperPassWord ='" & txtpwd & "'"
    Set rs = DB.OpenRecordset(sql, dbOpenDynaset)
    If rs.RecordCount > 0 Then
        sql = "DELETE UserName ,SuperUser,SuperPassWord FROM LOGINTABLE WHERE USERNAME = '" & txtUser & "' and SuperUser= '" & txtUser & "' and SuperPassWord ='" & txtpwd & "'"
        DB.Execute (sql)
        MsgBox "User removed Successfully", vbInformation, gblstrPrjTitle
        txtUser = ""
        txtpwd = ""
        txtCPWD = ""
        txtUser.SetFocus
        rs.Close
        Exit Sub
    End If
    txtUser = ""
    txtpwd = ""
    txtCPWD = ""
    txtUser.SetFocus
    rs.Close
    MsgBox "Invalid Username and Password!", vbInformation, gblstrPrjTitle
    Exit Sub
    
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Me.Icon = frmMainform.Icon
    txtUser.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo err
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    Exit Sub
err:
    MsgBox err.Description & "- " & err.Number
End Sub


Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtUser <> "" Then txtpwd.SetFocus
End Sub
Private Sub txtPWD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtpwd <> "" Then txtCPWD.SetFocus
End Sub

Private Sub txtcpwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCPWD <> "" Then cmdRemove.SetFocus
End Sub

