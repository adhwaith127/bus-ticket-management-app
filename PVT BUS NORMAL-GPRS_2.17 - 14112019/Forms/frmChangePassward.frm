VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmChangePassword 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4560
   Icon            =   "frmChangePassward.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2805
      Left            =   -15
      TabIndex        =   0
      Top             =   -45
      Width           =   4590
      Begin JeweledBut.JeweledButton cmdCancel 
         Height          =   375
         Left            =   3330
         TabIndex        =   12
         Top             =   2205
         Width           =   1005
         _ExtentX        =   1773
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
         MICON           =   "frmChangePassward.frx":0CCA
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdOK 
         Height          =   375
         Left            =   2190
         TabIndex        =   11
         Top             =   2205
         Width           =   1005
         _ExtentX        =   1773
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
         MICON           =   "frmChangePassward.frx":0CE6
         BC              =   12632256
         FC              =   0
      End
      Begin VB.CheckBox chkAdmin 
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
         Height          =   315
         Left            =   435
         TabIndex        =   9
         Top             =   2295
         Width           =   1035
      End
      Begin VB.TextBox txtConPasswd 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2430
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1695
         Width           =   1920
      End
      Begin VB.TextBox txtNewPasswd 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2430
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1230
         Width           =   1920
      End
      Begin VB.TextBox txtOldPasswd 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2430
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   765
         Width           =   1920
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
         Height          =   330
         Left            =   2445
         MaxLength       =   10
         TabIndex        =   1
         Top             =   315
         Width           =   1905
      End
      Begin VB.Label Label4 
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
         Height          =   450
         Left            =   450
         TabIndex        =   8
         Top             =   1740
         Width           =   1710
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   450
         TabIndex        =   7
         Top             =   1290
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   450
         TabIndex        =   6
         Top             =   855
         Width           =   1395
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
         Height          =   225
         Left            =   450
         TabIndex        =   5
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change Password"
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
      Left            =   990
      TabIndex        =   10
      Top             =   -585
      Width           =   2835
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CON As DAO.Database
Private recs As DAO.Recordset
Private sql As String
Private UserID As String
Private OldPassword As String
Private NewPassword As String
Private ConPassword As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Msg As VbMsgBoxResult
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeOf Ctrl Is TextBox Then
            If Ctrl.Text = "" Then
                MsgBox "Some field missing!", vbInformation, gblstrPrjTitle
                Exit Sub
            End If
        End If
    Next
    If Trim(txtConPasswd) <> Trim(txtNewPasswd) Then
        MsgBox "New Password And Confirm Password is Not Same.", vbInformation, gblstrPrjTitle
        txtNewPasswd.SetFocus
        Exit Sub
    End If
    Msg = MsgBox("Do you want to Change Password ?", vbYesNo)
    If Msg = vbYes Then
        Msg = MsgBox("Are You Sure ?", vbYesNo)
        If Msg = vbNo Then Exit Sub
    Else
        Exit Sub
    End If
    UserID = Trim(txtUser)
    OldPassword = Trim(txtOldPasswd)
    NewPassword = Trim(txtNewPasswd)
    ConPassword = Trim(txtConPasswd)
    
    If chkAdmin.Value = 0 Then
        adminflag = True
        sql = "SELECT * FROM logintable WHERE UserName ='" & UserID & "' AND PassWord ='" & OldPassword & "'"
    ElseIf chkAdmin.Value = 1 Then
        sql = "SELECT * FROM logintable WHERE SuperUser ='" & UserID & "' AND SuperPassWord ='" & OldPassword & "'"
    End If
    If GetPassword = True Then Unload Me
End Sub

Public Function GetPassword() As Boolean
Dim cmpVal As Byte

    Set CON = OpenDatabase(App.Path & "\PVT", _
    dbDriverComplete, False, ";UID=;PWD=silbus")
    Set recs = CON.OpenRecordset(sql, dbOpenDynaset)
    If recs.RecordCount > 0 Then
        If chkAdmin.Value = 0 Then
            If StrComp(UserID, recs!Username) = 0 Then
                If StrComp(txtOldPasswd, recs!PassWord) = 0 Then
                    'If adminflag <> True Then
                    recs.Edit
                    recs!Username = Trim(txtUser)
                    recs!PassWord = Trim(txtNewPasswd)
                    recs.Update
                    MsgBox "Password Changed Successfully", vbInformation, "Password"
                    GetPassword = True
                    Exit Function
                   ' Else
'                        recs.Edit
'                        recs!UserName = Trim(txtUser)
'                        recs.Fields(2) = Trim(txtUser)
'                        recs!SuperPassword = Trim(txtNewPasswd)
'                        recs.Update
'                        MsgBox "Admin Password Changed Successfully", vbInformation, "Password"
'                        GetPassword = True
                    'End If
                End If
            End If
        ElseIf chkAdmin.Value = 1 Then
            If StrComp(UserID, recs!SUPERUSER) = 0 Then
                If StrComp(txtOldPasswd, recs!SuperPassword) = 0 Then
                        recs.Edit
                        recs!Username = Trim(txtUser)
                        recs.Fields(2) = Trim(txtUser)
                        recs!SuperPassword = Trim(txtNewPasswd)
                        recs.Update
                        MsgBox "Admin Password Changed Successfully", vbInformation, "Password"
                        GetPassword = True
                    Exit Function
                End If
            End If
        
        End If
    End If
    MsgBox "Invalid User Name and Password! ", vbExclamation, "Password"
    GetPassword = False
End Function


Private Sub Form_Load()
    Me.Icon = frmMainform.Icon
    If LoginStatus = 0 Then chkAdmin.Enabled = False Else: chkAdmin.Enabled = True
End Sub

Private Sub txtConPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtConPasswd <> "" Then cmdOK_Click
End Sub

Private Sub txtNewPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtNewPasswd <> "" Then txtConPasswd.SetFocus
End Sub

Private Sub txtOldPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtOldPasswd <> "" Then txtNewPasswd.SetFocus
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtUser <> "" Then txtOldPasswd.SetFocus
End Sub
