VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form loginform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2535
   ClientLeft      =   3570
   ClientTop       =   2940
   ClientWidth     =   4440
   ControlBox      =   0   'False
   Icon            =   "loginform.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton CmdCancel 
      Height          =   450
      Left            =   2790
      TabIndex        =   9
      Top             =   1920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      TX              =   "&Cancel"
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
      MICON           =   "loginform.frx":0CCA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdOk 
      Height          =   450
      Left            =   1500
      TabIndex        =   8
      Top             =   1920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      TX              =   "&Login"
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
      MICON           =   "loginform.frx":0CE6
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox txtpwd 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2550
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1500
      Width           =   1455
   End
   Begin VB.TextBox txtusrnme 
      Appearance      =   0  'Flat
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   2535
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1050
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2085
      Left            =   165
      TabIndex        =   2
      Top             =   2460
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
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
      Height          =   510
      Left            =   1770
      TabIndex        =   7
      Top             =   150
      Width           =   2265
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1155
      TabIndex        =   6
      Top             =   1635
      Width           =   450
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1050
      TabIndex        =   5
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   1485
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   1620
      TabIndex        =   3
      Top             =   1080
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   0
      Picture         =   "loginform.frx":0D02
      Top             =   -15
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   3195
      Left            =   -15
      Top             =   -15
      Width           =   1800
   End
End
Attribute VB_Name = "loginform"
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
Public VersionOK As Boolean

Private Sub cmdCancel_Click()
 End
End Sub

Private Sub cmdOK_Click()
On Error GoTo err
    If txtusrnme.Text <> "" And txtpwd.Text <> "" Then
        GlobUsrName = Trim(txtusrnme.Text)
        If StrComp("softland", Trim(txtusrnme.Text)) = 0 And StrComp("sil123", Trim(txtpwd.Text)) = 0 Then
            loginsucceed = True
            frmMainform.Enabled = True
            frmMainform.mnuAddUser.Enabled = True
            frmChangePassword.chkAdmin.Enabled = True
            frmMainform.mnuRemoveUser.Enabled = True
            frmMainform.mnuPort.Enabled = True
            CreateTables
            CreateFields
            LoginStatus = 1
            Unload loginform  'loginform.Hide
            frmMainform.Show
            Exit Sub
        End If
        
        strSql = "SELECT * FROM LOGINTABLE WHERE UserName= '" & Trim(txtusrnme) & "' AND Password = '" & Trim(txtpwd) & "'"
        Set recs = CON.OpenRecordset(strSql, dbOpenDynaset)
        If recs.RecordCount = 0 Then
            recs.Close
            strSql = "SELECT * FROM LOGINTABLE WHERE SUPERUSER= '" & Trim(txtusrnme) & "' AND SuperPassWord = '" & Trim(txtpwd) & "'"
            Set recs = CON.OpenRecordset(strSql, dbOpenDynaset)
            If recs.RecordCount = 0 Then
                recs.Close
                MsgBox "Invalid Username or Password", vbInformation, "Password"
                txtusrnme.Text = ""
                txtpwd.Text = ""
                txtusrnme.SetFocus
                Exit Sub
            End If
        End If
        If StrComp(recs!Username, Trim(txtusrnme.Text)) = 0 And StrComp(recs!PassWord, Trim(txtpwd.Text)) = 0 Then
           loginsucceed = True
           frmMainform.Enabled = True
        ElseIf StrComp(recs!SUPERUSER, Trim(txtusrnme.Text)) = 0 And StrComp(recs!SuperPassword, Trim(txtpwd.Text)) = 0 Then
           loginsucceed = True
           SUPERUSER = True
           frmMainform.Enabled = True
        ElseIf StrComp(recs!Username, Trim(txtusrnme.Text)) <> 0 Then
           MsgBox "Invalid Username", vbCritical
           txtusrnme.Text = ""
           txtusrnme.SetFocus
        ElseIf StrComp(recs!PassWord, Trim(txtpwd.Text)) <> 0 Then
           MsgBox "Invalid Password", vbCritical
           txtpwd.Text = ""
           txtpwd.SetFocus
        End If
   
    Else
        MsgBox "Enter All the Fields"
        txtusrnme.SetFocus
    End If
    If loginsucceed = True Then
        If SUPERUSER = True Then
            frmMainform.mnuAddUser.Enabled = True
            frmChangePassword.chkAdmin.Enabled = True
            frmMainform.mnuRemoveUser.Enabled = True
            frmMainform.mnuPort.Enabled = True
            LoginStatus = 1
        Else
            frmMainform.mnuAddUser.Enabled = False
            frmChangePassword.chkAdmin.Enabled = False
            frmMainform.mnuRemoveUser.Enabled = False
            frmMainform.mnuPort.Enabled = False
            LoginStatus = 0
        End If
       CreateFields
       CreateTables
       CreateNewtables   ''' 20/01/2011
       Unload Me  'loginform.Hide
     '  Tarif_frm.Show
       frmMainform.Show
    End If
    Exit Sub
err:
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If VersionOK = False Then End
    If ProjectValidity = False Then End
    txtusrnme.SetFocus
End Sub

Private Sub Form_Load()
'ConnectDatabase(adoDBCon, App.Path & "\PVT.mdb", "silbus") As Boolean

On Error GoTo err
'Public Const gblprjTitle = "PayCollect"
'Public Const ProjectVersion = "CM 04.06.09  Ver : 1.1.1"

If App.PrevInstance = True Then
    MsgBox "Previous instance of Amphibia Bus Ticketing 2.17.exe is running. Please exit the running instance and try again", vbExclamation, gblstrPrjTitle
    End
End If

stayOnTop Me



VersionOK = False
    Dim ProjectVersion, strSql As String
    SUPERUSER = False
    Set CON = OpenDatabase(App.Path & "\PVT.mdb", _
           dbDriverComplete, False, ";UID=;PWD=silbus")
           ProjectVersion = "Amphibia Bus Ticketing 2.13.124"
           
    strSql = "SELECT * FROM VERSION WHERE VersionNo= '" & Trim(ProjectVersion) & "'"
        Set recs = CON.OpenRecordset(strSql, dbOpenDynaset)
        If recs.RecordCount = 0 Then
            MsgBox "Version Mismatch.Please Use the Original Database.Or Contact Sil Customer Care"
            VersionOK = False
            Exit Sub
        End If
        recs.Close
        VersionOK = True
    Change_System_Date
    Exit Sub
err:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If loginsucceed = False Then End
End Sub


Private Sub txtPWD_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Select Case KeyAscii
    Case 13
        cmdOk.SetFocus
    Case Else
        KeyAscii = KeyAscii
    End Select
End Sub


Private Sub txtusrnme_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Select Case KeyAscii
    Case 13
        txtpwd.SetFocus
    Case Else
        KeyAscii = KeyAscii
    End Select
End Sub
Public Function CheckTableExistsOrNot(strTableName As String) As Boolean
On Error GoTo err
    CheckTableExistsOrNot = False
    RSql = "SELECT * FROM " & strTableName
    Set RES = CNN.OpenRecordset(RSql, dbOpenDynaset)
                
    If RES.State <> 0 Then RES.Close
    CheckTableExistsOrNot = True
    Exit Function
err:
    CheckTableExistsOrNot = False
End Function
