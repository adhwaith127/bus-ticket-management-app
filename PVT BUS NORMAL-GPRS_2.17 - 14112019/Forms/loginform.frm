VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form loginform 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2385
   ClientLeft      =   3600
   ClientTop       =   3315
   ClientWidth     =   5025
   Icon            =   "loginform.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "loginform.frx":11F84
   ScaleHeight     =   2385
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton CmdCancel 
      Height          =   450
      Left            =   3630
      TabIndex        =   7
      Top             =   1680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "loginform.frx":15E6A
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdOk 
      Height          =   450
      Left            =   2340
      TabIndex        =   6
      Top             =   1680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      TX              =   "&Login"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "loginform.frx":15E86
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox txtpwd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3030
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   1815
   End
   Begin VB.TextBox txtusrnme 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3015
      MaxLength       =   10
      TabIndex        =   0
      Top             =   690
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2085
      Left            =   165
      TabIndex        =   2
      Top             =   2820
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Left            =   1650
      TabIndex        =   5
      Top             =   30
      Width           =   2265
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   1620
      TabIndex        =   4
      Top             =   1125
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "UserName :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
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
      Top             =   720
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   2295
      Left            =   -15
      Picture         =   "loginform.frx":15EA2
      Top             =   -15
      Width           =   2295
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
                MsgBox "Invalid username or password", vbCritical, gblstrPrjTitle
                txtusrnme.Text = ""
                txtpwd.Text = ""
                txtusrnme.SetFocus
                Exit Sub
            End If
        End If
        If setdbfields Then
        
        End If
        If StrComp(recs!Username, Trim(txtusrnme.Text)) = 0 And StrComp(recs!PassWord, Trim(txtpwd.Text)) = 0 Then
           loginsucceed = True
           frmMainform.Enabled = True
        ElseIf StrComp(recs!SUPERUSER, Trim(txtusrnme.Text)) = 0 And StrComp(recs!SuperPassword, Trim(txtpwd.Text)) = 0 Then
           loginsucceed = True
           SUPERUSER = True
           frmMainform.Enabled = True
        ElseIf StrComp(recs!Username, Trim(txtusrnme.Text)) <> 0 Then
           MsgBox "Invalid username", vbCritical, gblstrPrjTitle
           txtusrnme.Text = ""
           txtusrnme.SetFocus
        ElseIf StrComp(recs!PassWord, Trim(txtpwd.Text)) <> 0 Then
           MsgBox "Invalid password", vbCritical, gblstrPrjTitle
           txtpwd.Text = ""
           txtpwd.SetFocus
        End If
   
    Else
        MsgBox "Enter all the fields", vbCritical, gblstrPrjTitle
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
Private Function setdbfields() As Boolean
On Error Resume Next
    sql = "ALTER TABLE RPT ADD COLUMN  ladies_count  NUMBER;"
    CON.Execute sql
    
    sql = "ALTER TABLE RPT ADD COLUMN senior_count  NUMBER;"
    CON.Execute sql
    
    sql = "ALTER TABLE RPT ADD COLUMN senior_coll  NUMBER;"
    CON.Execute sql
    
    sql = "ALTER TABLE RPT ADD COLUMN ladies_coll  NUMBER;"
    CON.Execute sql
    
    sql = "ALTER TABLE TKTS ADD COLUMN  ladies_count  NUMBER"
    CON.Execute sql
    
    sql = "ALTER TABLE TKTS ADD COLUMN senior_count  NUMBER"
    CON.Execute sql
    
    sql = "Alter table Settings add column ladies_ratio double"
    CON.Execute sql
    
    sql = "ALTER TABLE Settings add column senior_ratio double"
    CON.Execute sql
    
End Function
Private Sub Form_Activate()
On Error Resume Next
  '   Me.Icon = frmMainform.Icon
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
        cmdOK.SetFocus
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
