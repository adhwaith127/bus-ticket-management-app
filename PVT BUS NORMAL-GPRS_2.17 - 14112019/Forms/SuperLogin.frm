VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form SuperLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super User Login"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4200
   Icon            =   "SuperLogin.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   -45
      TabIndex        =   2
      Top             =   -15
      Width           =   4215
      Begin JeweledBut.JeweledButton CmdSupCancel 
         Height          =   375
         Left            =   2415
         TabIndex        =   7
         Top             =   1395
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         TX              =   "&Cancel"
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
         MICON           =   "SuperLogin.frx":0CCA
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSupLog 
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   1395
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         TX              =   "&Login"
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
         MICON           =   "SuperLogin.frx":0CE6
         BC              =   12632256
         FC              =   0
      End
      Begin VB.TextBox txtSupPwd 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2220
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtSupUsrnme 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2205
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   735
         TabIndex        =   4
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   270
         Width           =   1380
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Super User Login"
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
      Left            =   945
      TabIndex        =   5
      Top             =   -495
      Width           =   2715
   End
End
Attribute VB_Name = "SuperLogin"
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

Private Sub CmdSupCancel_Click()
On Error Resume Next
    DataTrans = False
    RouteAdd = False
    GraphEdit = False
    FareTableEdit = False
    RouteEdit = False
    StageEdit = False
    DeleteRoute = False
    Settings = False
    PCSettings = False
    Unload Me
End Sub

Private Sub cmdSupLog_Click()
On Error GoTo err
    If txtSupUsrnme.Text <> "" And txtSupPwd.Text <> "" Then
        GlobUsrName = Trim(txtSupUsrnme.Text)
        strSql = "SELECT * FROM LOGINTABLE"
        Set recs = CON.OpenRecordset(strSql, dbOpenDynaset)
        If recs.RecordCount > 0 Then
            Do While Not recs.EOF
            If recs!SUPERUSER = Trim(txtSupUsrnme.Text) And recs!SuperPassword = Trim(txtSupPwd.Text) Then
                PC_to_PMTC_Cntr = 1
'                Unload Me  'loginform.Hide
                Me.Hide
                If DataTrans = True Then
                    Load FrmTransfer
                    FrmTransfer.Show vbModal
                    DataTrans = False
                ElseIf RouteAdd = True Then 'main form
                    Load FareTableFrm
                    FareTableFrm.Show vbModal
                    RouteAdd = False
                ElseIf GraphEdit = True Then  'graph edit form
                    Load frmgraphedit
                    frmgraphedit.Show vbModal
                    GraphEdit = False
                ElseIf FareTableEdit = True Then  'Fare Table edit form
                    Load frmFareTableEdit
                    frmFareTableEdit.Show vbModal
                    FareTableEdit = False
                ElseIf RouteEdit = True Then  'Route edit form
                    Load frmRoute
                    frmRoute.Show vbModal
                    RouteEdit = False
                ElseIf StageEdit = True Then  'stage edit form
                    RouteID = ""
                    Load frmStage
                    frmStage.Show vbModal
                    StageEdit = False
                ElseIf DeleteRoute = True Then  ' RouteDelete form
                    Load frmRouteDelete
                    frmRouteDelete.Show vbModal
                    DeleteRoute = False
                ElseIf Settings = True Then       'Settings form
                    Load frmSettings
                    frmSettings.Show vbModal
                    Settings = False
                ElseIf PCSettings = True Then       'PCSettings form
                    Load frmPCSettings
                    frmPCSettings.Show vbModal
                    PCSettings = False
                End If
                
                Unload Me
                LoginStatus = 1
                Exit Do
            Else
                 LoginStatus = 0
                recs.MoveNext
            End If
            Loop
'            ElseIf StrComp(recs!SUPERUSER, Trim(txtSupUsrnme.Text)) <> 0 Then
'                LoginStatus = 0
'                MsgBox "Invalid SuperUser Name", vbCritical
'                txtSupUsrnme.Text = ""
'                txtSupPwd.Text = ""
'                txtSupUsrnme.SetFocus
'                 Exit Do
'            ElseIf StrComp(recs!SuperPassword, Trim(txtSupPwd.Text)) <> 0 Then
'                MsgBox "Invalid SuperUser Password", vbCritical
'                txtSupPwd.Text = ""
'                txtSupPwd.SetFocus
'                 Exit Do
'            End If
'            LoginStatus = 1
'            recs.MoveNext
            'Loop
            
        End If
        If LoginStatus = 0 Then
                
                MsgBox "Invalid SuperUser Name Or Password", vbCritical
                txtSupUsrnme.Text = ""
                txtSupPwd.Text = ""
                txtSupUsrnme.SetFocus
                Exit Sub
       End If
    Else
        MsgBox "Enter All the Fields"
    End If
    Exit Sub
err:
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Icon = frmMainform.Icon
     Set CON = OpenDatabase(App.Path & "\PVT", dbDriverComplete, False, ";UID=;PWD=silbus")
     LoginStatus = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    CON.Close
End Sub


Private Sub txtSupUsrnme_KeyPress(KeyAscii As Integer)
On Error Resume Next
     Select Case KeyAscii
        Case 13
            txtSupPwd.SetFocus
    End Select
End Sub

Private Sub txtSupPwd_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Select Case KeyAscii
        Case 13
            cmdSupLog.SetFocus
    End Select
End Sub

