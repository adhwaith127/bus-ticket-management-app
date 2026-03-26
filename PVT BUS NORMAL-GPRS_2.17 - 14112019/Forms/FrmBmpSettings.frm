VERSION 5.00
Begin VB.Form frmBmpSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMP Settings"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2100
   ScaleMode       =   0  'User
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   715
      Left            =   60
      TabIndex        =   8
      Top             =   1320
      Width           =   4695
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   4695
      Begin VB.CheckBox chkBold 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cmbFontSize 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cmbFont 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   0
         Text            =   "cmbFont"
         Top             =   240
         Width           =   3045
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Name "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Name "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmBmpSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CON As DAO.Database
Public recs As DAO.Recordset
Public strSql As String
Private Sub chkBold_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call SendKeys("{TAB}")
    ElseIf KeyAscii = vbKeyEscape Then
        Call SendKeys("+{TAB}")
    End If
End Sub
Private Sub cmbFont_GotFocus()
    Call SendKeys("{HOME}+{END}")
End Sub
Private Sub cmbFont_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cmbFont, KeyAscii)
    If KeyAscii = vbKeyReturn And cmbFont.ListIndex <> -1 Then
        Call SendKeys("{TAB}")
    ElseIf KeyAscii = vbKeyEscape Then
        Call SendKeys("+{TAB}")
    End If
End Sub

Private Sub cmbFontSize_GotFocus()
    Call SendKeys("{HOME}+{END}")
End Sub
Private Sub cmbFontSize_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cmbFontSize, KeyAscii)
    If KeyAscii = vbKeyReturn And cmbFontSize.ListIndex <> -1 Then
        Call SendKeys("{TAB}")
    ElseIf KeyAscii = vbKeyEscape Then
        Call SendKeys("+{TAB}")
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
On Error GoTo err
    If cmbFont.Text <> "" And cmbFontSize.Text <> "" Then
        strSql = "select * from BMP_Settings "
        Set RES = CON.OpenRecordset(strSql, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            RES.Edit
            RES!Font_Name = cmbFont.Text
            RES!Font_Size = cmbFontSize.Text
            RES!Bold_EnableOrDisable = (IIf(chkBold.Value = 1, 1, 0))
            RES.Update
            MsgBox "Settings Updated Successfully", vbInformation, App.ProductName
        Else
            RES.AddNew
            RES!Font_Name = cmbFont.Text
            RES!Font_Size = cmbFontSize.Text
            RES!Bold_EnableOrDisable = (IIf(chkBold.Value = 1, 1, 0))
            RES!Bmp_Width = 192 ' 384
            RES!Bmp_Height = 64
            RES.Update
            MsgBox "Settings Saved Successfully", vbInformation, App.ProductName
        End If
        RES.Close
        Unload Me
    Else
        MsgBox "Please select the Font Name and Font Size", vbInformation, App.ProductName
    End If
Exit Sub
err:
    MsgBox err.Number & "  " & err.Description & " in cmdSave_Click"
End Sub
Private Sub Form_Activate()
    LoadFont
End Sub
Public Sub LoadFont()
Dim iVal As Integer
    For iVal = 1 To Screen.FontCount - 1
        cmbFont.AddItem Screen.Fonts(iVal)
    Next
    If cmbFont.ListCount > 0 Then
        cmbFont.Text = cmbFont.List(0)
    End If
    For iVal = 12 To 30
        cmbFontSize.AddItem iVal
    Next
    If cmbFontSize.ListCount > 0 Then
        cmbFontSize.Text = cmbFontSize.List(0)
    End If
End Sub
Private Sub Form_Load()
    Call CenterForm(Me)
    Set CON = OpenDatabase(App.Path & "\PVT", _
        dbDriverComplete, False, ";UID=;PWD=silbus")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CON.Close
End Sub
