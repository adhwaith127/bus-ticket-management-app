VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogo 
   Caption         =   "Logo Setup"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLogoPreview 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Frame fraLogoDescription 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "<<"
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   4680
         Width           =   9375
         Begin VB.Label lblTestDesc 
            BackStyle       =   0  'Transparent
            Caption         =   "File Description"
            Height          =   915
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   9120
         End
      End
      Begin VB.Frame frapboxparent 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   9375
         Begin VB.PictureBox pboxLogoPreview 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5040
            Left            =   840
            ScaleHeight     =   336
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   482
            TabIndex        =   18
            Top             =   360
            Width           =   7230
         End
         Begin VB.Label lblCurrentLogo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logo"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   0
            TabIndex        =   20
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblPreviewAvailable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Preview available"
            Height          =   195
            Left            =   4080
            TabIndex        =   19
            Top             =   3000
            Width           =   1500
         End
      End
      Begin VB.Image imgLogoNext 
         Height          =   375
         Left            =   4200
         Picture         =   "frmLogos.frx":0000
         Stretch         =   -1  'True
         Top             =   5820
         Width           =   375
      End
      Begin VB.Image imgLogoBack 
         Height          =   375
         Left            =   3720
         Picture         =   "frmLogos.frx":04ED
         Stretch         =   -1  'True
         Top             =   5820
         Width           =   375
      End
      Begin VB.Image imgLogoClose 
         Height          =   375
         Left            =   5280
         Picture         =   "frmLogos.frx":09E0
         Stretch         =   -1  'True
         Top             =   5820
         Width           =   375
      End
      Begin VB.Label lblLogoDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Description"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   5400
         Width           =   1185
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   9840
         Y1              =   5640
         Y2              =   5640
      End
   End
   Begin JeweledBut.JeweledButton cmdLogoSave 
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Save Settings"
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
      MICON           =   "frmLogos.frx":0EBB
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox txtLogoPath1 
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
      Left            =   2160
      TabIndex        =   0
      Top             =   2370
      Width           =   6255
   End
   Begin VB.TextBox txtLogoPath2 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   3000
      Width           =   6255
   End
   Begin VB.TextBox txtLogoPath3 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   3570
      Width           =   6255
   End
   Begin VB.TextBox txtLogoPath4 
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
      Left            =   2160
      TabIndex        =   3
      Top             =   4170
      Width           =   6255
   End
   Begin MSComDlg.CommonDialog cmdlgFileName 
      Left            =   8880
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin JeweledBut.JeweledButton cmdLogoBrowse1 
      Height          =   375
      Left            =   8520
      TabIndex        =   5
      Top             =   2370
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmLogos.frx":0ED7
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdLogoBrowse2 
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   2970
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmLogos.frx":0EF3
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdLogoBrowse3 
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   3570
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmLogos.frx":0F0F
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdLogoBrowse4 
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   4170
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmLogos.frx":0F2B
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton JeweledButton1 
      Height          =   375
      Left            =   5400
      TabIndex        =   26
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Exit"
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
      MICON           =   "frmLogos.frx":0F47
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Select monochromatic bmp images"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   25
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label lblLogoPreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   9120
      TabIndex        =   24
      Top             =   4200
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   6420
      Left            =   120
      Top             =   120
      Width           =   9975
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Logo1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   15
      Top             =   2370
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Logo2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Logo3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   3570
      Width           =   1215
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Logo4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Top             =   4170
      Width           =   1215
   End
   Begin VB.Label lblLogoPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   9165
      TabIndex        =   11
      Top             =   2415
      Width           =   570
   End
   Begin VB.Label lblLogoPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   9165
      TabIndex        =   10
      Top             =   3015
      Width           =   570
   End
   Begin VB.Label lblLogoPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   9165
      TabIndex        =   9
      Top             =   3615
      Width           =   570
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum LogoHeader
    LOGO_INFO1 = 1
    LOGO_INFO2 = 2
    LOGO_INFO3 = 3
    LOGO_INFO4 = 4
End Enum
Dim intLogoIndex As Integer
Dim blnCancelLogoBrowse As Boolean

Public DB As DAO.Database
Public RES, res1 As DAO.Recordset

Public RSql As String
Dim rsRecord As DAO.Recordset
Dim lpath As String
Function SaveHardwareSettings(ByVal RecordNo As Long) As Boolean
Dim intHandle As Integer
Dim mycheck As Byte
On Error GoTo CatchError
mycheck = 0
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    Set rsRecord = DB.OpenRecordset("LOGOSETTING", dbOpenDynaset)
   
        If rsRecord.RecordCount > 0 Then
            rsRecord.MoveFirst
            rsRecord.Edit
        Else
           If txtLogoPath1.Text <> "" Or txtLogoPath2.Text <> "" Or txtLogoPath3.Text <> "" Or txtLogoPath4.Text <> "" Then
           rsRecord.AddNew
         Else
            If txtLogoPath1.Text = "" And txtLogoPath2.Text = "" And txtLogoPath3.Text = "" And txtLogoPath4.Text = "" Then
                DB.Execute "DELETE FROM LOGOSETTING"
                 MsgBox "No logo settings for save ", vbInformation, gblstrPrjTitle
                 Exit Function
            End If
         End If
        End If
            
            With rsRecord
             
                If txtLogoPath1.Text <> "" Then
                    
                    If GetLogoInfo(txtLogoPath1.Text) = True Then
                        rsRecord!Logo1 = txtLogoPath1.Text
                        rsRecord!LOGO_INFO1 = strLogo_Info1
                    Else
                        txtLogoPath1.SetFocus
                        Call SendKeys("{HOME}+{END}")
                        Exit Function
                    End If
                    
                Else
                    rsRecord!Logo1 = ""
                    rsRecord!LOGO_INFO1 = ""
                    mycheck = mycheck + 1
                End If
                
                
                If txtLogoPath2.Text <> "" Then
                    
                    If GetLogoInfo(txtLogoPath2.Text) = True Then
                        rsRecord!Logo2 = txtLogoPath2.Text
                        rsRecord!LOGO_INFO2 = strLogo_Info2
                    Else
                        txtLogoPath2.SetFocus
                        Call SendKeys("{HOME}+{END}")
                        Exit Function
                    End If
                    
                Else
                    rsRecord!Logo2 = ""
                    rsRecord!LOGO_INFO2 = ""
                    mycheck = mycheck + 1
                End If
                
                If txtLogoPath3.Text <> "" Then
                    
                    If GetLogoInfo(txtLogoPath3.Text) = True Then
                        rsRecord!Logo3 = txtLogoPath3.Text
                        rsRecord!LOGO_INFO3 = strLogo_Info3
                    Else
                        txtLogoPath3.SetFocus
                        Call SendKeys("{HOME}+{END}")
                        Exit Function
                    End If
                    
                Else
                    rsRecord!Logo3 = ""
                    rsRecord!LOGO_INFO3 = ""
                    mycheck = mycheck + 1
                End If
                
                If txtLogoPath4.Text <> "" Then
                    
                    If GetLogoInfo(txtLogoPath4.Text) = True Then
                        rsRecord!Logo4 = txtLogoPath4.Text
                        rsRecord!LOGO_INFO4 = strLogo_Info4
                    Else
                        txtLogoPath4.SetFocus
                        Call SendKeys("{HOME}+{END}")
                        Exit Function
                    End If
                    
                Else
                    rsRecord!Logo4 = ""
                    rsRecord!LOGO_INFO4 = ""
                    mycheck = mycheck + 1
                End If
                If mycheck = 4 And rsRecord.RecordCount > 0 Then
                    rsRecord.Update
                    rsRecord.Delete
                Else
                    rsRecord.Update
                End If
                
               
                MsgBox "Logo settings saved successfully", vbOKOnly, gblstrPrjTitle
            End With
            Exit Function
CatchError:
    MsgBox "Logo settings  not saved "
   ' If rsRecord.State = adStateOpen Then rsRecord.Close
   '' If intHandle > 0 Then Close #intHandle
End Function


Private Sub cmdLogoSave_Click()
'If txtLogoPath1.Text = "" And txtLogoPath2.Text = "" And txtLogoPath3.Text = "" And txtLogoPath4.Text = "" Then
'Exit Sub
'End If
Call SaveHardwareSettings(1)
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Icon = frmMainform.Icon
    Call AdjustFromPosition(Me)
    Set ConVar = OpenDatabase(App.Path & "\PVT.mdb", _
    dbDriverComplete, False, ";UID=;PWD=silbus")
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    Set rsRecord = DB.OpenRecordset("LOGOSETTING", dbOpenDynaset)
    If rsRecord.RecordCount > 0 Then
        rsRecord.MoveFirst
        strLogo1 = IIf(IsNull(rsRecord!Logo1), "", rsRecord!Logo1)
        
        If strLogo1 <> "" Then
            
            If GetLogoInfo(strLogo1) = True Then
                txtLogoPath1.Text = strLogo1
            Else
                txtLogoPath1.Text = ""
            End If
            
        End If
                
        strLogo2 = IIf(IsNull(rsRecord!Logo2), "", rsRecord!Logo2)
        
        If strLogo2 <> "" Then
        
            If GetLogoInfo(strLogo2) = True Then
                txtLogoPath2.Text = strLogo2
            Else
                txtLogoPath2.Text = ""
            End If
            
        End If
                
        strLogo3 = IIf(IsNull(rsRecord!Logo3), "", rsRecord!Logo3)
        
        If strLogo3 <> "" Then
        
            If GetLogoInfo(strLogo3) = True Then
                txtLogoPath3.Text = strLogo3
            Else
                txtLogoPath3.Text = ""
            End If
             
           
        End If
         strLogo4 = IIf(IsNull(rsRecord!Logo4), "", rsRecord!Logo4)
            
            If strLogo4 <> "" Then
                If GetLogoInfo(strLogo4) = True Then
                    txtLogoPath4.Text = strLogo4
                Else
                    txtLogoPath4.Text = ""
                End If
            End If
           
    End If
            
End Sub

Private Sub JeweledButton1_Click()
    'Me.Hide
    Unload Me
End Sub

Private Sub lblLogoPreview_Click(Index As Integer)
    intLogoIndex = 0
    intLogoIndex = Index + 1
    lblCurrentLogo.caption = "Logo " & intLogoIndex
    Call GeneratePreview(Index)
End Sub

Private Sub pboxLogoPreview_Click()
    If fraLogoDescription.Visible = True Then
        fraLogoDescription.Visible = False
    Else
        fraLogoDescription.Visible = True
    End If
End Sub

Private Sub pboxLogoPreview_Resize()
    If pboxLogoPreview.ScaleWidth > 384 Then
        pboxLogoPreview.Width = 300 * 15
    End If
End Sub


Private Function BrowseLogo() As String
On Error GoTo CatchError
    cmdlgFileName.CancelError = True
    cmdlgFileName.Filter = "Logo Files(*.bmp)|*.bmp;*.bmp"
    cmdlgFileName.InitDir = App.Path
    cmdlgFileName.ShowOpen
    BrowseLogo = cmdlgFileName.filename
    Exit Function
CatchError:
    If err.Number <> 32755 Then
        MsgBox "Cannot open the selected file. Enter qualified path of the file into the textarea.", vbInformation
    End If
    BrowseLogo = cmdlgFileName.filename
    blnCancelLogoBrowse = True
End Function

Private Sub imgLogoBack_Click()
    If intLogoIndex > 1 Then
        lblCurrentLogo.caption = "Logo " & intLogoIndex - 1
        Call GeneratePreview(intLogoIndex - 2)
    End If
End Sub

Private Sub imgLogoNext_Click()
    If intLogoIndex < 4 Then
        lblCurrentLogo.caption = "Logo " & intLogoIndex + 1
        Call GeneratePreview(intLogoIndex)
    End If
End Sub
Private Sub imgLogoClose_Click()
    fraLogoPreview.Visible = False
End Sub
Private Sub cmdLogoBrowse1_Click()
    blnCancelLogoBrowse = False
    lpath = BrowseLogo
    If lpath <> "" Then txtLogoPath1.Text = lpath
    
    If txtLogoPath1.Text <> "" And Me.ActiveControl.Name <> "cmdHSClose" And blnCancelLogoBrowse = False Then
        enmLogoHead = LOGO_INFO1
    '    frmSetHeader.Show vbModal
    End If
End Sub

Private Sub cmdLogoBrowse2_Click()
    blnCancelLogoBrowse = False
    lpath = BrowseLogo
    If lpath <> "" Then txtLogoPath2.Text = lpath
    If txtLogoPath2.Text <> "" And Me.ActiveControl.Name <> "cmdHSClose" And blnCancelLogoBrowse = False Then
        enmLogoHead = LOGO_INFO2
        'frmSetHeader.Show vbModal
    End If
End Sub

Private Sub cmdLogoBrowse3_Click()
    blnCancelLogoBrowse = False
   lpath = BrowseLogo
    If lpath <> "" Then txtLogoPath3.Text = lpath
    If txtLogoPath3.Text <> "" And Me.ActiveControl.Name <> "cmdHSClose" And blnCancelLogoBrowse = False Then
        enmLogoHead = LOGO_INFO3
       ' frmSetHeader.Show vbModal
    End If
End Sub

Private Sub cmdLogoBrowse4_Click()
    blnCancelLogoBrowse = False
   lpath = BrowseLogo
    If lpath <> "" Then txtLogoPath4.Text = lpath
    If txtLogoPath4.Text <> "" And Me.ActiveControl.Name <> "cmdHSClose" And blnCancelLogoBrowse = False Then
        enmLogoHead = LOGO_INFO4
       ' frmSetHeader.Show vbModal
    End If
End Sub
Private Sub GeneratePreview(ByVal Index As Integer)
On Error GoTo CatchError
    fraLogoPreview.Visible = True
    lblPreviewAvailable.Visible = False
    Select Case Index
        Case 0:
            intLogoIndex = 1
            If GetLogoInfo(txtLogoPath1.Text) = True Then
                If Dir(txtLogoPath1.Text, vbNormal) = "" Then
                    GoTo CatchError
                Else
                    pboxLogoPreview.Picture = LoadPicture(txtLogoPath1.Text)
                End If
            Else
                GoTo CatchError
            End If
        Case 1:
            intLogoIndex = 2
            If GetLogoInfo(txtLogoPath2.Text) = True Then
                If Dir(txtLogoPath2.Text, vbNormal) = "" Then
                    GoTo CatchError
                Else
                    pboxLogoPreview.Picture = LoadPicture(txtLogoPath2.Text)
                End If
            Else
                GoTo CatchError
            End If
        Case 2:
            intLogoIndex = 3
            If GetLogoInfo(txtLogoPath3.Text) = True Then
                If Dir(txtLogoPath3.Text, vbNormal) = "" Then
                    GoTo CatchError
                Else
                    pboxLogoPreview.Picture = LoadPicture(txtLogoPath3.Text)
                End If
            Else
                GoTo CatchError
            End If
        Case 3:
            intLogoIndex = 4
            If GetLogoInfo(txtLogoPath4.Text) = True Then
                If Dir(txtLogoPath4.Text, vbNormal) = "" Then
                    GoTo CatchError
                Else
                    pboxLogoPreview.Picture = LoadPicture(txtLogoPath4.Text)
                End If
            Else
                GoTo CatchError
            End If
    End Select
    pboxLogoPreview.Left = frapboxparent.Left + (frapboxparent.Width / 2) - (pboxLogoPreview.Width / 2)
    pboxLogoPreview.Top = frapboxparent.Top + (frapboxparent.Height / 2) - (pboxLogoPreview.Height / 2)
    pboxLogoPreview.Visible = True
    Exit Sub
CatchError:
    pboxLogoPreview.Visible = False
    lblPreviewAvailable.Visible = True
End Sub
Private Function GetLogoInfo(ByVal LogoFile As String) As Boolean
Dim intReadHandle As Integer
Dim udtHead As BMP_HEAD, udtInfo As BMP_INFO
On Error GoTo CatchError
    lblTestDesc.caption = ""
    If Dir(LogoFile, vbNormal) = "" Or Len(LogoFile) <= 0 Then
        lblTestDesc.caption = "No file selected or Selected file is invalid"
       GoTo CatchError
    Else
        intReadHandle = FreeFile
        Open LogoFile For Binary Access Read As #intReadHandle
        Get #intReadHandle, , udtHead
        If EOF(intReadHandle) Then
            MsgBox "Error While reading from logo file", vbInformation, gblstrPrjTitle
            Close #intReadHandle
            lblTestDesc.caption = "Not a valid bitmap file. BMP file header not found."
            GoTo CatchError
        End If
        Get #intReadHandle, , udtInfo
        If EOF(intReadHandle) Then
            lblTestDesc.caption = "Not a valid bitmap file. DIB header not found."
            MsgBox "Error While reading from logo file", vbInformation, gblstrPrjTitle
            Close #intReadHandle
            GoTo CatchError
        End If
        If ((udtInfo.BMPHeight + 1) * 48) / 1024 > 15.5 Then
            lblTestDesc.caption = "Selected file is too large. File size must be less than 15 KB."
            MsgBox "Selected file is too large", vbInformation, gblstrPrjTitle
            Close #intReadHandle
            GoTo CatchError
        End If
        If udtInfo.BitsperPixel <> 1 Then
            lblTestDesc.caption = "Selected file is not a Monochrome Bitmap file"
            MsgBox "Selected file is not a Monochrome Bitmap file.", vbInformation, gblstrPrjTitle
            Close #intReadHandle
            GoTo CatchError
        End If
        If udtInfo.BMPHeight Mod 8 <> 0 Or udtInfo.BMPWidth Mod 8 <> 0 Then
            lblTestDesc.caption = "Dimensions of the selected file must be the multiple of 8."
            MsgBox "Dimensions of the selected file must be the multiple of 8.", vbInformation, gblstrPrjTitle
            Close #intReadHandle
            GoTo CatchError
        End If
    End If
    lblTestDesc.caption = "File Location : " & LogoFile & _
                        Space(5) & "Dimension : " & udtInfo.BMPWidth & " X " & udtInfo.BMPHeight & _
                        Space(5) & "File Size" & Space(5) & Format(udtHead.FileSize / 1024, "0.00") & " KB"
    GetLogoInfo = True
    Exit Function
CatchError:
    If intReadHandle > 0 Then Close #intReadHandle
    lblTestDesc.caption = lblTestDesc.caption & Space(5)
    If err.Number > 0 Then lblTestDesc.caption = lblTestDesc.caption & err.Number & Space(5) & err.Description
End Function


Private Sub ResetLogoInfo(ByRef LogoInfo As ADODB.Recordset)
Dim strLogoInfo As String
Dim intCount As Integer
    If LogoInfo.State = adStateOpen And Not LogoInfo.EOF Then
        strLogoInfo = IIf(IsNull(LogoInfo!LOGO_INFO1), "", LogoInfo!LOGO_INFO1)
        Call SetHeaderCaption(strLogoInfo, 1)
        strLogoInfo = IIf(IsNull(LogoInfo!LOGO_INFO2), "", LogoInfo!LOGO_INFO2)
        Call SetHeaderCaption(strLogoInfo, 2)
        strLogoInfo = IIf(IsNull(LogoInfo!LOGO_INFO3), "", LogoInfo!LOGO_INFO3)
        Call SetHeaderCaption(strLogoInfo, 3)
        strLogoInfo = IIf(IsNull(LogoInfo!LOGO_INFO4), "", LogoInfo!LOGO_INFO4)
        Call SetHeaderCaption(strLogoInfo, 4)
    End If
End Sub

Private Sub SetHeaderCaption(ByVal caption As String, Index As Integer)
On Error Resume Next
    Select Case caption
        Case "H1":
            txtHeader1.Text = "LOGO" & Index
        Case "H2":
            txtHeader2.Text = "LOGO" & Index
        Case "H3":
            txtHeader3.Text = "LOGO" & Index
        Case "H4":
            txtHeader4.Text = "LOGO" & Index
        Case "F1":
            txtFooter1.Text = "LOGO" & Index
        Case "F2":
            txtFooter2.Text = "LOGO" & Index
    End Select
End Sub

Sub AdjustFromPosition(ByRef objForm As Form)
On Error Resume Next
    If objForm.Name <> "frmLogin" And objForm.Name <> "frmMain" And objForm.Name <> "frmGetServer" Then
        objForm.Icon = frmMain.Icon
    End If
    objForm.Left = (Screen.Width / 2) - (objForm.Width / 2)
    objForm.Top = (Screen.Height / 2) - (objForm.Height / 2)
End Sub
''Private Function GetLogoInfo(ByVal LogoFile As String) As Boolean
''Dim intReadHandle As Integer
''Dim udtHead As BMP_HEAD, udtInfo As BMP_INFO
''On Error GoTo CatchError
''    lblTestDesc.Caption = ""
''    If Dir(LogoFile, vbNormal) = "" Then
''        lblTestDesc.Caption = "No file selected or Selected file is invalid"
''       GoTo CatchError
''    Else
''        intReadHandle = FreeFile
''        Open LogoFile For Binary Access Read As #intReadHandle
''        Get #intReadHandle, , udtHead
''        If EOF(intReadHandle) Then
''            MsgBox "Error While reading from logo file", vbInformation, gblstrPrjTitle
''            Close #intReadHandle
''            lblTestDesc.Caption = "Not a valid bitmap file. BMP file header not found."
''            GoTo CatchError
''        End If
''        Get #intReadHandle, , udtInfo
''        If EOF(intReadHandle) Then
''            lblTestDesc.Caption = "Not a valid bitmap file. DIB header not found."
''            MsgBox "Error While reading from logo file", vbInformation, gblstrPrjTitle
''            Close #intReadHandle
''            GoTo CatchError
''        End If
''        If ((udtInfo.BMPHeight + 1) * 48) / 1024 > 15 Then
''            lblTestDesc.Caption = "Selected file is too large. File size must be less than 15 KB."
''            MsgBox "Selected file is too large", vbInformation, gblstrPrjTitle
''            Close #intReadHandle
''            GoTo CatchError
''        End If
''        If udtInfo.BitsperPixel <> 1 Then
''            lblTestDesc.Caption = "Selected file is not a Monochrome Bitmap file"
''            MsgBox "Selected file is not a Monochrome Bitmap file.", vbInformation, gblstrPrjTitle
''            Close #intReadHandle
''            GoTo CatchError
''        End If
''        If udtInfo.BMPHeight Mod 8 <> 0 And udtInfo.BMPWidth Mod 8 <> 0 Then
''            lblTestDesc.Caption = "Dimensions of the selected file must be the multiple of 8."
''            MsgBox "Dimensions of the selected file must be the multiple of 8.", vbInformation, gblstrPrjTitle
''            Close #intReadHandle
''            GoTo CatchError
''        End If
''    End If
''    lblTestDesc.Caption = "File Location : " & LogoFile & _
''                        Space(5) & "Dimension : " & udtInfo.BMPWidth & " X " & udtInfo.BMPHeight & _
''                        Space(5) & "File Size" & Space(5) & Format(udtHead.FileSize / 1024, "0.00") & " KB"
''    GetLogoInfo = True
''    Exit Function
''CatchError:
''    If intReadHandle > 0 Then Close #intReadHandle
''    lblTestDesc.Caption = lblTestDesc.Caption & Space(5)
''    If err.Number > 0 Then lblTestDesc.Caption = lblTestDesc.Caption & err.Number & Space(5) & err.Description
''End Function
