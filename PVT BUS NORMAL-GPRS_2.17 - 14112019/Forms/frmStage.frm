VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmStage 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stages Entry"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   Icon            =   "frmStage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6690
      Left            =   30
      TabIndex        =   7
      Top             =   -15
      Width           =   11880
      Begin JeweledBut.JeweledButton cmdCancel 
         Height          =   375
         Left            =   10470
         TabIndex        =   6
         Top             =   5835
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
         MICON           =   "frmStage.frx":0CCA
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSave 
         Height          =   375
         Left            =   9195
         TabIndex        =   5
         Top             =   5835
         Width           =   1200
         _ExtentX        =   2117
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
         MICON           =   "frmStage.frx":0CE6
         BC              =   12632256
         FC              =   0
      End
      Begin VB.TextBox txtLanguage 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   9270
         MaxLength       =   23
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2640
         Width           =   2265
      End
      Begin VB.TextBox txtDistance 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9270
         MaxLength       =   8
         TabIndex        =   4
         Top             =   3720
         Width           =   1665
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmStage.frx":0D02
         Left            =   1410
         List            =   "frmStage.frx":0D04
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   1320
      End
      Begin VB.TextBox txtStage 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9270
         MaxLength       =   11
         TabIndex        =   2
         Top             =   2145
         Width           =   1665
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlex 
         Height          =   5760
         Left            =   90
         TabIndex        =   1
         Top             =   765
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   10160
         _Version        =   393216
         BackColorFixed  =   12632256
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLanguage 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Local Language"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   8190
         TabIndex        =   12
         Top             =   2805
         Width           =   1095
      End
      Begin VB.Label lblDistance 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Distance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8190
         TabIndex        =   11
         Top             =   3765
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "StageName"
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
         Left            =   8190
         TabIndex        =   9
         Top             =   2175
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route Code"
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
         Left            =   165
         TabIndex        =   8
         Top             =   345
         Width           =   1110
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stages Entry"
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
      Left            =   4980
      TabIndex        =   10
      Top             =   -540
      Width           =   2325
   End
End
Attribute VB_Name = "frmStage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As DAO.Database
Dim Id As Integer
Dim Edit As Boolean
Dim DOT_FLAG As Boolean
Private Sub cmdSave_Click()
Dim rs As DAO.Recordset
Dim rsRd As DAO.Recordset
Dim sql As String
Dim Buf As String
Dim SqlRd As String
Dim filehd As Integer
Dim Flag As Boolean
Dim myFlag As Boolean
On Error GoTo err
    Flag = False
    myFlag = False
    If Edit = True Then
       If txtStage.Text <> "" Or txtDistance <> "" Then
         Set cn = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
         sql = "Select * from stage where id=" & Id
         Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
         If rs.RecordCount > 0 Then
             If txtStage <> "" Then
                cn.Execute "update stage set stagename='" & Trim(UCase(txtStage.Text)) & "' where id=" & rs!Id & " and route ='" & rs!Route & "'"
                myFlag = True
             End If
             'If val(txtDistance) <> 0 Then
                cn.Execute "update stage set DISTANCE=" & val(txtDistance.Text) & " where id=" & rs!Id & " and route ='" & rs!Route & "'"
                myFlag = True
            ' End If
             
             If txtLanguage <> "" Then
                cn.Execute "update stage set STG_LOCAL_LANGUAGE='" & Replace(Trim((txtLanguage.Text)), "'", "''") & "' where id=" & rs!Id & " and route ='" & rs!Route & "'"   ''' 20/01/2011
                myFlag = True
             End If
             If myFlag = True Then
                MsgBox " Stage  saved", vbInformation, App.ProductName
             End If
         End If
         rs.Close
       End If
    End If
    txtStage.Text = ""
    txtDistance = ""
    txtLanguage = ""
    'Combo1.Text = ""
    Edit = False
    Grid
    If Combo1.Text = "" Then
        MsgBox "No route code selected!", vbExclamation, gblstrPrjTitle
        Exit Sub
    End If
    Call Combo1_Click
    Exit Sub
err:
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
Dim rs As DAO.Recordset
Dim sql As String
Dim I As Integer
On Error GoTo err
    I = 1
    Set cn = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    sql = "Select * from stage where route='" & Combo1.Text & "' order by id"
    Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            With MSFlex
                .Rows = I + 1
                .TextMatrix(I, 0) = rs.Fields!Id
                 .TextMatrix(I, 1) = rs.Fields!StageName
                 .TextMatrix(I, 2) = rs.Fields!Distance
'
               
                If LocalLanguage > 0 Then
                    If rs.Fields!STG_LOCAL_LANGUAGE <> "" Then     ''' 20/01/2011 vaisakh
                        .TextMatrix(I, 3) = IIf(rs.Fields!STG_LOCAL_LANGUAGE = "20-20-20-20", "", rs.Fields!STG_LOCAL_LANGUAGE)
                        If LocalLanguage > 0 Then 'LANG
                            .Col = 3
                            .row = I
                            .CellFontSize = 12
                            If LocalLanguage = 1 Then .CellFontName = "senthamil"
                            If LocalLanguage = 2 Then .CellFontName = "ML-TTKarthika"
                  
                        End If
                    End If
                 End If

'''                    If rs.Fields!BmpFile <> "" Then  ''' 20/01/2011
'''                        .TextMatrix(i, 3) = IIf(IsNull(rs.Fields!BmpFile), "", rs.Fields!BmpFile)
'''                    End If
                'End If
            End With
            I = I + 1
            rs.MoveNext
        Loop
    End If
    rs.Close
    Exit Sub
err:
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim rs As DAO.Recordset
Dim sql As String
Dim I As Integer
On Error GoTo err
    If KeyAscii = 13 Then
        I = 1
        Set cn = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
        sql = "Select * from stage where route='" & Combo1.Text & "' order by id"
        Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do While Not rs.EOF
                With MSFlex
                    .Rows = I + 1
                    .TextMatrix(I, 0) = rs.Fields!Id
                    .TextMatrix(I, 1) = CStr(rs.Fields!StageName)
                    .TextMatrix(I, 2) = rs.Fields!Distance
                    If LocalLanguage > 0 Then
                        .TextMatrix(I, 3) = rs.Fields!STG_LOCAL_LANGUAGE   ''' 20/01/2011
'''                        .TextMatrix(i, 3) = IIf(IsNull(rs.Fields!BmpFile), "", rs.Fields!BmpFile) ''' 20/01/2011
                    End If
                End With
                I = I + 1
                rs.MoveNext
            Loop
        End If
        rs.Close
    End If
    Exit Sub
err:
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Me.Icon = frmMainform.Icon
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
    If RouteID <> "" Then
         Combo1.Text = RouteID
    End If
    'Combo1.SetFocus
    DOT_FLAG = False
End Sub

Private Sub Form_Load()
Dim sql As String
Dim rs As DAO.Recordset
On Error GoTo err
    If RouteID = "" Then
        sql = "select rutcode from route"
        Set cn = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
        sql = "SELECT DISTINCT Route FROM stage "
        Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
        If rs.RecordCount <> 0 Then rs.MoveFirst
        With rs
            Do While Not .EOF
                Combo1.AddItem .Fields(0)
                .MoveNext
            Loop
        End With
        rs.Close
        RouteID = Combo1.List(0)
     End If

    If LocalLanguage = 0 Then '
        txtLanguage.Visible = False
        lblLanguage.Visible = False
        
        txtDistance.Top = 2640
        lblDistance.Top = 2675
    Else
        txtLanguage.Visible = True
        lblLanguage.Visible = True
    
        txtDistance.Top = 3720
        lblDistance.Top = 3765
    End If
    Call Grid
    Exit Sub
err:
End Sub

Private Sub MSFlex_Click()
On Error Resume Next
    If MSFlex.row > 0 Then
        Edit = True
        Id = MSFlex.TextMatrix(MSFlex.row, 0)
        txtStage = MSFlex.TextMatrix(MSFlex.row, 1)
        txtDistance = MSFlex.TextMatrix(MSFlex.row, 2)
'LAN   '
    If LocalLanguage > 0 Then
            txtLanguage = MSFlex.TextMatrix(MSFlex.row, 3)
             If LocalLanguage = 1 Then
                txtLanguage.FontName = "senthamil"
            ElseIf LocalLanguage = 2 Then
                txtLanguage.FontName = "ML-TTKarthika"
            End If
        End If
    End If
End Sub



Private Sub Grid()
On Error Resume Next
    With MSFlex
        If LocalLanguage > 0 Then '
            .Cols = 4
            .FormatString = "ID|Stage Name| KM |STAGE IN "  '' 20/01/2011
'''            .FormatString = "ID|Stage Name| KM |BMP File Name "   '' 20/01/2011
        Else
            .FormatString = "ID|Stage Name| KM "
            .Cols = 3
        End If
        .Rows = 1
        .Width = 8000
        .ColWidth(0) = 500
        .ColWidth(1) = 3000
        .ColWidth(2) = 800
        
        If LocalLanguage > 0 Then 'LAN
            .ColWidth(3) = 3400
        End If
    End With
End Sub



Private Sub txtDistance_GotFocus()
    txtDistance.BackColor = &HC0EF00
End Sub

Private Sub txtDistance_KeyPress(KeyAscii As Integer)
On Error Resume Next

    Static LastText As String
    Static SecondTime As Boolean
    Const MaxDecimal As Integer = 1
    Const MaxWhole As Integer = 4
  
    If TextBoxValidityNumeric(KeyAscii) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    With txtDistance
        If Not SecondTime Then
            If .Text Like "*[!0-9.]*" _
            Or .Text Like "*.*.*" _
            Or .Text Like "*." & String$(1 + MaxDecimal, "#") _
            Or .Text Like String$(MaxWhole, "#") & "[!.]" Then
                SecondTime = True
                .Text = LastText
                .SelStart = Len(.Text)
            Else
                LastText = .Text
            End If
        End If
    End With
    SecondTime = False
    If KeyAscii = 13 Then cmdSave_Click
End Sub

Private Sub txtDistance_LostFocus()
On Error Resume Next
    txtDistance.BackColor = &H80000005
End Sub

Private Sub txtLanguage_DblClick()
On Error Resume Next
    Editing = True
    RouteID = Combo1.Text
    txtLanguage_KeyPress (13)
End Sub

Private Sub txtLanguage_GotFocus()
On Error Resume Next
    strBmpName = ""    '''' 20/01/2011
    strBmpName = Trim(txtStage.Text)
    txtLanguage.BackColor = &HC0EF00
End Sub


Private Sub txtLanguage_KeyPress(KeyAscii As Integer)   ''' 20/01/2011  Original
On Error Resume Next
    If KeyAscii = 13 Then
        txtDistance.SetFocus
        'KeyAscii = 0
'        If PROJECT = SRILANKA Then
'            strLanguageStage = txtLanguage
'            Load H_Convert
'            H_Convert.Show vbModal
'            txtLanguage = strLanguageStage
'            strLanguageStage = ""
'            cmdSave_Click
'        End If
    End If
End Sub

'''Private Sub txtLanguage_KeyPress(KeyAscii As Integer)
'''On Error Resume Next
'''    If KeyAscii = 13 Then
'''        strBmpName = ""    '''' 20/01/2011
'''        strBmpName = Trim(txtStage.Text)
'''        KeyAscii = 0
'''        If PROJECT = SRILANKA Then
'''            strLanguageStage = txtLanguage
'''            txtLanguage.Text = ""
'''            gblBMPName = ""
'''            frmBmpEntry.Show vbModal
'''            If gblBMPName = "" Then
'''                gblBMPName = "0"
'''            End If
'''            txtLanguage = gblBMPName
'''            strLanguageStage = ""
'''            gblBMPName = ""
'''            cmdSave_Click
'''        End If
'''    End If
'''End Sub

Private Sub txtLanguage_LostFocus()
On Error Resume Next
    txtLanguage.BackColor = &H80000005
End Sub


Private Sub txtStage_GotFocus()
    txtStage.BackColor = &HC0EF00
End Sub


Private Sub txtStage_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        If txtLanguage.Visible = True Then
            txtLanguage.SetFocus
        Else
            txtDistance.SetFocus
        End If
    Exit Sub
    End If 'cmdSave_Click
End Sub

Private Sub txtStage_LostFocus()
    txtStage.BackColor = &H80000005
    strBmpName = ""    '''' 20/01/2011
    strBmpName = Trim(txtStage.Text)
End Sub
