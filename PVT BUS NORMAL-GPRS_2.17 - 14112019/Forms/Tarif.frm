VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form Tarif_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FARE  "
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6375
   Begin VB.Frame Frm1 
      Height          =   855
      Left            =   360
      TabIndex        =   18
      Top             =   960
      Width           =   5295
      Begin VB.TextBox from_txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox To_txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Amount_txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   1920
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "fare"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.ComboBox cmb_Vtype 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
   Begin VB.PictureBox FormSize_Pctr 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   120
      ScaleHeight     =   3180
      ScaleWidth      =   2625
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2655
      Begin VB.TextBox slab_txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox chk_night 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Night Tariff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Close_Cmd 
         BackColor       =   &H80000009&
         Height          =   495
         Left            =   8400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Tarif.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid card_FlxGrid1 
         DragMode        =   1  'Automatic
         Height          =   255
         Left            =   720
         TabIndex        =   8
         ToolTipText     =   "Double Click to Edit. Press Delete button for Deleting"
         Top             =   5400
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         FixedCols       =   0
         RowHeightMin    =   315
         BackColorFixed  =   -2147483645
         ForeColorFixed  =   0
         BackColorSel    =   -2147483646
         ForeColorSel    =   16777215
         BackColorBkg    =   14737632
         GridColor       =   0
         GridColorFixed  =   0
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total slab         hours"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   6120
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Msg_Lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   2280
         TabIndex        =   9
         Top             =   2400
         Width           =   3885
      End
   End
   Begin JeweledBut.JeweledButton Clear_cmd 
      Height          =   315
      Left            =   2040
      TabIndex        =   14
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      TX              =   "Clea&r"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Tarif.frx":263A
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton save_cmd 
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      TX              =   "&Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Tarif.frx":2656
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton Search_cmd 
      Height          =   315
      Left            =   4440
      TabIndex        =   15
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      TX              =   "S&earch"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Tarif.frx":2672
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton Delete_cmd 
      Height          =   315
      Left            =   840
      TabIndex        =   16
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      TX              =   "&Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Tarif.frx":268E
      BC              =   12632256
      FC              =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid item_FlxGrid 
      Height          =   4590
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Double Click to Edit. Press Delete button for Deleting"
      Top             =   2520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8096
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorFixed  =   -2147483645
      ForeColorFixed  =   0
      BackColorSel    =   -2147483646
      ForeColorSel    =   16777215
      BackColorBkg    =   14737632
      GridColor       =   0
      GridColorFixed  =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bus Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   600
      TabIndex        =   13
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label FormName_Lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Form Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "Tarif_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RecSet As New ADODB.Recordset
Public DB_ADODB As New ADODB.Connection
Private Type FARESLAB
    SLNO As Long
    BUSID As Long
    STARTKM As Single
    ENDKM As Single
    FARE As Single
End Type
Dim FARESLAB_pk As Long
Private Sub chk_Enable_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    save_cmd.SetFocus
End If
End Sub
Private Sub Amount_txt_KeyPress(KeyAscii As Integer)
    ValidationMode = FloatingPointValue
    ValidateKeyPress Amount_txt, KeyAscii
End Sub

Private Sub chk_night_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    save_cmd.SetFocus
End If

End Sub

Private Sub Clear_cmd_Click()
Clear
cmb_Vtype.ListIndex = 0
ShowActivity
End Sub
Private Sub Clear()

from_txt = ""
To_txt = ""
slab_txt = ""
Amount_txt = ""
FARESLAB_pk = 0
chk_night.Value = 0
'cardno_Txt.SetFocus
End Sub
Private Sub Close_Cmd_Click()
Unload Me
End Sub



Private Sub cmb_Vtype_Click()
'slab_txt = val(getvalueQuery("select dmrc_slab from dmrc_vehicle_tab where dmrc_vehicle_status=0 and dmrc_vehicle_pk=" & cmb_Vtype.ItemData(cmb_Vtype.ListIndex) & ""))

End Sub

Private Sub cmb_Vtype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    from_txt.SetFocus
End If
End Sub

Private Sub Delete_cmd_Click()
On Error GoTo erromod
    If IsNumeric(item_FlxGrid.TextMatrix(item_FlxGrid.row, 0)) Then
        If MsgBox("Do you want to delete the selected Fare ?", vbYesNo + vbQuestion + vbDefaultButton2, gblstrPrjTitle) = vbYes Then
            If chkdelete(item_FlxGrid.TextMatrix(item_FlxGrid.row, 0)) = "" Then
                DeleteActivity CLng(item_FlxGrid.TextMatrix(item_FlxGrid.row, 0))
                'seteditflag
             Else
                MsgBox "Could not Remove. Fare Details Exists in (" & chkdelete(item_FlxGrid.TextMatrix(item_FlxGrid.row, 0)) & ")", vbExclamation, gblstrPrjTitle
                Exit Sub
             End If
            Clear
            ShowActivity
        End If
    Else
        MsgBox "Select a Fare from the list.", vbExclamation, gblstrPrjTitle
        item_FlxGrid.SetFocus
    End If

Exit Sub
erromod:
MsgBox "Error occured Due to " & err.Description, vbExclamation, gblstrPrjTitle

End Sub

Private Sub DeleteActivity(dmrc_tariff_pk As Long)
On Error GoTo erromod

    If dmrc_tariff_pk > 0 Then
        sql = " delete from FARESLAB where SLNO=" & dmrc_tariff_pk
        DB_ADODB.Execute (sql)
        MsgBox "Fare Deleted successfully ", vbInformation, gblstrPrjTitle
    End If

Exit Sub
erromod:
MsgBox "Error occured Due to " & err.Description, vbExclamation, gblstrPrjTitle
End Sub


Private Function chkdelete(ByVal chkid As Double) As String
On Error GoTo lblErr
    chkdelete = ""
    'If val(getvalueQuery("SELECT COUNT(shift_id) FROM cms_shift WHERE shift_isdelete<>" & INVALID_STATUS & " AND shift_Activity=" & chkid)) > 0 Then
      'chkdelete = "Shit Details"
    'End If
Exit Function
lblErr:
End Function

Private Sub Form_Load()
On Error Resume Next
     Me.Icon = frmMainform.Icon
    FormName_Lbl = UCase(Me.caption)
   ' Bus_tariff_pk = 0
    ShowActivity
    FillBusTypeName
    Call Clear_cmd_Click
    Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
End Sub

Private Sub Form_Resize()
On Error Resume Next
If WindowState <> 1 And WindowState <> 2 Then
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = ((Screen.Height - 1200) / 2) - (Me.Height / 2)
End If
End Sub

Private Sub ShowActivity()

On Error GoTo ErrorMod
Dim RecSet As New ADODB.Recordset
Dim sql As String

setUpHeaderForGrid

If DB_ADODB.State <> 0 Then DB_ADODB.Close
                      
DB_ADODB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source= " & sDataBase & "; Jet OLEDB:Database Password =silbus"


sql = "SELECT STARTKM,ENDKM,SLNO,BUSID,FARE,name as type  FROM FARESLAB f,bustype b where b.id=f.BUSID "

If cmb_Vtype.ListIndex > 0 Then
    sql = sql & " and BUSID = " & cmb_Vtype.ItemData(cmb_Vtype.ListIndex) & ""
End If

sql = sql & " ORDER BY  b.name,STARTKM"
    
Set RecSet = DB_ADODB.Execute(sql)

With item_FlxGrid


If RecSet.EOF = True Then
    Msg_Lbl = "No Fare details Found"
    Exit Sub
Else
    
    Msg_Lbl = ""
    While Not RecSet.EOF
        
        .TextMatrix(.Rows - 1, 1) = .Rows - 1 'IIf(IsNumeric(RecSet("SLNO")), RecSet("SLNO"), 0)
        .TextMatrix(.Rows - 1, 0) = IIf(IsNumeric(RecSet("SLNO")), RecSet("SLNO"), 0)
        .TextMatrix(.Rows - 1, 2) = IIf(IsNull(RecSet("TYPE")), "", RecSet("TYPE"))
        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(RecSet("STARTKM")), "", RecSet("STARTKM"))
        .TextMatrix(.Rows - 1, 4) = IIf(IsNull(RecSet("ENDKM")), "", RecSet("ENDKM"))
        .TextMatrix(.Rows - 1, 5) = IIf(IsNull(RecSet("FARE")), "", RecSet("FARE"))
      
        .Rows = .Rows + 1
        RecSet.MoveNext
      
    Wend
End If
End With

Exit Sub
ErrorMod:
MsgBox "Error Occured due to " & err.Description, vbExclamation, gblstrPrjTitle
End Sub


Private Sub setUpHeaderForGrid()
On Error Resume Next

With item_FlxGrid
    .Clear
    .Rows = 2
    .Cols = 6
    
    .ColWidth(0) = 0
    .ColWidth(1) = 800
    .ColWidth(2) = 2000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1500

  
     
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignLeftCenter
    .ColAlignment(5) = flexAlignLeftCenter



    .TextMatrix(0, 0) = "SLNo."
    .TextMatrix(0, 1) = "SLNo."
    .TextMatrix(0, 2) = "TYPE"
    .TextMatrix(0, 3) = "START"
    .TextMatrix(0, 4) = "END"
    .TextMatrix(0, 5) = "FARE"
 
   
End With
End Sub


Private Sub from_txt_KeyPress(KeyAscii As Integer)
    ValidationMode = FloatingPointValue
    ValidateKeyPress from_txt, KeyAscii
End Sub

Private Sub item_FlxGrid_DblClick()
If item_FlxGrid.row > 0 And IsNumeric(item_FlxGrid.TextMatrix(item_FlxGrid.row, 0)) Then
        FARESLAB_pk = item_FlxGrid.TextMatrix(item_FlxGrid.row, 0)
        If FARESLAB_pk > 0 Then
            Loadtariff
           ' Loadroute
        End If
    End If
End Sub


Private Sub Loadtariff()
On Error GoTo erromod

sql = "SELECT SLNO,BUSID,STARTKM,ENDKM,FARE FROM FARESLAB WHERE SLNO=" & FARESLAB_pk
Set rs = DB_ADODB.Execute(sql)
If rs.EOF = True Then
    MsgBox "No such distence Details saved.", vbInformation, gblstrPrjTitle
    Exit Sub
Else
    While Not rs.EOF
        FARESLAB_pk = IIf(IsNumeric(rs("SLNO")), rs("SLNO"), 0)
        selectMatchDropdown cmb_Vtype, rs("BUSID")
        from_txt = IIf(IsNull(rs("STARTKM")), "", rs("STARTKM"))
        To_txt = IIf(IsNull(rs("ENDKM")), "", rs("ENDKM"))
        Amount_txt = IIf(IsNull(rs("FARE")), "", rs("FARE"))
        'chk_night.Value = IIf(IsNumeric(rs("dmrc_night")), rs("dmrc_night"), 0)
        rs.MoveNext
    Wend
End If
Exit Sub
erromod:
MsgBox "Error occured Due to " & err.Description, vbExclamation, gblstrPrjTitle
End Sub


Private Sub FillBusTypeName()
Dim RES As ADODB.Recordset
   
    TSQL = ""
    TSQL = "SELECT * FROM BUSTYPE"
    Set RES = DB_ADODB.Execute(TSQL)
    If Not RES.EOF Then
        RES.MoveFirst
        cmb_Vtype.Clear
        cmb_Vtype.AddItem "Select"
        cmb_Vtype.ItemData(cmb_Vtype.NewIndex) = 0
        Do While Not RES.EOF
            cmb_Vtype.AddItem (RES("Name"))
            cmb_Vtype.ItemData(cmb_Vtype.NewIndex) = RES("Id")
            RES.MoveNext
            Loop
    End If
End Sub

Private Sub save_cmd_Click()
On Error GoTo erromod
Dim mytime As String
Dim TariffTable As FARESLAB

If Trim(cmb_Vtype.Text) = "Select" Then
    MsgBox "Bus Type can't be empty.", vbExclamation, gblstrPrjTitle
    cmb_Vtype.SetFocus
    Exit Sub
End If
If (from_txt) = "" Then
    MsgBox "From Km  can't be empty.", vbExclamation, gblstrPrjTitle
    from_txt.SetFocus
    Exit Sub

End If
If (To_txt) = "" Then
    MsgBox "To Km  can't be empty.", vbExclamation, gblstrPrjTitle
    To_txt.SetFocus
    Exit Sub
End If
If val(To_txt) <= 0 Then
    MsgBox "To Km must be greater than zero.", vbExclamation, gblstrPrjTitle
    To_txt.SetFocus
    Exit Sub

End If
If val(To_txt) <= val(from_txt) Then
    MsgBox "To Km must be greater than from Km.", vbExclamation, gblstrPrjTitle
    To_txt.SetFocus
    Exit Sub
End If

TariffTable.SLNO = FARESLAB_pk
TariffTable.BUSID = cmb_Vtype.ItemData(cmb_Vtype.ListIndex)
TariffTable.STARTKM = val(from_txt)
TariffTable.ENDKM = val(To_txt)
TariffTable.FARE = val(Amount_txt.Text)

If TariffTable.SLNO > 0 Then
     If val(getvalueQuery("select count(*) from FARESLAB where (" & from_txt & " between STARTKM and ENDKM and slno <>" & FARESLAB_pk & " and BUSID=" & cmb_Vtype.ItemData(cmb_Vtype.ListIndex) & ") or (" & To_txt & " between STARTKM and ENDKM and slno <>" & FARESLAB_pk & " and BUSID=" & cmb_Vtype.ItemData(cmb_Vtype.ListIndex) & ")")) > 0 Then
        MsgBox "Distance already covered."
        Exit Sub
    ElseIf val(getvalueQuery("select * from FARESLAB where ENDKM between " & from_txt & " and " & To_txt & " and slno <>" & FARESLAB_pk * " and BUSID=" & cmb_Vtype.ItemData(cmb_Vtype.ListIndex))) > 0 Then
        MsgBox "Distance already covered."
        Exit Sub
    End If


'    If val(getvalueQuery("select count(SLNO) from FARESLAB where ENDKM>=" & val(from_txt) & " and SLNO<>" & FARESLAB_pk & " and BUSID=" & cmb_Vtype.ItemData(cmb_Vtype.ListIndex))) > 0 Then
'        MsgBox "Distance already covered."
'        Exit Sub
'    End If
    If UpdateTariffTable(TariffTable) Then   'Insering into table
        MsgBox "Fare Details edited successfully ", vbInformation, gblstrPrjTitle
        'seteditflag
    End If
ElseIf TariffTable.SLNO = 0 Then
     If val(getvalueQuery("select count(*) from FARESLAB where (" & from_txt & " between STARTKM and ENDKM  and BUSID=" & cmb_Vtype.ItemData(cmb_Vtype.ListIndex) & ") or (" & To_txt & " between STARTKM and ENDKM and BUSID=" & cmb_Vtype.ItemData(cmb_Vtype.ListIndex) & ")")) > 0 Then
        MsgBox "Distance already covered."
        Exit Sub
    ElseIf val(getvalueQuery("select * from FARESLAB where ENDKM between " & from_txt & " and " & To_txt & " and BUSID=" & cmb_Vtype.ItemData(cmb_Vtype.ListIndex))) > 0 Then
        MsgBox "Distance already covered."
        Exit Sub
    End If
'    If val(getvalueQuery("select count(SLNO) from FARESLAB where ENDKM>=" & val(from_txt) & " and BUSID=" & cmb_Vtype.ItemData(cmb_Vtype.ListIndex))) > 0 Then
'        MsgBox "Distance already covered."
'        Exit Sub
'    End If
    If InsertIntoTariffTable(TariffTable) Then 'Updating Table
        MsgBox "Fare Details saved successfully ", vbInformation, gblstrPrjTitle
        'seteditflag
    End If
End If

Clear

ShowActivity

cmb_Vtype.SetFocus
Exit Sub

erromod:

MsgBox "Error occured Due to " & err.Description, vbExclamation, gblstrPrjTitle
End Sub

Private Function InsertIntoTariffTable(TariffTableREF As FARESLAB) As Boolean
On Error GoTo ErrorMod

    sql = "INSERT INTO FARESLAB " _
    & "(BUSID,STARTKM,ENDKM,FARE) VALUES (" _
    & TariffTableREF.BUSID & ",'" & TariffTableREF.STARTKM & "','" & TariffTableREF.ENDKM & "','" & TariffTableREF.FARE & "')"
    
    DB_ADODB.Execute sql
    InsertIntoTariffTable = True
    
Exit Function
ErrorMod:
InsertIntoTariffTable = False

MsgBox "Error occured Due to " & err.Description, vbExclamation, gblstrPrjTitle
End Function


Private Function UpdateTariffTable(TariffTableREF As FARESLAB) As Boolean
On Error GoTo ErrorMod
    
    sql = "UPDATE FARESLAB SET " _
    & "BUSID='" & TariffTableREF.BUSID & "',STARTKM='" & TariffTableREF.STARTKM & "',ENDKM='" & TariffTableREF.ENDKM _
    & "',FARE ='" & TariffTableREF.FARE _
    & "' WHERE SLNO=" & FARESLAB_pk
    
DB_ADODB.Execute sql
UpdateTariffTable = True
Exit Function

ErrorMod:
UpdateTariffTable = False
MsgBox "Error occured Due to :  " & err.Description, vbExclamation, gblstrPrjTitle


End Function


Private Sub Search_cmd_Click()
On Error Resume Next
    ShowActivity
    cmb_Vtype.SetFocus
End Sub

Private Sub To_txt_KeyPress(KeyAscii As Integer)
    ValidationMode = FloatingPointValue
    ValidateKeyPress To_txt, KeyAscii
End Sub
