VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmAddBusType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bus Type"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar stbarMassage 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4080
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11271
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7935
      TabIndex        =   8
      Top             =   3585
      Width           =   7935
      Begin JeweledBut.JeweledButton cmdClose 
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
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
         MICON           =   "frmAddBusRoute.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdRemove 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   661
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
         MICON           =   "frmAddBusRoute.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSave 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
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
         MICON           =   "frmAddBusRoute.frx":0038
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdClear 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   661
         TX              =   "&Clear"
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
         MICON           =   "frmAddBusRoute.frx":0054
         BC              =   12632256
         FC              =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3900
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Width           =   8055
      Begin VB.TextBox txtbustype 
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
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   0
         Top             =   480
         Width           =   4575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshflxBustype 
         Height          =   2490
         Left            =   120
         TabIndex        =   1
         Top             =   1140
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   4392
         _Version        =   393216
         ForeColor       =   0
         Rows            =   9
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   15846563
         ForeColorFixed  =   0
         BackColorSel    =   16247257
         ForeColorSel    =   4210752
         GridColorFixed  =   0
         GridColorUnpopulated=   0
         TextStyle       =   3
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         SelectionMode   =   1
         AllowUserResizing=   3
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
         FontWidth       =   4
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontWidthFixed  =   5
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).TextStyleBand=   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bus Type"
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
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmAddBusType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intBustypeID As Integer
Dim blnEdit As Boolean
Private Sub cmdClear_Click()
On Error Resume Next
    txtbustype.Text = ""
    blnEdit = False
    cmdRemove.Enabled = False
End Sub
Private Sub cmdClear_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to clear the window"
End Sub
Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to clear the window"
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to close the window"
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to close the window"
End Sub

Private Sub cmdRemove_Click()
Dim strbustype As String
Dim rsbustype As ADODB.Recordset
Dim intBustypeID As Integer
On Error GoTo CatchError
    If mshflxBustype.TextMatrix(mshflxBustype.row, 0) = "" Or mshflxBustype.row = 0 Then Exit Sub
    strbustype = mshflxBustype.TextMatrix(mshflxBustype.row, 2)
    intBustypeID = mshflxBustype.TextMatrix(mshflxBustype.row, 0)
    If (MsgBox("Do you want to remove Bustype " & strbustype & " from Bustype details ?", vbQuestion + vbYesNo)) = vbYes Then
        gbladoCon.Execute "DELETE * FROM [bustype] WHERE id=" & intBustypeID
        Set rsbustype = New ADODB.Recordset
        rsbustype.Open "SELECT * FROM [bustype] WHERE [id]=" & intBustypeID, gbladoCon, adOpenDynamic, adLockOptimistic
        If rsbustype.State = adStateOpen Then
            If Not rsbustype.EOF Then
                MsgBox "Error while trying to remove Bustype  " & strbustype & " of Bustype details is not removed", vbExclamation, App.ProductName
                cmdclear.SetFocus
                Exit Sub
            End If
            Call cmdClear_Click
            Call FillbustypeView
            MsgBox "Bustype " & strbustype & " successfully removed from Bustype list", vbInformation, App.ProductName
            txtbustype.SetFocus
            rsbustype.Close
        End If
    End If
    cmdRemove.Enabled = False
Exit Sub
CatchError:
    MsgBox "Error : " & err.Number & vbTab & err.Description & vbCrLf & _
        "Bustype " & strbustype & " of Bustype details is not removed", vbExclamation, App.ProductName
        cmdclear.SetFocus
End Sub

Private Sub cmdRemove_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to remove a service type"
End Sub

Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to remove Bus Type details"
End Sub

Private Sub cmdSave_Click()
On Error GoTo CatchError
    If Trim(txtbustype.Text) = "" Then
        MsgBox "Bustype is empty! Please enter a Bustype", vbInformation, App.ProductName
        txtbustype.SetFocus
        Exit Sub
    End If
    Dim rsbustype As New ADODB.Recordset
    If blnEdit = False Then
        rsbustype.Open "SELECT * FROM [bustype] WHERE [name]='" & txtbustype.Text & "'", gbladoCon, adOpenDynamic, adLockOptimistic
        If rsbustype.State = adStateOpen Then
            If Not rsbustype.EOF Then
                MsgBox "Bustype already in use! Please specify another Bustype!", vbInformation, App.ProductName
                txtbustype.SetFocus
                Call SendKeys("{HOME}+{END}")
                Exit Sub
            Else
                rsbustype.AddNew
                rsbustype!Name = txtbustype.Text
                rsbustype.Update
                MsgBox "New Bustype details added successfully", vbInformation, App.ProductName
            End If
        Else
            MsgBox "Database Error! Unable to get Bustype details", vbInformation, App.ProductName
            Exit Sub
        End If
        rsbustype.Close
    Else
        rsbustype.Open "SELECT * FROM [bustype] WHERE [name]='" & txtbustype.Text & "' AND [id]<>" & intBustypeID, gbladoCon, adOpenDynamic, adLockOptimistic
        If rsbustype.State = adStateOpen Then
            If Not rsbustype.EOF Then
                MsgBox "Bustype code already in use! Please specify another or Bustype", vbInformation, App.ProductName
                txtbustype.SetFocus
                Call SendKeys("{HOME}+{END}")
                Exit Sub
            End If
            rsbustype.Close
        End If
        rsbustype.Open "SELECT * FROM [bustype] WHERE [id]=" & intBustypeID, gbladoCon, adOpenDynamic, adLockOptimistic
        If rsbustype.State = adStateOpen Then
            If Not rsbustype.EOF Then
                rsbustype!Name = txtbustype.Text
                rsbustype.Update
                MsgBox "Bustype details updated Successfully", vbInformation, App.ProductName
            Else
                MsgBox "Current record not found! Unable to update", vbInformation, App.ProductName
                Exit Sub
            End If
            rsbustype.Close
        Else
            MsgBox "Current record not found! Unable to update", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    Call cmdClear_Click
    Call FillbustypeView
    txtbustype.SetFocus
Exit Sub
CatchError:
    MsgBox "Error : " & err.Number & vbTab & err.Description & vbCrLf & _
           "Could not perform the requested operation. Adding or updating Bustype details failed", vbExclamation, App.ProductName
    txtbustype.SetFocus
End Sub

Private Sub cmdSave_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to add or update a Bustype details"
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to add or update a Bustype details"
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Icon = frmMainform.Icon
    Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
    Call FillbustypeView
    Call cmdClear_Click
End Sub

Private Sub FillbustypeView()
Dim rsbustype As ADODB.Recordset
Dim intCount As Integer
On Error GoTo CatchError
    mshflxBustype.Clear
    mshflxBustype.FormatString = "^RecordID   |^Sl.No               |<Bus Type Name                                                      "
    Set rsbustype = New ADODB.Recordset
    rsbustype.Open "SELECT * FROM [bustype]", gbladoCon, adOpenDynamic, adLockOptimistic
    intCount = 1
    If rsbustype.State = adStateOpen Then
        While Not rsbustype.EOF
            If intCount >= mshflxBustype.Rows Then mshflxBustype.Rows = mshflxBustype.Rows + 1
            mshflxBustype.TextMatrix(intCount, 0) = rsbustype!Id
            mshflxBustype.TextMatrix(intCount, 1) = intCount
            mshflxBustype.TextMatrix(intCount, 2) = rsbustype!Name
            intCount = intCount + 1
            rsbustype.MoveNext
        Wend
    End If
    If rsbustype.State = adStateOpen Then rsbustype.Close
    mshflxBustype.ColWidth(0) = 0
Exit Sub
CatchError:
    MsgBox err.Number & vbCrLf & err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = ""
End Sub
Private Sub mshflxBustype_DblClick()
Dim intCont As Integer
On Error Resume Next
    cmdRemove.Enabled = True
    Me.MousePointer = vbHourglass
    If mshflxBustype.row = 0 Or mshflxBustype.TextMatrix(mshflxBustype.row, 0) = "" Then Exit Sub
    txtbustype.Text = mshflxBustype.TextMatrix(mshflxBustype.row, 2)
    intBustypeID = mshflxBustype.TextMatrix(mshflxBustype.row, 0)
    blnEdit = True
    txtbustype.SetFocus
    Me.MousePointer = vbNormal
End Sub
Private Sub mshflxBustype_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Double click or press Ctrl+Enter to update selected Bus type details"
    mshflxBustype.HighLight = flexHighlightWithFocus
End Sub

Private Sub mshflxBustype_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn And Shift = 2 Then
        Call mshflxBustype_DblClick
    ElseIf KeyCode = vbKeyReturn Then
        Call SendKeys("{TAB}")
    ElseIf KeyCode = vbKeyEscape Then
        Call SendKeys("+{TAB}")
    End If
End Sub

Private Sub mshflxBustype_LostFocus()
On Error Resume Next
    mshflxBustype.HighLight = flexHighlightNever
End Sub

Private Sub mshflxBustype_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Double click or press Ctrl+Enter to update selected Bustype details"
End Sub

Private Sub txtbustype_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to Enter Bustype Name"
End Sub

Private Sub txtbustype_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Call SendKeys("{TAB}")
    ElseIf KeyAscii = vbKeyEscape Then
        Call SendKeys("+{TAB}")
    End If
End Sub
Private Sub txtbustype_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to enter Bustype Name"
End Sub
