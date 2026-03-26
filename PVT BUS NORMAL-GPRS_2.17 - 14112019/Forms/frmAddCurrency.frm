VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmAddCurrency 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Currency"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar stbarMassage 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4080
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13891
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
      TabIndex        =   10
      Top             =   3585
      Width           =   7935
      Begin JeweledBut.JeweledButton cmdClose 
         Height          =   375
         Left            =   5280
         TabIndex        =   4
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
         MICON           =   "frmAddCurrency.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdRemove 
         Height          =   375
         Left            =   960
         TabIndex        =   5
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
         MICON           =   "frmAddCurrency.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSave 
         Height          =   375
         Left            =   2400
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
         MICON           =   "frmAddCurrency.frx":0038
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdClear 
         Height          =   375
         Left            =   3840
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
         MICON           =   "frmAddCurrency.frx":0054
         BC              =   12632256
         FC              =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3900
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Width           =   8055
      Begin VB.TextBox txtCurrencyCode 
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
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtCountryName 
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
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   0
         Top             =   480
         Width           =   3015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshflxCurrency 
         Height          =   2490
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   7665
         _ExtentX        =   13520
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
         FormatString    =   "Sl.No               |County Name                                  |Currency Code         "
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Code"
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
         Left            =   4200
         TabIndex        =   9
         Top             =   480
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
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
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAddCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intCurrencyID As Integer
Dim blnEdit As Boolean
Private Sub cmdClear_Click()
On Error Resume Next
    txtCountryName.Text = ""
    txtCurrencyCode.Text = ""
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
Dim strCurrency As String, strCountry As String
Dim rsCurrency As ADODB.Recordset
Dim intCurrencyID As Integer
On Error GoTo CatchError
    If mshflxCurrency.TextMatrix(mshflxCurrency.row, 0) = "" Or mshflxCurrency.row = 0 Then Exit Sub
    strCurrency = mshflxCurrency.TextMatrix(mshflxCurrency.row, 3)
    strCountry = mshflxCurrency.TextMatrix(mshflxCurrency.row, 2)
    intCurrencyID = mshflxCurrency.TextMatrix(mshflxCurrency.row, 0)
    If (MsgBox("Do you want to remove Currency Code " & strCurrency & " of Country " & strCountry & " from Currency details ", vbQuestion + vbYesNo)) = vbYes Then
        gbladoCon.Execute "DELETE * FROM [Currency] WHERE RecordID=" & intCurrencyID
        Set rsCurrency = New ADODB.Recordset
        rsCurrency.Open "SELECT * FROM [Currency] WHERE [RecordID]=" & intCurrencyID, gbladoCon, adOpenDynamic, adLockOptimistic
        If rsCurrency.State = adStateOpen Then
            If Not rsCurrency.EOF Then
                MsgBox "Error while trying to remove Currency Code " & strCurrency & " of Country " & strCountry & ". Currency details not removed", vbExclamation, App.ProductName
                cmdclear.SetFocus
                Exit Sub
            End If
            Call cmdClear_Click
            Call FillCurrencyView
            MsgBox "Currency Code " & strCurrency & " of Country " & strCountry & " successfully removed from Currency list", vbInformation, App.ProductName
            txtCountryName.SetFocus
            rsCurrency.Close
        End If
    End If
    cmdRemove.Enabled = False
Exit Sub
CatchError:
    MsgBox "Error : " & err.Number & vbTab & err.Description & vbCrLf & _
        "Currency Code " & strCurrency & " of Country " & strCountry & " not removed", vbExclamation, App.ProductName
    cmdclear.SetFocus
End Sub

Private Sub cmdRemove_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to remove a service type"
End Sub

Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to remove Currency details"
End Sub

Private Sub cmdSave_Click()
On Error GoTo CatchError
    If Trim(txtCountryName.Text) = "" Then
        MsgBox "Country Name is empty! Please enter a Country Name", vbInformation, App.ProductName
        txtCountryName.SetFocus
        Exit Sub
    End If
    If Trim(txtCurrencyCode.Text) = "" Then
        MsgBox "Currency Code is empty! Please enter a Currency Code", vbInformation, App.ProductName
        txtCurrencyCode.SetFocus
        Exit Sub
    End If
    Dim rsCurrency As New ADODB.Recordset
    If blnEdit = False Then
        rsCurrency.Open "SELECT * FROM [Currency] WHERE [Country]='" & txtCountryName.Text & "'", gbladoCon, adOpenDynamic, adLockOptimistic
        If rsCurrency.State = adStateOpen Then
            If Not rsCurrency.EOF Then
                MsgBox "Country Name  already in use! Please specify another Country Name", vbInformation, App.ProductName
                txtCountryName.SetFocus
                Call SendKeys("{HOME}+{END}")
                Exit Sub
            Else
                rsCurrency.AddNew
                rsCurrency!Country = txtCountryName.Text
                rsCurrency!Currency = txtCurrencyCode.Text
                rsCurrency.Update
                MsgBox "New Currency details added Successfully", vbInformation, App.ProductName
            End If
        Else
            MsgBox "Database Error! Unable to get Currency details", vbInformation, App.ProductName
            Exit Sub
        End If
        rsCurrency.Close
    Else
        rsCurrency.Open "SELECT * FROM [Currency] WHERE [Country]='" & txtCountryName.Text & "' AND [RecordID]<> " & intCurrencyID & "", gbladoCon, adOpenDynamic, adLockOptimistic
        If rsCurrency.State = adStateOpen Then
            If Not rsCurrency.EOF Then
                MsgBox "Country Name already in use! Please specify another Country Name", vbInformation, App.ProductName
                txtCountryName.SetFocus
                Call SendKeys("{HOME}+{END}")
                Exit Sub
            End If
        End If
        rsCurrency.Close
        rsCurrency.Open "SELECT * FROM [Currency] WHERE [RecordID]=" & intCurrencyID, gbladoCon, adOpenDynamic, adLockOptimistic
        If rsCurrency.State = adStateOpen Then
            If Not rsCurrency.EOF Then
                rsCurrency!Country = txtCountryName.Text
                rsCurrency!Currency = txtCurrencyCode.Text
                rsCurrency.Update
                MsgBox "Currency details updated Successfully", vbInformation, App.ProductName
            Else
                MsgBox "Current record not found! Unable to update", vbInformation, App.ProductName
                Exit Sub
            End If
            rsCurrency.Close
        Else
            MsgBox "Current record not found! Unable to update", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    Call cmdClear_Click
    Call FillCurrencyView
    txtCountryName.SetFocus
Exit Sub
CatchError:
    MsgBox "Error : " & err.Number & vbTab & err.Description & vbCrLf & _
        "Could not perform the requested operation. Adding or updating Currency details failed", vbExclamation, App.ProductName
    txtCountryName.SetFocus
End Sub

Private Sub cmdSave_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to add or update a Currency details"
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to add or update a Currency details"
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Icon = frmMainform.Icon
    Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
    Call FillCurrencyView
    Call cmdClear_Click
End Sub

Private Sub FillCurrencyView()
Dim rsCurrency As ADODB.Recordset
Dim intCount As Integer
On Error GoTo CatchError
    mshflxCurrency.Clear
    mshflxCurrency.FormatString = "^RecordID   |^Sl.No               |<Country Name                         |<Currency Code                  "
    Set rsCurrency = New ADODB.Recordset
    rsCurrency.Open "SELECT * FROM [Currency]", gbladoCon, adOpenDynamic, adLockOptimistic
    intCount = 1
    If rsCurrency.State = adStateOpen Then
        While Not rsCurrency.EOF
            If intCount >= mshflxCurrency.Rows Then mshflxCurrency.Rows = mshflxCurrency.Rows + 1
            mshflxCurrency.TextMatrix(intCount, 0) = rsCurrency!RecordID
            mshflxCurrency.TextMatrix(intCount, 1) = intCount
            mshflxCurrency.TextMatrix(intCount, 2) = rsCurrency!Country
            mshflxCurrency.TextMatrix(intCount, 3) = rsCurrency!Currency
            intCount = intCount + 1
            rsCurrency.MoveNext
        Wend
    End If
    If rsCurrency.State = adStateOpen Then rsCurrency.Close
    mshflxCurrency.ColWidth(0) = 0
Exit Sub
CatchError:
    MsgBox err.Number & vbCrLf & err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = ""
End Sub

Private Sub mshflxCurrency_DblClick()
Dim intCont As Integer
On Error Resume Next
    cmdRemove.Enabled = True
    Me.MousePointer = vbHourglass
    If mshflxCurrency.row = 0 Or mshflxCurrency.TextMatrix(mshflxCurrency.row, 0) = "" Then Exit Sub
    txtCurrencyCode.Text = mshflxCurrency.TextMatrix(mshflxCurrency.row, 3)
    txtCountryName.Text = mshflxCurrency.TextMatrix(mshflxCurrency.row, 2)
    intCurrencyID = mshflxCurrency.TextMatrix(mshflxCurrency.row, 0)
    blnEdit = True
    txtCountryName.SetFocus
    Me.MousePointer = vbNormal
End Sub

Private Sub mshflxCurrency_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Double click or press Ctrl+Enter to update selected Currency details"
    mshflxCurrency.HighLight = flexHighlightWithFocus
End Sub

Private Sub mshflxCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn And Shift = 2 Then
        Call mshflxCurrency_DblClick
    ElseIf KeyCode = vbKeyReturn Then
        Call SendKeys("{TAB}")
    ElseIf KeyCode = vbKeyEscape Then
        Call SendKeys("+{TAB}")
    End If
End Sub

Private Sub mshflxCurrency_LostFocus()
On Error Resume Next
    mshflxCurrency.HighLight = flexHighlightNever
End Sub

Private Sub mshflxCurrency_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Double click or press Ctrl+Enter to update selected Currency details"
End Sub

Private Sub txtCurrencyCode_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to enter Currency Code"
End Sub

Private Sub txtCurrencyCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Trim(txtCurrencyCode.Text) <> "" Then
        Call SendKeys("{TAB}")
    ElseIf KeyAscii = vbKeyEscape Then
        Call SendKeys("+{TAB}")
    End If
End Sub

Private Sub txtCurrencyCode_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to enter Currency Code"
End Sub

Private Sub txtCountryName_GotFocus()
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to Enter Country Name"
End Sub

Private Sub txtCountryName_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Call SendKeys("{TAB}")
    ElseIf KeyAscii = vbKeyEscape Then
        Call SendKeys("+{TAB}")
    End If
End Sub

Private Sub txtCountryName_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    stbarMassage.Panels(1).Text = "Click here to enter Country Name"
End Sub
