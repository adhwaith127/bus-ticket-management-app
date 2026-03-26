VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmSelectRoute 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   8055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSelectRoute.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton cmdClose 
      Height          =   450
      Left            =   6300
      TabIndex        =   8
      Top             =   3700
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      TX              =   "&Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSelectRoute.frx":0CCA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdDownload 
      Height          =   450
      Left            =   5000
      TabIndex        =   2
      Top             =   3700
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      TX              =   "&Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSelectRoute.frx":0CE6
      BC              =   12632256
      FC              =   0
   End
   Begin VB.CheckBox ChkSelect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select All"
      Height          =   300
      Left            =   105
      TabIndex        =   0
      Top             =   3660
      Width           =   1710
   End
   Begin VB.ListBox lstRoute 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   90
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1155
      Width           =   7890
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Route"
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
      Height          =   495
      Left            =   2715
      TabIndex        =   7
      Top             =   165
      Width           =   2835
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fare Type"
      Height          =   285
      Index           =   2
      Left            =   4155
      TabIndex        =   6
      Top             =   870
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stages"
      Height          =   285
      Index           =   3
      Left            =   3285
      TabIndex        =   5
      Top             =   870
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Route Name"
      Height          =   285
      Index           =   1
      Left            =   975
      TabIndex        =   4
      Top             =   870
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Route"
      Height          =   285
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   870
      Width           =   855
   End
End
Attribute VB_Name = "FrmSelectRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim Rs_MRoute As DAO.Recordset
Dim NoOfStages As String * 2
Dim InterState As String * 2
Dim StartingPlace As String
Dim BusType As String * 2
Private Sub chkselect_Click()
Dim I As Integer
    If ChkSelect.Value = 1 Then
        For I = 0 To lstRoute.ListCount - 1
            lstRoute.Selected(I) = True 'True changed by deej 0n 03-05-05
            lstRoute.SetFocus
        Next I
    Else
            For I = 0 To lstRoute.ListCount - 1
            lstRoute.Selected(I) = False 'True                                                    changed by deej 0n 03-05-05
        Next I
    End If
End Sub

Private Sub ChkSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDownload.SetFocus
    End If
End Sub

Private Sub cmdClose_Click()
    CancelFlag = False
    Unload Me
End Sub

Private Sub cmdDownload_Click()
On Error Resume Next
Dim I As Integer
Dim strRecord As String
Dim temp As String
    SeletedRouteCount = 0
    If lstRoute.SelCount > 0 Then
        ReDim SeletedRoute(lstRoute.ListCount)
        For I = 0 To lstRoute.ListCount - 1
            lstRoute.ListIndex = I
            If lstRoute.Selected(I) = True Then
                strRecord = lstRoute.List(I)
                temp = Trim(Left(strRecord, 4))
                SeletedRoute(I) = Right(temp, 4) '[
                SeletedRouteCount = SeletedRouteCount + 1
            End If
        Next I
        Unload Me
        CancelFlag = True
    End If
End Sub


Private Sub Form_Load()
Dim RouteNo As String
Dim StPlace As String
Dim NofStages As String
Dim BusType As String
Dim I As Integer
    Me.Icon = frmMainform.Icon
    CONNECT_DB
    'Set Rs_MRoute = DB.OpenRecordset("ROUTE", dbOpenDynaset)------
     sql = "select * from route order by rutcode"
     Set Rs_MRoute = DB.OpenRecordset(sql, dbOpenDynaset)
    
    If Rs_MRoute.RecordCount = 0 Then
        MsgBox "Route File not Available...", , "SILTRANS"
        cmdDownload.Enabled = False
        Exit Sub
    End If
            
    With Rs_MRoute
        .MoveFirst
        Do While .EOF = False
            RouteNo = Space(4)
            RSet RouteNo = .Fields("rutcode")
            StPlace = Space(20)
            LSet StPlace = "  " & .Fields("rutname")
            NofStages = Space(4)
            RSet NofStages = " " & .Fields("nostage")
            BusType = Space(8)
            RSet BusType = " " & .Fields("faretype")
            lstRoute.AddItem RouteNo & StPlace & NofStages & BusType
        
            .MoveNext
        
        Loop
        For I = 0 To lstRoute.ListCount - 1
            lstRoute.Selected(I) = False
        Next I
    End With
  '  ChkSelect.SetFocus
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Rs_MRoute.Close
    Set Rs_MRoute = Nothing
End Sub


Private Sub lstRoute_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        cmdDownload.SetFocus
    End If
End Sub
