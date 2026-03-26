VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form NEWROUTEFORM 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NEW ROUTE"
   ClientHeight    =   9240
   ClientLeft      =   795
   ClientTop       =   1440
   ClientWidth     =   14130
   ControlBox      =   0   'False
   FillColor       =   &H8000000A&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "NEWROUTEFORM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9240
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4455
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtminfare 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   41
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox cmbBustype 
         Height          =   345
         ItemData        =   "NEWROUTEFORM.frx":0CCA
         Left            =   4920
         List            =   "NEWROUTEFORM.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1170
         Width           =   2055
      End
      Begin VB.ComboBox cmbfaretype 
         Height          =   345
         ItemData        =   "NEWROUTEFORM.frx":0CCE
         Left            =   4920
         List            =   "NEWROUTEFORM.frx":0CD8
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   570
         Width           =   2055
      End
      Begin VB.TextBox txtnostage 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   33
         Top             =   1770
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtrutcode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   32
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtrutname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   16
         TabIndex        =   31
         Top             =   1170
         Width           =   1575
      End
      Begin VB.CheckBox chkpass 
         BackColor       =   &H80000016&
         Caption         =   " ALLOW PASS"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6000
         Picture         =   "NEWROUTEFORM.frx":0CEA
         TabIndex        =   30
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CheckBox chkAdjust 
         BackColor       =   &H80000016&
         Caption         =   "ADJUST"
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
         Left            =   3000
         TabIndex        =   29
         Top             =   3360
         Width           =   1290
      End
      Begin VB.CheckBox chkConc 
         BackColor       =   &H80000016&
         Caption         =   "CONCESSION"
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
         Left            =   4320
         TabIndex        =   28
         Top             =   3360
         Width           =   1530
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000016&
         Caption         =   "Minimum Fare"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   42
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblBustype 
         BackStyle       =   0  'Transparent
         Caption         =   "Bus Type"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fare Type"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   39
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No of Stages"
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
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1770
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "RouteName"
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
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "RouteCode"
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
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   570
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdDel1 
      Caption         =   "Delete Row"
      Height          =   495
      Left            =   12000
      Picture         =   "NEWROUTEFORM.frx":18281
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdcleargrid 
      Caption         =   "CLEAR FARE"
      Height          =   495
      Left            =   12000
      MaskColor       =   &H00808080&
      Picture         =   "NEWROUTEFORM.frx":74753
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtDistance 
      Height          =   285
      Left            =   3960
      MaxLength       =   7
      TabIndex        =   14
      Top             =   10440
      Width           =   1200
   End
   Begin VB.TextBox txtStgName 
      Height          =   285
      Left            =   3960
      MaxLength       =   11
      TabIndex        =   13
      Top             =   9960
      Width           =   1800
   End
   Begin VB.OptionButton optallow 
      BackColor       =   &H00CAB7B4&
      Caption         =   "FULL"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.OptionButton optallow 
      BackColor       =   &H00CAB7B4&
      Caption         =   "HALF"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.OptionButton optallow 
      BackColor       =   &H00CAB7B4&
      Caption         =   "LUGG"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.OptionButton optallow 
      BackColor       =   &H00CAB7B4&
      Caption         =   "PH"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.OptionButton optallow 
      BackColor       =   &H00CAB7B4&
      Caption         =   "STUDENT"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtStage1 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8880
      MaxLength       =   15
      TabIndex        =   25
      Top             =   2160
      Width           =   1530
   End
   Begin VB.TextBox txtStage2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   11520
      MaxLength       =   6
      TabIndex        =   12
      Top             =   2160
      Width           =   1530
   End
   Begin VB.TextBox txtFare1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   720
      MaxLength       =   7
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtFare2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2010
      MaxLength       =   6
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "&CLEAR ALL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      MaskColor       =   &H00E0E0E0&
      Picture         =   "NEWROUTEFORM.frx":7589C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      Picture         =   "NEWROUTEFORM.frx":769E5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdFareEntry 
      Caption         =   "&FINISH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      Picture         =   "NEWROUTEFORM.frx":77B2E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7920
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid stageGrid 
      Height          =   3615
      Left            =   8400
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   4
      ForeColor       =   2568952
      BackColorFixed  =   64
      ForeColorFixed  =   -2147483643
      BackColorBkg    =   -2147483627
      ScrollTrack     =   -1  'True
      Enabled         =   -1  'True
      TextStyle       =   1
      GridLines       =   2
      PictureType     =   1
      Appearance      =   0
      FormatString    =   "              |         STAGE     |    DISTANCE  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid FareTypeGrid 
      Height          =   4050
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   7144
      _Version        =   393216
      ForeColor       =   255
      BackColorFixed  =   64
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   -2147483627
      GridColor       =   16777215
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004080&
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   5775
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Height          =   4335
      Left            =   120
      Top             =   4800
      Width           =   11055
   End
   Begin VB.Label lblallowtype 
      Caption         =   "allowtype"
      Height          =   375
      Left            =   8520
      TabIndex        =   23
      Top             =   10200
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "FARE ENTRY FOR "
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Label lblnoofstage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   9960
      Width           =   765
   End
   Begin VB.Label lblnoofentry 
      BackColor       =   &H00E0E0E0&
      Caption         =   "entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7755
      TabIndex        =   20
      Top             =   10680
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Entries :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6915
      TabIndex        =   19
      Top             =   10680
      Width           =   765
   End
   Begin VB.Label lblEntered 
      BackColor       =   &H00E0E0E0&
      Caption         =   "entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   8655
      TabIndex        =   18
      Top             =   10680
      Width           =   495
   End
   Begin VB.Label Label19 
      BackColor       =   &H8000000A&
      Caption         =   "Distance      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   2400
      TabIndex        =   17
      Top             =   10440
      Width           =   1500
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000A&
      Caption         =   "Stage Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   2400
      TabIndex        =   16
      Top             =   9960
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FARE ENTRY FOR........"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Height          =   4095
      Left            =   11280
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   2775
   End
End
Attribute VB_Name = "NEWROUTEFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As DAO.Database
Dim TDB As DAO.Database
Dim RES As DAO.Recordset
Dim TRES As DAO.Recordset
Dim rs As DAO.Recordset
Dim FareType As Integer
Dim chkarray(10) As String
Dim arraylen As Integer
Dim arrayindex As Integer
Dim c As Integer
Dim R As Integer
Dim EntryCount As Integer
Dim EnteredCount As Long
Dim NoOfEntries As Long
Dim chkno As Integer
Dim checkflag As Integer
Dim rowno As Integer
Dim MaxRow As Integer
Dim Id As Long
Dim clickflag As Integer
Dim rrow As Integer
Dim CCol As Integer
Dim dotc As Integer
Dim FARESAVE As Integer
Dim oldfare As Single
Dim stgclick As Boolean


Public Sub CenterForm(pobjForm As Form)
 On Error Resume Next
   With pobjForm
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub
Public Function ADD_FARE_TO_TABLE() 'for add fare in same order as previous prjct 20.12 vaisakh
    
    RSql = "SELECT * FROM " & optallow(Indexno).caption
    Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
       If RES.RecordCount > 0 Then RES.MoveFirst
        With FareTypeGrid
            .row = 2
            .col = 1
            R = 1
            c = 1
            'While R < NoOfEntries
                While c < NOSTGS
                    
                    While R < NOSTGS
                        RES.AddNew
                        RES!row = R
                        RES!col = c
                        RES!FARE = Round((.TextMatrix(R + 1, c)), 2)
                        RES!Route = txtrutcode.Text
                        RES.Update
                        R = R + 1
                    Wend
                    c = c + 1
                    R = c
                Wend
        End With
    
    
End Function

Public Function RESETOPTION()
    optallow(0).Enabled = True
    optallow(1).Enabled = True
    optallow(2).Enabled = True
    optallow(3).Enabled = True
    optallow(4).Enabled = True
    
    
    optallow(0).Value = False
    optallow(1).Value = False
    optallow(2).Value = False
    optallow(3).Value = False
    optallow(4).Value = False
    
    cmdcleargrid.Visible = False
    txtFare1.Visible = False

    'Indexno = 6
    
End Function
Public Function RESETSTAGEGRID()
    If stageGrid.Enabled = False Then stageGrid.Enabled = True
    If cmdDel1.Enabled = -False Then cmdDel1.Enabled = True
    stageGrid.Clear
    stageGrid.Rows = 2
    stageGrid.Cols = 2
'    stageGrid.FormatString = "SL.NO.|STAGE NAME|DISTANCE |LANGUAGE" '30.12.2010
    stageGrid.FormatString = "SL.NO.|STAGE NAME|DISTANCE "

    stageGrid.ColWidth(1) = 2500
    stageGrid.ColWidth(2) = 1500
    stageGrid.ColWidth(2) = 1500
    stageGrid.RowHeightMin = 340
End Function

Public Function RESETFAREGRID()

    FareTypeGrid.Clear
        FareTypeGrid.Cols = 2
        FareTypeGrid.Rows = 2
        txtFare1 = ""
        txtFare2 = ""
        flxgrdCnt = 1
        n = 1
        R = 1
        c = 1
        EnteredCount = 1
        EntryCount = 1
        NOS = NOSTGS - 1
    Label5.Visible = True
    Label11.Visible = True
    lblallowtype.Visible = True
    lblnoofentry.Visible = True
    lblEntered.Visible = True
    FareTypeGrid.Enabled = True
    
    cmdFareEntry.Enabled = True
    cmdFareEntry.Visible = True
    
End Function
Public Function showcol()
'    If stageGrid.Col = 3 Then
            txtStage1.Width = stageGrid.CellWidth
            txtStage1.Height = stageGrid.CellHeight
            txtStage1.Left = stageGrid.Left + stageGrid.CellLeft
            txtStage1.Top = stageGrid.Top + stageGrid.CellTop
            txtStage1.Visible = True
            txtStage1.Text = stageGrid.TextMatrix(stageGrid.row, stageGrid.col)
            txtStage1.SetFocus
        clickflag = 1
        rrow = stageGrid.row
        CCol = stageGrid.col
            ''Text2.Locked = True
            txtStage1.SelStart = 0
            txtStage1.SelLength = Len(txtStage1)
'        Else
'''''        txtStage1.Width = stageGrid.CellWidth
'''''        txtStage1.Height = stageGrid.CellHeight
'''''        txtStage1.Left = stageGrid.Left + stageGrid.CellLeft
'''''        txtStage1.Top = stageGrid.Top + stageGrid.CellTop
'''''        txtStage1.Visible = True
'''''        txtStage1.Text = stageGrid.TextMatrix(stageGrid.row, stageGrid.Col)
'''''
'''''
'''''        txtStage1.SetFocus   ''Text1.Locked = True
'''''        txtStage1.SelStart = 0
'''''        txtStage1.SelLength = Len(txtStage1)
      '  End If

End Function
Public Function DELETETABLES()
        RSql = "DELETE * FROM ROUTE"
        TDB.Execute (RSql)
         RSql = "DELETE * FROM STAGE"
        TDB.Execute (RSql)
        RSql = "DELETE * FROM FULL"
        TDB.Execute (RSql)
        RSql = "DELETE * FROM HALF"
        TDB.Execute (RSql)
        RSql = "DELETE * FROM LUGG"
        TDB.Execute (RSql)
        RSql = "DELETE * FROM PH"
        TDB.Execute (RSql)
''''        RSql = "DELETE * FROM PASS"
''''        TDB.Execute (RSql)
        RSql = "DELETE * FROM STUDENT"
        TDB.Execute (RSql)
End Function

'Public Function CHECKALLOWABLES()
'
'    For chkno = 0 To 4
'    If chkallowable(chkno).Value = 1 Then
'        checkflag = 1
'        RES.Close
'        TSQL = "SELECT * FROM " & chkallowable(chkno).Caption
'        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'        'chkno = chkno + 1
'        FareTypeGrid.Clear
'        FareTypeGrid.Cols = 2
'        FareTypeGrid.Rows = 2
'         flxgrdCnt = 1
'        n = 1
'    R = 1
'    c = 1
'    EnteredCount = 1
'
'        FARELISTENTRY
'        txtFare1.Enabled = True
'        txtFare1.Visible = True
'        txtFare1.SetFocus
'    End If
'
'    Next chkno
''    Else
''        cmdFareEntry.Visible = True
''        cmdFareEntry.SetFocus
'    checkflag = 0
'
'End Function


Public Function FARELISTENTRY()
    With FareTypeGrid
    If cmbfaretype.Text = "GRAPH" Then
        .Cols = NOSTGS + 1
        .Rows = NOSTGS + 1
        .col = 0
        .RowHeight(.col) = 300
        
        R = 0
        While R < .Cols
            .TextMatrix(0, R) = R
            .TextMatrix(R, 0) = R
            R = R + 1
        Wend
           
            .row = 1
            .col = 1
            rrow = 1
            CCol = 1
            .RowHeight(rrow) = 300
            .ColWidth(CCol) = 1400
            .CellForeColor = vbRed
            .TextMatrix(rrow, CCol) = stageGrid.TextMatrix(rrow, CCol)
            rrow = 2
            Do While rrow < NOSTGS + 1
                CCol = 1
                While CCol <= rrow - 1
                    .RowHeight(rrow) = 300
                    .ColWidth(CCol) = 1400
                    .row = rrow
                    .col = CCol
                    .TextMatrix(rrow, CCol) = "0.00"
                    .CellBackColor = &HF1CCA3
                    
                    CCol = CCol + 1
                Wend
                .row = rrow
                .col = CCol
                .RowHeight(rrow) = 300
                .ColWidth(CCol) = 1400
                .CellForeColor = vbRed
                .TextMatrix(rrow, rrow) = stageGrid.TextMatrix(rrow, 1)
                rrow = rrow + 1
            Loop
            
             .row = 2
                .col = 1
    
    Else
        .Cols = NOSTGS + 1
        
        .col = 0
        Do While .col < NOSTGS '- 1
            .col = .col + 1
            
            .TextMatrix(0, .col) = Trim(stageGrid.TextMatrix(.col, 1))
        Loop
        CCol = 1
        .RowHeight(0) = 300
        .RowHeight(1) = 300
        Do While CCol < .Cols
            .TextMatrix(1, CCol) = "0.00"

            .Gridlines = flexGridNone
            .GridLineWidth = 1
            .col = CCol
            .ColWidth(CCol) = 1400
            .CellBackColor = vbWhite
            CCol = CCol + 1
        Loop
        .row = 1
        .col = 1
    End If
    End With
    'End If
    
    FareTypeGrid.col = 1
    txtFare1.Visible = True
    txtFare1.Top = FareTypeGrid.CellTop + FareTypeGrid.Top
    txtFare1.Left = FareTypeGrid.CellLeft + FareTypeGrid.Left
    txtFare1.Width = FareTypeGrid.CellWidth
    txtFare1.Height = FareTypeGrid.CellHeight
    If FareTypeGrid.TextMatrix(FareTypeGrid.row, FareTypeGrid.col) <> "" Then txtFare1 = Round(FareTypeGrid.TextMatrix(FareTypeGrid.row, FareTypeGrid.col), 2)
    txtFare1.SelStart = 0
    txtFare1.SelLength = Len(txtFare1)
    Call SendKeys("{HOME}+{END}")
    txtFare1.Enabled = True
    txtFare1.Visible = True
    txtFare1.SetFocus
     
End Function

Public Function STAGENamesOneByOne()   ' To Store the Datas by pressing ENTER key
    If Trim(txtStgName.Text) <> "" Then
        If StgNameCnt < NOSTGS Then
            StageNameBuff = StageNameBuff + Trim(txtStgName.Text) + ","
            With stageGrid
                .AddItem ""
                .TextMatrix(StgNameCnt, 0) = StgNameCnt
                .TextMatrix(StgNameCnt, 1) = Trim(txtStgName.Text)
            
            End With
            If StgNameCnt > 9 Then
                stageGrid.TopRow = StgNameCnt - 1
            End If
            StgName(StgNameCnt - 1) = Trim(txtStgName.Text) & ","
            'txtStgName.Text = ""
            txtStgName.SetFocus
            txtStgName.BackColor = &HC0EF00
            toListBox = toListBox + 1
        End If
        
        
'MALAYALAM IS NOT AVAILABLE FOR NOW LOCAL LANGUAGE VALUE FOR MALAYALAM IS 1
'********************For Language Only****************************
        If LocalLanguage > 0 And LocalLanguage <> 4 Then 'AND PART ADDED BY SYAM<>1
            strLanguageStage = ""
'            Load H_Convert
'            H_Convert.Show vbModal
            'frmBMPEditor.Show vbModal
            If strLanguageStage = "" Then
                strLanguageStage = " "
            End If
            stageGrid.TextMatrix(StgNameCnt, 3) = strLanguageStage
        End If
'********************For Language Only****************************
        
        If StgNameCnt < stageGrid.Rows Then
            stageGrid.row = StgNameCnt
        End If
        stageGrid.col = 1
        stageGrid.CellBackColor = &H80000005
        stageGrid.col = 2
        stageGrid.CellBackColor = &H900FFF
        txtStgName.Enabled = False
        txtDistance.Enabled = True
        
        txtDistance.SetFocus
        txtDistance.BackColor = &HC0EF00
        If LocalLanguage = 0 Then StgNameCnt = StgNameCnt + 1
        
    Else
        MsgBox "Enter the stage name", vbInformation
        
        txtStgName.SetFocus
        txtStgName.BackColor = &HC0EF00
        Exit Function
    End If
    
    If StgNameCnt > NOSTGS And COUNTER >= 1 Then
        StageNameBuff = StageNameBuff + Trim(txtStgName.Text)
        'stagelist.AddItem Trim(txtStgName.Text), toListBox
        With stageGrid
            .AddItem ""
            .TextMatrix(StgNameCnt - 1, 0) = StgNameCnt - 1
            .TextMatrix(StgNameCnt - 1, 1) = Trim(txtStgName.Text)
            .TextMatrix(StgNameCnt - 1, 3) = strLanguageStage
        End With
        StgName(StgNameCnt - 2) = Trim(txtStgName.Text)
        txtStgName.Text = ""
        toListBox = 0
        txtStgName.Enabled = False
        stageGrid.Rows = stageGrid.Rows - 1
        'STAGENAMEENTRY
        Exit Function
    End If


End Function



Private Sub Image2_Click()

End Sub

Private Sub chkHalf_Click()

End Sub

Private Sub chkpass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then
        stageGrid.Enabled = True
        stageGrid_Click
    End If
End Sub

Private Sub cmbBustype_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'optallow(0).SetFocus
            stageGrid_Click
    End If
    
End Sub

Private Sub cmbfaretype_KeyPress(KeyAscii As Integer)
    If cmbfaretype.Text = "TABLE" Then
            txtFare1.Text = ""
            'FareTypeGrid.Cols = 2  'to set the rows to flexgrid
'            Buffer = Buffer & "1" & ","
            FrType = 1
        ElseIf cmbfaretype.Text = "GRAPH" Then
            'FareTypeGrid.Cols = 2  'to set the rows to flexgrid
'            Buffer = Buffer & "4" & ","
            FrType = 2
        End If
    If KeyAscii = 13 Then
            cmbBusType.Enabled = True
         cmbBusType.SetFocus
    End If
End Sub
'
'Private Sub cmdallow_Click(Index As Integer)
'
'
'
'Select Case Index
'
'Case 0
'        'RES.Close
'
'        RSql = "DELETE * FROM TMPFARE WHERE ROUTE = '" & txtrutcode & "'"
'        TDB.Execute (RSql)
'        TSQL = "SELECT * FROM TMPFARE" '& cmdallow(Index).Caption
'        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'Case 1
'        TSQL = "SELECT * FROM " & optallow(Index).Caption
'        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'Case 2
'        TSQL = "SELECT * FROM " & opallow(Index).Caption
'        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'Case 3
'        TSQL = "SELECT * FROM " & cmdallow(Index).Caption
'        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'Case 4
'        TSQL = "SELECT * FROM " & cmdallow(Index).Caption
'        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'Case 5
'        TSQL = "SELECT * FROM " & cmdallow(Index).Caption
'        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'
'End Select
'        FareTypeGrid.Clear
'        FareTypeGrid.Cols = 2
'        FareTypeGrid.Rows = 2
'        txtFare1 = ""
'        txtFare2 = ""
'        flxgrdCnt = 1
'        n = 1
'        R = 1
'        c = 1
'        EnteredCount = 1
'
'    Label5.Visible = True
'    Label11.Visible = True
'    lblallowtype.Visible = True
'    lblnoofentry.Visible = True
'    lblEntered.Visible = True
'    FareTypeGrid.Enabled = True
'
'
'
'        FARELISTENTRY
'        txtFare1.Enabled = True
'        txtFare1.Visible = True
'        txtFare1.SetFocus
'End Sub

Private Sub cmdClear_Click()
    
 Msg = MsgBox("Are you sure to clear all ?..", vbYesNo)
 If Msg = vbYes Then
    DELETETABLES
    txtrutcode = ""
    txtrutname = ""
    txtnostage = ""
    cmbfaretype.Enabled = True
    cmbBusType.Enabled = True
    cmbfaretype.Text = cmbfaretype.List(0)
    cmbBusType.Text = cmbBusType.List(0)
'    stageGrid.Clear
'    stageGrid.Cols = 3
'    stageGrid.Rows = 2
    RESETSTAGEGRID
    txtStage1 = ""
    txtStage2.Visible = False
    txtStage2 = ""
    FareTypeGrid.Clear
    RESETFAREGRID
    txtFare1 = ""
    txtFare1.Enabled = False
    txtFare1.Visible = False
    txtFare2.Enabled = False
    txtFare2.Visible = False
    txtFare2 = ""
    chkpass.Value = False
    RESETOPTION
    txtrutcode.SetFocus
    Indexno = 6
 End If
    
End Sub

Private Sub cmdcleargrid_Click()
    Msg = MsgBox("Are you sure to clear fare ?..", vbYesNo)
    If Msg = vbYes Then
        RSql = "DELETE * FROM " & optallow(Indexno).caption
        TDB.Execute (RSql)
        RESETOPTION
        RESETFAREGRID
        txtFare1 = ""
        txtFare1.Enabled = False
        txtFare1.Visible = False
        txtFare2.Enabled = False
        txtFare2.Visible = False
        txtFare2 = ""
        FARESAVE = 1
        
        Indexno = 6
        Exit Sub
    End If
End Sub

Private Sub cmdDel1_Click()
Dim I As Integer
    With stageGrid 'the msflexgrid
        If .RowSel <> 0 And (.row <> 0 And .row <> 1) And stgclick = True Then 'check if there is a selected row
        Msg = MsgBox("Are you sure to delete ?", vbYesNo)
        If (Msg = vbNo) Then
            Exit Sub
        End If
            For I = .RowSel To .Rows - 2 'loop from selected row to the las row
                .TextMatrix(I, 0) = I 'set rows with 1 back
                .TextMatrix(I, 1) = .TextMatrix(I + 1, 1)
                .TextMatrix(I, 2) = .TextMatrix(I + 1, 2)
                '.TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
            Next I
                .Rows = .Rows - 1 'make the rows 1 less
                
        Else
            If .row = 1 Or row = 0 Then
            MsgBox "Select row can't delete!!!", vbExclamation
            Else
            MsgBox "Selecet row to delete!!!", vbExclamation
            End If
        End If
        stgclick = False
    End With
End Sub

Private Sub cmdFareEntry_Click()
Dim Strstg As String
Dim Gpos As Integer
Dim tmpbuf As String
Dim stagecode As Integer
    
    For Each Ctrl In Me.Controls
        If TypeOf Ctrl Is TextBox Then
            If Ctrl.Name = "txtrutcode" Or Ctrl.Name = "txtrutname" Or Ctrl.Name = "txtnostage" Then
                If Ctrl.Text = "" Then
                    MsgBox "Some field missing!", vbInformation, gblstrPrjTitle
                    Exit Sub
                End If
            End If
        End If
    Next
    
    'If stageGrid.Rows <> val(txtnostage) + 1 Then
    If stageGrid.TextMatrix(NOSTGS, 1) = "" Then
        MsgBox "Please enter the stage details !", vbInformation
        txtnostage_KeyPress (13)
        Exit Sub
    'End If
    End If
    
    RSql = "SELECT * FROM FULL"
    Set rs = TDB.OpenRecordset(RSql, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        MsgBox "You should enter fare for full", vbInformation, gblstrPrjTitle
        optallow(0).SetFocus
        Exit Sub
    rs.Close
    End If
    
    
    
    'RouteCode = Mid$(DatFileName, 1, InStr(1, DatFileName, ".") - 1)
    
    Msg = MsgBox("Going to Create Route" & vbCrLf & "Do you want to continue ?", vbYesNo)
    
    If Msg = vbYes Then
        'Msg = MsgBox("Are You Sure ?", vbYesNo)
    If Msg = vbNo Then Exit Sub
    
    RES.Close
    RSql = "SELECT * FROM ROUTE"
        Set RES = cn.OpenRecordset(RSql, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            RES.MoveLast
            Id = RES!Id
        Else
            Id = 0
        End If
    RES.Close
     
        
    
    RSql = "SELECT * FROM ROUTE"
        Set RES = cn.OpenRecordset(RSql, dbOpenDynaset)
        'Id = RES!Id
        
        If RES.RecordCount > 0 Then RES.MoveLast
            RES.AddNew
            Id = Id + 1
            RES!Id = Id
            RES!RUTCODE = txtrutcode.Text
            RES!rutname = txtrutname.Text
            RES!nostage = val(txtnostage.Text)
            RES!nostop = val(txtnostage.Text)
            RES!MinFare = 0
    
            RES!FareType = cmbfaretype.ListIndex + 1
            RES!BusType = cmbBusType.ListIndex
            RES!UseStop = 0
        
        RSql = "SELECT * FROM FULL"
        Set rs = TDB.OpenRecordset(RSql, dbOpenDynaset)
        Set TRES = cn.OpenRecordset(RSql, dbOpenDynaset)
        If rs.RecordCount > 0 Then
        If TRES.RecordCount > 0 Then TRES.MoveLast
        Do While rs.EOF <> True
            TRES.AddNew
            TRES!row = rs!row
            TRES!col = rs!col
            TRES!FARE = rs!FARE
            TRES!Route = rs!Route
            TRES.Update
            rs.MoveNext
        Loop
            TRES.Close
        rs.Close
        End If
            
        RSql = "SELECT * FROM HALF"
        Set rs = TDB.OpenRecordset(RSql, dbOpenDynaset)
        Set TRES = cn.OpenRecordset(RSql, dbOpenDynaset)
        If rs.RecordCount > 0 Then
            RES!Half = 1
        If TRES.RecordCount > 0 Then TRES.MoveLast
        Do While rs.EOF <> True
            TRES.AddNew
            TRES!row = rs!row
            TRES!col = rs!col
            TRES!FARE = rs!FARE
            TRES!Route = rs!Route
            TRES.Update
            rs.MoveNext
        Loop
            TRES.Close
        rs.Close
        End If
        
        RSql = "SELECT * FROM LUGG"
        Set rs = TDB.OpenRecordset(RSql, dbOpenDynaset)
        Set TRES = cn.OpenRecordset(RSql, dbOpenDynaset)
        If rs.RecordCount > 0 Then
            RES!Luggage = 1
        If TRES.RecordCount > 0 Then TRES.MoveLast
        Do While rs.EOF <> True
            TRES.AddNew
            TRES!row = rs!row
            TRES!col = rs!col
            TRES!FARE = rs!FARE
            TRES!Route = rs!Route
            TRES.Update
            rs.MoveNext
        Loop
            TRES.Close
        rs.Close
        End If
        RSql = "SELECT * FROM STUDENT"
        Set rs = TDB.OpenRecordset(RSql, dbOpenDynaset)
        Set TRES = cn.OpenRecordset(RSql, dbOpenDynaset)
        If rs.RecordCount > 0 Then
            RES!student = 1
        If TRES.RecordCount > 0 Then TRES.MoveLast
        Do While rs.EOF <> True
            TRES.AddNew
            TRES!row = rs!row
            TRES!col = rs!col
            TRES!FARE = rs!FARE
            TRES!Route = rs!Route
            TRES.Update
            rs.MoveNext
        Loop
            TRES.Close
        rs.Close
        End If
'        RSql = "SELECT * FROM PH"            ''''''''PH DISABLED 13112010
'        Set rs = TDB.OpenRecordset(RSql, dbOpenDynaset)
'        Set TRES = cn.OpenRecordset(RSql, dbOpenDynaset)
'        If rs.RecordCount > 0 Then
'            RES!ph = 1
'        If TRES.RecordCount > 0 Then TRES.MoveLast
'        Do While rs.EOF <> True
'            TRES.AddNew
'            TRES!row = rs!row
'            TRES!Col = rs!Col
'            TRES!Fare = rs!Fare
'            TRES!Route = rs!Route
'            TRES.Update
'            rs.MoveNext
'        Loop
'            TRES.Close
'        rs.Close
'        End If
''''        RSql = "SELECT * FROM PASS"
''''        Set rs = TDB.OpenRecordset(RSql, dbOpenDynaset)
''''        Set TRES = cn.OpenRecordset(RSql, dbOpenDynaset)
''''        If rs.RecordCount > 0 Then
''''            RES!PASSALLOW = 1
''''        If TRES.RecordCount > 0 Then TRES.MoveLast
''''        Do While rs.EOF <> True
''''            TRES.AddNew
''''            TRES!row = rs!row
''''            TRES!Col = rs!Col
''''            TRES!Fare = rs!Fare
''''            TRES!Route = rs!Route
''''            TRES.Update
''''            rs.MoveNext
''''        Loop
''''            TRES.Close
''''        rs.Close
''''        End If
            RES!ph = 1
            RES!PASSALLOW = chkpass.Value
            RES!Adjust = chkAdjust.Value
            RES!Conc = chkConc.Value
            RES!StartFrom = 0
                
''''''                For i = 0 To txtnostage - 1
''''''                If i <> txtnostage - 1 Then
''''''                    Strstg = Strstg + StrCNTST(i) + ","
''''''                ElseIf i = txtnostage - 1 Then
''''''                    Strstg = Strstg + StrCNTST(i)
''''''                End If
''''''                Next
''''''                Gpos = InStr(Strstg, ",")
''''''                If Gpos > 0 Then
''''''                    tmpbuf = Mid(Strstg, 1, Gpos - 1) & "-"
''''''                End If
''''''                Gpos = 0
''''''                Gpos = InStrRev(Strstg, ",")
''''''                If Gpos > 0 Then
''''''                    tmpbuf = tmpbuf & Mid(Strstg, Gpos + 1, Len(Strstg) - Gpos)
''''''                End If
''''''                RES!StageCount = tmpbuf  'Strstg
          
            
            RES!StageCount = val(txtnostage.Text)
            RES!BusTypeName = cmbBusType.Text
            RES.Update
            RES.Close
       End If
        rrow = 1
        RSql = "SELECT * FROM STAGE"
        Set RES = cn.OpenRecordset(RSql, dbOpenDynaset)
            'RES.MoveFirst
            If RES.RecordCount > 0 Then
                RES.MoveLast
                Id = RES!Id + 1
                stagecode = RES!stagecode + 1
                
            Else
                Id = 0
                stagecode = 0
            End If
            'Id = Id + 1
        RES.Close
        RSql = "SELECT * FROM STAGE"
        'Set rs = TDB.OpenRecordset(RSql, dbOpenDynaset)
        Set TRES = cn.OpenRecordset(RSql, dbOpenDynaset)
        'stagecode = 0
        With stageGrid
        If TRES.RecordCount > 0 Then TRES.MoveLast
        Do While NOSTGS >= rrow
            TRES.AddNew
            
            TRES!StageName = Trim(.TextMatrix(rrow, 1))
            TRES!STG_LOCAL_LANGUAGE = .TextMatrix(rrow, 3)
            TRES!STG_LOCAL_LANGUAGE = 0
            TRES!Distance = 0
            TRES!Route = Trim(txtrutcode.Text)
            TRES!stagecode = stagecode
            TRES!Id = Id
            
            TRES.Update
            'rs.MoveNext
            rrow = rrow + 1
            Id = Id + 1
            stagecode = stagecode + 1
        Loop
        End With
            TRES.Close
        'rs.Close
        
        
        DELETETABLES
            
          
    
    
Unload Me
'frmMainform.Show
End Sub

Private Sub FareTypeGrid_Scroll()
    txtFare1.Enabled = False
    txtFare1.Visible = False
End Sub

Private Sub optallow_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index = 4 Then
        cmdFareEntry.Visible = True
        
        cmdFareEntry.SetFocus
    End If
End Sub

Private Sub optallow_LostFocus(Index As Integer)
    'RESETOPTION
End Sub

Private Sub stageGrid_DblClick()
'''    With stageGrid
'''        txtStage2.Left = .CellLeft + .Left
'''        txtStage2.Width = .CellWidth
'''        txtStage2.Height = .CellHeight
'''        txtStage2.Top = .CellTop + .Top
'''        txtStage2.Visible = True
''''        If .TextMatrix(.row, .Col) <> "" Then
''''            Text2.Text = .TextMatrix(.row, .Col)
''''            Text2.SetFocus
''''            Text2.SelStart = 0
''''            Text2.SelLength = Len(Text2)
''''        Else
''''            Text2.Visible = False
''''            If Text1.Enabled = True Then
''''                Text1.SetFocus
''''            End If
''''        End If
'''    End With
End Sub

Private Sub stageGrid_GotFocus()
    stageGrid.Enabled = True
    With stageGrid
        stageGrid.Enabled = True
        txtStage1.Enabled = True
        txtStage1.Visible = True
        txtStage1 = ""
        txtStage1.Top = .CellTop + .Top
        txtStage1.Left = .CellLeft + .Left
        txtStage1.Width = .CellWidth
        txtStage1.Height = .CellHeight
        If .TextMatrix(.row, .col) <> "" Then
         txtStage1 = .TextMatrix(.row, .col)
        End If
        txtStage1.SelStart = 0
        txtStage1.SelLength = Len(txtStage1)
        Call SendKeys("{HOME}+{END}")
        txtStage1.Visible = True
        txtStage1.SetFocus
    End With
End Sub

Private Sub txtFare1_Click()
'    With FareTypeGrid
'        .row = R
'        .Col = c
'        NOS = .Cols - .row
'    End With
End Sub

Private Sub txtFare1_GotFocus()
    txtFare1.Enabled = True
    txtFare1.Visible = True
    
End Sub

Private Sub txtFare1_LostFocus()
        
'   If FARESAVE = 0 Then
'                Msg = MsgBox("DO YOU WANT TO ESCAPE ?..", vbYesNo)
'                If Msg = vbYes Then
'                    RSql = "DELETE * FROM " & optallow(Indexno).Caption
'                    TDB.Execute (RSql)
'                    FARESAVE = 1
'                    Indexno = 6
'                    RESETFAREGRID
'                    RESETOPTION
'
'                Else
'                    txtFare1.SetFocus
'                    Exit Sub
'                End If
'    End If

    
''    With FareTypeGrid
''        If FrType = 1 Then
'''            If FARESAVE = 0 Then
'''                Msg = MsgBox("DO YOU WANT TO ESCAPE ?..", vbYesNo)
'''                If Msg = vbYes Then
'''                    RSql = "DELETE * FROM " & optallow(Indexno).Caption
'''                    TDB.Execute (RSql)
'''                    FARESAVE = 1
'''                Else
'''                    Exit Sub
'''                End If
''
''
''            If c > NOSTGS Then
''                RSql = "DELETE * FROM " & optallow(Indexno).Caption
''                TDB.Execute (RSql)
''            End If
''        ElseIf FrType = 2 Then
''            If R > NOSTGS Then
''                RSql = "DELETE * FROM " & optallow(Indexno).Caption
''                TDB.Execute (RSql)
''            End If
''        End If
''    End With
End Sub

Private Sub txtnostage_LostFocus()
'''    If cmbfaretype.Text = "TABLE" Then
'''        If val(txtnostage) < 2 Then
'''            MsgBox "NUMBER OF STAGES MUST BE MINIMUM OF TWO", vbInformation
'''            txtnostage.SetFocus
'''            Exit Sub
'''        End If
'''    End If
'''    If txtnostage <> "" Then
'''            stageGrid.Rows = val(txtnostage) + 1
''            stageGrid.Cols = 3
'            stageGrid.FormatString = "SL.NO.|STAGE NAME|DISTANCE"
'            stageGrid.ColWidth(1) = 2500
'            stageGrid.ColWidth(2) = 1500
'
'''    End If
End Sub

Private Sub cmdQuit_Click()
Msg = MsgBox("Are you sure to cancel ?", vbYesNo)
    
    If Msg = vbYes Then
    DELETETABLES
    Unload Me
    'frmMainform.Show
    Exit Sub
    Else
        txtrutcode.SetFocus
    End If
End Sub

Private Sub FareTypeGrid_Click()
'Dim MaxRow As Integer
'Dim MaxCol As Integer
Dim CRow As Integer
Dim CCol As Integer
Dim ACol As Integer
'
'    MaxRow = NOSTGS - 1
'    MaxCol = NOSTGS - 1
    
    FrType = cmbfaretype.ListIndex + 1
    With FareTypeGrid
        CRow = .row
        CCol = .col
        'ACol = NOSTGS - CRow
        If FrType = 2 Then

            ACol = CRow - .col
            If ACol <= 0 Then
                '.Col = 1
                'txtFare1.SetFocus
                Exit Sub
            End If
        ElseIf FrType = 1 Then
            If .Cols < 3 Then
                txtFare1.Visible = False
                Exit Sub
            End If
             
        End If
        
        txtFare1.Text = ""
        txtFare1.Visible = True
        txtFare1.Top = .CellTop + .Top
        txtFare1.Left = .CellLeft + .Left
        txtFare1.Width = .CellWidth
        txtFare1.Height = .CellHeight
        If .TextMatrix(.row, .col) <> "" Then txtFare1 = Round(.TextMatrix(.row, .col), 2)
                    
                    txtFare1.SelStart = 0
                    txtFare1.SelLength = Len(txtFare1)
                    txtFare1.Enabled = True
                    txtFare1.Visible = True
                    
                    txtFare1.SetFocus
                    clickflag = 1
        R = .row
        c = .col
    NOS = .Cols - .row
    End With


End Sub

Private Sub Form_Activate()
txtrutcode.SetFocus
RESETOPTION
clickflag = 0
RESETSTAGEGRID
optallow(4).Visible = False
optallow(4).Value = False
optallow(4).Enabled = False

'FARESAVE = 1
'If FARESAVE = 1 Then
'    stageGrid.Enabled = False
'End If
'


End Sub

Private Sub Form_Load()
    
    flxgrdCnt = 1
    COUNTER = 1
    StgNameCnt = 1
    toListBox = 0
    n = 1
    R = 1
    c = 1
    FARESAVE = 1
    First = False
    EnteredCount = 1
    checkflag = 0
    chkno = 0
    dotc = 0
    rowno = 1
    Indexno = 6
    stgclick = False
    txtStage1 = ""
    txtStage2 = ""
    txtStage1.Visible = False
    txtStage1.Enabled = False
    txtStage2.Visible = False
    txtFare1.Enabled = False
    Label5.Visible = False
    Label11.Visible = False
    lblallowtype.Visible = False
    lblnoofentry.Visible = False
    lblEntered.Visible = False
    FareTypeGrid.Enabled = False
    chkpass.Value = 1
    'cmdsavefare.Visible = False
   
    'CenterForm (Me)
    
    Dim BTYPE As DAO.Database
    Dim BTYPEREC As DAO.Recordset
    Set BTYPE = DAO.OpenDatabase(App.Path & "\pvt.mdb")
     
    TSQL = "SELECT * FROM BUSTYPE"
    Set BTYPEREC = BTYPE.OpenRecordset(TSQL, dbOpenDynaset)
    If BTYPEREC.RecordCount > 0 Then
        BTYPEREC.MoveFirst
        cmbBusType.Clear
        Do While Not BTYPEREC.EOF
            cmbBusType.AddItem (BTYPEREC!Name)
            ''MsgBox BTYPEREC!Name
            BTYPEREC.MoveNext
        Loop
    End If
        cmbfaretype.Text = cmbfaretype.List(0)
        cmbBusType.Text = cmbBusType.List(0)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set cn = DAO.OpenDatabase(App.Path & "\PVT.Mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
        Set TDB = DAO.OpenDatabase(App.Path & "\GBackUp.mdb")
        DELETETABLES
        'cmdnextfare.Visible = False
        
         TSQL = "SELECT * FROM " & optallow(Index).caption
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
End Sub



Private Sub optallow_Click(Index As Integer)
j = 0
clickflag = 0

''Do While j <= 4
''    If j <> Index Then
''        optallow(j).Enabled = False
''    End If
''    j = j + 1
''Loop
If Index = Indexno Then Exit Sub
If FARESAVE = 0 Then
    Msg = MsgBox("Do you want to cancel ?..", vbYesNo)
    If Msg = vbYes Then
        RSql = "DELETE * FROM " & optallow(Indexno).caption
        TDB.Execute (RSql)
        FARESAVE = 1
        
    Else
        'optallow(Indexno).Value = True
        optallow(Indexno).SetFocus
        
        Exit Sub
    End If
End If

If txtrutcode.Text <> "" Then
    TSQL = "SELECT * FROM " & optallow(Index).caption & " WHERE Route ='" & txtrutcode.Text & "'"
    Set rs = TDB.OpenRecordset(TSQL, dbOpenDynaset)
    If rs.RecordCount > 0 Then
        Msg = MsgBox("Do you want to replace existing fare details ? ", vbYesNo)
        If Msg = vbYes Then
           RSql = "DELETE * FROM " & optallow(Index).caption
            TDB.Execute (RSql)
            'RESETOPTION
            'RESETFAREGRID
            rs.Close
            'Exit Sub
        ElseIf Msg = vbNo Then
            RESETOPTION
            RESETFAREGRID
            txtFare1.Visible = False
            
            Indexno = 6
            rs.Close
            Exit Sub
        End If
        
    
    'Exit Sub
    End If
Else
    MsgBox "Please enter Routecode", vbInformation
    RESETOPTION
    Indexno = 6
    txtrutcode.SetFocus
    Exit Sub
'RESETOPTION
End If
optallow(Index).Value = True
rrow = 1
R = 0
Do While stageGrid.TextMatrix(rrow, 0) <> "" And stageGrid.TextMatrix(rrow, 1) <> ""
     
        R = R + 1
        
        If stageGrid.Rows > rrow + 1 Then
        rrow = rrow + 1
        Else
            Exit Do
        End If
Loop
    
stageGrid.Rows = R + 1
    If R = 0 Then
        R = 1
        stageGrid.Rows = 2
    End If
    
rrow = R
R = 1
'Do While R < stageGrid.Rows
'    If stageGrid.TextMatrix(R, 1) = "" Then
'        MsgBox "PLEASE ENTER STAGES NAME ", vbInformation
'        stageGrid.SetFocus
'        Exit Sub
'    Else
'        R = R + 1
'    End If
'Loop
txtnostage.Text = rrow
NOSTGS = rrow
Indexno = Index


''''If val(txtnostage.Text) < 2 Then
    If NOSTGS < 3 Then
        MsgBox "Please enter stages more than two", vbInformation
        RESETOPTION
        stageGrid.Enabled = True
        stageGrid.SetFocus
        
        Exit Sub
    End If
 RES.Close
cmdDel1.Enabled = False
 

'        RSql = "DELETE * FROM " & optallow(Index).Caption
'        TDB.Execute (RSql)
        
Select Case Index
       

Case 0
        TSQL = "SELECT * FROM " & optallow(Index).caption
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
Case 1
        TSQL = "SELECT * FROM " & optallow(Index).caption
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
Case 2
        TSQL = "SELECT * FROM " & optallow(Index).caption
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
Case 3
        TSQL = "SELECT * FROM " & optallow(Index).caption
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
Case 4
        TSQL = "SELECT * FROM " & optallow(Index).caption
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)

End Select
'        FareTypeGrid.Clear
'        FareTypeGrid.Cols = 2
'        FareTypeGrid.Rows = 2
'        txtFare1 = ""
'        txtFare2 = ""
'        flxgrdCnt = 1
'        n = 1
'        R = 1
'        c = 1
'        EnteredCount = 1
'        EntryCount = 1
'        NOS = NOSTGS
'    Label5.Visible = True
'    Label11.Visible = True
'    lblallowtype.Visible = True
'    lblnoofentry.Visible = True
'    lblEntered.Visible = True
'    FareTypeGrid.Enabled = True
    
        RESETFAREGRID
        If stageGrid.TextMatrix(txtnostage.Text, 1) <> "" Then
        stageGrid.Enabled = False
        FARELISTENTRY
        txtFare1.Enabled = True
        txtFare1.Visible = True
        txtFare1.SetFocus
        
        Else
            MsgBox "Please enter stage details first", vbInformation, gblstrPrjTitle
            RESETOPTION
            
        Exit Sub
        End If
End Sub

Private Sub stageGrid_Click()
On err GoTo err

'''    NOSTGS = val(txtnostage.Text)
'''    If val(txtnostage.Text) = 0 Then
        'stageGrid.Enabled = False
'''        txtnostage.SetFocus
'''        Exit Sub
'''    Else
'''      stageGrid.Enabled = True
'''    End If
    stageGrid.Enabled = True
    With stageGrid
        If .col = 2 Then
            If .TextMatrix(.row, 1) = "" Then
            txtStage1.Visible = False
            Exit Sub
            End If
        End If
    End With
    
    With stageGrid
        If .col = 0 And .row <> 0 Then
            txtStage1.Visible = False
            Exit Sub
                  
        End If
    End With
    
    With stageGrid
        stgclick = True
        stageGrid.Enabled = True
        txtStage1.Enabled = True
        txtStage1.Visible = True
        txtStage1 = ""
        txtStage1.Top = .CellTop + .Top
        txtStage1.Left = .CellLeft + .Left
        txtStage1.Width = .CellWidth
        txtStage1.Height = .CellHeight
        If .TextMatrix(.row, .col) <> "" Then
         txtStage1 = .TextMatrix(.row, .col)
        
        txtStage1.SelStart = 0
        txtStage1.SelLength = Len(txtStage1)
        Call SendKeys("{HOME}+{END}")
        End If
        txtStage1.SetFocus
    End With
    
    
    ';showcol
    Exit Sub
err:
MsgBox err.Number & "," & err.Description
End Sub

Private Sub txtDistance_KeyPress(KeyAscii As Integer)
    
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
    
    
    Select Case KeyAscii
        Case 48 To 57, 8, 46
            KeyAscii = KeyAscii
        Case 13
            stageGrid.TextMatrix(StgNameCnt - 1, 2) = val(txtDistance.Text)
            StageDistance(StgNameCnt - 1) = val(txtDistance)

            txtDistance = ""
            If StgNameCnt <= val(txtnostage) Then
                txtStgName.Enabled = True
                txtStgName.Text = ""
                txtStgName.SetFocus
                txtStgName.BackColor = &HC0EF00
            End If
            txtDistance.Enabled = False
            lblnoofstage.caption = StgNameCnt - 1 & "/" & NOSTGS
            If StgNameCnt > NOSTGS Then
                'cmdStgSave.Enabled = True
                'cmdStgSave.SetFocus
               
'                FARELISTENTRY            // 281010
'                txtFare1.Enabled = True
'                txtFare1.Visible = True
'                txtFare1.SetFocus
                optallow(0).SetFocus
                
            End If
            stageGrid.CellBackColor = &H80000005
            
            If StgNameCnt <= NOSTGS Then
                stageGrid.row = StgNameCnt
                stageGrid.col = 1
                stageGrid.CellBackColor = &H900FFF
            End If
            txtDistance.BackColor = &H80000005
        Case Else
            KeyAscii = 0
    End Select

End Sub




Private Sub txtFare1_KeyPress(KeyAscii As Integer)
    
Static LastText As String
Static SecondTime1 As Boolean
Const MaxDecimal As Integer = 1
Const MaxWhole As Integer = 4
'optallow(Index).Value = True
    If cmbfaretype.Text = "TABLE" Then FrType = 1
    If cmbfaretype.Text = "GRAPH" Then FrType = 2
        NOSTGS = val(txtnostage.Text)
'        TEST = NOSTGS
'        NOS = NOSTGS
'        NSatge = NOS
'        NoOfEntries = (NSatge * (NSatge - 1)) / 2
'
    If TextBoxValidityNumeric(KeyAscii) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    

    If KeyAscii = 9 Then
        If txtFare1.Visible = True Then txtFare1.Visible = False
    End If
    
'    If KeyAscii = 46 Then
'        dotc = dotc + 1
'        If dotc > 1 Then
'        KeyAscii = 0
'        'txtFare1.Text = ""
'        'dotc = 0
'        Exit Sub
'        End If
'    End If
'

    With txtFare1
        If Not SecondTime1 Then
            If .Text Like "*[!0-9.]*" _
            Or .Text Like "*.*.*" _
            Or .Text Like String$(MaxWhole, "#") & "[!.]" Then
                SecondTime1 = True
                .Text = LastText
                .SelStart = Len(.Text)
            Else
                LastText = .Text
            End If
        End If
    End With

    SecondTime1 = False

'
'
'    If KeyAscii = 13 Then
''        TSQL = "SELECT * FROM TMPFARE"
''        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'
'        If val(txtFare1) < val(txtminfare) Then
'            MsgBox "Minimum Fare is " & txtminfare, vbInformation, "Route"
'            txtFare1.SelStart = 0
'            txtFare1.SelLength = Len(txtFare1)
'            txtFare1.SetFocus
'            Exit Sub
'        End If
'        With FareTypeGrid
'            If Not SAVEFLAG = False Or FARESAVEFLAG = False Then
'
'                If val(txtFare1) < val(txtminfare) Then
'                    MsgBox "Minimum fare is  " & txtminfare, vbExclamation
'                    txtFare1 = ""
'                    Exit Sub
'                End If
'                If FrType = 2 Then
'
'                    .TextMatrix(.row, .Col) = Round(Trim(txtFare1), 2)
'                    If First = False And FARESAVEFLAG = True Then
'                        RES.AddNew
'                        RES!row = .row
'                        RES!Col = .Col
'                        RES!Fare = Round((.TextMatrix(.row, .Col)), 2)
'                        RES!Route = txtrutcode.Text
'                        RES.Update
''                        TRES.AddNew
''                        TRES!Save = False
''                        TRES!NoOfEntries = EnteredCount
''                        TRES.Update
'                        lblEntered.Caption = EnteredCount
'                        First = True
'                        .Cols = .Cols + 1
'                        .Col = .Col + 1
'                        c = c + 1
'                        EnteredCount = EnteredCount + 1
'                        EntryCount = EntryCount + 1
'                        txtFare2.Top = .CellTop + .Top
'                        txtFare2.Left = .CellLeft + .Left
'                        txtFare1.Width = .CellWidth
'                        txtFare1.Height = .CellHeight
'                    Else
'                        RES.Close
'                        TSQL = "SELECT * FROM TMPFARE WHERE Row = " & .row & " AND Col = " & .Col & " AND route= '" & lblShowRoute & "'"
'                        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'                        If RES.RecordCount <> 0 Then
'                        RES.Edit
'                        RES!Fare = Round((.TextMatrix(.row, .Col)), 2)
'                        RES.Update
'                        End If
'                        RES.Close
'                        .row = R
'                        .Col = c
'                        TSQL = "SELECT * FROM TMPFARE"
'                        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'                        RES.MoveLast
'                    End If
'                    If EntryCount > NoOfEntries Or EnteredCount > NoOfEntries Then
'                        txtFare1 = ""
'                        txtFare1.Visible = False
'                        Exit Sub
'                    End If
'                    txtFare1 = ""
'                    txtFare1.Visible = False
'                    txtFare2.Visible = True
'                    txtFare2.SetFocus
'
'                ElseIf FrType = 1 Then
'                    If First = False And FARESAVEFLAG = True Then
'                        First = True
'                        c = c + 1
'                        .TextMatrix(.row, .Col) = 0
'                        RES.AddNew
'                        RES!row = .row
'                        RES!Col = .Col
'                        RES!Fare = Round((.TextMatrix(.row, .Col)), 2)
'                        RES!Route = txtrutcode.Text
'                        RES.Update
''                        TRES.AddNew
''                        TRES!Save = False
'                        .Cols = .Cols + 2
'                        .Col = c
'
'                        .TextMatrix(.row, c) = val(txtFare1)
''                        TRES!NoOfEntries = 2
''                        TRES.Update
'                        RES.AddNew
'                        RES!row = .row
'                        RES!Col = .Col
'                        RES!Fare = Round((.TextMatrix(.row, c)), 2)
'                        RES!Route = txtrutcode.Text
'                        RES.Update
'
'                        c = c + 1
'                        .Col = c
'
'                        EnteredCount = EnteredCount + 1
'                        lblEntered.Caption = EnteredCount
'                        EntryCount = EntryCount + 1
'                        If NOSTGS = 2 Then
'                            txtFare1 = ""
'                            txtFare1.Visible = False
'                            cmdFareEntry.Enabled = True
'                            cmdFareEntry.SetFocus
'                            c = c - 1
'                            .Cols = .Cols - 1
'                            Exit Sub
'                        End If
'                        txtFare1.Visible = False
'                        txtFare2.Visible = True
'                        txtFare2.Top = .CellTop + .Top
'                        txtFare2.Left = .CellLeft + .Left
'                        txtFare2.SetFocus
'                    Else
'                        If EntryCount < NOS Then
'                            .TextMatrix(.row, .Col) = Round(val(txtFare1), 2)
'                            lblEntered.Caption = EnteredCount
'
'                            RES.Close
'                            TSQL = "SELECT * FROM TMPFARE WHERE Row = " & .row & " AND Col = " & .Col & " AND route= '" & lblShowRoute & "'"
'                            Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'                            If RES.RecordCount > 0 Then
'                                RES.Edit
'                            Else
'                                RES.AddNew
'                            End If
'                            RES!Fare = Round((.TextMatrix(.row, .Col)), 1)
'                            RES.Update
'                            RES.Close
'                            TSQL = "SELECT * FROM TMPFARE"
'                            Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'
'                            .Col = c
'                            .row = R
'                            txtFare2.Visible = True
'                            txtFare1.Visible = False
'                            txtFare2.Top = .CellTop + .Top
'                            txtFare2.Left = .CellLeft + .Left
'                            txtFare2.SetFocus
'                        Else
'                            .TextMatrix(.row, .Col) = Round(val(txtFare1), 2)
'
'                            RES.Close
'                            TSQL = "SELECT * FROM TMPFARE WHERE Row = " & .row & " AND Col = " & .Col & " AND Route= '" & lblShowRoute & "'"
'                            Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'                            RES.Edit
'                            RES!Fare = Round((.TextMatrix(.row, .Col)), 2)
'                            RES.Update
'                            RES.Close
'
'                            TSQL = "SELECT * FROM TMPFARE"
'                            Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'
'                            txtFare1 = ""
'                            txtFare1.Visible = False
'                            Exit Sub
'                        End If
'                    End If
'                    txtFare1.Visible = False
'                End If
'            Else
'                .TextMatrix(.row, .Col) = Round(val(txtFare1), 2)
'                txtFare1.Visible = False
'
'                TSQL = "SELECT * FROM TMPGRAPH WHERE ROW= " & .row & "AND COL = " & .Col
'                Set TRES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
'                If TRES.RecordCount > 0 Then
'                    TRES.Edit
'                    TRES!Fare = Round((val(.TextMatrix(.row, .Col))), 2)
'                Else
'                    TRES.AddNew
'                    TRES!Fare = Round((val(.TextMatrix(.row, .Col))), 2)
'                End If
'                TRES.Update
'
'            End If
'        End With
'    End If

With FareTypeGrid
        
       
    
        If KeyAscii = 13 And val(txtFare1.Text) <> 0 Then
        
            If txtFare1.Text = "." And Len(txtFare1.Text) < 2 Then
                KeyAscii = 0
                Exit Sub
            End If

        
        'if dotc > 1 then KeyAscii = 0 end sub end if
        txtFare1.Text = Round(txtFare1.Text, 2)
        cmdcleargrid.Visible = True
        cmdDel1.Enabled = False
        
        
'        cmdFareEntry.Visible = False '30.12.2010
'        cmdFareEntry.Enabled = False '30.12.2010
        FARESAVE = 0    ' to ensure the fare entry completed
            '.TextMatrix(.row, .Col) = txtFare1.Text
            'txtFare1.Text = ""
            If FrType = 1 Then
                    RSql = "SELECT * FROM " & optallow(Indexno).caption & " where row =" & .row & " and col =" & .col & " and route ='" & txtrutcode.Text & "'"
                    Set TRES = TDB.OpenRecordset(RSql, dbOpenDynaset)
                    If TRES.RecordCount > 0 Then
                        If .col = 1 Then
                                    .TextMatrix(1, 1) = 0
                                    TRES.Edit
                                    TRES!row = .row
                                    TRES!col = .col
                                    TRES!FARE = Round((.TextMatrix(.row, .col)), 2)
                                    TRES!Route = txtrutcode.Text
                                    TRES.Update
                                    
                                    .col = .col + 1
                                    .TextMatrix(.row, .col) = txtFare1.Text
                                   
                                    TRES.Edit
                                    TRES!row = .row
                                    TRES!col = .col
                                    TRES!FARE = Round((.TextMatrix(.row, .col)), 2)
                                    TRES!Route = txtrutcode.Text
                                    TRES.Update
                                    
                            Else
'
'                                    If (Round((txtFare1.Text), 2) < oldfare) Then
'                                        MsgBox "Can't enter less amount than previous"
'                                        txtFare1.Text = ""
'                                        txtFare1.SetFocus
'                                        Exit Sub
'                                    End If
'                                    oldfare = Round((.TextMatrix(.row, .Col)), 2) '14.12.2010
                                    .TextMatrix(.row, .col) = txtFare1.Text
                                   
                                    TRES.Edit
                                    TRES!row = .row
                                    TRES!col = .col
                                    TRES!FARE = Round((.TextMatrix(.row, .col)), 2)
                                    TRES!Route = txtrutcode.Text
                                    TRES.Update
                                    
                            End If
                        Else
                            If .col = 1 Then
                                    .TextMatrix(1, 1) = 0
                                    RES.AddNew
                                    RES!row = .row
                                    RES!col = .col
                                    RES!FARE = Round((.TextMatrix(.row, .col)), 2)
                                    RES!Route = txtrutcode.Text
                                    RES.Update
                                    
                                    .col = .col + 1
''                                    oldfare = Round((.TextMatrix(.row, .Col)), 2) '14.12.2010
''                                    If (Round((txtFare1.Text), 2) < oldfare) Then
''                                        MsgBox "Can't enter less amount than previous"
''                                        txtFare1.Text = ""
''                                        txtFare1.SetFocus
''                                        Exit Sub
''                                    End If
                                   .TextMatrix(.row, .col) = txtFare1.Text
                                   
                                    RES.AddNew
                                    RES!row = .row
                                    RES!col = .col
                                    RES!FARE = Round((.TextMatrix(.row, .col)), 2)
                                    RES!Route = txtrutcode.Text
                                    RES.Update
                                    
                            Else
                                   
''                                    If (Round((txtFare1.Text), 2) < oldfare) Then
''                                        MsgBox "Can't enter less amount than previous"
''                                        txtFare1.Text = ""
''                                        txtFare1.SetFocus
''                                        Exit Sub
''                                    End If
                                    
                                    .TextMatrix(.row, .col) = txtFare1.Text
''                                    oldfare = Round((.TextMatrix(.row, .Col)), 2) '14.12.2010
                                    RES.AddNew
                                    RES!row = .row
                                    RES!col = .col
                                    RES!FARE = Round((.TextMatrix(.row, .col)), 2)
                                    RES!Route = txtrutcode.Text
                                    RES.Update
                            End If
                        End If
                        TRES.Close
                If .col < NOSTGS Then
                        If clickflag = 0 Then
                            .col = .col + 1
                        
                            txtFare1.Text = ""
                            txtFare1.Enabled = True
                            txtFare1.Visible = True
                            txtFare1.Top = .CellTop + .Top
                            txtFare1.Left = .CellLeft + .Left
                            txtFare1.Width = .CellWidth
                            txtFare1.Height = .CellHeight
                            If .TextMatrix(.row, .col) <> "" Then txtFare1 = Round(.TextMatrix(.row, .col), 2)
                            txtFare1.Enabled = True
                            txtFare1.Visible = True
                            txtFare1.SelStart = 0
                            txtFare1.SelLength = Len(txtFare1)
                            Call SendKeys("{HOME}+{END}")
                            txtFare1.SetFocus
                            R = .row
                            c = .col
                        Else
                            .row = R
                            .col = c
                            .col = .col + 1
                            txtFare1.Text = ""
                            txtFare1.Enabled = True
                            txtFare1.Visible = True
                            txtFare1.Top = .CellTop + .Top
                            txtFare1.Left = .CellLeft + .Left
                            txtFare1.Width = .CellWidth
                            txtFare1.Height = .CellHeight
                            If .TextMatrix(.row, .col) <> "" Then txtFare1 = Round(.TextMatrix(.row, .col), 2)
                            txtFare1.Enabled = True
                            txtFare1.Visible = True
                            txtFare1.SelStart = 0
                            txtFare1.SelLength = Len(txtFare1)
                            Call SendKeys("{HOME}+{END}")
                            txtFare1.SetFocus
                            
                            
                        End If
                Else
                    
                    FARESAVE = 1
                    txtFare1.Text = ""
                    txtFare1.Visible = False
                'cmd1.SetFocus
                    cmdFareEntry.Enabled = True
                     cmdFareEntry.Visible = True
                    Msg = MsgBox("Are you sure to save ? ", vbYesNo, "SAVE FARE")
                    If Msg = vbNo Then
                        RSql = "DELETE * FROM " & optallow(Indexno).caption
                        TDB.Execute (RSql)
                         
                        'Exit Sub
                    End If
                    RESETFAREGRID
                    RESETOPTION
                    'If Indexno < 4 Then
                    If Indexno < 3 Then
                         optallow(Indexno + 1).SetFocus
                         'optallow_Click (Indexno + 1)
                    Else
                        RESETFAREGRID
                        cmdFareEntry.SetFocus
                    End If
                    cmbBusType.Enabled = False
                    cmbfaretype.Enabled = False
                    clickflag = 0
                    Exit Sub
                End If
        ElseIf FrType = 2 Then
                
'            If .row + .Col <= NOSTGS Then 2112
              If .row + 1 - .col > 0 Then
                        .TextMatrix(.row, .col) = txtFare1.Text
                    RSql = "SELECT * FROM " & optallow(Indexno).caption & " where row =" & .row & " and col =" & .col & " and route ='" & txtrutcode.Text & "'"
                    Set TRES = TDB.OpenRecordset(RSql, dbOpenDynaset)
                    If TRES.RecordCount > 0 Then
                        TRES.Edit
                        TRES!row = .row
                        TRES!col = .col
                        TRES!FARE = Round((.TextMatrix(.row, .col)), 2)
                        TRES!Route = txtrutcode.Text
                        TRES.Update
                   Else
                        RES.AddNew
                        RES!row = .row
                        RES!col = .col
                        RES!FARE = Round((.TextMatrix(.row, .col)), 2)
                        RES!Route = txtrutcode.Text
                        RES.Update
                    End If
                    TRES.Close
                
                'If .Col < .row - 1 Then '23122010
                If .row < NOSTGS Then
                        If clickflag = 0 Then
                            '.Col = .Col + 1 '2312
                            .row = .row + 1
                            txtFare1.Text = ""
                            txtFare1.Enabled = True
                            txtFare1.Visible = True
                            txtFare1.Top = .CellTop + .Top
                            txtFare1.Left = .CellLeft + .Left
                            txtFare1.Width = .CellWidth
                            txtFare1.Height = .CellHeight
                            If .TextMatrix(.row, .col) <> "" Then txtFare1 = Round(.TextMatrix(.row, .col), 2)
                            txtFare1.SelStart = 0
                            txtFare1.SelLength = Len(txtFare1)
                            txtFare1.Enabled = True
                            txtFare1.Visible = True
                            txtFare1.SetFocus
                            R = .row
                            c = .col
                        Else
                            txtFare1.Visible = False
                            .row = R
                            .col = c
                            '.Col = .Col + 1 '23122010
                            .row = .row + 1
                            txtFare1.Text = ""
                            txtFare1.Enabled = True
                            txtFare1.Visible = True
                            txtFare1.Top = .CellTop + .Top
                            txtFare1.Left = .CellLeft + .Left
                            txtFare1.Width = .CellWidth
                            txtFare1.Height = .CellHeight
                            If .TextMatrix(.row, .col) <> "" Then txtFare1 = Round(.TextMatrix(.row, .col), 2)
                            txtFare1.Enabled = True
                            txtFare1.Visible = True
                            txtFare1.SelStart = 0
                            txtFare1.SelLength = Len(txtFare1)
                            txtFare1.SetFocus
                          End If
                Else
                 
'                    .Col = 1
'                    If .row < NOSTGS Then
'                    .row = .row + 1
'                    NOS = .Cols - .row
                    If .col < NOSTGS - 1 Then
    '                    .row = .row + 1 211210
                    .col = .col + 1
                    .row = .col + 1
                    NOS = .Cols - .row
                
                    txtFare1.Text = ""
                    txtFare1.Enabled = True
                    txtFare1.Visible = True
                    txtFare1.Top = .CellTop + .Top
                    txtFare1.Left = .CellLeft + .Left
                    txtFare1.Width = .CellWidth
                    txtFare1.Height = .CellHeight
                    If .TextMatrix(.row, .col) <> "" Then txtFare1 = Round(.TextMatrix(.row, .col), 2)
                        txtFare1.SelStart = 0
                        txtFare1.SelLength = Len(txtFare1)
                        txtFare1.Enabled = True
                        txtFare1.Visible = True
                        txtFare1.SetFocus
                    Else
                        FARESAVE = 1
                        txtFare1.Text = ""
                        txtFare1.Visible = False
    '''''                    'cmd1.SetFocus
                        cmdFareEntry.Visible = True
                        cmdFareEntry.Enabled = True
                        Msg = MsgBox("Are you sure to save ? ", vbYesNo, "SAVE FARE")
                            If Msg = vbNo Then
                                RSql = "DELETE * FROM " & optallow(Indexno).caption
                                TDB.Execute (RSql)
                                'optallow_Click (Indexno)
                                'Exit Sub
                            ElseIf Msg = vbYes Then
                                RSql = "DELETE * FROM " & optallow(Indexno).caption
                                TDB.Execute (RSql)
                                ADD_FARE_TO_TABLE
                            
                            End If
                            RESETFAREGRID
                            RESETOPTION
                            'If Indexno < 4 Then
                            If Indexno < 3 Then
                                optallow(Indexno + 1).SetFocus
                                'optallow_Click (Indexno + 1)
                            Else
                                cmdFareEntry.SetFocus
                            End If
                            cmbBusType.Enabled = False
                            cmbfaretype.Enabled = False
                    End If
                End If
             End If
        End If
            clickflag = 0
        End If
    End With
End Sub




Private Sub txtFare2_KeyPress(KeyAscii As Integer)

Static LastText As String
Static SecondTime As Boolean
Const MaxDecimal As Integer = 1
Const MaxWhole As Integer = 4
  
    If TextBoxValidityNumeric(KeyAscii) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    With txtFare2
        If Not SecondTime Then
            If .Text Like "*[!0-9.]*" _
            Or .Text Like "*.*.*" _
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
    
    
    
    
    If KeyAscii = 13 Then
        
        With FareTypeGrid
            
            If val(txtFare2) < val(txtminfare) Then
                MsgBox "Minimum fare is  " & txtminfare, vbExclamation
                txtFare2 = ""
                Exit Sub
            End If
            If FrType = 2 Then
            
                .TextMatrix(R, c) = Round(Trim(txtFare2), 2)
                If RES.RecordCount > 0 Then RES.MoveLast
                RES.AddNew
                RES!row = .row
                RES!col = .col
                RES!FARE = Round((.TextMatrix(R, c)), 2)
                RES!Route = txtrutcode.Text
                RES.Update
'                TRES.MoveFirst
'                TRES.Edit
'                TRES!Save = False
'                TRES!NoOfEntries = EnteredCount
'                TRES.Update
                
                lblEntered.caption = EnteredCount
                If EntryCount < NoOfEntries Then
                    EntryCount = EntryCount + 1
                    If c >= NOS - 1 Then
                        R = R + 1
                        c = 1
                        NOS = NOS - 1
                        .Rows = .Rows + 1
                        .col = 1
                        .row = R
                        EnteredCount = EnteredCount + 1
                    Else
                        If R = 1 Then
                            .Cols = .Cols + 1
                            EnteredCount = EnteredCount + 1
                        Else
                            EnteredCount = EnteredCount + 1
                        End If
                        .col = c + 1
                        .row = R
                        c = c + 1
                    End If
                    txtFare2.Top = .CellTop + .Top
                    txtFare2.Left = .CellLeft + .Left
                    txtFare2 = ""
                    txtFare1.Visible = False
                    txtFare2.SetFocus
                Else
                    NOS = NOS - 1
                End If
                txtFare2 = ""
''''''''''''''''''''''''CHECKING THE END ''''''''''''''''''''''''

                If NOS <= 1 Then
                    EntryCount = EntryCount + 1
                    txtFare2.Visible = False
'                    cmdFareEntry.Enabled = True
'                    cmdFareEntry.SetFocus
                    MsgBox "Fare for '" & optallow(Index).caption & "' saved", vbInformation, gblstrPrjTitle
                    FareTypeGrid.Clear
                    FareTypeGrid.Cols = 2
                    FareTypeGrid.Rows = 2
                    txtFare1 = ""
                    txtFare2 = ""
                    txtFare1.Visible = False
                    txtFare2.Visible = False
                    RESETOPTION
'                    End If
                    Exit Sub
                End If
            
            ElseIf FrType = 1 Then
                .TextMatrix(.row, c) = Round(val(txtFare2), 2)
                
                RES.AddNew
                RES!row = .row
                RES!col = .col
                RES!FARE = Round((.TextMatrix(R, c)), 2)
                RES!Route = txtrutcode.Text
                RES.Update
                lblEntered.caption = EnteredCount + 1
                
                EntryCount = EntryCount + 1
''''''''''''''''''''''''CHECKING THE END ''''''''''''''''''''''''
                If EntryCount > NOS - 1 Then
                    txtFare2 = ""
                    txtFare2.Visible = False
'                    CHECKALLOWABLES
'                    If Not checkflag = 1 Then
'                        txtFare1.Enabled = True
'                        txtFare1.Visible = True
'                        txtFare1.SetFocus
'                    Else
'                        cmdsavefare.Visible = True
'                        cmdsavefare.Enabled = True
'                        cmdsavefare.SetFocus
                    MsgBox "Fare for '" & optallow(Index).caption & "' saved", vbInformation, gblstrPrjTitle
                    FareTypeGrid.Clear
                    FareTypeGrid.Cols = 2
                    FareTypeGrid.Rows = 2
                    txtFare1 = ""
                    txtFare2 = ""
                    txtFare1.Visible = False
                    txtFare2.Visible = False
                    RESETOPTION
'                    End If
                    Exit Sub
                Else
                    c = c + 1
                    .Cols = .Cols + 1
                    .col = c
                    EnteredCount = EnteredCount + 1
                    txtFare2 = ""
                    txtFare2.Top = .CellTop + .Top
                    txtFare2.Left = .CellLeft + .Left
                    txtFare2.SetFocus
                End If
            End If
            
        End With
    End If
End Sub


Private Sub txtnostage_KeyPress(KeyAscii As Integer)
       
    If KeyAscii = 13 Then
        If txtnostage <> "" Then
            stageGrid.Rows = val(txtnostage) + 1
'            stageGrid.Cols = 3
'            stageGrid.FormatString = "SL.NO.|STAGE NAME|DISTANCE"
'            stageGrid.ColWidth(1) = 2500
'            stageGrid.ColWidth(2) = 1500
          
            
            NOSTGS = val(txtnostage)
            cmbfaretype.SetFocus
        End If
    End If
End Sub

Private Sub txtrutcode_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 39 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        If txtrutcode.Text <> "" Then
            TSQL = "SELECT * FROM Route WHERE rutcode ='" & txtrutcode.Text & "'"
            Set rs = cn.OpenRecordset(TSQL, dbOpenDynaset)
            If rs.RecordCount > 0 Then
            MsgBox "ROUTE CODE ALREADY EXISTS", vbInformation
            txtrutcode.Text = ""
            txtrutcode.SelLength = Len(txtrutcode.Text)
            Call SendKeys("{HOME}+{END}")
            txtrutcode.SetFocus
            Exit Sub
            End If
            rs.Close
        
        End If
        txtrutcode.Text = UCase(txtrutcode.Text)
        txtrutname.SetFocus
    End If
End Sub

Private Sub txtrutname_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 39 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        txtrutname.Text = UCase(txtrutname.Text)
        cmbfaretype.Enabled = True
        cmbfaretype.SetFocus
    End If
End Sub


Private Sub txtStage1_KeyPress(KeyAscii As Integer)
 
 With stageGrid
                If .col = 2 Then
                    If TextBoxValidityNumeric(KeyAscii) > 0 Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
 
 
'
'        If .TextMatrix(.row, .Col) <> "" Then
'            txtStage1 = .TextMatrix(.row, .Col)
'        End If
'        txtStage1.SelStart = 0
'        txtStage1.SelLength = Len(txtStage1)
'
        
        If KeyAscii = 13 Then
            If .col = 1 And Trim(txtStage1.Text) = "" Then Exit Sub
                
                .TextMatrix(.row, 0) = .row
                .TextMatrix(.row, 2) = 0
                If .col = 1 Then
                .TextMatrix(.row, .col) = UCase(Trim(txtStage1.Text))
                ElseIf .col = 2 Then
                .TextMatrix(.row, .col) = val(txtStage1.Text)
                End If
                txtStage1.Text = ""
                If .col = 1 Then
                    .col = .col + 1
                    txtStage1 = ""
                    txtStage1.Top = .CellTop + .Top
                    txtStage1.Left = .CellLeft + .Left
                    txtStage1.Width = .CellWidth
                    txtStage1.Height = .CellHeight - 10
                    If .TextMatrix(.row, .col) <> "" Then
                        txtStage1 = .TextMatrix(.row, .col)
                    End If
                    txtStage1.SelStart = 0
                    txtStage1.SelLength = Len(txtStage1)
                    Call SendKeys("{HOME}+{END}")
                    txtStage1.SetFocus
               
               ElseIf .col = 2 Then
                'MALAYALAM IS NOT AVAILABLE FOR NOW LOCAL LANGUAGE VALUE FOR MALAYALAM IS 1
'********************For Language Only****************************
        If LocalLanguage > 0 And LocalLanguage <> 4 Then 'AND PART ADDED BY SYAM<>1
            strLanguageStage = ""
'            Load H_Convert
'            H_Convert.Show vbModal
            'frmBMPEditor.Show vbModal
            If strLanguageStage = "" Then
                strLanguageStage = " "
            End If
            stageGrid.TextMatrix(.row, 3) = strLanguageStage
        End If
                '.row = .row + 1
                txtStage1.SetFocus
'********************For Language Only****************************
               
                
                    .col = 1
'''                    rrow = .row
'''                    If rrow <> .Rows - 1 Then
                        .Rows = .Rows + 1
                        .row = .row + 1
                        txtStage1 = ""
                        txtStage1.Top = .CellTop + .Top
                        txtStage1.Left = .CellLeft + .Left
                        txtStage1.Width = .CellWidth
                        txtStage1.Height = .CellHeight - 10
                        If .TextMatrix(.row, .col) <> "" Then
                            txtStage1 = .TextMatrix(.row, .col)
                        End If
                        txtStage1.SelStart = 0
                        txtStage1.SelLength = Len(txtFare1)
                        Call SendKeys("{HOME}+{END}")
                        txtStage1.Visible = True
                        txtStage1.SetFocus
'''                    Else
'''                        txtStage1.Visible = False
'''                        txtStage1.Enabled = False
'''                        cmbBustype.SetFocus
'''                        Exit Sub
'''                    End If
                End If
            
        End If
    End With
End Sub



Private Sub txtStage1_LostFocus()
    txtStage1.Visible = False
End Sub

Private Sub txtStage2_KeyPress(KeyAscii As Integer)
    
    With stageGrid
        .row = rowno
    If MaxRow >= rowno Then
        
        txtStage2.Top = .CellTop + .Top
        txtStage2.Left = .CellLeft + .Left
        txtStage2.Width = .CellWidth
        txtStage2.Height = .CellHeight - 50
         If KeyAscii = 13 Then
    
           .TextMatrix(rowno, 2) = Round(val(txtStage2.Text))
           
           txtStage2.SelStart = 0
           txtStage2.SelLength = Len(txtFare2)
           Call SendKeys("{HOME}+{END}")
            
           txtStage2 = ""
           txtStage2.Visible = False
           txtStage1.Visible = True
           rowno = rowno + 1
        If MaxRow >= rowno Then
            .col = c + 1
            .row = rowno
            txtStage1.Top = .CellTop + .Top
            txtStage1.Left = .CellLeft + .Left
            txtStage1.Width = .CellWidth
            txtStage1.Height = .CellHeight - 50
            
'MALAYALAM IS NOT AVAILABLE FOR NOW LOCAL LANGUAGE VALUE FOR MALAYALAM IS 1
'********************For Language Only****************************
        If LocalLanguage > 0 And LocalLanguage <> 4 Then 'AND PART ADDED BY SYAM<>1
            strLanguageStage = ""
'            Load H_Convert
'            H_Convert.Show vbModal
            frmBMPEditor.Show vbModal
            If strLanguageStage = "" Then
                strLanguageStage = " "
            End If
            stageGrid.TextMatrix(rowno, 3) = strLanguageStage
        End If
'********************For Language Only****************************
            
            
            
            
            
            
           txtStage1.SetFocus
           Else
           txtStage2.Enabled = False
           txtStage2.Visible = False
           txtStage1.Enabled = False
           txtStage1.Visible = False
            'optallow(0).SetFocus
        End If
            
           
            
        End If
    End If
     End With
End Sub

Private Sub txtStgName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
      Case Else
        KeyAscii = 0
    End Select
    If KeyAscii = 13 Then
    
      STAGENamesOneByOne
      stageGrid.TextMatrix(0, 0) = "S.No"
      stageGrid.TextMatrix(0, 1) = "STAGE NAME"
      stageGrid.TextMatrix(0, 2) = "KM"
      If LocalLanguage > 0 Then
        stageGrid.TextMatrix(0, 3) = " STG"
      End If
      'txtStgName.BackColor = &H80000005
End If
End Sub


