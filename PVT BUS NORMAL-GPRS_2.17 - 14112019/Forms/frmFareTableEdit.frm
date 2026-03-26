VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmFareTableEdit 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fare Table"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11190
   Icon            =   "frmFareTableEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton Command1 
      Height          =   375
      Left            =   8565
      TabIndex        =   8
      Top             =   4725
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
      MICON           =   "frmFareTableEdit.frx":0CCA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton Command2 
      Height          =   375
      Left            =   9900
      TabIndex        =   7
      Top             =   4725
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
      MICON           =   "frmFareTableEdit.frx":0CE6
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4350
      Left            =   135
      TabIndex        =   1
      Top             =   150
      Width           =   10995
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   350
         Width           =   1320
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   405
         Width           =   810
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1365
         TabIndex        =   3
         Top             =   1110
         Visible         =   0   'False
         Width           =   960
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2235
         Left            =   570
         TabIndex        =   0
         Top             =   1005
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   3942
         _Version        =   393216
         Cols            =   5
         ForeColor       =   16711808
         BackColorFixed  =   14737632
         TextStyle       =   1
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Stages"
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
         Left            =   5490
         TabIndex        =   5
         Top             =   390
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Route:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1635
         TabIndex        =   4
         Top             =   420
         Width           =   1650
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fare Table"
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
      Left            =   4485
      TabIndex        =   6
      Top             =   -630
      Width           =   2325
   End
End
Attribute VB_Name = "frmFareTableEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnn_faretable As Database
Dim rs_faretable As DAO.Recordset
Dim RS_FARETABLE_TEMP As DAO.Recordset
Dim rs_faretable_count As DAO.Recordset
Dim sql As String
Dim sql1 As String
Dim TOTAFARE As Integer
Dim Handle As Integer
Dim MinFare As Integer
Dim myflg  As Boolean

Private Sub Combo1_Click()
    RSql = "SELECT RUTCODE,NOSTAGE,MINFARE FROM ROUTE WHERE RUTCODE = '" & Combo1.Text & "'"
    Set DB = DAO.OpenDatabase(App.Path & "\Pvt.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
    If RES.RecordCount > 0 Then
        NOSTGS = RES!nostage
        MinFare = RES!MinFare
        RES.Close
        Text1.Text = NOSTGS
        LoadFareTable (NOSTGS)
    End If
End Sub
Private Sub Command1_Click()
On Error GoTo err
Dim RES3 As DAO.Recordset
Dim sql3 As String
Dim Row1 As Integer
Dim Count As Integer
Dim flexrowcount As Integer
Dim flexcolcount As Integer
Dim totalcount As Integer
Dim tempcolcount As Integer
    
    sql3 = "SELECT FARE FROM FARE WHERE ROUTE='" & Combo1.Text & "'"
    Set RES3 = cnn_faretable.OpenRecordset(sql3, dbOpenDynaset)
    If RES3.RecordCount <> 0 Then
        With RES3
            .MoveFirst
            Do While .EOF = False
                .Delete
                .MoveNext
            Loop
        End With
    End If
    If Text1.Text = "" Then
        MsgBox "No Route Details!", vbExclamation, "Fare Edit"
        Exit Sub
    End If
    totalcount = 1
    flexrowcount = Text1.Text
    flexcolcount = TOTAFARE
    tempcolcount = flexrowcount
    Row1 = 1
    Count = 1
    
    Set rs_faretable = cnn_faretable.OpenRecordset("FARE", dbOpenDynaset)
    If rs_faretable.RecordCount <> 0 Then rs_faretable.MoveLast
    rs_faretable.AddNew
    rs_faretable!row = Row1
    rs_faretable!Col = Count
    rs_faretable!FARE = 0
    rs_faretable!Route = Combo1.Text
    rs_faretable.Update

    Do While totalcount < flexcolcount
    
        With MSFlexGrid1
            
            With rs_faretable
                .AddNew
                !row = Row1
                !Col = Count + 1
                !FARE = MSFlexGrid1.TextMatrix(Row1, Count)
                !Route = Combo1.Text  'RouteNo
                .Update
            End With
            Count = Count + 1
        End With
        totalcount = totalcount + 1
    Loop
    rs_faretable.Close
    myflg = False
    MsgBox "Fare Saved Successfully", vbInformation, App.ProductName
    'Unload Me
    Exit Sub
err:
    MsgBox err.Number & "," & err.Description
     
    Close #Handle
    Exit Sub
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Me.Icon = frmMainform.Icon
    Set cnn_faretable = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    'OpenDatabase(App.Path & "\PVT.mdb", _
           dbDriverComplete, False, ";UID=;PWD=silbus")
    sql = "SELECT DISTINCT Rutcode FROM Route where faretype=1"
    Set rs_faretable = cnn_faretable.OpenRecordset(sql, dbOpenDynaset)
    If rs_faretable.RecordCount <> 0 Then rs_faretable.MoveFirst
    With rs_faretable
    Do While Not .EOF
        Combo1.AddItem rs_faretable.Fields(0)
        .MoveNext
    Loop
    End With
    myflg = True
    rs_faretable.Close
    MSFlexGrid1.ColWidth(0) = 500
   
End Sub
Private Sub MSFlexGrid1_Click()
    With MSFlexGrid1
        Text2.Left = .CellLeft + .Left
        Text2.Width = .CellWidth
        Text2.Height = .CellHeight
        Text2.Top = .CellTop + .Top
        Text2.Visible = True
        If .TextMatrix(.row, .Col) <> "" Then
            Text2.Text = .TextMatrix(.row, .Col)
            Text2.SetFocus
            Text2.SelStart = 0
            Text2.SelLength = Len(Text2)
        Else
            Text2.Visible = False
            If Text1.Enabled = True Then
                'Text1.Visible = True
                Text1.SetFocus
            End If
        End If
    End With
End Sub
Private Sub MSFlexGrid1_Scroll()
    Text2.Text = ""
    Text2.Visible = False
End Sub

Private Sub Text2_DblClick()
    myflg = True
    Text2_KeyPress (13)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)

Static LastText As String
Static SecondTime As Boolean
Const MaxDecimal As Integer = 1
Const MaxWhole As Integer = 4
   
    If TextBoxValidityNumeric(KeyAscii) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    With Text2
        If Not SecondTime Then
            If .Text Like "*[!0-9.]*" _
            Or .Text Like "*.*.*" _
            Or .Text Like String$(1 + MaxDecimal, "#") & "[!.]" Then
                SecondTime = True
                .Text = LastText
                .SelStart = Len(.Text)
            Else
                LastText = .Text
            End If
        End If
    End With
    myflg = False
    SecondTime = False


If KeyAscii = 13 Then   '''vaisakh 31.03.11
    myflg = True
     
    If Text2.Text = "." Then  'And Len(Text2.Text) = 1 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If Trim$(Text2.Text) <> "" And val(Text2) >= MinFare Then
        With MSFlexGrid1
             If val(.TextMatrix(.row, .Col - 1)) > val(Trim(Text2)) Then
                MsgBox "Fare must be greater than previous fare! ", vbExclamation
                Text2.Text = ""
                Text2.Visible = False
                Exit Sub
            End If
            .TextMatrix(.row, .Col) = Format(Text2.Text, "0.00") 'Format(Round(Text2.Text, 2), "0.00")
             'MsgBox myflg
            'DoEvents
            Text2.Text = ""
            Text2.Visible = False
            DoEvents
            'MsgBox "1"
            Command1_Click        ' 15.12.2010 vaisakh
           
            If .Col < TOTAFARE - 1 Then
                .Col = .Col + 1
                MSFlexGrid1_Click
                'Text2.SetFocus
            End If
             
        End With
'''    If Trim$(Text2.Text) <> "" And val(Text2) >= MinFare Then
'''        With MSFlexGrid1
'''            .TextMatrix(.row, .Col) = Format(Text2.Text, "0.00")
'''            Text2.Text = ""
'''            Text2.Visible = False
'''            Text1.SetFocus
'''            'If txtfarelist.Enabled = True Then txtfarelist.SetFocus
'''            '.SetFocus
'''        End With
    Else
         Text2.Text = ""
         Text2.Visible = False
         MsgBox "Minimum Fare is " & MinFare, vbExclamation, "faretableEdit"
         'If txtfarelist.Enabled = True Then txtfarelist.SetFocus
    End If
    
    myflg = False
End If
End Sub
Private Sub Text2_LostFocus()
    'Text2.Visible = False
    'MsgBox "lost"
     If myflg = False Then
        If MsgBox("Do you want to save the fare?" & vbCrLf & "Press enter key to save the fare.", vbYesNo, gblstrPrjTitle) = vbYes Then
            Text2.SetFocus
        Else
            Text2.Visible = False
        End If
        myflg = True
    End If
End Sub
Public Function ClearGrid() As Boolean
Dim CurCol As Integer
Dim CurRow As Integer
    With MSFlexGrid1
        For CurCol = 1 To .Cols - 1
            .Col = CurCol
            For CurRow = 1 To .Rows - 1
                .row = CurRow
                .TextMatrix(.row, .Col) = ""
            Next CurRow
        Next CurCol
    End With
End Function

Public Function LoadFareTable(NumOfStage As Integer)
On Error GoTo err
Dim sql2 As String
Dim COUNT2 As Integer
Dim SCHEDULE As Integer
Dim j As Integer ', TOTAFARE As Integer
Dim flexcolcount As Integer
Dim totalcount As Integer
Dim TEST1 As Integer
Dim COUNT1 As Integer
Dim temp As Long
Dim TTTT As DAO.Recordset
Dim TTT As String
    
    ClearGrid
    With MSFlexGrid1
        .Rows = 2
        .Cols = 5
        .row = 1
        .Col = 1
    End With
    If MSFlexGrid1.Enabled = False Then MSFlexGrid1.Enabled = True
    
    SCHEDULE = NumOfStage
    TOTAFARE = NumOfStage
    MSFlexGrid1.Cols = TOTAFARE
    
    sql2 = "SELECT * FROM FARE WHERE ROUTE='" & Combo1.Text & "'"
    Set RES = cnn_faretable.OpenRecordset(sql2, dbOpenDynaset)
    If RES.RecordCount <> 0 Then RES.MoveFirst
    Do While RES.EOF = False
        COUNT2 = COUNT2 + 1
        RES.MoveNext
    Loop
    If COUNT2 <> TOTAFARE Then
        MsgBox "You Cannot edit this faretable fare, Mismatch in Number of Stages and Total Entry", vbInformation, gblstrPrjTitle
        MSFlexGrid1.Enabled = False
        Exit Function
    End If
    
    totalcount = 1
    flexcolcount = TOTAFARE
    COUNT1 = 1
    RES.MoveFirst
           
    For temp = 1 To TOTAFARE
        With MSFlexGrid1
            .TextMatrix(0, temp - 1) = temp
        End With
    Next temp
        
    MSFlexGrid1.Cols = flexcolcount
        
    Do While RES.EOF = False
        If totalcount <= flexcolcount + 1 Then
            If MSFlexGrid1.row = 1 Then
                MSFlexGrid1.TextMatrix(1, 0) = 1
            End If
            With MSFlexGrid1
                TTT = "Select * from FARE where row=" & .row & " and col=" & COUNT1 & " and ROUTE='" & Combo1.Text & "'"
                Set TTTT = cnn_faretable.OpenRecordset(TTT, dbOpenSnapshot)
                If TTTT!FARE = 0 Then
                    TTTT.Close
                Else
                    .TextMatrix(.row, COUNT1 - 1) = TTTT!FARE 'res.Fields(2)
                    TTTT.Close
                End If
                COUNT1 = COUNT1 + 1
                RES.MoveNext
            End With
        End If
        totalcount = totalcount + 1
    
    Loop
    RES.Close
    Exit Function
err:
    MsgBox err.Number & "," & err.Description
    Exit Function

End Function
