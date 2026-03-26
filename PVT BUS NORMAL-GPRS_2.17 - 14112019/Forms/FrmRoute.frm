VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmRoute 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Route Edit"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9135
   Icon            =   "FrmRoute.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6720
      Left            =   -75
      TabIndex        =   9
      Top             =   -105
      Width           =   9180
      Begin VB.CheckBox chkAllowPass 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Allow Pass"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4140
         TabIndex        =   15
         Top             =   5580
         Width           =   1425
      End
      Begin JeweledBut.JeweledButton cmdCancel 
         Height          =   375
         Left            =   7815
         TabIndex        =   14
         Top             =   5745
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
         MICON           =   "FrmRoute.frx":0CCA
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command1 
         Height          =   375
         Left            =   6525
         TabIndex        =   13
         Top             =   5745
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
         MICON           =   "FrmRoute.frx":0CE6
         BC              =   12632256
         FC              =   0
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
         ItemData        =   "FrmRoute.frx":0D02
         Left            =   1275
         List            =   "FrmRoute.frx":0D04
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   4125
         Width           =   2175
      End
      Begin VB.TextBox txtName 
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
         Height          =   405
         Left            =   1260
         MaxLength       =   22
         TabIndex        =   1
         Top             =   4650
         Width           =   3780
      End
      Begin VB.CheckBox chkph 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PH"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4155
         TabIndex        =   7
         Top             =   5235
         Width           =   630
      End
      Begin VB.CheckBox chkConc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Concession"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2475
         TabIndex        =   6
         Top             =   5595
         Width           =   1425
      End
      Begin VB.CheckBox chkadjust 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjust"
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
         Left            =   1230
         TabIndex        =   5
         Top             =   5595
         Width           =   975
      End
      Begin VB.CheckBox chkstudent 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Student"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   5925
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CheckBox chkluggage 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Luggage"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2475
         TabIndex        =   3
         Top             =   5145
         Width           =   1080
      End
      Begin VB.CheckBox chkHalf 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Half"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1230
         TabIndex        =   2
         Top             =   5175
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlex 
         Height          =   3660
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   6456
         _Version        =   393216
         BackColorFixed  =   14737632
         AllowUserResizing=   1
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
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Code"
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
         Left            =   600
         TabIndex        =   11
         Top             =   4215
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Name"
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
         Left            =   570
         TabIndex        =   10
         Top             =   4680
         Width           =   555
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Route Edit"
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
      Height          =   450
      Left            =   3510
      TabIndex        =   12
      Top             =   -435
      Width           =   2490
   End
End
Attribute VB_Name = "frmRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As DAO.Database
Dim rs As DAO.Recordset
Dim Id As Integer
Dim Fso As New FileSystemObject
Dim Edit As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub Combo1_Click()
  Dim sql As String
    sql = "select * from route where rutcode='" & Combo1.Text & "'"
    Set cn = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    ', dbDriverComplete, False, ";UID=;PWD=softland")
    Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
 If rs.RecordCount <> 0 Then rs.MoveFirst
    With rs
    Id = rs!Id
     txtName = rs!rutname
     If rs!Half = 1 Then
        chkHalf.Value = 1
     Else
        chkHalf.Value = 0
     End If
     
      If rs!Luggage = 1 Then
        chkluggage.Value = 1
     Else
        chkluggage.Value = 0
     End If
    
     If rs!student = 1 Then
         chkstudent.Value = 1
     Else
         chkstudent.Value = 0
     End If
     
     If rs!Adjust = 1 Then
          chkAdjust.Value = 1
     Else
          chkAdjust.Value = 0
     End If
     
     If rs!PASSALLOW = 1 Then  '04/01/2010
          chkAllowPass.Value = 1
     Else
          chkAllowPass.Value = 0
     End If
         
     If rs!Conc = 1 Then
          chkConc.Value = 1
     Else
         chkConc.Value = 0
     End If
     
     If rs!ph = 1 Then
         chkph.Value = 1
     Else
         chkph.Value = 0
     End If
     
    End With

 rs.Close

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Dim sql As String
  If Combo1.ListIndex = -1 Then Exit Sub
    sql = "select * from route where rutcode='" & Combo1.Text & "'"
    Set cn = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    ', dbDriverComplete, False, ";UID=;PWD=softland")
    Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
 If rs.RecordCount > 0 Then rs.MoveFirst
    With rs
    Id = rs!Id
     txtName = rs!rutname
     If rs!Half = 1 Then
        chkHalf.Value = 1
     Else
        chkHalf.Value = 0
     End If
     
      If rs!Luggage = 1 Then
        chkluggage.Value = 1
     Else
        chkluggage.Value = 0
     End If
    
     If rs!student = 1 Then
         chkstudent.Value = 1
     Else
         chkstudent.Value = 0
     End If
     
     If rs!Adjust = 1 Then
          chkAdjust.Value = 1
     Else
          chkAdjust.Value = 0
     End If
     
     If rs!PASSALLOW = 1 Then  '04/01/2010
          chkAllowPass.Value = 1
     Else
          chkAllowPass.Value = 0
     End If
     
     If rs!Conc = 1 Then
          chkConc.Value = 1
     Else
         chkConc.Value = 0
     End If
     
     If rs!ph = 1 Then
         chkph.Value = 1
     Else
         chkph.Value = 0
     End If
     
    End With

 rs.Close
End If
End Sub
Private Sub Command1_Click()
Dim sql As String
Dim ts As TextStream
Dim file As Integer
Dim Buf As String
Dim Flag As Boolean
Flag = False
Dim RsR As DAO.Recordset
On Error GoTo err
If Trim(txtName.Text) = "" Then
   txtName.SetFocus
   Exit Sub
End If
sql = "select * from route where id=" & Id
Set cn = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
', dbDriverComplete, False, ";UID=;PWD=softland")

Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
If rs.RecordCount > 0 Then
rs.Edit
 rs!rutname = Trim(txtName)
 rs!Half = chkHalf.Value
 rs!Luggage = chkluggage.Value
 rs!student = chkstudent.Value
 rs!Adjust = chkAdjust.Value
 rs!PASSALLOW = chkAllowPass.Value ' 04/01/2010
 rs!Conc = chkConc.Value
 rs!ph = chkph.Value
 rs.Update
    If Dir$(App.Path & "\RouteLst.Txt") <> "" Then
      Kill App.Path & "\RouteLst.Txt"
    End If
    Dim RSql As String
    RSql = "select id,rutcode,rutname,nostage,nostop,minfare,faretype,usestop,half,luggage,student,adjust,PASSALLOW,conc,ph,startfrom from route order by id "  '04/01/2010
    Set RsR = cn.OpenRecordset(RSql, dbOpenDynaset)
    file = FreeFile()
    Open App.Path & "\RouteLst.Txt" For Append As #file
    If RsR.RecordCount > 0 Then
     RsR.MoveFirst
     Do While Not RsR.EOF
      Buf = RsR!RUTCODE & "," & RsR!rutname & "," & RsR!nostage & "," & RsR!nostop & "," & RsR!MinFare & "," & RsR!FareType & "," & RsR!UseStop & "," & RsR!Half & "," & RsR!Luggage & "," & RsR!student & "," & RsR!Adjust & "," & RsR!PASSALLOW & "," & RsR!Conc & "," & RsR!ph & "," & RsR!StartFrom  '04/01/2010
      Print #file, Buf
      RsR.MoveNext
     Loop
     Flag = True
    End If
    Close #file
    RsR.Close
'    RouteLST
    If Flag = True Then
       MsgBox "Edited route saved", vbInformation, gblstrPrjTitle
    End If
    createGrid
End If
rs.Close
'Combo1.Text = ""
txtName.Text = ""
chkHalf.Value = False
chkluggage.Value = False
chkstudent.Value = False
chkAdjust.Value = False
chkConc.Value = False
chkph.Value = False
chkAllowPass.Value = False
Exit Sub
err:
End Sub
' Buffer = Buffer & "0" & "," & chkHalf.Value & "," & chkluggage.Value & "," & chkstudent.Value & "," & chkadjust.Value & "," & chkconc.Value & "," & chkph.Value & "," & "0"
Private Sub Form_Load()
Dim sql As String
Me.Icon = frmMainform.Icon
sql = "select rutcode from route"
Set cn = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
', dbDriverComplete, False, ";UID=;PWD=softland")
 'SQL = "SELECT DISTINCT Route FROM Fare "
 sql = "select distinct rutcode from route"
 Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
 If rs.RecordCount <> 0 Then rs.MoveFirst
 With rs
    Do While Not .EOF
        Combo1.AddItem .Fields(0)
        .MoveNext
    Loop
 End With
 rs.Close
 With MSFlex
    .Rows = 1
    .Cols = 10
    '.Width = 7000
    .FormatString = "id|Code|Name|NoStage|Half|Luggage|student|Adjust|Conc|Ph|AllowPass" '04/01/2010
    .ColWidth(0) = 0
    .ColWidth(1) = (.Width - 350) * 0.11
    .ColWidth(2) = (.Width - 350) * 0.21
    .ColWidth(3) = (.Width - 350) * 0.12
    .ColWidth(4) = (.Width - 350) * 0.1
    .ColWidth(5) = (.Width - 350) * 0.1
    '.ColWidth(6) = (.Width - 350) * 0.1
    .ColWidth(6) = 0
    .ColWidth(7) = (.Width - 350) * 0.1
     .ColWidth(8) = (.Width - 350) * 0.1
      .ColWidth(9) = (.Width - 350) * 0.08
      .ColWidth(10) = (.Width - 350) * 0.12 '04/01/2010
      
 End With
 createGrid
End Sub
Private Sub createGrid()
Dim sql As String
Dim I As Integer
Dim rs As DAO.Recordset
sql = "Select * from Route"
Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
I = 1
 MSFlex.Rows = 1
If rs.RecordCount > 0 Then
   rs.MoveFirst
  Do While Not rs.EOF
    With MSFlex
    .Rows = I + 1
     .TextMatrix(I, 0) = rs!Id
      .TextMatrix(I, 1) = rs!RUTCODE
      .TextMatrix(I, 2) = rs!rutname
      .TextMatrix(I, 3) = rs!nostage
      If rs!Half = 1 Then
      .TextMatrix(I, 4) = "Yes"
      Else
      .TextMatrix(I, 4) = "No"
      End If
     If rs!Luggage = 1 Then
      .TextMatrix(I, 5) = "Yes"
      Else
      .TextMatrix(I, 5) = "No"
      End If
     
      If rs!student = 1 Then
      .TextMatrix(I, 6) = "Yes"
      Else
      .TextMatrix(I, 6) = "No"
      End If
       
      If rs!Adjust = 1 Then
      .TextMatrix(I, 7) = "Yes"
      Else
      .TextMatrix(I, 7) = "No"
      End If
      
      If rs!Conc = 1 Then
      .TextMatrix(I, 8) = "Yes"
      Else
      .TextMatrix(I, 8) = "No"
      End If
      
           
      If rs!ph = 1 Then
      .TextMatrix(I, 9) = "Yes"
      Else
      .TextMatrix(I, 9) = "No"
      End If
      
      If rs!PASSALLOW = 1 Then  '04/01/2010
      .TextMatrix(I, 10) = "Yes"
      Else
      .TextMatrix(I, 10) = "No"
      End If
  
     
    End With
    rs.MoveNext
      I = I + 1
Loop
End If
rs.Close


End Sub

Private Sub MSFlex_Click()
 'Edit = True
 If MSFlex.Rows > 1 Then
   MSFlex.SelectionMode = flexSelectionByRow
   MSFlex.HighLight = flexHighlightAlways
    MSFlex.SelectionMode = flexSelectionByRow
    Combo1.Text = MSFlex.TextMatrix(MSFlex.row, 1)
    txtName.Text = MSFlex.TextMatrix(MSFlex.row, 2)
    If MSFlex.TextMatrix(MSFlex.row, 4) = "No" Then
     chkHalf.Value = 0
    Else
        chkHalf.Value = 1
    End If
    If MSFlex.TextMatrix(MSFlex.row, 5) = "No" Then
        chkluggage.Value = 0
    Else
        chkluggage.Value = 1
    End If
'    If MSFlex.TextMatrix(MSFlex.row, 6) = "No" Then
        chkstudent.Value = 0
'    Else
'        chkstudent.Value = 1
'    End If
     If MSFlex.TextMatrix(MSFlex.row, 7) = "No" Then
        chkAdjust.Value = 0
    Else
        chkAdjust.Value = 1
    End If
    If MSFlex.TextMatrix(MSFlex.row, 8) = "No" Then
        chkConc.Value = 0
    Else
        chkConc.Value = 1
    End If
    If MSFlex.TextMatrix(MSFlex.row, 9) = "No" Then
        chkph.Value = 0
    Else
        chkph.Value = 1
    End If
    If MSFlex.TextMatrix(MSFlex.row, 10) = "No" Then '04/01/2010
        chkAllowPass.Value = 0
    Else
        chkAllowPass.Value = 1
    End If
    
 End If
End Sub
