VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmRouteDelete 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Route"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4950
   Icon            =   "frmRouteDelete.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1755
      Left            =   165
      TabIndex        =   2
      Top             =   150
      Width           =   4620
      Begin JeweledBut.JeweledButton cmdCancel 
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   1020
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
         MICON           =   "frmRouteDelete.frx":0CCA
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdDelete 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   495
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         TX              =   "&Delete"
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
         MICON           =   "frmRouteDelete.frx":0CE6
         BC              =   12632256
         FC              =   0
      End
      Begin VB.ComboBox cmbRoute 
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
         Left            =   195
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Route Code"
         Top             =   525
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Route"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   255
         Width           =   1590
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete Route"
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
      Left            =   1230
      TabIndex        =   1
      Top             =   -615
      Width           =   2490
   End
End
Attribute VB_Name = "frmRouteDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As DAO.Database
Dim rs As DAO.Recordset
Dim sql As String
Dim EMPTYFLAG As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim I As Integer
Dim Msg As VbMsgBoxResult
    Reset
    If cmbRoute.Text <> "" Then
        Msg = MsgBox("Do you want to delete route " & cmbRoute.Text & " ?", vbYesNo)
        If Msg = vbYes Then
            Msg = MsgBox("Are you sure ?", vbYesNo)
            If Msg = vbNo Then Exit Sub
        Else
            Exit Sub
        End If
    Else
        MsgBox "No route found!", vbExclamation, gblstrPrjTitle
        Exit Sub
    End If
    If EMPTYFLAG = False Then
        sql = "DELETE * FROM ROUTE WHERE rutcode='" & cmbRoute.Text & "'"
        cn.Execute (sql)
        
        sql = "DELETE * FROM FARE WHERE route='" & cmbRoute.Text & "'"
        cn.Execute (sql)
        
        sql = "DELETE * FROM STAGE WHERE route='" & cmbRoute.Text & "'"
        cn.Execute (sql)
        
        If Dir(App.Path & "\" & cmbRoute.Text & ".dat") <> "" Then Kill App.Path & "\" & cmbRoute.Text & ".dat"
        
        sql = "SELECT ID FROM STAGE"
        Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
        I = 0
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do While Not rs.EOF
                rs.Edit
                rs!Id = I
                I = I + 1
                rs.Update
                rs.MoveNext
            Loop
            rs.Close
        End If
        
        
'**********************************************************************
    Dim fname As String
    Dim tname As String
    Dim FHndl As Integer
    Dim THndl As Integer
    Dim StgCount As Integer
    Dim RouteCode As String
    Dim sSQL As String
    
    fname = App.Path & "\LOCAL_LANGUAGE.DAT"
    tname = App.Path & "\TEMP.DAT"
    If Dir(tname, vbNormal) <> "" Then Kill tname
    If Dir(fname, vbNormal) <> "" Then
        FHndl = FreeFile()
        Open fname For Binary Access Read As #FHndl
        THndl = FreeFile()
        Open tname For Binary Access Write As #THndl
        Do While Not EOF(FHndl)
            Get #FHndl, , LSTAG
            RouteCode = Mid(LSTAG.RouteCode, 1, InStr(1, LSTAG.RouteCode, Chr(0)) - 1)
            
            sSQL = "SELECT count(*) FROM ROUTE WHERE rutcode='" & RouteCode & "'"
            Set rs = cn.OpenRecordset(sSQL, dbOpenDynaset)
            If rs.Fields(0) > 0 Then
                If RouteCode <> RouteCode Then
                    If LSTAG.LocalLanguageStageName(0) <> &H0 Then
                        Put #THndl, , LSTAG
                    End If
                End If
            End If
            rs.Close
        Loop
        Close #FHndl
        Close #THndl
        Kill fname
        THndl = FreeFile()
        Open tname For Binary Access Read As #THndl
        FHndl = FreeFile()
        StgCount = 0
        Open fname For Binary Access Write As #FHndl
            Do While Not EOF(THndl)
                Get #THndl, , LSTAG
                LSTAG.stagecode = StgCount
                If LSTAG.LocalLanguageStageName(0) <> &H0 Then
                    Put #FHndl, , LSTAG
                    StgCount = StgCount + 1
                End If
            Loop
            Close #THndl
        Close #FHndl
    End If
'**********************************************************************
        
        MsgBox "Route - " & cmbRoute.Text & vbCrLf & "Removed Successfully", vbInformation, "Delete Route"
        cmbRoute.Clear
        
        sql = "SELECT rutcode FROM ROUTE"
        Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            
            While Not rs.EOF
                cmbRoute.AddItem rs!RUTCODE
                rs.MoveNext
            Wend
            cmbRoute.Text = cmbRoute.List(0)
        End If
        rs.Close
    Else
        MsgBox "No routes found!", vbExclamation, gblstrPrjTitle
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMainform.Icon
    Set cn = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbOpenDynaset, False, ";UID=;PWD=silbus")
    sql = "SELECT rutcode FROM ROUTE"
    Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
    cmbRoute.Clear
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            cmbRoute.AddItem rs!RUTCODE
            rs.MoveNext
        Wend
        cmbRoute.Text = cmbRoute.List(0)
        EMPTYFLAG = False
    Else
        EMPTYFLAG = True
    End If
    rs.Close
End Sub
