VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{011424E4-FAAB-4D1D-B936-E5C631EEDC26}#1.0#0"; "SmartProgressBar.ocx"
Begin VB.Form Transfer 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9960
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Transfer.frx":0000
   ScaleHeight     =   6915
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   1650
      Top             =   -480
   End
   Begin VB.Timer Timer3 
      Interval        =   600
      Left            =   0
      Top             =   -420
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   495
      Top             =   -420
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5655
      Top             =   -2160
   End
   Begin VB.FileListBox FileContainer 
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Top             =   -360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPacketNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   -480
      Width           =   1125
   End
   Begin VB.TextBox txtTotalPacket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   -390
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Files in Palmtec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2520
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   9585
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   1680
         ItemData        =   "Transfer.frx":83D14
         Left            =   150
         List            =   "Transfer.frx":83D16
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   285
         Width           =   9285
      End
      Begin JeweledBut.JeweledButton cmdUpload 
         Height          =   375
         Left            =   7920
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "&Upload"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Transfer.frx":83D18
         BC              =   12632256
         FC              =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Files in PC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2520
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   9585
      Begin VB.CheckBox ChkSelect 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select All"
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1710
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   1680
         ItemData        =   "Transfer.frx":83D34
         Left            =   120
         List            =   "Transfer.frx":83D36
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   255
         Width           =   9285
      End
      Begin JeweledBut.JeweledButton cmdDownload 
         Height          =   375
         Left            =   7920
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "&Download"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Transfer.frx":83D38
         BC              =   12632256
         FC              =   0
      End
   End
   Begin JeweledBut.JeweledButton Command1 
      Height          =   510
      Left            =   7680
      TabIndex        =   3
      Top             =   600
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   900
      TX              =   "&Safely Remove Palmtec Amphibia"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Transfer.frx":83D54
      BC              =   12632256
      FC              =   0
   End
   Begin SmartProgressBar.SmartPrgress pBar 
      Height          =   255
      Left            =   240
      Top             =   3840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      ForeColor       =   12582912
      BorderColor     =   12582912
      BackColor       =   16777215
      TextColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSCommLib.MSComm SerialCom 
      Left            =   4095
      Top             =   -2085
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Shape Shape1 
      Height          =   5535
      Left            =   120
      Top             =   1200
      Width           =   9735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   5865
      TabIndex        =   8
      Top             =   3915
      Width           =   3825
   End
   Begin VB.Label lblUSBStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Palmtec Communication Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   810
      Width           =   5850
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Transfer"
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
      Height          =   465
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   3420
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drag the file names you need to upload or download into the respective folders."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   60
      TabIndex        =   0
      Top             =   7095
      Visible         =   0   'False
      Width           =   9975
   End
End
Attribute VB_Name = "Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim strpath As String

Private Type PORTSETUP
    Port As Integer
    baud As String * 6
End Type

Private Type CREATE_CUR
    CurString As String * 8
End Type

Public DragFlg1 As Boolean
Public DragFlg2 As Boolean
Public pmtcID As String
Public scduID As String
Dim Fso As New FileSystemObject
Dim PSetup As PORTSETUP
Dim FileRpt As Integer
Dim I As Single
Dim j As Integer
Dim Msg As String
Dim ldP As Boolean
Dim SysDate As Date
Dim DiffDate As Date
'Dim gd As DATETYPE
'Dim gt As TIMETYPE
Dim Pathname As String
Private PortFlag As Boolean
Public WayBillFlg As Boolean
Dim rs_MachineIssue As DAO.Recordset
Dim blnDownload As Boolean

Private Sub Check1_Click()
On Error Resume Next
    If Check1.Value = 1 Then
        chkBttTrans.Value = 0
        Timer1.Enabled = False
        Timer4.Enabled = False
    End If
    CommandP.Value = True
End Sub

Private Sub Check2_Click()
    
End Sub

Private Sub chkBttTrans_Click()
Dim intThreshold As Integer
    If chkBttTrans.Value = 1 Then
        Check1.Value = 0
        intThreshold = 1
        Timer1.Enabled = True
        Timer4.Enabled = True
    Else
        intThreshold = 0
        Timer1.Enabled = False
        Timer4.Enabled = False
    End If
    If Not InitPort(PSetup.Port, PSetup.baud, intThreshold) Then
        MsgBox TransMsg
        PortFlag = False
    Else
        lblUSBStatus.caption = "Amphibia is Ready for Communication  through PORT " & PSetup.Port
        txtTotalPacket.Visible = False
        txtPacketNo.Visible = False
        Label1.Visible = False
        Check1.Visible = True
        Label2.Visible = False
        lblUSBStatus.ForeColor = vbBlue
        List1.Enabled = True
        List2.Enabled = True
        Timer2.Enabled = False
        TextP.Text = PSetup.Port
    End If
End Sub

Private Sub chkselect_Click()
On Error GoTo lblErr
Dim I As Integer
    If ChkSelect.Value = 1 Then
        For I = 0 To List2.ListCount - 1
            List2.Selected(I) = True 'True changed by deej 0n 03-05-05
            List2.SetFocus
        Next I
    Else
            For I = 0 To List2.ListCount - 1
            List2.Selected(I) = False 'True changed by deej 0n 03-05-05
        Next I
    End If
Exit Sub
lblErr:
End Sub

Private Sub cmdDownload_Click()
On Error Resume Next
Me.Enabled = False
    Call List2_DblClick
Me.Enabled = True
End Sub

Private Sub cmdUpload_Click()
On Error Resume Next
    Call List1_DblClick
End Sub

Private Sub Command1_Click()
On Error Resume Next
    If blnDownload Then Call Write_USB("SHUTDOWN")
    If Mode = "USB" Then Disconnect_USB
    Mode = "USB"
    Unload Me
End Sub

Private Sub CommandP_Click()
On Error GoTo err

    If TextP.Text = "" Then
        MsgBox "Enter Port Number", vbInformation, "Serial Port"
        Exit Sub
    End If
    
    CONNECT_DB
    Set RES = DB.OpenRecordset("PORT", dbOpenDynaset)
    
    If RES.RecordCount > 0 Then
        PSetup.Port = RES!Port
        PSetup.baud = RES!Boud
    Else
        PSetup.Port = TextP.Text
        PSetup.baud = 115200
    End If
    
    RES.Close

    If SerialComm.PortOpen = True Then SerialComm.PortOpen = False
    If SerialCom.PortOpen = True Then SerialCom.PortOpen = False
    
    If Not InitPort(val(PSetup.Port), PSetup.baud) Then
         MsgBox TransMsg
         Exit Sub
    Else
         FHndl = FreeFile()
         Open UCase(App.Path & "\trans.dat") For Binary Access Write As #FHndl
         Put FHndl, , PSetup
         Close #FHndl
         MsgBox "Port No.: " & PSetup.Port & Chr(13) & "Baud Rate :" & PSetup.baud & Chr(13) & "Port opened Successfully", vbInformation, "Admin"
         lblUSBStatus.caption = "Amphibia is Ready for Communication  through PORT " & PSetup.Port
         Mode = "COM"
         txtTotalPacket.Visible = False
         txtPacketNo.Visible = False
         Label1.Visible = False
         Label2.Visible = False
         lblUSBStatus.ForeColor = vbBlue
         List1.Enabled = True
         Check1.Visible = True
         List2.Enabled = True
         Timer2.Enabled = False
    End If
    
    If TextP.Enabled = True And ldP = True Then TextP.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
err:
End Sub

Private Sub Form_Activate()
On Error Resume Next
    WayBillFlg = False
    Me.Icon = frmMainform.Icon
    unvFlg = False
    ldP = True
    pBar.MinValue = 0
    pBar.MaxValue = 32000
    pBar.Value = 0
    TransPath = App.Path '& "\Transfer"
    TextP.Locked = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim idxCnt As Integer
    If KeyAscii = 27 Then
        For idxCnt = 0 To List1.ListCount - 1
            List1.Selected(idxCnt) = False
        Next idxCnt
        For idxCnt = 0 To List2.ListCount - 1
            List2.Selected(idxCnt) = False
        Next idxCnt
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
'    Form_KeyPress (27)
    DragFlg1 = False
    DragFlg2 = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
'    If USB_FLAG = True Then
'        frmMainform.Timer2.Enabled = True
'    End If
    
'    DbKSRTC.Close
'    End
End Sub

Private Sub List1_DblClick()
Dim msgStatus As VbMsgBoxResult '06/01/2010
On Error GoTo err: ''LAST BUG
Dim iCnt As Integer
Dim Lcnt As Integer
Dim lstcount As Integer
Dim lstcount1 As Integer
    lstcount = 0
    lstcount1 = 0
    iCnt = 0
    Lcnt = List1.ListCount
    USB_Path = TransPath

    While (iCnt < Lcnt)
        If List1.Selected(iCnt) = True And UCase(Mid$(List1.List(iCnt), 1, 8)) = UCase("SCHEDULE") Then
            lstcount = 1
        End If
        iCnt = iCnt + 1
    Wend
    iCnt = 0
    If lstcount = 1 Then
        While (iCnt < Lcnt)
            If UCase(Mid$(List1.List(iCnt), 1, 8)) = UCase("SCHEDULE") Then
                List1.Selected(iCnt) = True
            End If
            iCnt = iCnt + 1
        Wend
    End If
    iCnt = 0
    Do While (iCnt < Lcnt)
        
        'If List1.Selected(iCnt) = True Then
            If List1.Selected(iCnt) = True And UCase(List1.List(iCnt)) = UCase("Collection") Then
                Label3.caption = "Uploading Please wait ......."
                setEnable False
                
                '''''''''added by syam
                sleep (100)
                
                ''''''''''''''
                Label3.caption = "Version file Reading started..."
                If Dir(App.Path & "\VERSION.DAT", vbNormal) <> "" Then Call Kill(App.Path & "\VERSION.DAT")
                'USB_Path = App.Path
                If Read_USB("VERSION.DAT") = True Then
                    If CheckVersion(USB_Path) = False Then
                        MsgBox "Version mismatch", vbInformation, gblstrPrjTitle
                        Command1.Enabled = True
                        lstRefresh 2
                        Exit Sub
                    Else
                        Call Write_USB("VERISON_SUCESS")
                    End If
                Else
                    MsgBox "Version file reading failed!", vbInformation, gblstrPrjTitle
                    Command1.Enabled = True
                    lstRefresh 2
                    Exit Sub
                End If
                 If TripClosed = False Then
                    MsgBox "Trip status verification failed! Unable to Upload", vbInformation, App.ProductName
                    Command1.Enabled = True
                    lstRefresh 2
                    Exit Sub
                End If
                Label3.caption = "Uploading Setup Please wait ......."
                If UploadSetup = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                Label3.caption = "Uploading Collection Please wait ......."
                If UploadRPT = False Then setEnable True: lstRefresh 1: MsgBox "RPT Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If CovertRPT = False Then setEnable True: lstRefresh 1: MsgBox "Data Conversion Aborted", vbInformation: Label3.caption = "": Exit Sub
                lstRep 2, "Collection"
                MsgBox "Collection Uploaded Successfully", vbInformation, gblstrPrjTitle
                              
                If UploadODMTR = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If Dir$(App.Path & "\ODOMETER.DAT") <> "" Then
                    Label3.caption = "Uploading Odometer Please wait ......."
                    DBODMTR (FnameUp)
                    MsgBox "OdometerDetails Uploaded Successfully", vbInformation, gblstrPrjTitle
                End If
                If UploadExpense = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If Dir$(App.Path & "\EXPENSE.DAT") <> "" Then
                    Label3.caption = "Uploading Expense Please wait ......."
                    DBEXPENSE (FnameUp)
                    MsgBox "Expense Details Uploaded ", vbInformation, gblstrPrjTitle
                End If
                If UploadInsptr = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                    If Dir$(App.Path & "\INSPECTOR.DAT") <> "" Then
                        Label3.caption = "Uploading Inspectordet Please wait ......."
                        DBINSPR (FnameUp)
                        MsgBox "Inspector Details Uploaded ", vbInformation, gblstrPrjTitle
                    End If
             ElseIf UCase(Mid$(List1.List(iCnt), 1, 8)) = UCase("SCHEDULE") Then
                If UploadTKT(Format(iCnt, "00")) = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If UploadPASS = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If CovertColln(Format(iCnt, "00")) = False Then setEnable True: lstRefresh 1: MsgBox "Data Conversion Aborted", vbInformation: Label3.caption = "": Exit Sub
                Call AddPasscount
                Call Write_USB("UPLOADSUCCESS")
                lstRep 2, "SCHEDULE  " & Format(iCnt, "00")
''                If iCnt = 1 Then
''                    MsgBox "SCHEDULE1 Uploaded Successfully", vbInformation, gblstrPrjTitle
''                ElseIf iCnt = 2 Then
''                    MsgBox "SCHEDULE2 Uploaded Successfully", vbInformation, gblstrPrjTitle
''                ElseIf iCnt = 3 Then
''                    MsgBox "SCHEDULE3 Uploaded Successfully", vbInformation, gblstrPrjTitle
''                ElseIf iCnt = 4 Then
''                    MsgBox "SCHEDULE4 Uploaded Successfully", vbInformation, gblstrPrjTitle
''                ElseIf iCnt = 5 Then
''                    MsgBox "SCHEDULE5 Uploaded Successfully", vbInformation, gblstrPrjTitle
''                ElseIf iCnt = 6 Then
''                    MsgBox "SCHEDULE6 Uploaded Successfully", vbInformation, gblstrPrjTitle
''                ElseIf iCnt = 7 Then
''                    MsgBox "SCHEDULE7 Uploaded Successfully", vbInformation, gblstrPrjTitle
''                End If
                MsgBox "Ticket Details Uploaded Successfully", vbInformation, gblstrPrjTitle
               '''''''''''''''''''
               If iCnt + 1 = Lcnt And UCase(Mid$(List1.List(iCnt), 1, 8)) = UCase("SCHEDULE") Then
                If IsTicketRemoveEnabled = True Then   '06/01/2010
                    msgStatus = MsgBox("Do You Want to Remove Data? " & vbCrLf & "Click Yes to Remove ", vbQuestion + vbYesNo + vbDefaultButton2)
                    If msgStatus = vbYes Then
                        Call Write_USB("DELETETKTS")   ' 06/01/2010
                    End If
                End If
                MsgBox "All Details Uploaded Successfully", vbInformation, gblstrPrjTitle
               End If
               
               '''''''''''''''''''
                
            End If
        'End If
          If iCnt < Lcnt Then iCnt = iCnt + 1
    Loop
    
    lstRefresh 1
'    If Lcnt = 1 Then
'        CmdUpload.SetFocus
'    End If
    Label3.caption = ""
    DragFlg1 = False
    setEnable True
    List1.Enabled = True
    List2.Enabled = True
   
'    If UploadInsptr = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.Caption = "": Exit Sub
'    DBINSPR (FnameUp)
'    MsgBox "Inspector Details Uploaded ", vbInformation, gblstrPrjTitle
'
'    If UploadExpense = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.Caption = "": Exit Sub
'    DBEXPENSE (FnameUp)
'    MsgBox "Expense Details Uploaded ", vbInformation, gblstrPrjTitle
'
'    MsgBox "All Details Uploaded Successfully", vbInformation, gblstrPrjTitle
    Label3.caption = " "
    Exit Sub
err:

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    DragFlg1 = False
    List1.Enabled = True
    List2.Enabled = True
End Sub

Private Sub List2_DblClick()
  Dim iCnt As Integer
  Dim Lcnt As Integer
  Dim gblstrTransferPath As String 'LOGO
  On err GoTo err
    iCnt = 0
    Lcnt = List2.ListCount
    If Lcnt > 0 Then
        If Dir(App.Path & "\VERSION.DAT", vbNormal) <> "" Then Call Kill(App.Path & "\VERSION.DAT")
        USB_Path = TransPath
        Label3.caption = "Version file Reading started..."
        If Read_USB("VERSION.DAT") = True Then
            If CheckVersion(USB_Path) = False Then
               MsgBox "Version mismatch", vbInformation, gblstrPrjTitle
               Command1.Enabled = True
               lstRefresh 2
               Exit Sub
            Else
               Call Write_USB("VERISON_SUCESS")
            End If
        Else
            MsgBox "Version file reading failed!", vbInformation, gblstrPrjTitle
            Command1.Enabled = True
            lstRefresh 2
            Exit Sub
        End If
    End If
    Do While (iCnt < Lcnt)
        If List2.Selected(iCnt) = True Then
            If UCase(List2.List(iCnt)) = UCase("Schedule") Then
                Label3.caption = "Downloading Please wait ......."
                USB_Path = TransPath
                If ScheduleClosed = False Then
                    MsgBox "Schedule status verification failed! Unable to download", vbInformation, App.ProductName
                    Command1.Enabled = True
                    lstRefresh 2
                    Label3.caption = ""
                    Exit Sub
                End If
                CancelFlag = False
                FrmSelectRoute.Show vbModal
                If CancelFlag = False Then
                    Label3.caption = ""
                    lstRefresh 2
                    DragFlg2 = False
                    setEnable True
                    List1.Enabled = True
                    List2.Enabled = True
                    Exit Sub
                End If
                CreateSchedule
                CancelFlag = False
                setEnable False
                USB_Path = TransPath
                
                If unvFlg = False Then
                    blnDownload = True
                    Label3.caption = "Downloading Route List..."
                    subSetDateTime   ''' 21/01/2011
                    If DownSTGNM = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                    Label3.caption = "Downloading Stage Name List..."
                    'sleep (200)
                    If DownSTGE = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                    'If LANGAGE_ENABLED = 1 Then
                     If LocalLanguage > 0 Then
'''                        If DownGRSTAGE = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.Caption = "": Exit Sub     ''''' 20/01/2011
                        If DownLanguage = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                    End If
                    MsgBox "Stage Name List Downloaded Successfully", vbInformation, gblstrPrjTitle
                    Label3.caption = "Downloading Route Details..."
                    'sleep (200)
                    'sleep (200)
                    If DownRTE = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
'''                    If IsTicketRemoveEnabled = True Then Call Write_USB("DELETETKTS")  ' 06/01/2010
                    lstRep 1, "Schedule"
                    MsgBox "Route Details Downloaded Successfully", vbInformation, gblstrPrjTitle
                End If
                blnDownload = True
                
            ElseIf UCase(List2.List(iCnt)) = UCase("Crew Details") Then
            dbcrew
                If Dir$(App.Path & "\CREW.DAT") <> "" Then
                    Label3.caption = "Downloading Crew Details..."
                    subSetDateTime   ''' 21/01/2011
                    setEnable False
                    USB_Path = TransPath
                    
                    'dbcrew
                    'sleep (200)
                    If DownCREW = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                    lstRep 1, "Crew Details"
                    MsgBox "Crew Details Downloaded Successfully", vbInformation, gblstrPrjTitle
                    blnDownload = True
                Else
                    MsgBox "Crew Details Not Found ", vbInformation, gblstrPrjTitle
                End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''RNC
            
            ElseIf UCase(List2.List(iCnt)) = UCase("Vehicle Details") Then
            dbvehicle
                If Dir$(App.Path & "\VEHICLE.DAT") <> "" Then
                    Label3.caption = "Downloading Vehicle Details..."
                    subSetDateTime   ''' 21/01/2011
                    setEnable False
                    USB_Path = TransPath
                    
                    'dbvehicle()
                    'sleep (200)
                    If DownVEHICLE = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                    lstRep 1, "Vehicle Details"
                    MsgBox "Vehicle Details Downloaded Successfully", vbInformation, gblstrPrjTitle
                    blnDownload = True
                Else
                    MsgBox "Vehicle Details Not Found ", vbInformation, gblstrPrjTitle
                End If
            
            ElseIf UCase(List2.List(iCnt)) = UCase("Expense Details") Then
            'dbexpdet
            
                dbexpdet
                If Dir$(App.Path & "\EXPENSEDET.DAT") <> "" Then
                    Label3.caption = "Downloading Expense Details..."
                    subSetDateTime   ''' 21/01/2011
                    setEnable False
                    USB_Path = TransPath
                    'sleep (300)
                    If DownEXPDET = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                    lstRep 1, "EXPENSE Details"
                    MsgBox "Expense Details Downloaded Successfully", vbInformation, gblstrPrjTitle
                    blnDownload = True
                Else
                    MsgBox "Expense Details Not Found ", vbInformation, gblstrPrjTitle
                End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ElseIf UCase(List2.List(iCnt)) = UCase("Settings") Then
                Label3.caption = "Downloading Settings..."
                subSetDateTime   ''' 21/01/2011
                setEnable False
                USB_Path = TransPath
               ' Call CreateCurrency
                'sleep (200)
                'If DownCURRENCY = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If UploadSetup = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If DownSetup = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Settings Aborted", vbInformation: Label3.caption = "": Exit Sub
                'sleep (300)
              '  MsgBox "2"
                If DownSTP = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                ' MsgBox "3"
                blnDownload = True
                lstRep 1, "Settings"
               
                
                '''''''''''''''Now Download Logo..Sccratchpad1.dat
                Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
                Set rsRecord = DB.OpenRecordset("LOGOSETTING", dbOpenDynaset)
                If rsRecord.RecordCount > 0 Then
                    rsRecord.MoveFirst
                    USB_Path = App.Path
                    gblstrTransferPath = App.Path
                    Call ConvertLanBMP(IIf(IsNull(rsRecord!Logo1), "", rsRecord!Logo1), gblstrTransferPath & "\SCRATCHPAD1.DAT", False)
                    Call ConvertLanBMP(IIf(IsNull(rsRecord!Logo2), "", rsRecord!Logo2), gblstrTransferPath & "\SCRATCHPAD1.DAT", True)
                    Call ConvertLanBMP(IIf(IsNull(rsRecord!Logo3), "", rsRecord!Logo3), gblstrTransferPath & "\SCRATCHPAD1.DAT", True)
                    Call ConvertLanBMP(IIf(IsNull(rsRecord!Logo4), "", rsRecord!Logo4), gblstrTransferPath & "\SCRATCHPAD1.DAT", True)
                    If Dir(gblstrTransferPath & "\SCRATCHPAD1.DAT", vbNormal) <> "" Then
                    If Write_USB(UCase("SCRATCHPAD1.DAT")) = False Then
                        MsgBox ("Communication failed ! Logo not downloaded " & vbTab & vbTab & vbTab & vbTab & vbTab & Now & vbCrLf)
                        MsgBox "Communication failed ! Logo file not downloaded " & vbCrLf & _
                               "Downloading aborted. Ensure Palmtec is in Transfer Mode and try again", vbInformation, gblstrPrjTitle
                        lstRefresh 2
                        setEnable True
                        Exit Sub
                    End If
                    MsgBox ("Logo file  downloaded successfully")
                    Else
                    MsgBox ("Logo information not found...")
                    End If
                  End If
                
                '''''''''''''''''''''''''''''''''''''''''''''''''
                 MsgBox "Settings Downloaded Successfully", vbInformation, gblstrPrjTitle
            End If
        End If
        If iCnt < Lcnt Then iCnt = iCnt + 1
    Loop

    Label3.caption = ""
    lstRefresh 2
    DragFlg2 = False
    setEnable True
    List1.Enabled = True
    List2.Enabled = True
    Exit Sub
    
err:
  MsgBox "Error" & err.Number & err.Description
End Sub



Public Sub subSetDateTime()
Dim strDateTime As String
On Error GoTo CatchError

     strDateTime = Format(Day(Now), "00") & "/" & _
                Format(Month(Now), "00") & "/" & _
                Format(Year(Now), "0000") & " " & _
                Format(Hour(Now), "00") & ":" & _
                Format(Minute(Now), "00") & ":" & _
                Format(Second(Now), "00")

    strDateTime = Replace(strDateTime, "/", "")
    strDateTime = Replace(strDateTime, ":", "")
    strDateTime = Replace(strDateTime, " ", "")
    Call Write_USB("DATEANDTIME-" & strDateTime)
    Exit Sub
CatchError:

End Sub



Private Sub List2_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 27 Then
        If ScheduleClosed = False Then
            MsgBox "Schedule status verification failed! Unable to download", vbInformation, App.ProductName
            Command1.Enabled = True
            lstRefresh 2
            Exit Sub
        End If
    End If
End Sub

Private Sub List2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
'    DragFlg2 = False
    List1.Enabled = True
    List2.Enabled = True
End Sub


Private Sub SerialCom_OnComm()
On Error GoTo err
    If SerialCom.InBufferCount > 0 Then
        bRcvdByte = Asc(SerialCom.Input)
        If lRcvdByteCount = 0 And bRcvdByte <> &H2 Then Exit Sub
        bDataReadyTimer = 1
        sStatus = sStatus & Chr(bRcvdByte)
        sRcvData = sRcvData & Format(Hex(bRcvdByte), "00")
        lRcvdByteCount = lRcvdByteCount + 1
        
        If lRcvdByteCount = 3 Then
            lBytesToBeRead = Asc(Mid$(sStatus, 3, 1)) * 256 + Asc(Mid$(sStatus, 2, 1))
            blnDataReady = False
            'Debug.Print "lBytesToBeRead : " & lBytesToBeRead
        End If
        
        If lBytesToBeRead = lRcvdByteCount Then blnDataReady = True
        
    End If
    Exit Sub
err:
    MsgBox "Error! " & vbCrLf & "serialcom_OnComm() " & vbCrLf & "Error No: " & err.Number & vbCrLf & err.Description
End Sub

Private Sub TextP_Change()
On Error Resume Next
    If Trim(TextP.Text) <> "" Then
        If IsNumeric(TextP.Text) = False Then
            MsgBox "Only numeric values allowed", vbInformation, "Admin"
            TextP.Text = ""
        End If
    If val(TextP.Text) > 128 Then
        TextP.Text = Mid$(Trim(TextP.Text), 1, Len(TextP.Text) - 1)
        SendKeys "{End}"
    End If
    End If
End Sub

Private Sub TextP_GotFocus()
On Error Resume Next
   TextP.SetFocus
    SendKeys "{Home}+{End}"
End Sub

Private Sub TextP_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 13 Then
        Select Case KeyAscii
        Case 48 To 57:
            KeyAscii = KeyAscii
        Case 8:
            KeyAscii = KeyAscii
        Case Else:
        KeyAscii = 0
        End Select
    Else
        CommandP.SetFocus
    End If
End Sub

Private Sub TextP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SendKeys "{Home}+{End}"
End Sub

'Private Sub Form_Activate1()
'Dim LCnts As Integer
'On Error GoTo ErrLoc
'    If PortFlag = False Then Exit Sub
'     If CmdRefresh = False Then
''         MsgBox "Directory List not Found", vbCritical
'         Exit Sub
'     End If
'  For LCnts = 0 To LstPalmTec.ListCount - 1
'    LstPalmTec.ListIndex = LCnts
'   If CmdUpload = False Then
'        Exit Sub
'   End If
'  Next LCnts
'ErrLoc:
'    MsgBox err.Description, vbCritical, "BUS Data Transfer"
'    Frame1.Enabled = True
'    Frame2.Enabled = True
'    Close #FileRpt
'    Close #DirRPT
'    Exit Sub
'End Sub

Private Sub Form_Load()
Dim Filehdl As Integer
    CONNECT_DB
    Mode = "USB"
    ldP = False
    TransPath = App.Path & "\Transfer"
    DragFlg1 = False
    DragFlg2 = False
    PortFlag = True
    ChkSelect.Value = 0
    Set Fso = CreateObject("Scripting.FileSystemObject")
    FileContainer.Path = App.Path
    ArryCount = 0
    Timer2.Enabled = False
    USB_FLAG = False
    
    If USB_FLAG = False Then
    
        If GetDevices = True Then
           Timer2.Enabled = False
'           Timer1.Enabled = False
           frmMainform.Timer2.Enabled = False
           lblUSBStatus.caption = "Amphibia is Ready for Communication through USB"
           lblUSBStatus.ForeColor = vbBlue
           List1.Enabled = True
           List2.Enabled = True  'test
           'TextP.Visible = False
          ' Check1.Visible = False
           'CommandP.Visible = False
           'Label12.Visible = False
           txtTotalPacket.Visible = True
           txtPacketNo.Visible = True
           'Label2.Visible = True
           'Label1.Visible = True
           Mode = "USB"
           USB_FLAG = True
        Else
            TransPath = App.Path
            lblUSBStatus.caption = "Amphibia not Detected"
            lblUSBStatus.ForeColor = vbRed
            List1.Enabled = False
            List2.Enabled = False
            cmdDownload.Enabled = False
            cmdUpload.Enabled = False
            Mode = "USB"
        End If
    
    Else
        lblUSBStatus.caption = "Amphibia is Ready for Communication  through PORT " & PSetup.Port
        txtTotalPacket.Visible = False
        txtPacketNo.Visible = False
        Label1.Visible = False
        Check1.Visible = True
        Label2.Visible = False
        lblUSBStatus.ForeColor = vbBlue
        List1.Enabled = True
        List2.Enabled = True
        Timer2.Enabled = False
        TextP.Text = PSetup.Port
        If PortFlag = False Then
            lblUSBStatus.caption = "Amphibia not Detected"
            lblUSBStatus.ForeColor = vbRed
            List1.Enabled = False
            List2.Enabled = False
            cmdDownload.Enabled = False
            cmdUpload.Enabled = False
            Mode = "USB"
        End If
    End If
    'If mode = "USB" Then
    '   TextP.Visible = False
    '   CommandP.Visible = False
    '   Label12.Visible = False
    '   lblUSBStatus.Visible = True
    ''   Timer2.Enabled = True
    'Else
    '   TextP.Visible = True
    '   CommandP.Visible = True
    '   Label12.Visible = True
    '   lblUSBStatus.Visible = False
    'End If
    List1.AddItem "Collection"
'    List1.AddItem "Odometer"
'    List1.AddItem "Expense"
'    List1.AddItem "Inspector Details"
    List2.AddItem "Crew Details"
    List2.AddItem "Vehicle Details"
    List2.AddItem "Settings"
    List2.AddItem "Schedule"
    List2.AddItem "Expense Details"
   ' List2.AddItem "Expensedet Details"
    
    '    List2.AddItem "Date & Time"
    '    List2.AddItem "Way Bill"
        TransferPath = App.Path & "\Transfer\"
    '    Set DbKSRTC = OpenDatabase(App.Path & "\KSRTC.MDB", dbDriverComplete, False, ";UID=;PWD=etm" & PPWD)
    Exit Sub
chk:
    Select Case err.Number
        Case 70
            MsgBox "   Database already running    " & vbCrLf & "   Please close database  ", vbInformation, "Bus"
        Case Else
            MsgBox "Error No: " & err.Number & "     " & err.Description, vbInformation, "Bus"
    End Select
End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    List1.Enabled = False
    DragFlg1 = True
    lstRefresh (2)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If DragFlg2 = True Then
  Dim iCnt As Integer
  Dim Lcnt As Integer
  On Error GoTo err
    iCnt = 0
    Lcnt = List2.ListCount
    DragFlg2 = False
    Do While (iCnt < Lcnt)
        If List2.Selected(iCnt) = True Then
            If UCase(List2.List(iCnt)) = UCase("Schedule") Then
                Label3.caption = "Downloading Please wait ......."
                If ScheduleClosed = False Then
                    MsgBox "Schedule status verification failed! Unable to download", vbInformation, App.ProductName
                    lstRefresh 2
                    Exit Sub
                End If
                CancelFlag = False
                FrmSelectRoute.Show vbModal
                If CancelFlag = False Then
                    Label3.caption = ""
                    DragFlg2 = False
                    setEnable True
                    List1.Enabled = True
                    List2.Enabled = True
                    Exit Sub
                End If
                CreateSchedule
                CancelFlag = True
                setEnable False
                USB_Path = TransPath
                If unvFlg = False Then
                    Label3.caption = "Downloading Route List..."
                    If DownSTGNM = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Exit Sub
                    Label3.caption = "Downloading Stage Name List..."
                    If DownSTGE = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Exit Sub
                    'If LANGAGE_ENABLED = 1 Then
                     If LocalLanguage > 0 Then
                        If DownLanguage = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Exit Sub
                    End If
                    Label3.caption = "Downloading Route Details..."
                    If DownRTE = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Exit Sub
                    If IsTicketRemoveEnabled = True Then Call Write_USB("DELETETKTS")
                    lstRep 1, "Schedule"
                    MsgBox "Route Details Downloaded Successfully", vbInformation, gblstrPrjTitle
                End If
                blnDownload = True
            ElseIf UCase(List2.List(iCnt)) = UCase("Crew Details") Then
                Label3.caption = "Downloading Crew Details..."
                setEnable False
                USB_Path = TransPath
                If DownCREW = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Exit Sub
                lstRep 1, "Crew Details"
                MsgBox "Crew Details Downloaded Successfully", vbInformation, gblstrPrjTitle
                blnDownload = True
            ElseIf UCase(List2.List(iCnt)) = UCase("Settings") Then
                Label3.caption = "Downloading Settings..."
                setEnable False
                USB_Path = TransPath
                Call CreateCurrency
                If DownCURRENCY = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If UploadSetup = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Exit Sub
                If DownSetup = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Exit Sub
                If DownSTP = False Then setEnable True: lstRefresh 2: MsgBox "Downloading Aborted", vbInformation: Exit Sub
                blnDownload = True
                lstRep 1, "Settings"
                
                
                '''''''''''''''Now Download Logo..Sccratchpad1.dat
                Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
                Set rsRecord = DB.OpenRecordset("LOGOSETTING", dbOpenDynaset)
                If rsRecord.RecordCount > 0 Then
                    rsRecord.MoveFirst
                    USB_Path = App.Path
                    gblstrTransferPath = App.Path
                    Call ConvertLanBMP(IIf(IsNull(rsRecord!Logo1), "", rsRecord!Logo1), gblstrTransferPath & "\SCRATCHPAD1.DAT", False)
                    Call ConvertLanBMP(IIf(IsNull(rsRecord!Logo2), "", rsRecord!Logo2), gblstrTransferPath & "\SCRATCHPAD1.DAT", True)
                    Call ConvertLanBMP(IIf(IsNull(rsRecord!Logo3), "", rsRecord!Logo3), gblstrTransferPath & "\SCRATCHPAD1.DAT", True)
                    Call ConvertLanBMP(IIf(IsNull(rsRecord!Logo4), "", rsRecord!Logo4), gblstrTransferPath & "\SCRATCHPAD1.DAT", True)
                    If Dir(gblstrTransferPath & "\SCRATCHPAD1.DAT", vbNormal) <> "" Then
                    If Write_USB(UCase("SCRATCHPAD1.DAT")) = False Then
                        MsgBox ("Communication failed ! Logo not downloaded " & vbTab & vbTab & vbTab & vbTab & vbTab & Now & vbCrLf)
                        MsgBox "Communication failed ! Logo file not downloaded " & vbCrLf & _
                               "Downloading aborted. Ensure Palmtec is in Transfer Mode and try again", vbInformation, gblstrPrjTitle
                        lstRefresh 2:
                        setEnable True
                        Exit Sub
                    End If
                    MsgBox ("Logo file  downloaded successfully")
                    Else
                    MsgBox ("Logo information not found...")
                    End If
                  End If
                
                '''''''''''''''''''''''''''''''''''''''''''''''''
                
               MsgBox "Settings Downloaded Successfully", vbInformation, gblstrPrjTitle
            End If
        End If
        If iCnt < Lcnt Then iCnt = iCnt + 1
    Loop
    Label3.caption = ""
    lstRefresh 2
    DragFlg2 = False
    setEnable True
    List1.Enabled = True
    List2.Enabled = True
  End If
  Exit Sub
err:
End Sub

'Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    DragFlg = False
'End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'     List2.Enabled = False
     DragFlg2 = True
    lstRefresh (1)
End Sub


Private Sub List2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
 If DragFlg1 = True Then
  Dim iCnt As Integer
  Dim Lcnt As Integer
  
    iCnt = 0
    Lcnt = List1.ListCount
    DragFlg1 = False
    USB_Path = TransPath
    Do While (iCnt < Lcnt)
        If List1.Selected(iCnt) = True Then
            If UCase(List1.List(iCnt)) = UCase("Collection") Then
                Label3.caption = "Uploading Please wait ......."
                setEnable False
                
                Label3.caption = "Uploading Setup Please wait ......."
                If UploadSetup = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                Label3.caption = "Uploading Collection Please wait ......."
                If UploadRPT = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If CovertRPT = False Then setEnable True: lstRefresh 1: MsgBox "Data Conversion Aborted", vbInformation: Label3.caption = "": Exit Sub
                lstRep 2, "Collection"
                MsgBox "Collection Uploaded Successfully", vbInformation, gblstrPrjTitle
            
            ElseIf UCase(Mid$(List1.List(iCnt), 1, 8)) = UCase("SCHEDULE") Then
                If UploadTKT(Format(iCnt, "00")) = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If UploadPASS = False Then setEnable True: lstRefresh 1: MsgBox "Uploading Aborted", vbInformation: Label3.caption = "": Exit Sub
                If CovertColln(Format(iCnt, "00")) = False Then setEnable True: lstRefresh 1: MsgBox "Data Conversion Aborted", vbInformation: Label3.caption = "": Exit Sub
                Call Write_USB("UPLOADSUCCESS")
                lstRep 2, "SCHEDULE  " & Format(iCnt, "00")
                MsgBox "Ticket Details Uploaded Successfully", vbInformation, gblstrPrjTitle
            End If
        End If
          If iCnt < Lcnt Then iCnt = iCnt + 1
    Loop
    
    
    lstRefresh 1
    Label3.caption = ""
    DragFlg1 = False
    setEnable True
    List1.Enabled = True
    List2.Enabled = True
 End If
 
End Sub

'Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  Dim iCnt As Integer
'  Dim Lcnt As Integer
'    iCnt = 0
'    Lcnt = List2.ListCount - 1
'    For iCnt = 0 To Lcnt
'        List2.Selected(iCnt) = False
'    Next iCnt
'End Sub
'Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    DragFlg = False
'End Sub

Function DownPRM() As Boolean
On Error Resume Next
Dim ret As Long
Dim StrTrans As String
    
    If Dir$(TransPath & "\" & "PRM") <> "" Then
        FileCopy TransPath & "\" & "PRM", TransferPath & "PRM1"
'        If ChangeFile1("PRM.TXT") = False Then
'            DownPRM = False
'            Exit Function
'        End If
        If Mode = "USB" Then
            If Write_USB("PRM") = True Then
'                Timer1.Enabled = True
                DownPRM = True
            End If
        Else
            If Trans("2 PRM") = True Then
'                Timer1.Enabled = True
                DownPRM = True
            End If
        End If
    Else
        MsgBox "Settings File Not Created " & vbCrLf & _
            "Please Enter all the Details in EMT Setting..", vbCritical, "ERROR!"
            DownPRM = False
    End If
    
End Function

Function DownDT(Optional AfterPRM As Boolean = False) As Boolean
Dim ret As Long
Dim StrTrans As String
On Error Resume Next
    If AfterPRM Then
        strDate = LPad(DTTrans.Day, "0", 2) & LPad(DTTrans.Month, "0", 2) & LPad(DTTrans.Year, "0", 2) & DTTrans.DayOfWeek
        strTime = LPad(Hour(Time), "0", 2) & LPad(Minute(Time), "0", 2) & LPad(Second(Time), "0", 2)
    End If
    If strDate <> "" Then
        If Trans("5 " & strDate) = False Then
            MsgBox TransMsg, vbCritical, "ERROR!"
            DownDT = False
'            Timer1.Enabled = True
'            Unload Me
            Exit Function
        Else
            If Trans("4 " & strTime) = True Then
'                Timer1.Enabled = True
                DownDT = True
            Else
                MsgBox TransMsg, vbCritical, "ERROR!"
                DownDT = False
'                Timer1.Enabled = True
                'Unload Me
                Exit Function
            
            End If
        End If
    Else
        MsgBox "Please set date...", vbInformation, "ERROR!"
        DownDT = False
    End If
    
End Function

Function UpPrmForDelete() As Boolean
On Error Resume Next
    If Dir$(App.Path & "\Transfer\PRM1") <> "" Then
        Kill App.Path & "\Transfer\PRM1"
    End If
    If Trans("1 PRM1") = False Then
        If StrComp(TransMsg, "File Not Found") = 0 Then
            PRM_EMPTY_FLAG = True
            UpPrmForDelete = False
            Exit Function
        End If
        MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
        UpPrmForDelete = False
        Exit Function
    End If
'    Timer1.Enabled = True
    UpPrmForDelete = True
    
End Function

Public Function BforSCH()
Dim Fhnd As Integer
Dim StrTrans As String
On Error Resume Next
    SCH_FLAG = True
    StrTrans = "SOFTLAND"
    Fhnd = FreeFile()
    If Dir(TransferPath & "Bustype.dat") <> "" Then Kill TransferPath & "Bustype.dat"
    Open TransferPath & "Bustype.dat" For Binary Access Write As #Fhnd
        Put #Fhnd, , StrTrans
    Close #Fhnd
    If Trans("2 Bustype.dat") = True Then
'        Timer1.Enabled = True
    Else
        checkflag = False
'        Timer1.Enabled = True
    End If
End Function

Function DownSTGNM() As Boolean
On Error Resume Next
    If Dir$(App.Path & "\RouteLst.LST") <> "" Then
'        If ChangeFile1("SCH.TXT") = False Then
'            DownSCH = False
'            Exit Function
'        End If
        If Mode = "USB" Then
            If Write_USB("ROUTELST.LST") = True Then
'                Timer1.Enabled = True
                DownSTGNM = True
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("RouteLst.LST") & " ") = True Then
'                Timer1.Enabled = TrueRouteLst
                DownSTGNM = True
                Exit Function
            End If
        End If
    End If
    MsgBox "No Schedule File Created ...", vbInformation, "BUS"
    DownSTGNM = False
End Function
Function DownSTP() As Boolean
On Error Resume Next
    If Dir$(App.Path & "\BUS.DAT") <> "" Then
'        If ChangeFile1("SCH.TXT") = False Then
'            DownSCH = False
'            Exit Function
'        End If
        If Mode = "USB" Then
            If Write_USB("BUS.DAT") = True Then
'                Timer1.Enabled = True
                DownSTP = True
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("BUS.DAT") & " ") = True Then
'                Timer1.Enabled = True
                DownSTP = True
                Exit Function
            End If
        End If
    End If
    MsgBox "No Schedule File Created ...", vbInformation, "BUS"
    DownSTP = False
End Function



Function DownLanguage() As Boolean
On Error Resume Next
    If Dir$(App.Path & "\LANGUAGE.DAT") <> "" Then
        If Mode = "USB" Then
            If Write_USB("LANGUAGE.DAT") = True Then
                DownLanguage = True
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("LANGUAGE.DAT") & " ") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownLanguage = False
                Unload Me
                Exit Function
            End If
         End If
        DownLanguage = False
    Else
        MsgBox "Donwloading Language Details Error!.", vbInformation, "BUS"
        DownLanguage = False
    End If
    
End Function

Function DownSTGE() As Boolean
On Error Resume Next
    If Dir$(App.Path & "\STAGE.LST") <> "" Then
        If Mode = "USB" Then
            If Write_USB("STAGE.LST") = True Then
                DownSTGE = True
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("STAGE.LST") & " ") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownSTGE = False
                Unload Me
                Exit Function
            End If
         End If
        DownSTGE = False
    Else
        MsgBox "No Schedule File Created ...", vbInformation, "BUS"
        DownSTGE = False
    End If
    
End Function


Function DownGRSTAGE() As Boolean   ''' 21/01/2011
On Error Resume Next
    If Dir$(App.Path & "\GRSTAGE.LST") <> "" Then
        If Mode = "USB" Then
            If Write_USB("GRSTAGE.LST") = True Then
                DownGRSTAGE = True
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("GRSTAGE.LST") & " ") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownGRSTAGE = False
                Unload Me
                Exit Function
            End If
         End If
        DownGRSTAGE = False
    Else
        MsgBox "No Schedule File Created ...", vbInformation, "BUS"
        DownGRSTAGE = False
    End If
    
End Function

Function DownRTE() As Boolean
On Error Resume Next
    If Dir$(App.Path & "\RTE.DAT") <> "" Then
        If Mode = "USB" Then
            If Write_USB("RTE.DAT") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownRTE = False
                Unload Me
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("RTE.DAT") & " ") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownRTE = False
                Unload Me
                Exit Function
            End If
        End If
        DownRTE = True
    Else
        MsgBox "No Schedule File Created ...", vbInformation, "BUS"
        DownRTE = False
    End If
End Function

Public Sub CreateCurrency()
On Error GoTo err
Dim intHandle As Integer
Dim temp As CREATE_CUR
Dim DB As DAO.Database
'Dim strsql As String
Dim RES As DAO.Recordset
    If Dir(App.Path & "\CURRENCY.DAT", vbNormal) <> "" Then Call Kill(App.Path & "\CURRENCY.DAT")
    intHandle = FreeFile()
    Open App.Path & "\CURRENCY.DAT" For Binary Access Write As #intHandle
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    'strsql = "select * from CURRENCY"
    Set RES = DB.OpenRecordset("SELECT * FROM [CURRENCY]", dbOpenDynaset)
    If RES.RecordCount > 0 Then
        While RES.EOF <> True
            temp.CurString = Mid(IIf(IsNull(RES!Currency), "", RES!Currency), 1, 7) & Chr(0)
            Put #intHandle, , temp
            RES.MoveNext
        Wend
        Close #intHandle
    End If
    Exit Sub
err:
End Sub

Function DownCURRENCY() As Boolean
On Error GoTo err
   If Dir$(App.Path & "\CURRENCY.DAT") <> "" Then
        If Mode = "USB" Then '
            If Write_USB("CURRENCY.DAT") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownCURRENCY = False
                Unload Me
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("CURRENCY.DAT") & " ") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownCURRENCY = False
                Unload Me
                Exit Function
            End If
        End If
        DownCURRENCY = True
    Else
        MsgBox "Currency File not found...", vbInformation, "BUS"
        DownCURRENCY = False
    End If
    Exit Function
err:
End Function
Function DownCREW() As Boolean
On Error Resume Next
   If Dir$(App.Path & "\CREW.DAT") <> "" Then
        If Mode = "USB" Then '
            If Write_USB("CREW.DAT") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
'                If Mode = "USB" Then Call Disconnect_USB
                DownCREW = False
                Unload Me
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("CREW.DAT") & " ") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownCREW = False
                Unload Me
                Exit Function
            End If
        End If
        DownCREW = True
    Else
        MsgBox "No Crew File Created ...", vbInformation, "BUS"
        DownCREW = False
        
    End If
    
End Function
Function DownVEHICLE() As Boolean
On Error Resume Next
   If Dir$(App.Path & "\VEHICLE.DAT") <> "" Then
        If Mode = "USB" Then '
            If Write_USB("VEHICLE.DAT") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
'                If Mode = "USB" Then Call Disconnect_USB
                DownVEHICLE = False
                Unload Me
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("VEHICLE.DAT") & " ") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownVEHICLE = False
                Unload Me
                Exit Function
            End If
        End If
        DownVEHICLE = True
    Else
        MsgBox "No Vehicle File Created ...", vbInformation, "BUS"
        DownVEHICLE = False
        
    End If
    
End Function

Function DownEXPDET() As Boolean                       ''''''''''''RNC
On Error Resume Next
   If Dir$(App.Path & "\EXPENSEDET.DAT") <> "" Then
        If Mode = "USB" Then '
            If Write_USB("EXPENSEDET.DAT") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
'                If Mode = "USB" Then Call Disconnect_USB
                DownEXPDET = False
                Unload Me
                Exit Function
            End If
        Else
            If Trans(" 2 " & UCase("EXPENSEDET.DAT") & " ") = False Then
                MsgBox "Error in  Communication with ETM", vbCritical, "ERROR!"
                DownEXPDET = False
                Unload Me
                Exit Function
            End If
        End If
        DownEXPDET = True
    Else
        MsgBox "No ExpenseDet File Created ...", vbInformation, "BUS"
        DownEXPDET = False
        
    End If
    
End Function


Public Function ConvFileL(FCname As String) As Boolean
On Error GoTo errLn
Dim fsoSg As Integer

Dim Fso As New FileSystemObject
Dim textStm As TextStream
Dim strBuf As String

Dim FlotN As Single
    fsoSg = FreeFile()
    Open TransferPath & "\LSAMT" For Binary Access Write As #fsoSg
    Set textStm = Fso.OpenTextFile(TransPath & "\" & FCname, ForReading, False)
    Do While Not textStm.AtEndOfStream
        strBuf = textStm.ReadLine
        FlotN = CSng(val(strBuf))
        Put #fsoSg, , FlotN
    Loop
    textStm.Close
    Close #fsoSg
    ConvFileL = True
    Exit Function
errLn:
    ConvFileL = False
End Function



Private Sub Timer1_Timer()
On Error GoTo err
    If bDataReadyTimer > 0 Then
        bDataReadyTimer = bDataReadyTimer - 1
    Else
        If blnDataReady = True Then
            If sRcvData <> "" Then
                UpdateData (sRcvData & vbCrLf)
                blnStatus = True
                sRcvData = ""
                lRcvdByteCount = 0
                blnDataReady = False
            End If
        End If
    End If
    Exit Sub
err:
    MsgBox "Error! " & vbCrLf & "Timer1_Timer() " & vbCrLf & "Error No: " & err.Number & vbCrLf & err.Description
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
    If GetDevices = True Then
    '   If USB_FLAG = True Then
    '    Disconnect_USB
    '   End If
       
       Timer2.Enabled = False
       lblUSBStatus.caption = "Amphibia is Ready for Communication"
       lblUSBStatus.ForeColor = vbBlue
       List1.Enabled = True
       List2.Enabled = True
       
    Else
       lblUSBStatus.caption = "Amphibia not Detected"
       lblUSBStatus.ForeColor = vbRed
       List1.Enabled = False
       List2.Enabled = False
       Timer2.Enabled = False
       USB_FLAG = False
    ' End If
    End If
End Sub
Private Sub lstRefresh(x As Integer)
On Error Resume Next
Dim idxCnt As Integer
     If x = 1 Then
       
        For idxCnt = 0 To List1.ListCount - 1
           List1.Selected(idxCnt) = False
        Next idxCnt
     Else
        For idxCnt = 0 To List2.ListCount - 1
           List2.Selected(idxCnt) = False
        Next idxCnt
       
    End If
End Sub

Private Sub lstRep(x As Integer, str As String)
On Error Resume Next
Dim idxCnt As Integer
     If x = 1 Then
       
        For idxCnt = 0 To List1.ListCount - 1
          If List1.List(idxCnt) = str Then Exit Sub
        Next idxCnt
        List1.AddItem str
        
     Else
        For idxCnt = 0 To List2.ListCount - 1
           If List2.List(idxCnt) = str Then Exit Sub
        Next idxCnt
        List2.AddItem str
    End If
End Sub


Function UploadTKT(Fint As String) As Boolean
    On Error GoTo errLn
    Dim FnameUp As String
    Dim SysD As String
    Dim SysT As String
   
    sql = "SELECT PALMTECID FROM SETTINGS"
        Set res1 = DB.OpenRecordset(sql, dbOpenDynaset)
            PalmID = Trim(res1!PalmtecID) '& Chr(0)
        'RES.Close
    SysD = Format(DateValue(Date), "DDMMYY")
    'SysD = Format(Day(Date) & Month(Date) & Year(Date), "DDMMYY")
    SysT = Time
    FnameUp = "TKTS" & Fint & ".DAT"
    '& SysD & "-" & SysT & "-" & PalmId
    If Dir$(TransPath & "\" & FnameUp) <> "" Then Kill TransPath & "\" & FnameUp
        If Mode = "USB" Then
            If Read_USB(FnameUp) = True Then
'                Timer1.Enabled = True
                UploadTKT = True
                Exit Function
            End If
        ElseIf chkBttTrans.Value = 1 Then
            If UploadFile(TransPath & "\", FnameUp) = True Then
                UploadTKT = True
                Exit Function
            End If
        Else
            If Trans(" 1 " & UCase(FnameUp) & " ") = True Then
'                Timer1.Enabled = True
                UploadTKT = True
                Exit Function
            End If
         End If
'    End If
                UploadTKT = False
                'RES.Close
                CNN.Close
                Exit Function
errLn:
    MsgBox "No Ticket file Uploaded ...", vbInformation, "BUS"
    UploadTKT = False
End Function

Function UploadPASS() As Boolean
    On Error GoTo errLn
    If Dir$(App.Path & "\PASS.PAS") <> "" Then Kill App.Path & "\PASS.PAS"
        If Mode = "USB" Then
            If Read_USB("PASS.PAS") = True Then
'                Timer1.Enabled = True
                UploadPASS = True
                Exit Function
            End If
        ElseIf chkBttTrans.Value = 1 Then
            If UploadFile(TransPath & "\", "PASS.PAS") = True Then
                UploadPASS = True
                Exit Function
            End If
        Else
            If Trans(" 1 " & UCase("PASS.PAS") & " ") = True Then
'                Timer1.Enabled = True
                UploadPASS = True
                Exit Function
            End If
                UploadPASS = False
                Exit Function
        End If
'    End If
                UploadPASS = False
                Exit Function
errLn:
    MsgBox "No Pass file Uploaded ...", vbInformation, "BUS"
    UploadPASS = False
End Function

Function UploadCON() As Boolean
    On Error GoTo errLn
    If Dir$(App.Path & "\Transfer\T.CON") <> "" Then Kill App.Path & "\Transfer\T.CON"
        If Mode = "USB" Then
            If Read_USB("T.CON") = True Then
'                Timer1.Enabled = True
                UploadCON = True
                Exit Function
            End If
        ElseIf chkBttTrans.Value = 1 Then
            If UploadFile(TransPath & "\", "T.CON") = True Then
                UploadCON = True
                Exit Function
            End If
        Else
            If Trans("1 T.CON") = True Then
 '               Timer1.Enabled = True
                UploadCON = True
                Exit Function
            End If
        End If
'    End If
                UploadCON = True
                Exit Function
errLn:
    MsgBox "No Concession file Uploaded ...", vbInformation, "BUS"
    UploadCON = False
End Function

Function UploadRPT() As Boolean
    On Error GoTo errLn
    If Dir$(App.Path & "\RPT01.DAT") <> "" Then Kill App.Path & "\RPT01.DAT"
        If Mode = "USB" Then
            If Read_USB("RPT01.DAT") = True Then
'                Timer1.Enabled = True
                UploadRPT = True
                'Exit Function
            End If
             If Read_USB("FAREWISE.DAT") = True Then
'                Timer1.Enabled = True
                UploadRPT = True
                Exit Function
            End If
        ElseIf chkBttTrans.Value = 1 Then
            If UploadFile(TransPath & "\", "RPT01.DAT") = True Then
                UploadRPT = True
                Exit Function
            End If
        Else
            If Trans(" 1 " & UCase("RPT01.DAT") & " ") = True Then
'                Timer1.Enabled = True
                UploadRPT = True
                Exit Function
            End If
                UploadRPT = False
                Exit Function
        End If
'    End If
                UploadRPT = False
                Exit Function
errLn:
    MsgBox "No Report file Uploaded ...", vbInformation, "BUS"
    UploadRPT = False
End Function


Function UploadSetup() As Boolean

 '<---DEBUG
                    On Error GoTo errLn
                    If Dir$(App.Path & "\BUS.DAT") <> "" Then Kill App.Path & "\BUS.DAT"
                        If Mode = "USB" Then
                            If Read_USB("BUS.DAT") = True Then
                '                Timer1.Enabled = True
                                UploadSetup = True
                                Exit Function
                            End If
                        ElseIf chkBttTrans.Value = 1 Then
                            If UploadFile(TransPath & "\", "BUS.DAT") = True Then
                                UploadSetup = True
                                Exit Function
                            End If
                        Else
                            If Trans(" 1 " & UCase("BUS.DAT") & " ") = True Then
                '                Timer1.Enabled = True
                                UploadSetup = True
                                Exit Function
                            End If
                                UploadSetup = False
                                Exit Function
                        End If
                '    End If
                                UploadSetup = False
                                Exit Function
errLn:
                    MsgBox "No Report file Uploaded ...", vbInformation, "BUS"
                    UploadSetup = False

                    
   '--GUBED>
                    
                    
                    
'
'
'
''   '''''''''''''''''''''syamkrishna
'                    On Error GoTo errLn
'                    If Dir$(App.Path & "\BUS.DAT") <> "" Then
'                        If MODE = "USB" Then
'                            If Read_USB("BUS.DAT") = True Then
'                '                Timer1.Enabled = True
'                                UploadSetup = True
'                                Exit Function
'                            End If
'                        Else
'                            If Trans(" 1 " & UCase("BUS.DAT") & " ") = True Then
'                '                Timer1.Enabled = True
'                                UploadSetup = True
'                                Exit Function
'                            End If
'                                UploadSetup = False
'                                Exit Function
'                        End If
'                    Else
'                    MsgBox "Setting file Not Found in Your Machine.Please update the settings"
'
'                    End If
'
'                                UploadSetup = False
'                                Exit Function
'
'errLn:
'                    MsgBox "No Report file Uploaded ...", vbInformation, "BUS"
'                    UploadSetup = False
'

End Function
Function UploadODMTR() As Boolean
    On Error GoTo errLn
    'If Dir$(App.Path & "\ODOMETER.DAT") <> "" Then Kill App.Path & "\ODOMETER.DAT"
        If Mode = "USB" Then
            If Read_USB("ODOMETER.DAT") = True Then
'                Timer1.Enabled = True
                UploadODMTR = True
                Exit Function
            End If
'        ElseIf chkBttTrans.Value = 1 Then
'            If UploadFile(TransPath & "\", "ODOMETER.DAT") = True Then
'                UploadODMTR = True
'                Exit Function
'            End If
'        Else
'            If Trans(" 1 " & UCase("ODOMETER.DAT") & " ") = True Then
''                Timer1.Enabled = True
'                UploadODMTR = True
'                Exit Function
'            End If
'                UploadODMTR = False
'                Exit Function
        End If
'    End If
                UploadODMTR = False
                Exit Function
errLn:
    MsgBox "No Odometer file Uploaded ...", vbInformation, "BUS"
    UploadODMTR = False
End Function

Function UploadInsptr() As Boolean

 '<---DEBUG
                    On Error GoTo errLn
                    If Dir$(App.Path & "\INSPECTOR.DAT") <> "" Then Kill App.Path & "\INSPECTOR.DAT"
                        If Mode = "USB" Then
                            If Read_USB("INSPECTOR.DAT") = True Then
                '                Timer1.Enabled = True
                                UploadInsptr = True
                                Exit Function
                            End If
'                        ElseIf chkBttTrans.Value = 1 Then
'                            If UploadFile(TransPath & "\", "INSPECTOR.DAT") = True Then
'                                UploadInsptr = True
'                                Exit Function
'                            End If
'                        Else
'                            If Trans(" 1 " & UCase("INSPECTOR.DAT") & " ") = True Then
'                '                Timer1.Enabled = True
'                                UploadInsptr = True
'                                Exit Function
'                            End If
'                                UploadInsptr = False
'                                Exit Function
                        End If
'                    End If
                                UploadInsptr = False
                                Exit Function
errLn:
                    MsgBox "No Inspector file Uploaded ...", vbInformation, "BUS"
                    UploadInsptr = False
End Function

Function UploadExpense() As Boolean

 '<---DEBUG
                    On Error GoTo errLn
                    If Dir$(App.Path & "\EXPENSE.DAT") <> "" Then Kill App.Path & "\EXPENSE.DAT"
                        If Mode = "USB" Then
                            If Read_USB("EXPENSE.DAT") = True Then
                '                Timer1.Enabled = True
                                UploadExpense = True
                                Exit Function
                            End If
'                        ElseIf chkBttTrans.Value = 1 Then
'                            If UploadFile(TransPath & "\", "EXPENSE.DAT") = True Then
'                                UploadExpense = True
'                                Exit Function
'                            End If
'                        Else
'                            If Trans(" 1 " & UCase("EXPENSE.DAT") & " ") = True Then
'                '                Timer1.Enabled = True
'                                UploadExpense = True
'                                Exit Function
'                            End If
'                                UploadExpense = False
'                                Exit Function
                        End If
'                    End If
                                UploadExpense = False
                                Exit Function
errLn:
                    MsgBox "No Expense file Uploaded ...", vbInformation, "BUS"
                    UploadExpense = False
End Function

Public Sub setEnable(x As Boolean)
On Error Resume Next
      Command1.Enabled = x
      CommandP.Enabled = x
      TextP.Enabled = x
      Frame1.Enabled = x
      Frame2.Enabled = x
End Sub

Public Function CovertColln(Fint As String) As Boolean
 On Error GoTo errLn
    Dim FS As New FileSystemObject
    Dim TcketPath As String
    Dim FHndl As Integer
    Dim fShndl As Integer
    Dim tktK As PTicket
    Dim fBuff As String
    Dim FnameUp As String
    Dim PFname As String
    Dim pHandle As Integer
    Dim sHandle As Integer
    Dim gPass As PASSCONC
    '''gPassCount added by syam
    Dim gPassCount As Long
    Dim iFull As Integer
    Dim iHalf As Integer
    Dim iPhy As Integer
    Dim iLugg As Integer
    Dim iSt As Integer, isc As Long, ilad As Long
    Dim lTotPassenger As Long
    Dim fTotAmount As Single
    Dim fTotLuggAmount As Single
    Dim strYear As String
    Dim SysD, SysT, PID As String, line_length As Integer
    gPassCount = 0
    line_length = 96
        TSQL = "SELECT * FROM PCSETUP"
        Set RES = DB.OpenRecordset(TSQL, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            TcketPath = RES!TICKET_PATH
            TransPath = RES!TRANSFER_PATH
        End If
        RES.Close
        
        sHandle = FreeFile
         Open App.Path & "\BUS.DAT" For Binary Access Read As #sHandle
            Get #sHandle, , HStr
        Close #sHandle
        
        PalmID = Mid$(HStr.PalmtecID, 1, InStr(1, HStr.PalmtecID, Chr(0)) - 1)
        
'        sql = "SELECT PALMTECID FROM SETTINGS"
'        Set RES = DB.OpenRecordset(sql, dbOpenDynaset)
'        If RES.RecordCount > 0 Then
'            PalmID = RES!PalmtecID
'        End If
'        RES.Close
        
        FnameUp = "TKTS" & Fint & ".DAT"
        'If Dir(TransPath & "\" & FnameUp) <> "" Then
        SysD = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
        'SysT = Replace(Time, ":", ".")
        SysT = Format(Time, "hhmmAM/PM")
        PID = Replace(PalmID, Chr(0), "")
        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
        TcketPath = TcketPath & "\" & SysD
        If FS.FolderExists(TcketPath) = False Then FS.CreateFolder TcketPath
        FnameUp = "SCHEDULE" & Fint & "-" & SysT & PID & ".TXT"
        If Dir(TcketPath & "\" & FnameUp) <> "" Then Kill TcketPath & "\" & FnameUp
        fShndl = FreeFile()
            Open TcketPath & "\" & FnameUp For Binary Access Write As #fShndl
                fBuff = ""
                fBuff = String(line_length, "_") & vbCrLf  '05/01/2010
                Put #fShndl, , fBuff
                fBuff = Format("TICKET No| ", "@@@@@@@@@@@")
                fBuff = fBuff & Format("FL| ", "@@@@")
                fBuff = fBuff & Format("HF| ", "@@@@")
                fBuff = fBuff & Format("LG| ", "@@@@")
                fBuff = fBuff & Format("PH| ", "@@@@")
                fBuff = fBuff & Format("ST| ", "@@@@")
                
                fBuff = fBuff & Format("Ladies| ", "@@@@@@@@")
                fBuff = fBuff & Format("SC| ", "@@@@")
                
                fBuff = fBuff & Format("PASS No|", "@@@@@@@@@")
                fBuff = fBuff & Format("FROM| ", "@@@@@")
                fBuff = fBuff & Format("TO| ", "@@@@@")
                fBuff = fBuff & Format("AMOUNT|", "@@@@@@@")
                fBuff = fBuff & Format("LUG_AMT| ", "@@@@@@@@")
                fBuff = fBuff & Format("TIME|", "@@@@@@")
                fBuff = fBuff & Format("DATE|", "@@@@@@@@@@@@") & vbCrLf
                Put #fShndl, , fBuff
                fBuff = String(line_length, "_") & "|" & vbCrLf  '05/01/2010
                Put #fShndl, , fBuff
                                
                
        FnameUp = "TKTS" & Fint & ".DAT"
        PFname = App.Path & "\PASS.PAS"
        FHndl = FreeFile()
            Open TransPath & "\" & FnameUp For Binary Access Read As #FHndl
               Do While Not EOF(FHndl)
                Dim str As String
                Get #FHndl, , tktK
                If tktK.TicketNo = -1 Then Exit Do
                If EOF(FHndl) = True Then Exit Do
                 With tktK
                    fBuff = Format(.TicketNo, " 00000000") & "| "
                    fBuff = fBuff & Format(.Full, "00") & "| "
                    fBuff = fBuff & Format(.Half, "00") & "| "
                    fBuff = fBuff & Format(.Lugg, "00") & "| "
                    fBuff = fBuff & Format(.Phy, "00") & "| "
                    fBuff = fBuff & Format(.st, "00") & "| "
                    
                    fBuff = fBuff & Format(.ladies_count, "000000") & "| "
                    fBuff = fBuff & Format(.seniar_count, "00") & "| "
                    
                    iFull = iFull + .Full
                    iHalf = iHalf + .Half
                    iLugg = iLugg + .Lugg
                    iPhy = iPhy + .Phy
                    
                    isc = isc + .seniar_count
                    ilad = ilad + .ladies_count
                    
                    iSt = iSt + .st
                    
                    If .Typ = 32 Then
                        If Dir(PFname) <> "" Then
                            pHandle = FreeFile()
                            Open PFname For Binary Access Read As pHandle
                            Do While Not EOF(pHandle)
                                Get #pHandle, , gPass
                                If .TicketNo = gPass.TicketNo Then Exit Do
                            Loop
                            Close #pHandle
                            str = TrimChr(gPass.PassNo)
                            'MsgBox str
                        Else
                            str = "  "
                        End If
                        '''''''''''SYAM ADDED
                        If str <> "  " Then
                        gPassCount = gPassCount + 1
                        End If
                        
                        
                        
                        ''''''''''''''
                        
                        fBuff = fBuff & Format(str & "|", "@@@@@@@@@")
                    Else
                        fBuff = fBuff & String(7, " ") & "-|"
                    End If
                    fBuff = fBuff & Format(.From, " 000") & "|"
                    fBuff = fBuff & Format(.To, " 000") & "|"
                    str = Format(.Amount, "0.00")
                    fBuff = fBuff & Format(str, "@@@@@@@") & "|"
                    str = Format(.Luggage, "0.00")
                    fBuff = fBuff & Format(str, "@@@@@@@") & "|"
                    fBuff = fBuff & " " & Format(.Hr & ":" & .Minut, "HH:MM") & "|"
                    strYear = ""  '05/01/2010
                    strYear = DatePart("YYYY", Date) '05/01/2010
                    fBuff = fBuff & " " & Format(.Dy & "/" & .Mn & "/" & strYear, "DD/MM/YYYY") & "|" & vbCrLf  '05/01/2010
                    Put #fShndl, , fBuff
                     
                    
                    fTotAmount = fTotAmount + .Amount
                    fTotLuggAmount = fTotLuggAmount + .Luggage
                    fBuff = ""
                 End With
               Loop
                fBuff = String(line_length, "_") & "|" & vbCrLf '05/01/2010
                Put #fShndl, , fBuff
                lTotPassenger = iFull + iHalf + iPhy + iSt + gPassCount + ilad + isc
                
                
                
                ''''''''''''''''
                ''''''''''''''''''''
                Dim crpthndler As Integer
                crpthndler = FreeFile()
                Dim rpobj As PReport
                Dim ifullAmt As Double, iHalfAmt As Double, iPhyAmt As Double, istAmt As Double, adjustAmt As Double, ladamt As Double, scamt As Double
                Open App.Path & "\RPT01.DAT" For Binary Access Read As #crpthndler
                Do While Not EOF(crpthndler)
                    Get #crpthndler, , rpobj
                    If rpobj.SCHEDULE = Fint Then
                        ifullAmt = ifullAmt + rpobj.FullColl
                        iHalfAmt = iHalfAmt + rpobj.HalfColl
                        iPhyAmt = iPhyAmt + rpobj.PhyColl
                        istAmt = istAmt + rpobj.STColl
                        adjustAmt = adjustAmt + rpobj.AdjustColl
                        ladamt = ladamt + rpobj.ladies_coll
                        scamt = scamt + rpobj.seniar_coll
                    End If
                Loop
                 Close #crpthndler
                ''''''''''''''''''''
                '''''''''''''
                
                fTotAmount = fTotAmount - adjustAmt
                
                str = Format(ifullAmt, "0.00")
                fBuff = "TOTAL FULL       |" & Format(iFull, "@@@@@@@@@@@") & "|" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
                Put #fShndl, , fBuff
                
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
                
                str = Format(iHalfAmt, "0.00")
                fBuff = "TOTAL HALF       |" & Format(iHalf, "@@@@@@@@@@@") & "|" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
                Put #fShndl, , fBuff
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
                
                str = Format(iPhyAmt, "0.00")
                fBuff = "TOTAL PHY        |" & Format(iPhy, "@@@@@@@@@@@") & "|" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
                Put #fShndl, , fBuff
                
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
                
                str = Format(istAmt, "0.00")
                fBuff = "TOTAL ST         |" & Format(iSt, "@@@@@@@@@@@") & "|" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
                Put #fShndl, , fBuff
                'dd
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
                
                str = Format(scamt, "0.00")
                fBuff = "TOTAL SC         |" & Format(isc, "@@@@@@@@@@@") & "|" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
                Put #fShndl, , fBuff
                'dd
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
                
                str = Format(ladamt, "0.00")
                fBuff = "TOTAL Ladies     |" & Format(ilad, "@@@@@@@@@@@") & "|" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
                Put #fShndl, , fBuff
                'rr
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
                
                str = Format(fTotLuggAmount, "0.00")
                fBuff = "TOTAL LUGGAGE    |" & Format(iLugg, "@@@@@@@@@@@") & "|" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
                Put #fShndl, , fBuff
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
                
                
                 ''''''''''''''''''''''syam
                
                
               ' fBuff = "TOTAL PASS" & gPassCount
                '  Put #fShndl, , fBuff
                  
                fBuff = "TOTAL PASS       |" & Format(gPassCount, "@@@@@@@@@@@") & "|" & "           " & "|" & vbCrLf
                Put #fShndl, , fBuff
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
               ''''''''''''''''''''''''''
                
                fBuff = "TOTAL PASSENGER  |" & Format(lTotPassenger, "@@@@@@@@@@@") & "|" & "           " & "|" & vbCrLf
                Put #fShndl, , fBuff
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
'
'                str = Format(fTotLuggAmount, "0.00")
'                fBuff = "TOTAL LUGGAGE    |" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
'                Put #fShndl, , fBuff
'                fBuff = String(29, "_") & "|" & vbCrLf
 '               Put #fShndl, , fBuff
                
                str = Format(adjustAmt, "0.00")
                fBuff = "Adjust           |" & "           " & "|" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
                Put #fShndl, , fBuff
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
                
                
                str = Format(fTotAmount, "0.00")
                fBuff = "TOTAL AMOUNT     |" & "           " & "|" & Format(str, "@@@@@@@@@@@") & "|" & vbCrLf
                Put #fShndl, , fBuff
                fBuff = String(29, "_") & "|" & String(12, "_") & vbCrLf
                Put #fShndl, , fBuff
                
              DBTKTS (FnameUp)
                
                
            Close #FHndl
         Close #fShndl
         CovertColln = True
        Exit Function
      'End If
errLn:
  CovertColln = False
End Function

Public Function AddPasscount() As Boolean     'SANGEETHA
On Error GoTo err
Dim Sqlselect, sqlpass, sqlExp As String


Dim Filehdl As Integer
Dim rp As PReport
Dim gp As PASSCONC
Dim sql As String
Dim sHandle As Integer
Dim PalmID As String
Dim sqlRPT As String
   
    
    sHandle = FreeFile()
    Open App.Path & "\BUS.DAT" For Binary Access Read As #sHandle
        Get #sHandle, , HStr
    Close #sHandle
    
    PalmID = Mid$(HStr.PalmtecID, 1, InStr(1, HStr.PalmtecID, Chr(0)) - 1)
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")

    Filehdl = FreeFile()
    Open App.Path & "\" & "RPT01.DAT" For Binary Access Read As #Filehdl
    Do While Not EOF(Filehdl)
        
        Get #Filehdl, , rp
        'If EOF(Filehdl) Then Exit Do
        If rp.SCHEDULE = 0 Then Exit Do
        Dim DA As String, DAA As String
        DA = rp.StartD & "/" & rp.StartMO & "/" & rp.StartY
        DAA = rp.EndD & "/" & rp.EndMO & "/" & rp.EndY
'        sql = "DELETE * FROM RPT WHERE PALMID=" & Trim(PalmId) & " AND SCHEDULE= " & rp.SCHEDULE & " AND DATE='" & Format(Now, "DD/MM/YYYY") & "'"DATE=DateValue ('" & DTSchDate & "')"
        Sqlselect = "SELECT * FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & rp.SCHEDULE & " AND TRIPNO= " & rp.TripNo & " AND DATE BETWEEN DateValue('" & Format(DA, "DD/MM/YYYY") & "') AND DateValue('" & Format(DAA, "DD/MM/YYYY") & "')"
        
        't.CollectDt BETWEEN  DateValue('" & DTFrom.Value & "') AND DateValue('" & DTTo.Value & "')
        
       ' DB.Execute (sql)

        Set RES6 = DB.OpenRecordset(Sqlselect, dbOpenDynaset)
        Dim f1 As Boolean
       
        If RES6.RecordCount > 0 Then
        
           ' Do While Not RES6.EOF
            
                sqlpass = ""
                sqlpass = "SELECT count(PassNo) as PSC  from tkts where  PALMID='" & Trim(PalmID) & "' AND Schdule= " & rp.SCHEDULE & " AND TRIPNO= " & rp.TripNo & " AND DATE BETWEEN DateValue('" & Format(DA, "DD/MM/YYYY") & "') AND DateValue('" & Format(DAA, "DD/MM/YYYY") & "') and PassNo<>''"
                
                Set RES7 = DB.OpenRecordset(sqlpass, dbOpenDynaset)
                RES6.Edit
                RES6!pass = RES7!PSC
                
                RES6.Update
            'Loop
            
           
            RES6.MoveNext
            
        End If
        
         Loop
        
    Close #Filehdl
    If f1 Then RES.Close
    Exit Function
err:
    MsgBox "Error!" & vbCrLf & "Err No : " & err.Number & vbCrLf & err.Description, vbCritical
End Function
Public Function CovertRPT() As Boolean
 On Error GoTo errLn
    Dim FHndl As Integer
    Dim rptK As PReport
    Dim Nsche As Integer
    Dim rootRep As String
    Dim fShndl As Integer
    Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
        If Dir(TransPath & "\RPT01.DAT") <> "" Then
          fShndl = FreeFile()
           If Dir(TransPath & "\TRIP_RPT.TXT") <> "" Then Kill (TransPath & "\TRIP_RPT.TXT")
            Open TransPath & "\TRIP_RPT.TXT" For Binary Access Write As #fShndl
                rootRep = ""
                    rootRep = rootRep & Format("Full", "@@@@@@")
                    rootRep = rootRep & Format("FullColl", "@@@@@@@@@")
                    rootRep = rootRep & Format("Half", "@@@@@@")
                    rootRep = rootRep & Format("HalfColl", "@@@@@@@@@")
                    rootRep = rootRep & Format("Lugg", "@@@@@@")
                    rootRep = rootRep & Format("LuggColl", "@@@@@@@@@")
                    rootRep = rootRep & Format("Phy", "@@@@@@")
                    rootRep = rootRep & Format("PhyColl", "@@@@@@@@")
                    rootRep = rootRep & Format("st", "@@@@@@")
                    rootRep = rootRep & Format("STColl", "@@@@@@@")
                    rootRep = rootRep & Format("Adjust", "@@@@@@@")
                    rootRep = rootRep & Format("AdjustColl", "@@@@@@@@@@@")
                    rootRep = rootRep & Format("Conductor", "@@@@@@@@@@")
                    rootRep = rootRep & Format("Driver", "@@@@@@@")
                    rootRep = rootRep & Format("Cleaner", "@@@@@@@@")
                    rootRep = rootRep & Format("STicketNo", "@@@@@@@@@@")
                    rootRep = rootRep & Format("ETicketNo", "@@@@@@@@@@")
                    rootRep = rootRep & Format("Free", "@@@@@@")
                    rootRep = rootRep & Format("Date", "@@@@@@")
                    rootRep = rootRep & Format("SCHEDULE", "@@@@@@@@@")
                    rootRep = rootRep & Format("NoOfMisBill", "@@@@@@@@@@@")
                    rootRep = rootRep & Format("RouteCode", "@@@@@@@@@@")
                    rootRep = rootRep & Format("BusNo", "@@@@@@") & vbCrLf
                    Put #fShndl, , rootRep
                    
                                    ''''''''''''''''''''''''''''test syam
                  
                '''''''''''''''''''''''''
                
                    
        FHndl = FreeFile()
         
            Open TransPath & "\RPT01.DAT" For Binary Access Read As #FHndl
               Do While Not EOF(FHndl)
                
                
                Get #FHndl, , rptK
                If rptK.STicketNo <= 0 Then Exit Do
'                If EOF(fHndl) = True Then
'                    Exit Do
 '               Else
                
                With rptK
                
                    rootRep = rootRep & Format(.Full, "@@@@@@")
                    rootRep = rootRep & Format(.FullColl, "@@@@@@@@@")
                    rootRep = rootRep & Format(.Half, "@@@@@@")
                    rootRep = rootRep & Format(.HalfColl, "@@@@@@@@@")
                    rootRep = rootRep & Format(.Lugg, "@@@@@@")
                    rootRep = rootRep & Format(.LuggageColl, "@@@@@@@@@")
                    rootRep = rootRep & Format(.Phy, "@@@@@@")
                    rootRep = rootRep & Format(.PhyColl, "@@@@@@@@")
                    rootRep = rootRep & Format(.st, "@@@@@@")
                    rootRep = rootRep & Format(.STColl, "@@@@@@@")
                    rootRep = rootRep & Format(.Adjust, "@@@@@@")
                    rootRep = rootRep & Format(.AdjustColl, "@@@@@@@@@@@")
                    rootRep = rootRep & Format(.Conductor, "@@@@@@@@@@")
                    rootRep = rootRep & Format(.Driver, "@@@@@@@")
                    rootRep = rootRep & Format(.Cleaner, "@@@@@@@@")
                    rootRep = rootRep & Format(.STicketNo, "@@@@@@@@@@")
                    rootRep = rootRep & Format(.ETicketNo, "@@@@@@@@@@")
                    rootRep = rootRep & Format(.Free, "@@@@@@") & " "
                    rootRep = rootRep & Format(.EndD & "/" & .EndM & "/" & .EndY, " DD/MM/YY ")
                    rootRep = rootRep & Format(.SCHEDULE, "@@@@@@@@@")
                    rootRep = rootRep & Format(.FareType, "@@@@@@@@@@@@")
                    rootRep = rootRep & Format(.RouteCode, "@@@@@@@@@@")
                    rootRep = rootRep & Format(.Busno, "@@@@@@") & vbCrLf
                    Put #fShndl, , rootRep
                                    
                                    
                  ''''''''''''''''''''''''''''
                
                '''''''''''''''''''''''''
                
                    
                    rootRep = ""
                    Nsche = .SCHEDULE
                End With
 '               End If
               Loop
            Close #FHndl
            Close #fShndl
        End If
        Dim I As Integer
    For I = 1 To Nsche
        DoEvents
        List1.AddItem "SCHEDULE  " & Format(I, "00"), I
        DoEvents
    Next I
    For I = 0 To List1.ListCount - 1
        List1.Selected(I) = True
        'List1.Selected(I) = False
    Next I
    List1.Selected(0) = True
    'List1.Enabled = True
    Do While Not Nsche = List1.ListCount - 1
        List1.RemoveItem Nsche + 1
    Loop
    '***********Updating to Datbase
    DBRpt ("RPT01.DAT")
    
    CovertRPT = True
   Exit Function
errLn:
  CovertRPT = False
End Function


Function UploadPRM() As Boolean
    On Error GoTo errLn
    If Dir$(App.Path & "\Transfer\PRM") <> "" Then Kill App.Path & "\Transfer\PRM"
        If Mode = "USB" Then
            If Read_USB("PRM") = True Then
'                Timer1.Enabled = True
                UploadPRM = True
                Exit Function
            End If
        Else
            If Trans("1 PRM") = True Then
'                Timer1.Enabled = True
                UploadPRM = True
                Exit Function
            End If
                UploadPRM = False
                Exit Function
        End If
'    End If
                UploadPRM = True
                Exit Function
errLn:
    MsgBox "No WAYBILL file Uploaded ...", vbInformation, "BUS"
    UploadPRM = False
End Function

'Public Function TrimChr(ByVal Buf As String) As String
'    Dim tmp As Integer
'    Dim Buf1 As String
'    tmp = InStr(Buf, Chr$(0))
'    If tmp > 0 Then
'        TrimChr = Trim(Mid(Buf, 1, tmp - 1))
'    Else
'        TrimChr = Trim(Buf)
'    End If
'    'TrimChr = left(Buf, InStr(1, strString, Chr$(0)) - 1) 'Mid(Buf, 1, (InStr(Buf, Chr(0)) - 1))
'End Function

Private Sub Timer3_Timer()
    DoEvents
    If TimeOut > 0 Then
        TimeOut = TimeOut - 1
    End If
End Sub

Public Function EditSettings() As Boolean
Dim Fhandle As Integer

    Open App.Path & "\BUS.DAT" For Binary Access Read Write As #Fhandle
        Get #Fhandle, , HStr
        Get #Fhandle, , hardwaresettings
    Close #Fhandle
    
    Kill App.Path & "\BUS.DAT"
    Fhandle = FreeFile()
    Open App.Path & "\BUS.DAT" For Binary Access Write As #Fhandle
    Close #Fhandle

End Function

Private Function TripClosed() As Boolean
Dim intHandle As Integer
Dim bytStatus As Byte
On Error Resume Next
    If Mode <> "USB" Then
        If Trans(" 1 STATUS.DAT") = False Then Exit Function
    Else
        If Read_USB("STATUS.DAT") = False Then Exit Function
    End If
    If Dir(TransPath & "\STATUS.DAT", vbNormal) = "" Then Exit Function
    intHandle = FreeFile
    Open TransPath & "\STATUS.DAT" For Binary Access Read As #intHandle
    Get #intHandle, , bytStatus
    Close #intHandle
    TripClosed = IIf((bytStatus And 2) = 2, False, True)
End Function

Private Function ScheduleClosed() As Boolean
Dim intHandle As Integer
Dim bytStatus As Byte
On Error Resume Next
    If Mode <> "USB" Then
        If Trans(" 1 STATUS.DAT") = False Then Exit Function
    Else
        If Read_USB("STATUS.DAT") = False Then Exit Function
    End If
    If Dir(TransPath & "\STATUS.DAT", vbNormal) = "" Then Exit Function
    intHandle = FreeFile
    Open TransPath & "\STATUS.DAT" For Binary Access Read As #intHandle
    Get #intHandle, , bytStatus
    Close #intHandle
    ScheduleClosed = IIf((bytStatus And 1) = 1, False, True)
End Function

Private Function IsTicketRemoveEnabled() As Boolean
Dim rsRecord As DAO.Recordset
On Error GoTo CatchError
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    Set rsRecord = DB.OpenRecordset("SELECT * FROM Settings", dbOpenDynaset)
    If rsRecord.RecordCount > 0 Then
        If IIf(IsNull(rsRecord!RemoveTicketFlag), 0, rsRecord!RemoveTicketFlag) = 1 Then
            IsTicketRemoveEnabled = True
        End If
    End If
    rsRecord.Close
    Set rsRecord = Nothing
    Exit Function
CatchError:
End Function

Private Sub Timer4_Timer()
If lTimer > 0 Then
        lTimer = lTimer - 1
    Else
        lRcvdByteCount = 0
        blnDataReady = False
        blnStatus = False
        sRcvData = ""
        sStatus = ""
    End If
End Sub
