VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmSetDateTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set ETM Date and Time"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpSetDateTime 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH.mm.ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16576
      CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
      Format          =   111542275
      CurrentDate     =   40007
   End
   Begin JeweledBut.JeweledButton cmdClose 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   840
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
      TX              =   "C&lose"
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
      MICON           =   "frmSetDateTime.frx":0000
      BC              =   12632256
      FC              =   0
      Picture         =   "frmSetDateTime.frx":001C
   End
   Begin JeweledBut.JeweledButton cmdSave 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
      TX              =   "&Send"
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
      MICON           =   "frmSetDateTime.frx":2668
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Date and Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "frmSetDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim strDateTime As String
    strDateTime = Format(Day(dtpSetDateTime.Value), "00") & "/" & _
                Format(Month(dtpSetDateTime.Value), "00") & "/" & _
                Format(Year(dtpSetDateTime.Value), "0000") & " " & _
                Format(Hour(dtpSetDateTime.Value), "00") & ":" & _
                Format(Minute(dtpSetDateTime.Value), "00") & ":" & _
                Format(Second(dtpSetDateTime.Value), "00")
    If MsgBox("Do you want to Set ETM Date and Time to " & strDateTime, vbYesNo + vbQuestion, App.ProductName) = vbYes Then
        strDateTime = Replace(strDateTime, "/", "")
        strDateTime = Replace(strDateTime, ":", "")
        strDateTime = Replace(strDateTime, " ", "")
        If GetDevices = True Then
            If Write_USB("DATEANDTIME-" & strDateTime) = True Then
                MsgBox "Date and Time of ETM set successfully", vbInformation, App.ProductName
            Else
                MsgBox "Communication failure! Date and Time of ETM not set", vbInformation, App.ProductName
            End If
        Else
                MsgBox "Communication failure! Date and Time of ETM not set", vbInformation, App.ProductName
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Icon = frmMainform.Icon
    dtpSetDateTime.Value = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call Disconnect_USB
End Sub
