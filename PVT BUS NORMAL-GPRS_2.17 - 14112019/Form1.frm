VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{011424E4-FAAB-4D1D-B936-E5C631EEDC26}#1.0#0"; "SmartProgressBar.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Tfrm 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   7410
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9840
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Transferring through IR"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   135
         TabIndex        =   16
         Top             =   -1395
         Width           =   2640
      End
      Begin VB.CommandButton CommandP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SET P&ORT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   870
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtPacketNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8550
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   -2040
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTotalPacket 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8550
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1170
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
         Left            =   120
         TabIndex        =   9
         Top             =   4800
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
            Height          =   1680
            ItemData        =   "Form1.frx":0000
            Left            =   150
            List            =   "Form1.frx":0002
            Style           =   1  'Checkbox
            TabIndex        =   11
            Top             =   285
            Width           =   9285
         End
         Begin JeweledBut.JeweledButton cmdUpload 
            Height          =   375
            Left            =   7920
            TabIndex        =   10
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "Form1.frx":0004
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
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   9585
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
            Height          =   1680
            ItemData        =   "Form1.frx":0020
            Left            =   120
            List            =   "Form1.frx":0022
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   255
            Width           =   9285
         End
         Begin VB.CheckBox ChkSelect 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Select All"
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   1710
         End
         Begin JeweledBut.JeweledButton cmdDownload 
            Height          =   375
            Left            =   7920
            TabIndex        =   8
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "Form1.frx":0024
            BC              =   12632256
            FC              =   0
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   9165
         Top             =   0
      End
      Begin VB.TextBox TextP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1770
         TabIndex        =   4
         Top             =   -750
         Width           =   495
      End
      Begin VB.FileListBox FileContainer 
         Height          =   285
         Left            =   2970
         TabIndex        =   3
         Top             =   855
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Interval        =   5
         Left            =   4005
         Top             =   1740
      End
      Begin VB.Timer Timer3 
         Interval        =   600
         Left            =   3510
         Top             =   1740
      End
      Begin VB.CheckBox chkBttTrans 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Uploading through Wireless 2.4 GHz"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   -1080
         Width           =   3720
      End
      Begin VB.Timer Timer4 
         Interval        =   500
         Left            =   5160
         Top             =   1680
      End
      Begin SmartProgressBar.SmartPrgress pBar 
         Height          =   255
         Left            =   240
         Top             =   4440
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
      Begin JeweledBut.JeweledButton Command1 
         Height          =   510
         Left            =   3675
         TabIndex        =   2
         Top             =   855
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
         MICON           =   "Form1.frx":0040
         BC              =   12632256
         FC              =   0
      End
      Begin MSCommLib.MSComm SerialCom 
         Left            =   7605
         Top             =   75
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTTrans 
         Height          =   300
         Left            =   3750
         TabIndex        =   17
         Top             =   870
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MM-yy"
         Format          =   4063235
         CurrentDate     =   37863
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   3540
         TabIndex        =   23
         Top             =   75
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Packets"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7185
         TabIndex        =   22
         Top             =   -1335
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tx/Rx Packet No"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7185
         TabIndex        =   21
         Top             =   -1665
         Width           =   1365
      End
      Begin VB.Label lblUSBStatus 
         Alignment       =   1  'Right Justify
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
         Left            =   585
         TabIndex        =   20
         Top             =   1560
         Width           =   5490
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port No."
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   165
         TabIndex        =   19
         Top             =   -750
         Width           =   1545
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
         Left            =   5745
         TabIndex        =   18
         Top             =   4515
         Width           =   3825
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim san As New ADODB.Recordset
    If DatabaseADOB_Connection() = False Then MsgBox "cannot connect db"
    sql = "select * from Rpt"
    If san.State = 1 Then san.Close
    san.Open sql, adoc, adOpenDynamic, adLockOptimistic
    Do While Not san.EOF
        If san.EOF = True Then Exit Do
        Total = Total + (san!Full)
        san.MoveNext
    Loop
    If san.State = adStateOpen Then san.Close
End Sub
Public Function DatabaseADOB_Connection() As Boolean
    On Error GoTo CatchError
    If Dir(App.Path & "\Pvt.mdb", vbNormal) = "" Then
        MsgBox "Database file not found!", vbExclamation
        Exit Function
    End If
    If adoc.State = 1 Then adoc.Close
    If adoc.State <> 1 Then
        adoc.Provider = "Microsoft.Jet.OLEDB.4.0"
        adoc.Properties("Jet OLEDB:Database Password") = "silbus"
        adoc.ConnectionString = App.Path & "\Pvt.mdb"
        adoc.Open
    End If
    If adoc.State = adStateOpen Then
        DatabaseADOB_Connection = True
    Else
        DatabaseADOB_Connection = False
    End If
Exit Function
CatchError:
    MsgBox "Database Error! " & vbCrLf & "Error Number : " & err.Number & vbTab & "Description : " & err.Description, vbExclamation
End Function
