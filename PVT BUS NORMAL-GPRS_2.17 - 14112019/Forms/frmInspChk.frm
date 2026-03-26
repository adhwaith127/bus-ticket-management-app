VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmInspChk 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inspector Check Report"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frSelectSchedule 
      BackColor       =   &H00E0E0E0&
      Height          =   2835
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5115
      Begin VB.CheckBox chkEnableDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Date"
         Height          =   240
         Left            =   3500
         TabIndex        =   6
         Top             =   800
         Width           =   1245
      End
      Begin VB.OptionButton optReportType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Date Wise"
         Height          =   300
         Index           =   1
         Left            =   2415
         TabIndex        =   5
         Top             =   225
         Width           =   1590
      End
      Begin VB.OptionButton optReportType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Schedule Wise"
         Height          =   300
         Index           =   0
         Left            =   285
         TabIndex        =   4
         Top             =   225
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.ComboBox cmbPalmID 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cmbShedule 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   450
         Left            =   3570
         TabIndex        =   7
         Top             =   2235
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   794
         TX              =   "&Exit"
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
         MICON           =   "frmInspChk.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdOdmtr 
         Height          =   450
         Left            =   2475
         TabIndex        =   8
         Top             =   2235
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   794
         TX              =   "&OK"
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
         MICON           =   "frmInspChk.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin MSComCtl2.DTPicker DTSchDate 
         Height          =   300
         Left            =   690
         TabIndex        =   9
         Top             =   1770
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   99418113
         CurrentDate     =   39536
      End
      Begin MSComCtl2.DTPicker DTEnd 
         Height          =   330
         Left            =   3300
         TabIndex        =   10
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         _Version        =   393216
         Format          =   99418113
         CurrentDate     =   39536
      End
      Begin MSComCtl2.DTPicker DTStart 
         Height          =   345
         Left            =   3285
         TabIndex        =   11
         Top             =   900
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   609
         _Version        =   393216
         Format          =   99418113
         CurrentDate     =   39536
      End
      Begin VB.Label lblEndDateOrSch 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "        Schedule :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblSdateOrID 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter PalmtecID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1770
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00B0524A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inspector Check Report"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   4320
   End
End
Attribute VB_Name = "FrmInspChk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Icon = frmMainform.Icon
End Sub
