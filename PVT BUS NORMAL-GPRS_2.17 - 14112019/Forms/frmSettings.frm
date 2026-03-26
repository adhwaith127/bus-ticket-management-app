VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   13440
   ControlBox      =   0   'False
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmSettings.frx":0CCA
   ScaleHeight     =   8580
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox sc_ratio_txt 
      Height          =   405
      Left            =   7440
      MaxLength       =   5
      TabIndex        =   23
      Top             =   4920
      Width           =   885
   End
   Begin VB.TextBox lad_ratio_txt 
      Height          =   405
      Left            =   5130
      MaxLength       =   5
      TabIndex        =   22
      Top             =   4920
      Width           =   885
   End
   Begin VB.CheckBox chk_refund 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11040
      TabIndex        =   65
      Top             =   7440
      Width           =   210
   End
   Begin VB.CheckBox chk_userpswd 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6720
      TabIndex        =   45
      Top             =   5640
      Width           =   210
   End
   Begin VB.CheckBox chk_AutoShutdown 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11040
      TabIndex        =   64
      Top             =   7080
      Width           =   210
   End
   Begin VB.TextBox txtsmsph3 
      Height          =   405
      Left            =   9960
      MaxLength       =   12
      TabIndex        =   33
      Top             =   4560
      Width           =   3375
   End
   Begin VB.ComboBox cmb_font 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5385
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4440
      Width           =   2145
   End
   Begin VB.CheckBox chk_inspector 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11040
      TabIndex        =   63
      Top             =   6360
      Width           =   210
   End
   Begin VB.CheckBox chk_multiple 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11040
      TabIndex        =   61
      Top             =   5640
      Width           =   210
   End
   Begin VB.CheckBox chk_simple 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6720
      TabIndex        =   60
      Top             =   7440
      Width           =   210
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   12360
      TabIndex        =   148
      Top             =   7920
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtSTRoundAmt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   147
      Text            =   "50"
      Top             =   5520
      Width           =   1005
   End
   Begin VB.CheckBox chkSTroundEnable 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   144
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox Chkinspect 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11040
      TabIndex        =   50
      Top             =   6720
      Width           =   210
   End
   Begin VB.CheckBox chkftp 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11040
      TabIndex        =   59
      Top             =   5280
      Width           =   210
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2640
      TabIndex        =   140
      Top             =   13200
      Width           =   210
   End
   Begin VB.CheckBox chkexp 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8880
      TabIndex        =   58
      Top             =   7440
      Width           =   210
   End
   Begin VB.CheckBox chksmart 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6720
      TabIndex        =   51
      Top             =   7800
      Width           =   210
   End
   Begin VB.CheckBox chkgprs 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8880
      TabIndex        =   52
      Top             =   5280
      Width           =   210
   End
   Begin VB.CheckBox chkgprsmsgenable 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8880
      TabIndex        =   53
      Top             =   5640
      Width           =   210
   End
   Begin VB.CheckBox chksendbillenable 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8880
      TabIndex        =   56
      Top             =   6720
      Width           =   210
   End
   Begin VB.CheckBox chktripsendenable 
      Caption         =   "Check2"
      Height          =   255
      Left            =   8880
      TabIndex        =   55
      Top             =   6360
      Width           =   210
   End
   Begin VB.CheckBox chkschedulesendenable 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8880
      TabIndex        =   54
      Top             =   6000
      Width           =   210
   End
   Begin VB.CheckBox chksendpendinbill 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8880
      TabIndex        =   57
      Top             =   7080
      Width           =   210
   End
   Begin VB.TextBox txtsmsph2 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9960
      MaxLength       =   12
      TabIndex        =   32
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox txtaccesspoint 
      Height          =   405
      Left            =   9960
      MaxLength       =   23
      TabIndex        =   24
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtftpurl 
      Height          =   405
      Left            =   9960
      MaxLength       =   31
      TabIndex        =   25
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtftpuname 
      Height          =   405
      Left            =   9960
      MaxLength       =   15
      TabIndex        =   26
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtftppswd 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   9960
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   27
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox txtuploadpath 
      Height          =   405
      Left            =   9960
      TabIndex        =   28
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtdwnldpath 
      Height          =   405
      Left            =   9960
      TabIndex        =   29
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txthttpurl 
      Height          =   405
      Left            =   9960
      MaxLength       =   63
      TabIndex        =   30
      Top             =   3120
      Width           =   3375
   End
   Begin VB.CheckBox chkSh 
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   37
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox chkTRIP 
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   36
      Top             =   7680
      Width           =   255
   End
   Begin VB.TextBox txtphno 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9960
      MaxLength       =   12
      TabIndex        =   31
      Text            =   " "
      Top             =   3600
      Width           =   3405
   End
   Begin VB.CheckBox chkCrewcheckEnable 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   49
      Top             =   7080
      Width           =   210
   End
   Begin VB.CheckBox chkTcktBold 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   43
      Top             =   7440
      Width           =   210
   End
   Begin VB.CheckBox chkOdmtrEnable 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   44
      Top             =   7800
      Width           =   210
   End
   Begin VB.CheckBox chkDefaultStage 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   47
      Top             =   6360
      Width           =   210
   End
   Begin VB.CheckBox chkStageUpdation 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   40
      Top             =   6360
      Width           =   210
   End
   Begin VB.CheckBox chkRemoveTicketEnable 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   46
      Top             =   6000
      Width           =   210
   End
   Begin VB.CheckBox chkBoldFontEnable 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   42
      Top             =   7080
      Width           =   210
   End
   Begin VB.CheckBox chkstenable 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   6120
      Width           =   255
   End
   Begin VB.CheckBox chkbigfontenable 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   41
      Top             =   6720
      Width           =   210
   End
   Begin VB.CheckBox chkcrewenable 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   48
      Top             =   6720
      Width           =   210
   End
   Begin VB.CheckBox chklogenable 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   67
      Top             =   7680
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox chkAdvertise 
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   66
      Top             =   7800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox chkRnFair 
      Height          =   255
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   38
      Top             =   5640
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin JeweledBut.JeweledButton cmdLoadF 
      Height          =   375
      Left            =   4800
      TabIndex        =   71
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Load "
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
      MICON           =   "frmSettings.frx":22DFA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdSaveF 
      Height          =   375
      Left            =   6240
      TabIndex        =   70
      Top             =   8160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      TX              =   "Save As"
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
      MICON           =   "frmSettings.frx":22E16
      BC              =   12632256
      FC              =   0
   End
   Begin VB.CheckBox chkPrint 
      Height          =   195
      Left            =   11040
      TabIndex        =   62
      Top             =   6000
      Width           =   330
   End
   Begin VB.ComboBox cmbLocalLanguage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmSettings.frx":22E32
      Left            =   5340
      List            =   "frmSettings.frx":22E3F
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3960
      Width           =   2160
   End
   Begin JeweledBut.JeweledButton cmdCancel 
      Height          =   375
      Left            =   9000
      TabIndex        =   69
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "&EXIT"
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
      MICON           =   "frmSettings.frx":22E5F
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdSave 
      Height          =   375
      Left            =   7560
      TabIndex        =   68
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "&SAVE"
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
      MICON           =   "frmSettings.frx":22E7B
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox txtLuggage 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      MaxLength       =   24
      TabIndex        =   34
      Text            =   "0"
      Top             =   -120
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.ComboBox cmdCurrency 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   22740
      Style           =   2  'Dropdown List
      TabIndex        =   72
      Top             =   4320
      Width           =   2145
   End
   Begin VB.TextBox txtFTR2 
      Height          =   405
      Left            =   5340
      MaxLength       =   31
      TabIndex        =   19
      Top             =   3390
      Width           =   2115
   End
   Begin VB.TextBox txtFTR1 
      Height          =   405
      Left            =   5340
      MaxLength       =   31
      TabIndex        =   18
      Top             =   2850
      Width           =   2115
   End
   Begin VB.TextBox txtMainDisp2 
      Height          =   405
      Left            =   5340
      MaxLength       =   22
      TabIndex        =   14
      Top             =   825
      Width           =   2115
   End
   Begin VB.TextBox txtHDR1 
      Height          =   405
      Left            =   5340
      MaxLength       =   31
      TabIndex        =   15
      Text            =   "LOGO1"
      Top             =   1260
      Width           =   2115
   End
   Begin VB.TextBox txtHDR3 
      Height          =   405
      Left            =   5340
      MaxLength       =   31
      TabIndex        =   17
      Top             =   2280
      Width           =   2115
   End
   Begin VB.TextBox txtHDR2 
      Height          =   405
      Left            =   5340
      MaxLength       =   31
      TabIndex        =   16
      Text            =   "LOGO2"
      Top             =   1740
      Width           =   2115
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   600
      Left            =   2040
      TabIndex        =   83
      Top             =   4485
      Width           =   1335
      Begin VB.OptionButton optRoundUD 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Down"
         Height          =   240
         Index           =   0
         Left            =   150
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   12
         Top             =   315
         Width           =   855
      End
      Begin VB.OptionButton optRoundUD 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Up"
         Height          =   225
         Index           =   1
         Left            =   150
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   11
         Top             =   45
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox txtRoundAmt 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   5340
      MaxLength       =   3
      TabIndex        =   13
      Top             =   360
      Width           =   2085
   End
   Begin VB.TextBox txtSTMin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   8
      Text            =   "2.50"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtMain 
      Height          =   405
      Left            =   2010
      MaxLength       =   17
      TabIndex        =   2
      Top             =   1740
      Width           =   1365
   End
   Begin VB.TextBox txtID 
      Height          =   405
      Left            =   2010
      MaxLength       =   5
      TabIndex        =   3
      Top             =   2280
      Width           =   1365
   End
   Begin VB.TextBox txtHalfRatio 
      Height          =   405
      Left            =   2010
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2790
      Width           =   1365
   End
   Begin VB.TextBox txtST 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "25"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtPH 
      Height          =   405
      Left            =   2010
      MaxLength       =   2
      TabIndex        =   6
      Top             =   3240
      Width           =   1365
   End
   Begin VB.TextBox txtSTMax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   7
      Text            =   "23.50"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtmstPWD 
      Height          =   405
      Left            =   2010
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1260
      Width           =   1365
   End
   Begin VB.TextBox txtUserPWD 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2010
      MaxLength       =   6
      TabIndex        =   0
      Top             =   720
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   600
      Left            =   2010
      TabIndex        =   82
      Top             =   3750
      Width           =   1380
      Begin VB.OptionButton optRoundoff 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   60
         Width           =   975
      End
      Begin VB.OptionButton optRoundoff 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   330
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.CheckBox chkReport 
      Height          =   195
      Left            =   3360
      TabIndex        =   39
      Top             =   6000
      Width           =   210
   End
   Begin JeweledBut.JeweledButton cmdViewExp 
      Height          =   375
      Left            =   3120
      TabIndex        =   120
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      TX              =   "View Expense"
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
      MICON           =   "frmSettings.frx":22E97
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label78 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00000080&
      Height          =   165
      Left            =   8505
      TabIndex        =   163
      Top             =   5100
      Width           =   225
   End
   Begin VB.Label Label77 
      BackStyle       =   0  'Transparent
      Caption         =   "SC Ratio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6360
      TabIndex        =   162
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label76 
      BackStyle       =   0  'Transparent
      Caption         =   "Ladies Ratio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3720
      TabIndex        =   161
      Top             =   4995
      Width           =   1215
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00000080&
      Height          =   165
      Left            =   6105
      TabIndex        =   160
      Top             =   5055
      Width           =   225
   End
   Begin VB.Label Label75 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Refund"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      TabIndex        =   159
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   2
      Left            =   1800
      TabIndex        =   158
      Top             =   1822
      Width           =   165
   End
   Begin VB.Label Label74 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   157
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label73 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Shutdown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      TabIndex        =   156
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label72 
      BackStyle       =   0  'Transparent
      Caption         =   "SMS No3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   155
      Top             =   4635
      Width           =   2175
   End
   Begin VB.Label Label71 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Font"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3720
      TabIndex        =   154
      Top             =   4440
      Width           =   1800
   End
   Begin VB.Label Label70 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10080
      TabIndex        =   153
      Top             =   4335
      Width           =   45
   End
   Begin VB.Label Label69 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector SMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      TabIndex        =   152
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label68 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Multiple Pass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      TabIndex        =   151
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label67 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Simple Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   150
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label66 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Modom Always ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12720
      TabIndex        =   149
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label64 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RoundOff Amount (in Paise)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   120
      TabIndex        =   146
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Label Label63 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "ST RoundOff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   145
      Top             =   5040
      Width           =   1275
   End
   Begin VB.Label Label62 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector Report "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      TabIndex        =   143
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label53 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FTP "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      TabIndex        =   142
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label52 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gprs Enable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   141
      Top             =   13200
      Width           =   1935
   End
   Begin VB.Label Label51 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Send "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   139
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label46 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Smart Card "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   138
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label45 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gprs "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   137
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label65 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gprs Message "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   136
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "Send Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9240
      TabIndex        =   135
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "Trip Send "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   134
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   133
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "SendPendingTicket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   132
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "SMS No2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   131
      Top             =   4155
      Width           =   1455
   End
   Begin VB.Label Label55 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Access Point"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8040
      TabIndex        =   130
      Top             =   315
      Width           =   1215
   End
   Begin VB.Label Label56 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ftp Url"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8040
      TabIndex        =   129
      Top             =   795
      Width           =   1335
   End
   Begin VB.Label Label57 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ftp User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8040
      TabIndex        =   128
      Top             =   1215
      Width           =   1455
   End
   Begin VB.Label Label58 
      BackStyle       =   0  'Transparent
      Caption         =   "Ftp Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   127
      Top             =   1755
      Width           =   1935
   End
   Begin VB.Label Label59 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ftp Upload Path"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8040
      TabIndex        =   126
      Top             =   2235
      Width           =   1575
   End
   Begin VB.Label Label60 
      BackStyle       =   0  'Transparent
      Caption         =   "Ftp Download Path"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   125
      Top             =   2715
      Width           =   1815
   End
   Begin VB.Label Label61 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Http Url"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8040
      TabIndex        =   124
      Top             =   3195
      Width           =   975
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMS Prompt      "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   123
      Top             =   8040
      Width           =   1425
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMS Sending"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   122
      Top             =   7680
      Width           =   1245
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SMS No1"
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
      Left            =   7560
      TabIndex        =   121
      Top             =   3667
      Width           =   1920
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   " Crew Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   119
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "TicketNo Bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   118
      Top             =   7440
      Width           =   1230
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Odometer Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   117
      Top             =   7800
      Width           =   1320
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Repeat Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   116
      Top             =   6360
      Width           =   1230
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Stage Updation Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   115
      Top             =   6360
      Width           =   2145
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Remove Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   114
      Top             =   6000
      Width           =   1305
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Stage Name Bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   113
      Top             =   7080
      Width           =   1485
   End
   Begin VB.Label Label34 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "ST Fare Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   112
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   " Header1 Bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   111
      Top             =   6720
      Width           =   1230
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   " Crew Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   110
      Top             =   6720
      Width           =   1140
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Logo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   109
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Advertisement in Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9240
      TabIndex        =   108
      Top             =   7800
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Next Fair RoundOff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   107
      Top             =   5640
      Width           =   1710
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "SETUP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   106
      Top             =   -360
      Width           =   2775
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Amount in Refund Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   105
      Top             =   6000
      Width           =   2760
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   " Pass Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11520
      TabIndex        =   104
      Top             =   6000
      Width           =   930
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Language"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Language"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   102
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Luggage Rate"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   -120
      TabIndex        =   101
      Top             =   840
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Round UP/Down"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   240
      TabIndex        =   100
      Top             =   4500
      Width           =   1740
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "RoundOff"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   99
      Top             =   3750
      Width           =   1740
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Header3"
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
      Left            =   3690
      TabIndex        =   98
      Top             =   2340
      Width           =   1800
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Header2"
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
      Left            =   3690
      TabIndex        =   97
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Header1"
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
      Left            =   3690
      TabIndex        =   96
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RoundOff Amount (in Paise)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   3480
      TabIndex        =   95
      Top             =   360
      Width           =   1800
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Display 2"
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
      Left            =   3690
      TabIndex        =   94
      Top             =   975
      Width           =   1800
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Footer1"
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
      Left            =   3690
      TabIndex        =   93
      Top             =   2895
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Footer2"
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
      Left            =   3690
      TabIndex        =   92
      Top             =   3450
      Width           =   1635
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*  -  Mandatory field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Index           =   5
      Left            =   90
      TabIndex        =   90
      Top             =   135
      Width           =   2100
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   4
      Left            =   21990
      TabIndex        =   89
      Top             =   4305
      Width           =   165
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   3
      Left            =   1665
      TabIndex        =   88
      Top             =   2362
      Width           =   105
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   1860
      TabIndex        =   87
      Top             =   1342
      Width           =   165
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   1800
      TabIndex        =   86
      Top             =   802
      Width           =   165
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00000080&
      Height          =   165
      Left            =   3465
      TabIndex        =   85
      Top             =   3375
      Width           =   225
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   3465
      TabIndex        =   84
      Top             =   2820
      Width           =   210
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Master Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   81
      Top             =   1342
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Palmtec ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   80
      Top             =   2362
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Half Ratio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   79
      Top             =   2872
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ST Ratio "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   240
      TabIndex        =   78
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PH Ratio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   77
      Top             =   3322
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ST Max. Amount (Paise)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   76
      Top             =   6480
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ST Min. Amount (Paise)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   75
      Top             =   7200
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   74
      Top             =   802
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Display 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   240
      TabIndex        =   73
      Top             =   1792
      Width           =   1575
   End
   Begin VB.Label Currency 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Currency :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   20970
      TabIndex        =   91
      Top             =   4335
      Width           =   930
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public fname As String
Public BHandle As Integer
Dim strpath As String

Private Type Footer
    FooterString As String * 31
End Type

Private Sub chkReport_Change()
cmdSave.SetFocus
End Sub

Private Sub chkstenable_Click()
On Error Resume Next
    If chkstenable.Value = 1 Then
        txtSTMax.Visible = False
        txtST.Visible = False
        txtSTMin.Visible = False
        Label3.Visible = False
        Label5.Visible = False
        Label7.Visible = False
    Else
        txtSTMax.Visible = True
        txtST.Visible = True
        txtSTMin.Visible = True
        Label3.Visible = True
        Label5.Visible = True
        Label7.Visible = True
    End If
End Sub

Private Sub chkSTroundEnable_Click()
 If chkSTroundEnable.Value = 0 Then
        txtSTRoundAmt.Visible = False
        Label64.Visible = False
        
    Else
        txtSTRoundAmt.Visible = True
        Label64.Visible = True
    End If
End Sub

Private Sub cmb_font_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    lad_ratio_txt.SetFocus
End If
End Sub

Private Sub cmbLocalLanguage_Change()
On Error Resume Next
    cmb_font.SetFocus
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    Unload Me
End Sub



Private Sub cmdCurrency_Change()
On Error Resume Next
    chkPrint.SetFocus
End Sub

Private Sub cmdLoadF_Click()
''//load settings from settings file
Dim ffname As String
On Error Resume Next
    CommonDialog1.DialogTitle = "Open Settings"
    CommonDialog1.Filter = "*.PSTP|*.pstp"
    CommonDialog1.CancelError = False
    CommonDialog1.ShowOpen
    ffname = CommonDialog1.filename
 If ffname <> "" Then
    printflag = 0
    If cmbLocalLanguage.ListCount > 0 Then cmbLocalLanguage.ListIndex = 0
    GetCurrency
    'ffname = App.Path & "\BUS.DAT"
    BHandle = FreeFile()
    If Dir(ffname) <> "" Then
        Open ffname For Binary Access Read As #BHandle
        Get #BHandle, , HStr
        Get #BHandle, , hardwaresettings
        Close #BHandle
        LoadValues
    End If
    If optRoundoff(0).Value = True Then
        Frame2.Enabled = False
        optRoundUD(0).Enabled = False
        optRoundUD(1).Enabled = False
        txtRoundAmt.Enabled = False
        txtRoundAmt.BackColor = &H80000004
    End If
    cmbLocalLanguage.Text = cmbLocalLanguage.List(LocalLanguage)
    
 Else
 MsgBox "No File Selected"
 End If
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    
'''''''''''''''''modification syam
'''''''''''new fields added in SETUP 13/3/09
    
    
printflag = 0
  
If chkReport.Value = 1 Then
printflag = printflag + 1
End If
If chkPrint.Value = 1 Then
printflag = printflag + 2
End If
''MsgBox printflag
    
    '''''''''''''''''''''''''''''''''''
    If Trim(txtUserPWD) = "" Then
        MsgBox "Please enter the User password", vbInformation, gblstrPrjTitle
        txtUserPWD.SetFocus
        Exit Sub
    End If
    If Trim(txtmstPWD) = "" Then
        MsgBox "please enter the master password", vbInformation, gblstrPrjTitle
        txtmstPWD.SetFocus
        Exit Sub
    End If
   
     If Trim(txtMain) = "" Then
        MsgBox "please enter the Main Display 1", vbInformation, gblstrPrjTitle
        txtMain.SetFocus
        Exit Sub
    End If
    If Trim(txtID) = "" Then
        MsgBox "Please enter the palmtec ID", vbInformation, gblstrPrjTitle
        txtID.SetFocus
        Exit Sub
    End If
    If cmdCurrency.Text = "" Then
        MsgBox "Select Currency", vbInformation, gblstrPrjTitle
        cmdCurrency.SetFocus
        Exit Sub
    End If
    If val(txtSTMax) <= val(txtSTMin) Then
        MsgBox "ST Max Amount must be greater than ST Min Amount!", vbInformation, gblstrPrjTitle
        txtSTMax.SetFocus
        Exit Sub
    End If
    SaveValues
    BHandle = FreeFile()
    If Dir(fname) <> "" Then Kill App.Path & "\" & "BUS.DAT "
    Open fname For Binary Access Write As #BHandle
        Put #BHandle, , HStr
        Put #BHandle, , hardwaresettings
    Close #BHandle
    DBSet ("BUS.DAT")
    MsgBox "Settings successfully saved"
    Unload Me
End Sub

Private Sub cmdSaveF_Click()
Dim ffname As String
On Error Resume Next

    CommonDialog1.DialogTitle = "Save Settings"
    CommonDialog1.Filter = "*.PSTP|*.pstp"
    CommonDialog1.CancelError = False
    CommonDialog1.ShowSave
    ffname = CommonDialog1.filename
    If ffname <> "" Then
        printflag = 0
        
        If chkReport.Value = 1 Then
        printflag = printflag + 1
        End If
        
        If chkPrint.Value = 1 Then
        printflag = printflag + 2
        End If
        
        If Trim(txtmstPWD) = "" Then
            MsgBox "please enter the master password", vbInformation, gblstrPrjTitle
            txtmstPWD.SetFocus
            Exit Sub
        End If
        
        If Trim(txtUserPWD) = "" Then
            MsgBox "Please enter the User password", vbInformation, gblstrPrjTitle
            txtUserPWD.SetFocus
            Exit Sub
        End If
        If Trim(txtID) = "" Then
            MsgBox "Please enter the palmtec ID", vbInformation, gblstrPrjTitle
            txtID.SetFocus
            Exit Sub
        End If
        
        If cmdCurrency.Text = "" Then
            MsgBox "Select Currency", vbInformation, gblstrPrjTitle
            cmdCurrency.SetFocus
            Exit Sub
        End If
        
        SaveValues
        
        BHandle = FreeFile()
        If Dir(ffname) <> "" Then Kill ffname
        Open ffname For Binary Access Write As #BHandle
            Put #BHandle, , HStr
            Put #BHandle, , hardwaresettings
        Close #BHandle
        'DBSet ("BUS.DAT")
        'Dim REP As Integer
        'REP=MsgBox "Settings successfully saved.Do You want to Save This settings as the Default Settings", vbYesNo,""
        Dim iResponce As Integer
        iResponce = MsgBox("Settings successfully saved.Do You want to save this as the Default Settings", vbYesNo + vbQuestion, "Settings")
        If iResponce = vbYes Then ' They Clicked YES!
            cmdSave.Enabled = True
            cmdSave_Click
        End If
    Else
         MsgBox "Enter a valid File Name"
    End If
End Sub

Private Sub cmdViewExp_Click()
Dim expense As String
On Error Resume Next
    cmdViewExp.Enabled = False
    frmExpense.Show vbModal
 
End Sub

Sub GetCurrency()
On Error GoTo err
Dim DB As DAO.Database
Dim RES As DAO.Recordset

Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
Set RES = DB.OpenRecordset("CURRENCY", dbOpenDynaset)
If RES.RecordCount > 0 Then
    RES.MoveFirst
    Do While Not RES.EOF
        cmdCurrency.AddItem RES.Fields(1)
        RES.MoveNext
    Loop
End If
RES.Close
DB.Close
Exit Sub
err:
    MsgBox "Error in opening DB" & vbCrLf & err.Number & " , " & err.Description
    Exit Sub
End Sub






Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
'SINHALA
'HINDI
    Me.Icon = frmMainform.Icon
    
    '''''''''''''''''modification syam
    '''''''''''new field added in SETUP 13/3/09
'
printflag = 0
    If cmbLocalLanguage.ListCount > 0 Then cmbLocalLanguage.ListIndex = 0

    ''''''''''''''''''''''
    cmb_font.Clear
    cmb_font.AddItem "Normal"
    cmb_font.ListIndex = 0
    cmb_font.AddItem "Condensed"
    cmb_font.ListIndex = 1
    GetCurrency
    fname = App.Path & "\BUS.DAT"
    BHandle = FreeFile()
    If Dir(fname) <> "" Then
        Open fname For Binary Access Read As #BHandle
        Get #BHandle, , HStr
        Get #BHandle, , hardwaresettings
        Close #BHandle
        LoadValues
    End If
    If optRoundoff(0).Value = True Then
        Frame2.Enabled = False
        optRoundUD(0).Enabled = False
        optRoundUD(1).Enabled = False
        txtRoundAmt.Enabled = False
        txtRoundAmt.BackColor = &H80000004
    End If
    cmbLocalLanguage.Text = cmbLocalLanguage.List(LocalLanguage)
End Sub



Private Sub lad_ratio_txt_KeyPress(KeyAscii As Integer)
On Error Resume Next
    ValidationMode = FloatingPointValue
    ValidateKeyPress lad_ratio_txt, KeyAscii, False, 2, 2
End Sub

Private Sub optRoundoff_Click(Index As Integer)
On Error Resume Next
    If optRoundoff(0).Value = True Then
        Frame2.Enabled = False
        optRoundUD(0).Enabled = False
        optRoundUD(1).Enabled = False
        txtRoundAmt.Enabled = False
        txtRoundAmt.BackColor = &H80000004
        chkRnFair.Enabled = True
    ElseIf optRoundoff(1).Value = True Then
        Frame2.Enabled = True
        optRoundUD(0).Enabled = True
        optRoundUD(1).Enabled = True
        txtRoundAmt.Enabled = True
        txtRoundAmt.BackColor = &H80000005
        chkRnFair.Value = vbUnchecked
        chkRnFair.Enabled = False
    End If

End Sub

Private Sub optRoundoff_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 And optRoundoff(1).Value = True Then optRoundUD(0).SetFocus
    If KeyAscii = 13 And optRoundoff(0).Value = True Then txtMainDisp2.SetFocus 'txtMain.SetFocus
    If optRoundoff(0).Value = True Then
        Frame2.Enabled = False
        optRoundUD(0).Enabled = False
        optRoundUD(1).Enabled = False
        txtRoundAmt.Enabled = False
    ElseIf optRoundoff(1).Value = True Then
        Frame2.Enabled = True
        optRoundUD(0).Enabled = True
        optRoundUD(1).Enabled = True
        txtRoundAmt.Enabled = True
    End If
    
End Sub


Private Sub optRoundUD_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then txtRoundAmt.SetFocus
End Sub

Private Sub Text1_Change()

End Sub

Private Sub sc_ratio_txt_KeyPress(KeyAscii As Integer)
On Error Resume Next
    ValidationMode = FloatingPointValue
    ValidateKeyPress sc_ratio_txt, KeyAscii, False, 2, 2
End Sub

Private Sub txtaccesspoint_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtaccesspoint <> "" Then
txtftpurl.SetFocus
Else
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = KeyAscii
    End Select
End If
End Sub

Private Sub txtdwnldpath_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtdwnldpath <> "" Then
txthttpurl.SetFocus
Else
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
           KeyAscii = KeyAscii
    End Select
End If
End Sub

Private Sub txtftppswd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtftppswd <> "" Then
txtuploadpath.SetFocus
Else
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = KeyAscii
    End Select
End If
End Sub

Private Sub txtftpuname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtftpuname <> "" Then
txtftppswd.SetFocus
Else
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
           KeyAscii = KeyAscii
    End Select
End If
End Sub

Private Sub txtftpurl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtftpurl <> "" Then
txtftpuname.SetFocus
Else
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
KeyAscii = KeyAscii
End Select
End If
End Sub

Private Sub txtFTR1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 And txtFTR1 <> "" Then
        If txtFTR2.Enabled = False Then
            cmdCurrency.SetFocus
        Else
            txtFTR2.SetFocus
        End If
        Exit Sub
    End If
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub


Private Sub txtFTR2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 And txtFTR2 <> "" Then cmdCurrency.SetFocus   'txtLuggage.SetFocus
   Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub txtHalfRatio_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 And txtHalfRatio <> "" Then
    txtPH.SetFocus  'txtST.SetFocus
    ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
    ElseIf KeyAscii = 43 Then
    ElseIf KeyAscii = 8 Then
    Else
    KeyAscii = 0
    End If
    Select Case KeyAscii
        Case 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select

End Sub

Private Sub txtHDR1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 And txtHDR1 <> "" Then txtHDR2.SetFocus
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub txtHDR2_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then txtHDR3.SetFocus
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub



Private Sub txtHDR3_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then txtFTR1.SetFocus
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub


Private Sub txthttpurl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txthttpurl <> "" Then
txtphno.SetFocus
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
           KeyAscii = KeyAscii
    End Select
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If TextBoxPalmValidity(KeyAscii) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 And txtID <> "" Then txtHalfRatio.SetFocus
End Sub

Private Sub txtLuggage_KeyPress(KeyAscii As Integer)
On Error Resume Next
Static LastText As String
  
Static SecondTime As Boolean
  
Const MaxDecimal As Integer = 1
  
Const MaxWhole As Integer = 4
  
    
    With txtLuggage
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
    If KeyAscii = 13 Then cmbLocalLanguage.SetFocus

End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 And txtMain <> "" Then txtID.SetFocus
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub
Private Sub txtMainDisp2_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 And txtMain <> "" Then txtHDR1.SetFocus
    Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub txtmstPWD_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If TextBoxValidity(KeyAscii) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 And txtmstPWD.Text <> "" Then txtMain.SetFocus
End Sub
Private Sub txtPH_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If TextBoxValidity(KeyAscii) > 0 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 And txtPH <> "" Then
    optRoundoff(0).SetFocus  'txtSTMax.SetFocus
    ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
    ElseIf KeyAscii = 43 Then
    ElseIf KeyAscii = 8 Then
    Else
    KeyAscii = 0
    End If

    
End Sub

Private Sub txtphno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtST <> "" Then
txtsmsph2.SetFocus
ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
    ElseIf KeyAscii = 43 Then
    ElseIf KeyAscii = 8 Then
    Else
    KeyAscii = 0
    End If
End Sub


Private Sub txtRoundAmt_KeyPress(KeyAscii As Integer)
    ValidationMode = FloatingPointValue
    ValidateKeyPress txtRoundAmt, KeyAscii, False
End Sub

Private Sub txtsmsph2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtsmsph2 <> "" Then
txtsmsph3.SetFocus
ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
    ElseIf KeyAscii = 43 Then
    ElseIf KeyAscii = 8 Then
    Else
    KeyAscii = 0
    End If
End Sub

Private Sub txtsmsph3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtsmsph3 <> "" Then
    cmdSave.SetFocus
ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 43 Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If

End Sub

'Private Sub txtRoundAmt_KeyPress(KeyAscii As Integer)
'On Error Resume Next
 '   If TextBoxValidity(KeyAscii) > 0 Then
  '      KeyAscii = 0
   ' End If
    'If KeyAscii = 13 And txtRoundAmt <> "" Then txtMainDisp2.SetFocus
'End Sub

Private Sub txtST_KeyPress(KeyAscii As Integer)
On Error Resume Next
    ValidationMode = FloatingPointValue
    ValidateKeyPress txtST, KeyAscii, False
    If KeyAscii = 13 And txtST <> "" Then txtPH.SetFocus
End Sub


Private Sub txtSTMax_KeyPress(KeyAscii As Integer)
On Error Resume Next
    ValidationMode = FloatingPointValue
    ValidateKeyPress txtSTMax, KeyAscii, False
    If KeyAscii = 13 And txtSTMax <> "" Then txtSTMin.SetFocus
End Sub

Private Sub txtSTMin_KeyPress(KeyAscii As Integer)
On Error Resume Next
    ValidationMode = FloatingPointValue
    ValidateKeyPress txtSTMin, KeyAscii, False
    If KeyAscii = 13 And txtSTMin <> "" Then optRoundoff(0).SetFocus
End Sub
Private Sub txtSTRoundAmt_KeyPress(KeyAscii As Integer)
    ValidationMode = FloatingPointValue
    ValidateKeyPress txtSTRoundAmt, KeyAscii, False
End Sub

Private Sub txtuploadpath_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtuploadpath <> "" Then
txtdwnldpath.SetFocus
Else
Select Case KeyAscii
        Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
          KeyAscii = KeyAscii
    End Select
End If
End Sub

Private Sub txtUserPWD_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If TextBoxValidity(KeyAscii) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 And txtUserPWD <> "" Then txtmstPWD.SetFocus
End Sub


Private Sub txtYear_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If TextBoxValidity(KeyAscii) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 And txtYear <> "" Then txtID.SetFocus
End Sub


Public Function LoadValues()
On Error Resume Next
Dim iHandle As Integer
Dim temp As Footer
Dim bytPrev As Byte, intLoopCount As Integer
    With HStr
        txtID = .PalmtecID
        txtHalfRatio = .HalfPer
        txtST = val(.ConPer)
        txtPH = .PhyPer
        txtSTMax = val(.STMaxAmt)
        txtSTMin = val(.STMinCon)
        chkstenable.Value = .ucbSTFareEdit
        If chkstenable.Value = 0 Then
            txtSTMax.Visible = True
            txtST.Visible = True
            txtSTMin.Visible = True
            Label3.Visible = True
            Label5.Visible = True
            Label7.Visible = True
        End If
        If .Roundoff = 1 Then optRoundoff(1).Value = True
        If .Roundoff = 0 Then optRoundoff(0).Value = True
        If .RoundUp = 1 Then optRoundUD(1).Value = True
        If .RoundUp = 0 Then optRoundUD(0).Value = True
        txtRoundAmt = .RoundAmt
        txtMain = .MainDisp
        txtMainDisp2 = .MainDisp2
        txtHDR1 = .bhl1
        txtHDR2 = .bhl2
        txtHDR3 = .bhl3
        txtFTR1 = .bfl1
        txtphno = .PhNo
'        If CheckPatchVerified = False Then
'            Dim strpath As String
'            strpath = checkfooter
'            If strpath <> "" Then
'                iHandle = FreeFile()
'                Open strpath For Binary Access Read As #iHandle
'                Get #iHandle, , temp
'                Close #iHandle
'                txtFTR2.Text = TrimChr(temp.FooterString)
'            Else
'                txtFTR2 = "Softland India Ltd"
'            End If
'            txtFTR2.Enabled = False
'        Else
'            txtFTR2.Enabled = True
'        End If
strFooter = LoadFooter
        If strFooter = "" Then
            txtFTR2 = "Softland India Ltd"
            'txtFTR2 = "SOFTLAND-J.T.M"
        Else
            txtFTR2.Text = strFooter
        End If

        txtFTR2.Enabled = False
        txtLuggage = .LuggageUnitRate
        cmdCurrency.Text = .Currency
        
        chkReport.Value = 0: chkPrint.Value = 0: chkAdvertise = 0
        
        bytPrev = 1
        If (.ReportFlag And bytPrev) = bytPrev Then
            chkReport.Value = 1
        End If
        For intLoopCount = 1 To 7
            bytPrev = (bytPrev * 2)
            Select Case bytPrev
                Case 2:
                    If (.ReportFlag And bytPrev) = bytPrev Then chkPrint.Value = 1
                Case 4:
                    If (.ReportFlag And bytPrev) = bytPrev Then chkAdvertise.Value = 1
                Case 8:
                    If (.ReportFlag And bytPrev) = bytPrev Then chklogenable.Value = 1
                Case 16:
                    If (.ReportFlag And bytPrev) = bytPrev Then chkcrewenable.Value = 1
                Case 32:
                    If (.ReportFlag And bytPrev) = bytPrev Then chkbigfontenable.Value = 1
                Case 64:
                    
            End Select
        Next
        chkDefaultStage.Value = .EnableStageDefault
        chkRnFair.Value = .NextFareRound '06-06-09
        If optRoundoff(1).Value = True Then
            chkRnFair.Enabled = False
        Else
            chkRnFair.Enabled = True
        End If
        
        chkRemoveTicketEnable.Value = .EnableRemoveTicket
        chkBoldFontEnable.Value = .EnableStageFont
        chkStageUpdation.Value = .UpdateStageMsg
        chkOdmtrEnable.Value = .OdometerEntry
        chkCrewcheckEnable.Value = .CrewCheck
        chkTcktBold.Value = .TicketNoBigFont
        chkTRIP.Value = .TripSMS
        chkSh.Value = .ScheduleSMS
        chksendbillenable.Value = .sendbillEnable
        chktripsendenable.Value = .TripsendEnable
        chkschedulesendenable.Value = .SchedulesendEnable
        chksendpendinbill.Value = .Sendpend
       
        txtsmsph2.Text = TrimChr(.PhNo2)
        txtaccesspoint.Text = TrimChr(.AccessPoint)
        txtftpurl.Text = TrimChr(.DestAdds)
        txtftpuname.Text = TrimChr(.Username)
        txtftppswd.Text = TrimChr(.PassWord)
        txtuploadpath.Text = TrimChr(.Uploadpath)
        txtdwnldpath.Text = TrimChr(.Downloadpath)
        txthttpurl.Text = TrimChr(.HttpUrl)
        chkgprs.Value = .GprsEnable
        chksmart.Value = .SmartCard
        chkexp.Value = .ExpEnable
        chkgprsmsgenable.Value = .MsgPrompt
        chkftp.Value = .FtpEnable
        Chkinspect.Value = .Inspectorreport
        txtSTRoundAmt = TrimChr(.StRoundoff_Amt)
        chkSTroundEnable.Value = .StRoundoff_E_D
        cmb_font.ListIndex = .ReportFONT
        chk_multiple.Value = .MultiplePass
        chk_simple.Value = .Simplereport
        chk_inspector.Value = .InspectorSMS
        txtsmsph3.Text = TrimChr(.PhNo3)
        If chkSTroundEnable.Value = 0 Then
            txtSTRoundAmt.Visible = False
            Label64.Visible = False
        Else
            txtSTRoundAmt.Visible = True
            Label64.Visible = True
        End If
        chk_AutoShutdown.Value = .AutoShutdownEnable
        chk_userpswd.Value = .UserPasswordEnable
        chk_refund.Value = .refund
        lad_ratio_txt = val(.ladis_per)
        sc_ratio_txt = val(.seniar_per)
    End With
    With hardwaresettings
       txtmstPWD = .MSR_PSWD
       txtUserPWD = .USR_PSWD
    End With
    
End Function
Public Function SaveValues()
On Error Resume Next
Dim bytPrev As Byte, intLoopCount As Integer
    With HStr
        .PalmtecID = Trim(txtID.Text) & Chr(0)
        .HalfPer = val(txtHalfRatio)
        .PhyPer = val(txtPH)
        If chkstenable.Value = 1 Then
            .ConPer = "25"
            .STMaxAmt = "23.50"
            .STMinCon = "2.50"
        Else
            .ConPer = val(txtST)
            .STMaxAmt = val(txtSTMax)
            .STMinCon = val(txtSTMin)
        End If
        .ucbSTFareEdit = chkstenable.Value
        If optRoundoff(1).Value = True Then .Roundoff = 1
        If optRoundoff(0).Value = True Then .Roundoff = 0
        If optRoundUD(1).Value = True And optRoundoff(1).Value = True Then .RoundUp = 1
        If optRoundUD(0).Value = True Then .RoundUp = 0
        
        .RoundAmt = val(txtRoundAmt.Text)
        .MainDisp = Trim(txtMain.Text) & Chr(0)
        .MainDisp2 = Trim(txtMainDisp2.Text) & Chr(0)
        .bhl1 = Trim(txtHDR1.Text) & Chr(0)
        .bhl2 = Trim(txtHDR2.Text) & Chr(0)
        .bhl3 = Trim(txtHDR3.Text) & Chr(0)
        .bfl1 = Trim(txtFTR1.Text) & Chr(0)
        .bfl2 = Trim(txtFTR2.Text) & Chr(0)
        .PaperFeed = &H2
        .DefaultFull = 1
        .LuggageUnitRateEdit = &H1
        .LuggageUnitRate = val(txtLuggage)
        .StageUpdation = 0
        .StageDisplayFont = 1
        .UseDuplicate = 1
        .UseDup1 = 0
        .Currency = Trim(Mid(cmdCurrency.Text, 1, 7)) & Chr(0)
        .PhNo = Trim(txtphno.Text) & Chr(0)
        .ReportFlag = 0
        .UpdateStageMsg = chkStageUpdation.Value
        bytPrev = 1
        If chkReport.Value = 1 Then .ReportFlag = 1
        
        For intLoopCount = 1 To 7
            bytPrev = (bytPrev * 2)
            Select Case bytPrev
                Case 2:
                    If chkPrint.Value = 1 Then .ReportFlag = (bytPrev Or .ReportFlag)
                Case 4:
                    If chkAdvertise.Value = 1 Then .ReportFlag = (bytPrev Or .ReportFlag)
                Case 8:
                    If chklogenable.Value = 1 Then .ReportFlag = (bytPrev Or .ReportFlag)
                Case 16:
                    If chkcrewenable.Value = 1 Then .ReportFlag = (bytPrev Or .ReportFlag)
                Case 32:
                    If chkbigfontenable.Value = 1 Then .ReportFlag = (bytPrev Or .ReportFlag)
                Case 64:
                    
            End Select
        Next
        
        .NextFareRound = chkRnFair.Value '06-06-09
        .EnableRemoveTicket = chkRemoveTicketEnable.Value
        .EnableStageFont = chkBoldFontEnable.Value
        .EnableStageDefault = chkDefaultStage.Value
        .OdometerEntry = chkOdmtrEnable.Value
        .CrewCheck = chkCrewcheckEnable.Value
        .ScheduleSMS = chkSh.Value
        .TripSMS = chkTRIP.Value
        .TicketNoBigFont = chkTcktBold.Value
        '''''''''''''''''''''''''''''''''''''''''''''''
        .sendbillEnable = chksendbillenable
        .TripsendEnable = chktripsendenable.Value
        .SchedulesendEnable = chkschedulesendenable.Value
        .Sendpend = chksendpendinbill.Value
        .MsgPrompt = chkgprsmsgenable.Value
        .PhNo2 = txtsmsph2.Text & Chr(0)
        .AccessPoint = txtaccesspoint.Text & Chr(0)
        .DestAdds = txtftpurl.Text & Chr(0)
        .Username = txtftpuname.Text & Chr(0)
        .PassWord = txtftppswd.Text & Chr(0)
        .Uploadpath = txtuploadpath.Text & Chr(0)
        .Downloadpath = txtdwnldpath.Text & Chr(0)
        .HttpUrl = txthttpurl.Text & Chr(0)
        .GprsEnable = chkgprs.Value
        .SmartCard = chksmart.Value
        .ExpEnable = chkexp.Value
        .FtpEnable = chkftp.Value
        .Inspectorreport = Chkinspect.Value
        .StRoundoff_Amt = val(txtSTRoundAmt)
        .StRoundoff_E_D = chkSTroundEnable.Value
        .Simplereport = chk_simple.Value
        .MultiplePass = chk_multiple.Value
        .InspectorSMS = chk_inspector.Value
        .ReportFONT = cmb_font.ListIndex
        .PhNo3 = txtsmsph3.Text & Chr(0)
        .AutoShutdownEnable = chk_AutoShutdown.Value
        .UserPasswordEnable = chk_userpswd.Value
        .refund = chk_refund.Value
        .ladis_per = val(IIf(IsNumeric(lad_ratio_txt.Text), lad_ratio_txt.Text, 0))
        .seniar_per = val(IIf(IsNumeric(sc_ratio_txt.Text), sc_ratio_txt.Text, 0))
    ''''''''''''''''''''''''''''''''''''''''''''''''''
        
        .ucTemp = Chr(0)
        
    End With
    
    With hardwaresettings
        .Ptime.ti_hour = Hour(Time)
        .Ptime.ti_min = Minute(Time)
        .Ptime.ti_sec = Second(Time)
        .Ptime.ti_hund = 0
        .Pdate.da_day = Day(Now)
        .Pdate.da_mon = Month(Now)
        .Pdate.da_year = Year(Now)
        
        .MSR_PSWD = Trim(txtmstPWD) & Chr(0)
        .USR_PSWD = Trim(txtUserPWD) & Chr(0)
        .SPR_PSWD = "10SIL9" & Chr(0)
        .val_contrast = &H4
        .val_brightness = &HA
        .screensaver_onoff = &H0
        .backlit_timer = &H3
        .keyhitdelay = &H8
        .boarder_en = &H0
        .dooropen_alert = &H1
        .paperout_alert = &H1
        .ucHalfPagePrinter = &H0
        .buzz_onoff = &H1
        .rs232_baud = &H1
        .ir_baud = &H1
        .rf_baud = &H1
        .connecting_medium = &H1
        .footer_stat = &H10
        .select_language = cmbLocalLanguage.ListIndex + 1 '&H1
        .login_mode = &H0
        .ucKPLight_opt = &H0
        .LangNo = cmbLocalLanguage.ListIndex + 1 '&H1
        .ucTemp = 1
    End With
    SetLocalLanguage (cmbLocalLanguage.ListIndex)
    
End Function

Private Function CheckPatchVerified() As Boolean
Dim strpath As String
Dim strBuffer As String * 250
On Error Resume Next
    Call GetSystemDirectory(strBuffer, Len(strBuffer))
    strBuffer = Mid(strBuffer, 1, InStr(1, strBuffer, Chr(0)) - 1)
    strpath = Trim(strBuffer)
    If Dir(strpath & "\Ftr.sys", vbHidden + vbSystem) <> "" Then
        CheckPatchVerified = True
    End If
End Function


