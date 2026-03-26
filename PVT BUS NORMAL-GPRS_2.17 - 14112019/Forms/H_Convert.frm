VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form H_Convert 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Sinhala Printer Pro                             Version : 1.0.3.D     DT.18.06.08"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   Icon            =   "H_Convert.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "H_Convert.frx":57E2
   ScaleHeight     =   7485
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text7 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   -30
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8385
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6750
      Width           =   1410
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6750
      Width           =   1410
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8385
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6210
      Width           =   1410
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear &All "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6210
      Width           =   1410
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7665
      Top             =   825
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Load"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9975
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5685
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5685
      Width           =   1410
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8415
      TabIndex        =   9
      Top             =   5265
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8415
      TabIndex        =   8
      Top             =   4710
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   8415
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4170
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   8415
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3600
      Width           =   3015
   End
   Begin VB.ComboBox Cref 
      Height          =   315
      Left            =   11610
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   1620
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton CShow 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show"
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
      Left            =   12165
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1455
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Next >>>"
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CPrev 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<< Previous"
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
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1455
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   8415
      MaxLength       =   24
      TabIndex        =   0
      Top             =   390
      Width           =   3015
   End
   Begin VB.CommandButton CSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
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
      Left            =   12060
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1395
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDA782&
      Caption         =   "Sinhala Fonts"
      Height          =   6300
      Left            =   150
      TabIndex        =   14
      Top             =   915
      Width           =   7695
      Begin MSComDlg.CommonDialog CD 
         Left            =   6825
         Top             =   4635
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Space"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3075
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   6015
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1710
         Left            =   6890
         ScaleHeight     =   1680
         ScaleWidth      =   705
         TabIndex        =   27
         Top             =   1250
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   6885
         ScaleHeight     =   795
         ScaleWidth      =   705
         TabIndex        =   35
         Top             =   3105
         Width           =   735
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   73
         Left            =   5040
         Picture         =   "H_Convert.frx":1C98C
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   72
         Left            =   4440
         Picture         =   "H_Convert.frx":1D4C6
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   71
         Left            =   3840
         Picture         =   "H_Convert.frx":1E000
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   70
         Left            =   3240
         Picture         =   "H_Convert.frx":1EB3A
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   69
         Left            =   2640
         Picture         =   "H_Convert.frx":1F674
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   68
         Left            =   2040
         Picture         =   "H_Convert.frx":201AE
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   67
         Left            =   1440
         Picture         =   "H_Convert.frx":20CE8
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   66
         Left            =   840
         Picture         =   "H_Convert.frx":21822
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   65
         Left            =   6240
         Picture         =   "H_Convert.frx":2235C
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   64
         Left            =   5640
         Picture         =   "H_Convert.frx":22E96
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   63
         Left            =   5040
         Picture         =   "H_Convert.frx":239D0
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   62
         Left            =   4440
         Picture         =   "H_Convert.frx":2450A
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   61
         Left            =   3840
         Picture         =   "H_Convert.frx":25044
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   60
         Left            =   3240
         Picture         =   "H_Convert.frx":25B7E
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   59
         Left            =   2640
         Picture         =   "H_Convert.frx":266B8
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   58
         Left            =   2040
         Picture         =   "H_Convert.frx":271F2
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   57
         Left            =   1440
         Picture         =   "H_Convert.frx":27D2C
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   56
         Left            =   840
         Picture         =   "H_Convert.frx":28866
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   55
         Left            =   240
         Picture         =   "H_Convert.frx":293A0
         Stretch         =   -1  'True
         Top             =   3975
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   7320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line6 
         X1              =   7320
         X2              =   7320
         Y1              =   720
         Y2              =   1080
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   120
         Y1              =   720
         Y2              =   1080
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   120
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line3 
         X1              =   7320
         X2              =   7320
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   7320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7320
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   2
         Left            =   840
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   1
         Left            =   480
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   3
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   20
         Left            =   120
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   39
         Left            =   6960
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   33
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   38
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   37
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   36
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   35
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   34
         Left            =   5160
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   32
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   31
         Left            =   4080
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   30
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   29
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   28
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   27
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   26
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   25
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   24
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   23
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   22
         Left            =   840
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   21
         Left            =   480
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   0
         Left            =   240
         Picture         =   "H_Convert.frx":29EDA
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   1
         Left            =   840
         Picture         =   "H_Convert.frx":2AA14
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   2
         Left            =   1440
         Picture         =   "H_Convert.frx":2B54E
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   3
         Left            =   2040
         Picture         =   "H_Convert.frx":2C088
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   4
         Left            =   2640
         Picture         =   "H_Convert.frx":2CBC2
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   5
         Left            =   3240
         Picture         =   "H_Convert.frx":2D6FC
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   6
         Left            =   3840
         Picture         =   "H_Convert.frx":2E236
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   7
         Left            =   4440
         Picture         =   "H_Convert.frx":2ED70
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   8
         Left            =   5040
         Picture         =   "H_Convert.frx":2F8AA
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   9
         Left            =   5640
         Picture         =   "H_Convert.frx":303E4
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   10
         Left            =   6240
         Picture         =   "H_Convert.frx":30F1E
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   11
         Left            =   240
         Picture         =   "H_Convert.frx":31A58
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   12
         Left            =   840
         Picture         =   "H_Convert.frx":32592
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   21
         Left            =   6240
         Picture         =   "H_Convert.frx":330CC
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   20
         Left            =   5640
         Picture         =   "H_Convert.frx":33C06
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   19
         Left            =   5040
         Picture         =   "H_Convert.frx":34740
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   18
         Left            =   4440
         Picture         =   "H_Convert.frx":3527A
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   17
         Left            =   3840
         Picture         =   "H_Convert.frx":35DB4
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   16
         Left            =   3240
         Picture         =   "H_Convert.frx":368EE
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   15
         Left            =   2640
         Picture         =   "H_Convert.frx":37428
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   14
         Left            =   2040
         Picture         =   "H_Convert.frx":37F62
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   13
         Left            =   1440
         Picture         =   "H_Convert.frx":38A9C
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   34
         Left            =   840
         Picture         =   "H_Convert.frx":395D6
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   33
         Left            =   240
         Picture         =   "H_Convert.frx":3A110
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   32
         Left            =   6240
         Picture         =   "H_Convert.frx":3AC4A
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   31
         Left            =   5640
         Picture         =   "H_Convert.frx":3B784
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   30
         Left            =   5040
         Picture         =   "H_Convert.frx":3C2BE
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   29
         Left            =   4440
         Picture         =   "H_Convert.frx":3CDF8
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   28
         Left            =   3840
         Picture         =   "H_Convert.frx":3D932
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   27
         Left            =   3240
         Picture         =   "H_Convert.frx":3E46C
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   26
         Left            =   2640
         Picture         =   "H_Convert.frx":3EFA6
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   25
         Left            =   2040
         Picture         =   "H_Convert.frx":3FAE0
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   24
         Left            =   1440
         Picture         =   "H_Convert.frx":4061A
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   23
         Left            =   840
         Picture         =   "H_Convert.frx":41154
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   22
         Left            =   240
         Picture         =   "H_Convert.frx":41C8E
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   43
         Left            =   6255
         Picture         =   "H_Convert.frx":427C8
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   42
         Left            =   5640
         Picture         =   "H_Convert.frx":43302
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   41
         Left            =   5040
         Picture         =   "H_Convert.frx":43E3C
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   40
         Left            =   4440
         Picture         =   "H_Convert.frx":44976
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   39
         Left            =   3840
         Picture         =   "H_Convert.frx":454B0
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   38
         Left            =   3240
         Picture         =   "H_Convert.frx":45FEA
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   37
         Left            =   2640
         Picture         =   "H_Convert.frx":46B24
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   36
         Left            =   2040
         Picture         =   "H_Convert.frx":4765E
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   35
         Left            =   1440
         Picture         =   "H_Convert.frx":48198
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   54
         Left            =   6240
         Picture         =   "H_Convert.frx":48CD2
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   53
         Left            =   5640
         Picture         =   "H_Convert.frx":4980C
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   52
         Left            =   5040
         Picture         =   "H_Convert.frx":4A346
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   51
         Left            =   4440
         Picture         =   "H_Convert.frx":4AE80
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   50
         Left            =   3840
         Picture         =   "H_Convert.frx":4B9BA
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   49
         Left            =   3240
         Picture         =   "H_Convert.frx":4C4F4
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   48
         Left            =   2640
         Picture         =   "H_Convert.frx":4D02E
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   47
         Left            =   2040
         Picture         =   "H_Convert.frx":4DB68
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   46
         Left            =   1440
         Picture         =   "H_Convert.frx":4E6A2
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   45
         Left            =   840
         Picture         =   "H_Convert.frx":4F1DC
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   44
         Left            =   240
         Picture         =   "H_Convert.frx":4FD16
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   4
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   5
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   6
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   7
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   8
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   9
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   10
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   11
         Left            =   4080
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   12
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   13
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   14
         Left            =   5160
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   15
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   16
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   17
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   18
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   19
         Left            =   6960
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   7320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6360
         Picture         =   "H_Convert.frx":50850
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   435
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add Hindi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   -825
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3165
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add Hindi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   -645
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2610
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   8415
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3045
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   8415
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   8415
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1950
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   8415
      MaxLength       =   24
      TabIndex        =   2
      Top             =   1425
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   8415
      MaxLength       =   24
      TabIndex        =   1
      Top             =   930
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 1 [ English ]"
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
      Left            =   8415
      TabIndex        =   33
      Top             =   135
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 4 [ Sinhala ]"
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
      Left            =   8415
      TabIndex        =   32
      Top             =   1710
      Width           =   3015
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 8 [ Sinhala ]"
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
      Left            =   8415
      TabIndex        =   31
      Top             =   3915
      Width           =   3015
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 10 [ English ]"
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
      Left            =   8415
      TabIndex        =   30
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 9 [ English ]"
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
      Left            =   8415
      TabIndex        =   29
      Top             =   4470
      Width           =   3015
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 7 [ Sinhala ]"
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
      Left            =   8415
      TabIndex        =   28
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Consumer Reference Number"
      Height          =   615
      Left            =   -675
      TabIndex        =   25
      Top             =   2505
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 6 [ Sinhala ]"
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
      Left            =   8415
      TabIndex        =   19
      Top             =   2805
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 5 [ Sinhala ]"
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
      Left            =   8415
      TabIndex        =   18
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 3  [ English ]"
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
      Left            =   8415
      TabIndex        =   17
      Top             =   1215
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line 2  [ English ]"
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
      Left            =   8415
      TabIndex        =   16
      Top             =   705
      Width           =   3015
   End
End
Attribute VB_Name = "H_Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dbcon As New ADODB.Connection
Private InsertPhoto As New ADODB.Recordset

Dim cn As DAO.Database
Dim RES As DAO.Recordset
Dim LanguageDB As DAO.Database
Dim LanguageRes As DAO.Recordset


Private SavA As Boolean
Private SavS As Boolean
Dim rs As Recordset, b() As Byte, f As Long, fn As String
Private CrefName As String
Dim IDxi As Integer
Dim SP_Cnt As Integer

Private Const ToutDly = 0.1
'''Private Const MALAYALAM = 1
'''Private Const TAMIL = 2
'''Private Const HINDI = 3
'''Private Const SINHALA = 4


Private Const TAMIL = 1
Private Const SINHALA = 2
Private Const HINDI = 3
Private Const MALAYALAM = 4

Private HT1 As Boolean
Private HT2 As Boolean
Private HT3 As Boolean
Private HT4 As Boolean
Private HT5 As Boolean

Private Const TEnd = &HE

Private filename As String
Private fsoo As New FileSystemObject
Private tsrm As TextStream

Private NoOfLocalLanguageChar As Integer

Private Type PORTSETUP
    Port As Byte 'String * 1
    baud As String * 6
End Type

Dim PSetup As PORTSETUP

Dim LanguageStg(23) As Byte



Dim LanguageCharCount As Byte
Dim LanguageValidChar As Byte
Dim cbyte1 As Byte
Dim SAVE_FLAG As Boolean
Dim Msg As String
Dim NofPixel As Byte

Private Sub CNext_Click()
On Error Resume Next
    If InsertPhoto.RecordCount > 0 And InsertPhoto.EOF = False Then
        InsertPhoto.MoveNext
        Call ShowData
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
    Frame1.Enabled = True
    Text3(0).Enabled = False
    Text3(1).Enabled = False
    Text3(2).Enabled = False
    Text3(Index).Enabled = True
    CShow.Enabled = False
    Cref.Enabled = False
End Sub

Private Sub Command10_Click()
Dim Msg As VbMsgBoxResult
On Error Resume Next
 
    If SAVE_FLAG = False And Text1.Text = "" And strLanguageStage = "" Then
        Msg = MsgBox("Inavlid Stage Name!" & vbCrLf & "YES to Continue with space" & vbCrLf & "NO to Try again", vbYesNo)
        If Msg = vbYes Then
            SaveLanguageStg (&H20)
        Else
            Cancel = True
            Exit Sub
        End If
    ElseIf Text1.Text <> "" Then
        Command2.Value = True
    End If
    
    If InsertPhoto.State = adStateOpen Then
        InsertPhoto.Close
    End If
    sltxt = Trim(Text3(0).Text)
    Unload Me
End Sub

Private Sub Command2_Click()
 'On Error GoTo errLn
   Dim tempW() As String
   Dim tempH() As String
   Dim Lstr As String
    Dim idxxx As Integer
    Dim wIdx As Integer
    Dim MWthBC As Integer
    Dim MWthAD As Integer
    Dim er As Integer
    Dim Lmt As Integer
    Dim tltxt(23) As Byte
    Dim I As Byte
    Frmflag = True
    If strLanguageStage <> "" And NofPixel = 0 Then
        Text1.Text = ""
        IDxi = 0
        NofPixel = 0
        Unload Me
        Exit Sub
    End If
    If NofPixel = 0 Then
        MsgBox "Invalid Stage Name! " & vbCrLf & "Please Enter a Valid Stage Name", vbInformation, "BUS"
        Exit Sub
    End If
    strLanguageStage = Text1.Text
    Lmt = 7
    Lstr = ""
    tempW() = Split(Trim(Text7.Text), "-")
    tempH() = Split(Trim(Text1.Text), "-")
    If UBound(tempW()) <> UBound(tempH()) Then MsgBox "Error": Exit Sub
    For wIdx = 0 To UBound(tempW())
        MWthBC = MWthBC + tempW(wIdx)
        If (wIdx + 1) <= UBound(tempW()) Then
            If tempW(wIdx + 1) <= Lmt Then
              MWthBC = MWthBC + tempW(wIdx + 1)
              MWthBC = MWthBC - tempW(wIdx + 1)
              er = 5
              MWthAD = MWthBC
              Lstr = Lstr & tempH(wIdx)
              er = 6
            Else
                MWthAD = MWthBC
                Lstr = Lstr & tempH(wIdx)
                er = 7
            End If
        Else
              MWthAD = MWthBC
              Lstr = Lstr & tempH(wIdx)
              er = 8
        End If
    Next wIdx
        
        '''''''''''''''''''
    
    
  '''''''''syam  MsgBox Lstr
    
    ''''''''''''''''''''
        
        If HT1 = True Then Text3(0).Text = Lstr: Lstr = ""  'Replace(Text1.Text, "-", "")
        If HT2 = True Then Text3(1).Text = Lstr: Lstr = ""  'Replace(Text1.Text, "-", "")
        If HT3 = True Then Text3(2).Text = Lstr: Lstr = ""  'Replace(Text1.Text, "-", "")
        If HT4 = True Then Text3(3).Text = Lstr: Lstr = ""  'Replace(Text1.Text, "-", "")
        If HT5 = True Then Text3(4).Text = Lstr: Lstr = ""  'Replace(Text1.Text, "-", "")
        Frame1.Enabled = False
   For idxxx = 0 To 39
        Image2(idxxx).Picture = Nothing
   Next idxxx
   Text7.Text = ""
        If HT1 = True Then Text3(0).SetFocus
        If HT2 = True Then Text3(1).SetFocus
        If HT3 = True Then Text3(2).SetFocus
        If HT4 = True Then Text3(3).SetFocus
        If HT5 = True Then Text3(4).SetFocus
         HT1 = False
         HT2 = False
         HT3 = False
         HT4 = False
         HT5 = False
   
   
   
    If SaveLanguageStg = False Then
'        MsgBox "Error in Saving Stage Name!" & vbCrLf & "Filled with Space", vbInformation
        SaveLanguageStg (&H20)
    End If
    For IDxi = 0 To UBound(LanguageStg)
       LanguageStg(IDxi) = &H0
    Next
    Text1.Text = ""
    IDxi = 0
    NofPixel = 0
    Unload Me

    ' MsgBox "1"
 
    Exit Sub
errLn:    MsgBox "Error :" & err.Number & vbCrLf & err.Description, vbCritical, "ERROR"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim x As Integer
    Command3.Enabled = False
    Command4.Enabled = False
'    FlushPort
    For x = 0 To 4
       Text3(x).Enabled = False
    Next x
    
    For x = 0 To 2
       Text2(x).Enabled = False
    Next x
    
    Text4.Enabled = False
    Text5.Enabled = False
    
    If Text2(0).Text <> "" Then sndPrintE (Text2(0).Text)
    If Text2(1).Text <> "" Then sndPrintE (Text2(1).Text)
    If Text2(2).Text <> "" Then sndPrintE (Text2(2).Text)
'     'SendByte &H18
     DWait (ToutDly)
    For x = 0 To 4
      If Text3(x).Text <> "" Then sndPrintH (Text3(x).Text)
    Next x
'     'SendByte &H19
'    'SendByte TEnd
    DWait (ToutDly)
    If Text4.Text <> "" Then sndPrintE (Text4.Text)
    If Text5.Text <> "" Then sndPrintE (Text5.Text)
    
    Command3.Enabled = True
    Command4.Enabled = True
    
    For x = 0 To 4
       Text3(x).Enabled = True
    Next x
    
    For x = 0 To 2
       Text2(x).Enabled = True
    Next x
    
    Text4.Enabled = True
    Text5.Enabled = True
    Frame1.Enabled = False
    Command7.Enabled = False
    
    For idxxx = 0 To 39
         Image2(idxxx).Picture = Nothing
    Next idxxx
    
   Text1.Text = ""
   Text7.Text = ""
   Text2(0).SetFocus
   SendKeys "{Home}+{End}"
'    FlushPort
End Sub

Private Sub Command4_Click()
On Error Resume Next
    CD.filename = ""
    CD.Filter = "[*.PRH]|*.PRH"
    CD.ShowOpen
    If CD.filename = "" Then
       Exit Sub
    End If
     If CD.filename <> "" Then
        '   If UCase(CD.filename) = UCase(App.Path & "\POSCheck.DAT") Then
        '        MsgBox "Default fileName overwriting not allowed", vbCritical, "Distribution"
        '   Else
          filename = CD.filename
          Set tsrm = fsoo.OpenTextFile(filename, ForReading, True)
          For x = 0 To 2
              Text2(x).Text = tsrm.ReadLine
          Next x
          For x = 0 To 4
             Text3(x).Text = tsrm.ReadLine
          Next x
          Text4.Text = tsrm.ReadLine
          Text5.Text = tsrm.ReadLine
          tsrm.Close
     End If
     Frame1.Enabled = False
     Command7.Enabled = False
     For idxxx = 0 To 39
         Image2(idxxx).Picture = Nothing
     Next idxxx
    Text1.Text = ""
    Text7.Text = ""
    Text2(0).SetFocus
    SendKeys "{Home}+{End}"
End Sub

Private Sub Command5_Click()
On Error Resume Next
'  If SP_Cnt = 0 Then
    If Image2.UBound >= IDxi Then
        CharWcnt = CharWcnt + val(6)
        NofPixel = CharWcnt
        If CharWcnt > 172 Or LanguageValidChar > 13 Or LanguageCharCount > 22 Then
           CharWcnt = CharWcnt - val(6)
           NofPixel = CharWcnt
           Exit Sub
        End If
        SavePath = App.Path & "\Space.bmp"
        Image2(IDxi).Picture = LoadPicture(SavePath)
        IDxi = IDxi + 1
        If Trim(Text1.Text) <> "" Then Text1.Text = Trim(Text1.Text) & "-"
        Text1.Text = Trim(Text1.Text) & "20"
        LanguageStg(LanguageCharCount) = &H20
        LanguageCharCount = LanguageCharCount + 1
        LanguageValidChar = LanguageValidChar + 1
        If Trim(Text7.Text) <> "" Then Text7.Text = Trim(Text7.Text) & "-"
        Text7.Text = Trim(Text7.Text) & "06"
    End If
End Sub

Private Sub Command6_Click()
On Error Resume Next

    For x = 0 To 4
       Text3(x).Text = ""
    Next x
    
    For x = 0 To 2
       Text2(x).Text = ""
    Next x
    
    Text4.Text = ""
    Text5.Text = ""
    
    For idxxx = 0 To 39
       Image2(idxxx).Picture = Nothing
    Next idxxx
    
    Text1.Text = ""
    Text7.Text = ""
    Text2(0).SetFocus
End Sub

Private Sub Command7_Click()
On Error Resume Next

    If HT1 = True Then Text3(0).Text = ""
    If HT2 = True Then Text3(1).Text = ""
    If HT3 = True Then Text3(2).Text = ""
    If HT4 = True Then Text3(3).Text = ""
    If HT5 = True Then Text3(4).Text = ""
    
    For idxxx = 0 To 39
         Image2(idxxx).Picture = Nothing
    Next idxxx
    
    Text1.Text = ""
    Text7.Text = ""
    
    If HT1 = True Then Text3(0).SetFocus
    If HT2 = True Then Text3(1).SetFocus
    If HT3 = True Then Text3(2).SetFocus
    If HT4 = True Then Text3(3).SetFocus
    If HT5 = True Then Text3(4).SetFocus
    IDxi = 0
End Sub

Private Sub Command8_Click()
    End
End Sub

Private Sub Command9_Click()
On Error Resume Next
    CD.filename = ""
    CD.Filter = "[*.PRH]|*.PRH"
    CD.ShowSave
    If CD.filename = "" Then
       Exit Sub
    End If
    If CD.filename <> "" Then
        Flag = vbYes
        '   If UCase(CD.filename) = UCase(App.Path & "\POSCheck.DAT") Then
        '        MsgBox "Default fileName overwriting not allowed", vbCritical, "Distribution"
        '   Else
        If Dir(UCase(CD.filename)) <> "" Then
            Flag = MsgBox("File allready exist" & Chr(13) & "Do you want to Overwrite", vbYesNo + vbQuestion, App.Title)
        End If
        If Flag = vbYes Then
            filename = CD.filename
            Set tsrm = fsoo.OpenTextFile(filename, ForWriting, True)
            For x = 0 To 2
               tsrm.WriteLine Text2(x).Text
            Next x
            For x = 0 To 4
               tsrm.WriteLine Text3(x).Text
            Next x
            tsrm.WriteLine Text4.Text
            tsrm.WriteLine Text5.Text
            tsrm.Close
        End If
    End If
    Frame1.Enabled = False
    Command7.Enabled = False
    For idxxx = 0 To 39
        Image2(idxxx).Picture = Nothing
    Next idxxx
    Text1.Text = ""
    Text7.Text = ""
    Text2(0).SetFocus
    SendKeys "{Home}+{End}"
End Sub

Private Sub CommandP_Click()
On Error Resume Next
'  Dim sor As New FileSystemObject
'  Dim fsrc As TextStream
    
    If TextP.Text = "" Then
        MsgBox "Enter Port Number", vbInformation, "Mahindra"
    Else
        PortNo = TextP.Text
        Set SerialComm = H_Convert.MSComm1
        PSetup.Port = PortNo
        PSetup.baud = 9600 '115200
    End If
    
    If H_Convert.MSComm1.PortOpen = True Then
       H_Convert.MSComm1.PortOpen = False
    End If
    
    If Not InitPort(val(PSetup.Port), PSetup.baud) Then
        MsgBox TransMsg
        Command3.Enabled = False
        Exit Sub
    Else
        FHndl = FreeFile()
        Open UCase(App.Path & "\trans.dat") For Binary Access Write As #FHndl
        Put FHndl, , PSetup
        Close #FHndl
        MsgBox "Port No.: " & PortNo & Chr(13) & "Baud Rate : 9600" & Chr(13) & "Port opened Successfully", vbInformation, "KOT"
        Command3.Enabled = True
    End If
    Frame1.Enabled = False
    Command7.Enabled = False
    For idxxx = 0 To 39
        Image2(idxxx).Picture = Nothing
    Next idxxx
    Text1.Text = ""
    Text7.Text = ""
    TextP.SetFocus
    SendKeys "{Home}+{End}"
End Sub

Private Sub CPrev_Click()
On Error Resume Next
    If InsertPhoto.RecordCount > 0 And InsertPhoto.BOF = False Then
        InsertPhoto.MovePrevious
        Call ShowData
    End If
End Sub

Private Sub CSave_Click()
On Error Resume Next
    If Cref.Text = "View All" Then Exit Sub
    If InsertPhoto.State = adStateOpen Then
        InsertPhoto.Close
    End If
    dbcon.Execute ("Update Input_Data set Name='" & Text2(0).Text & "', Add1='" & Text2(1).Text & "', Add2='" & Text2(2).Text & "',H_Name='" & Text3(0).Text & "', H_Add1='" & Text3(1).Text & "', H_Add2='" & Text3(2).Text & "' where  CONS_REF='" & CrefName & "' ")
    MsgBox "Saved Successfully", vbInformation, App.Title
    With InsertPhoto
        .ActiveConnection = dbcon
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Source = "Select CONS_REF, Name, H_Name, Add1, H_Add1, Add2, H_Add2 from Input_Data order by CONS_REF"
        .Open
    End With
    Call ShowData
End Sub

Private Sub CShow_Click()
On Error Resume Next
    If Cref.Text <> "View All" Then
        If InsertPhoto.State = adStateOpen Then
                InsertPhoto.Close
        End If
        CrefName = Cref.Text
        With InsertPhoto
            .ActiveConnection = dbcon
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Source = "Select CONS_REF, Name, H_Name, Add1, H_Add1, Add2, H_Add2 from Input_Data where CONS_REF='" & Cref.Text & "' "
            .Open
            .MoveFirst
        End With
        SavS = True
        Call ShowData
        SavS = False
    Else
        If InsertPhoto.State = adStateOpen Then
            InsertPhoto.Close
        End If
        With InsertPhoto
            .ActiveConnection = dbcon
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Source = "Select CONS_REF, Name, H_Name, Add1, H_Add1, Add2, H_Add2 from Input_Data order by CONS_REF"
            .Open
            .MoveFirst
        End With
        SavS = True
        Call ShowData
        SavS = False
    End If
End Sub

Private Sub Form_Deactivate()
On Error Resume Next
''''''''''''syam added
    If Editing Then
        Editing = False
        Unload Me
    Else
        Dim Msg As VbMsgBoxResult
        If SAVE_FLAG = False Then
            Msg = MsgBox("Inavlid Stage Name!           " & vbCrLf & "YES to Continue with space" & vbCrLf & "NO to Try again", vbYesNo)
            If Msg = vbYes Then
                SaveLanguageStg (&H20)
            Else
                Load H_Convert
                H_Convert.Show vbModal
                Exit Sub
            End If
        End If
            If InsertPhoto.State = adStateOpen Then
                InsertPhoto.Close
            End If
            sltxt = Trim(Text3(0).Text)
        
            Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If (KeyAscii >= 65 And KeyAscii <= (65 + NoOfLocalLanguageChar) - 1) Then
        Image1_Click (KeyAscii - 65)
    ElseIf KeyAscii = 8 Then
        Image3_Click
    ElseIf KeyAscii = 32 Then
        Command5.Value = True
    ElseIf KeyAscii = 13 Then
        Command2.Value = True
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    ''CloseDB
    Set SerialComm = Me.MSComm1
    If initconn = False Then Command3.Enabled = False
    connect
End Sub
Private Sub Form_Activate()
'On Error GoTo err
Me.Icon = frmMainform.Icon
Dim ProvSSN As String
Dim SavePath As String
Dim PicEdit As Picture
Dim iDx As Integer
    Set cn = DAO.OpenDatabase(App.Path & "\PVT.Mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    Set LanguageDB = DAO.OpenDatabase(App.Path & "\HCHAR.MDB", dbDriverComplete, False, ";UID=;PWD=siljvvnl")
    SAVE_FLAG = False
    IDxi = 0
    LanguageCharCount = 0
    LanguageValidChar = 0
    SavS = False
    SavA = False
    Frame1.Enabled = False
    SP_Cnt = 0
    CharWcnt = 0
    Maltxt = ""
    
    LoadLanguageStage (NOSTGS)

    If LocalLanguage = 0 Then
        Unload Me
        Exit Sub
    End If
    Me.caption = strLocalLanguage & " Printer Pro"

    For iDx = 0 To 23
        LanguageStg(iDx) = &H0
    Next
    If InsertPhoto.State = adStateOpen Then
            InsertPhoto.Close
    End If
    'setting cursor and open information for recordset table
    With InsertPhoto
        .ActiveConnection = dbcon
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Source = "Select Image from " & strLocalLanguage & " order by MapID" 'TamilChar" ' "INSERT INTO ArabChar (Image) VALUES ('" & b() & "') "  'OLE1.Picture
        .Open
    End With
    If InsertPhoto.RecordCount > 0 Then
        InsertPhoto.MoveFirst
        iDx = 0
        Do While Not InsertPhoto.EOF
            b() = InsertPhoto.Fields("Image")
            ProvSSN = temp
            If Dir(App.Path & "\Data", vbDirectory) = "" Then MkDir (App.Path & "\Data")
            SavePath = App.Path & "\Data\" & iDx + 1 & ".bmp"
            f = FreeFile()
            fn = SavePath
            Open fn For Binary Access Write As #f
                Put #f, , b()
            Close f
            Image1(iDx).Visible = True
            Image1(iDx).Picture = LoadPicture(SavePath) 'Works and loads image
            iDx = iDx + 1
            InsertPhoto.MoveNext
        Loop
    End If
    InsertPhoto.Close
    HT1 = False
    HT2 = False
    HT3 = False
    HT4 = False
    HT5 = False
    Select Case Index
        Case 0: HT1 = True
        Case 1: HT2 = True
        Case 2: HT3 = True
        Case 3: HT4 = True
        Case 4: HT5 = True
    End Select
    Frame1.Enabled = True
    SendKeys "{Home}+{End}"
    Command7.Enabled = True
    If Text3(Index).Text = "" Then Command7.Value = True
    If Text3(Index).Text <> "" Then
        For idxxx = 0 To 39
             Image2(idxxx).Picture = Nothing
        Next idxxx
        Text1.Text = ""
        Text7.Text = ""
        IDxi = 0
        AddFont Text3(Index).Text
    End If
    Exit Sub
err:
    If err.Number = -2147217865 Then
        MsgBox "Error!" & vbCrLf & strLocalLanguage & " Language Not Available!"
        Unload Me
    Else
        MsgBox "Error !" & vbCrLf & "Err No: " & err.Number & vbCrLf & err.Description, vbCritical
    End If
    
End Sub

Public Sub connect()
On Error Resume Next
    If dbcon.State <> 1 Then
        dbcon.Provider = "Microsoft.jet.oledb.4.0"
        dbcon.Properties("Jet OLEDB:Database Password") = "siljvvnl"
        dbcon.ConnectionString = App.Path & "\HCHAR.mdb"
        dbcon.Open
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim Msg As VbMsgBoxResult
    
    If SAVE_FLAG = False And Text1.Text = "" And strLanguageStage = "" Then
        Msg = MsgBox("Inavlid Stage Name!" & vbCrLf & "YES to Continue with space" & vbCrLf & "NO to Try again", vbYesNo)
        If Msg = vbYes Then
            SaveLanguageStg (&H20)
        Else
            Cancel = True
            Exit Sub
        End If
        'The code below causes error after the data get stored in the database
        'it triggers the click event of Command2 Button
        
    End If

    If InsertPhoto.State = adStateOpen Then
        InsertPhoto.Close
    End If
    sltxt = Trim(Text3(0).Text)
   ' sleep (200)
    
'    Unload Me
 '   FareTableFrm.Visible = True
    
End Sub

Private Sub Image1_Click(Index As Integer)
Dim SavePath As String
On Error GoTo err
    If InsertPhoto.State = adStateOpen Then ''''''''''insertphoto--BDCON
        InsertPhoto.Close
    End If
    With InsertPhoto
        .ActiveConnection = dbcon
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Source = "Select Iscii, Width from " & strLocalLanguage & " where MapId=" & Index + 1
        .Open
        If Image2.UBound >= IDxi Then
            If .RecordCount > 0 Then
                CharWcnt = CharWcnt + val(.Fields(1))
                NofPixel = CharWcnt
                If CharWcnt > 172 Or LanguageValidChar > 13 Or LanguageCharCount > 22 Then
                   CharWcnt = CharWcnt - val(.Fields(1))
                   NofPixel = CharWcnt
                   Exit Sub
                End If
            End If
            If val(.Fields(1)) <> 0 Then
                LanguageValidChar = LanguageValidChar + 1
            End If
            If Trim(Text1.Text) <> "" Then Text1.Text = Trim(Text1.Text) & "-"
            Text1.Text = Trim(Text1.Text) & .Fields(0)
            cbyte1 = CByte("&H" & Mid(.Fields(0), 1, 1))
            LanguageStg(LanguageCharCount) = ((cbyte1 * 16) + CByte("&H" & Mid(.Fields(0), 2, 1)))
            LanguageCharCount = LanguageCharCount + 1
            If Trim(Text7.Text) <> "" Then Text7.Text = Trim(Text7.Text) & "-"
            Text7.Text = Trim(Text7.Text) & Format(.Fields(1), "00")
            SavePath = App.Path & "\Data\" & Index + 1 & ".bmp"
            Image2(IDxi).Picture = LoadPicture(SavePath)
            IDxi = IDxi + 1
        End If
        .Close
    End With
    SP_Cnt = 0
    Exit Sub
err:
End Sub

Private Sub Image3_Click()
'On Error GoTo err
Dim RemovedCharLen As Byte
Dim endLoc As Byte
    endLoc = InStrRev(Trim(Text1.Text), "-")
    If endLoc <= 0 And Len(Text1.Text) = 0 Then
        Text1.Text = ""
        Text7.Text = ""
        If Not IDxi < 1 Then
            IDxi = IDxi - 1
            LanguageCharCount = LanguageCharCount - 1
            If LanguageValidChar > 0 Then
                LanguageValidChar = LanguageValidChar - 1
                If LanguageValidChar = 0 Then NofPixel = CharWcnt = 0
            End If
            Image2(IDxi).Picture = Nothing
        End If
        Exit Sub
    End If
    If Len(Text1.Text) = 2 Then
        RemovedCharLen = val(Mid$(Trim(Text7.Text), endLoc + 1, Len(Text7.Text)))
        Text1.Text = ""
        Text7.Text = ""
    Else
        Text1.Text = Mid$(Trim(Text1.Text), 1, endLoc - 1)
        RemovedCharLen = val(Mid$(Trim(Text7.Text), endLoc + 1, Len(Text7.Text)))
        Text7.Text = Mid$(Trim(Text7.Text), 1, endLoc - 1)
    End If
    CharWcnt = CharWcnt - RemovedCharLen
    NofPixel = CharWcnt
    IDxi = IDxi - 1
    LanguageCharCount = LanguageCharCount - 1
    LanguageStg(LanguageCharCount) = &H0
    If RemovedCharLen <> 0 Then
        LanguageValidChar = LanguageValidChar - 1
    End If
    Image2(IDxi).Picture = Nothing
    Exit Sub
err:
    MsgBox "Error!" & vbCrLf & err.Description, Bus
End Sub

Public Sub ShowData()
On Error Resume Next
    With InsertPhoto
        If .RecordCount > 0 Then
            .MoveNext
            If .EOF = True Then
                CNext.Enabled = False
            Else
                CNext.Enabled = True
            End If
            .MovePrevious
            .MovePrevious
            If .BOF = True Then
                CPrev.Enabled = False
            Else
                CPrev.Enabled = True
            End If
            .MoveNext
            Text2(0) = !Name
            Text2(1) = !Add1
            Text2(2) = !Add2
            Text3(0) = !H_Name
            Text3(1) = !H_Add1
            Text3(2) = !H_Add2
            Cref.Text = !CONS_REF
            CrefName = !CONS_REF
        End If
    End With
End Sub

Private Sub Text2_Change(Index As Integer)
On Error Resume Next
    If SavS = False Then SavA = True
End Sub

Private Sub Text2_GotFocus(Index As Integer)
On Error Resume Next
    Frame1.Enabled = False
    Command7.Enabled = False
    For idxxx = 0 To 39
        Image2(idxxx).Picture = Nothing
    Next idxxx
    Text1.Text = ""
    Text7.Text = ""
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 13 Then
        Select Case KeyAscii
          Case 48 To 57:
                  KeyAscii = KeyAscii
          Case 65 To 90:
                  KeyAscii = KeyAscii
          Case 97 To 122:
                  KeyAscii = KeyAscii
          Case 13:
                  KeyAscii = KeyAscii
          Case 8:
                  KeyAscii = KeyAscii
          Case 32:
                  KeyAscii = KeyAscii
        '    Case 91:
        '            KeyAscii = KeyAscii
        '    Case 93:
        '            KeyAscii = KeyAscii
          Case 40 To 41:
                  KeyAscii = KeyAscii
          Case 45 To 46:
                  KeyAscii = KeyAscii
          Case Else:
                  KeyAscii = 0
        End Select
    End If
End Sub

Private Sub Text3_Change(Index As Integer)
On Error Resume Next
    If SavS = False Then SavA = True
    SendKeys "{Home}+{End}"
End Sub

Private Sub Text3_GotFocus(Index As Integer)
On Error Resume Next
    HT1 = False
    HT2 = False
    HT3 = False
    HT4 = False
    HT5 = False
    Select Case Index
        Case 0: HT1 = True
        Case 1: HT2 = True
        Case 2: HT3 = True
        Case 3: HT4 = True
        Case 4: HT5 = True
    End Select
    Frame1.Enabled = True
    SendKeys "{Home}+{End}"
    Command7.Enabled = True
    If Text3(Index).Text = "" Then Command7.Value = True
    If Text3(Index).Text <> "" Then
        For idxxx = 0 To 39
             Image2(idxxx).Picture = Nothing
        Next idxxx
        Text1.Text = ""
        Text7.Text = ""
        IDxi = 0
        AddFont Text3(Index).Text
    End If
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then KeyCode = 0
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    SendKeys "{Home}+{End}"
End Sub

Private Sub Text4_GotFocus()
On Error Resume Next
    Frame1.Enabled = False
    Command7.Enabled = False
    For idxxx = 0 To 39
         Image2(idxxx).Picture = Nothing
    Next idxxx
    Text1.Text = ""
    Text7.Text = ""
End Sub

Private Sub Text5_GotFocus()
On Error Resume Next
    Frame1.Enabled = False
    Command7.Enabled = False
    For idxxx = 0 To 39
         Image2(idxxx).Picture = Nothing
    Next idxxx
    Text1.Text = ""
    Text7.Text = ""
End Sub

Private Sub TextP_Change()
    If Trim(TextP.Text) <> "" Then
        If IsNumeric(TextP.Text) = False Then
            MsgBox "Only numeric values allowed", vbInformation, "KOT"
            TextP.Text = ""
        End If
    If val(TextP.Text) > 128 Then
        TextP.Text = Mid$(Trim(TextP.Text), 1, Len(TextP.Text) - 1)
        SendKeys "{End}"
    End If
    End If
End Sub

Private Sub TextP_GotFocus()
    Frame1.Enabled = False
    Command7.Enabled = False
   For idxxx = 0 To 39
        Image2(idxxx).Picture = Nothing
   Next idxxx
   Text1.Text = ""
   Text7.Text = ""
   TextP.SetFocus
    SendKeys "{Home}+{End}"
End Sub

Private Sub TextP_KeyPress(KeyAscii As Integer)
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

Private Sub AddFont(Fstr As String)
  Dim I As Integer
  Dim Cvstr As String * 2
  Dim sBt As Byte
   For I = 1 To Len(Fstr) Step 2
     Cvstr = Mid$(Fstr, I, 2)
     If Cvstr = "20" Then
        Command5_Click
     ElseIf Cvstr <> "0D" Then
If InsertPhoto.State = adStateOpen Then
        InsertPhoto.Close
End If
With InsertPhoto
    .ActiveConnection = dbcon
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Source = "Select MapId from " & strLocalLanguage & " where Iscii='" & Cvstr & "'"
    .Open
    If .RecordCount > 0 Then
        Image1_Click (.Fields(0) - 1)
    End If
If InsertPhoto.State = adStateOpen Then
        InsertPhoto.Close
End If
End With
    End If
   Next I
End Sub

Public Function SaveLanguageStg(Optional FByte As Byte) As Boolean
'On Error GoTo errLn
    Dim FHndl As Integer
    Dim fname As String
    Dim THndl As Integer
    Dim tname As String
    Dim StgCount As Long
    Dim PixelToFill As Integer
    Dim j As Byte
    Dim k As Byte
    Dim SpaceCount As Byte
    
    SaveLanguageStg = False
    fname = App.Path & "\LOCAL_LANGUAGE.DAT"
    tname = App.Path & "\TEMP.DAT"
    RSql = "SELECT count(*) FROM STAGE"
    Set RES = cn.OpenRecordset(RSql, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            StgCount = RES.Fields(0)
            StgCount = StgCount + StgNameCnt - 1
        Else
            StgCount = StgNameCnt - 1
        End If
    RES.Close
    
    
    
    For I = 0 To 23
            LSTAG.LocalLanguageStageName(I) = &H0
    Next
    ''''''''''clarify
    If StgCount < 1 Then
        PixelToFill = 0
        If NofPixel < 166 And NofPixel <> 0 Then
            PixelToFill = ((172 - NofPixel) / 6) - 1
            If PixelToFill > 16 Then
                PixelToFill = 16
            End If
            If PixelToFill Mod 2 <> 0 Then PixelToFill = PixelToFill + 1
        End If
        
        If Dir(fname) <> "" Then Kill fname
        FHndl = FreeFile()
        
        Open fname For Binary Access Write As #FHndl
            j = 0
            For I = 0 To 22
                If PixelToFill <> 0 Then
                    For k = 0 To PixelToFill / 2
                        LSTAG.LocalLanguageStageName(I) = &H20
                        I = I + 1
                    Next
                    SpaceCount = PixelToFill / 2
                    PixelToFill = 0
                End If
                
                If FByte = &H20 Then
                    LSTAG.LocalLanguageStageName(I) = &H20
                Else
                    If LanguageStg(j) = &H0 Then
                        For k = 0 To SpaceCount
                            LSTAG.LocalLanguageStageName(I) = &H20
                            I = I + 1
                            If I = 22 Then Exit For
                        Next
                        Exit For
                    End If
                    LSTAG.LocalLanguageStageName(I) = LanguageStg(j)
                End If
                j = j + 1
            Next
            LSTAG.stagecode = StgCount
            LSTAG.RouteCode = Trim(FareTableFrm.txtrutcode) & Chr(0)
            Put #FHndl, , LSTAG
        Close #FHndl
        SAVE_FLAG = True
        SaveLanguageStg = True
        Unload Me
        Exit Function
    End If
    
    If Dir(fname) <> "" Then
        FHndl = FreeFile()
        Open fname For Binary Access Read As #FHndl
        If Dir(tname) <> "" Then Kill tname
        THndl = FreeFile()
        Open tname For Binary Access Write As #THndl
        Do While Not EOF(FHndl)
            Get #FHndl, , LSTAG
            If LSTAG.LocalLanguageStageName(0) <> &H0 Then
                Put #THndl, , LSTAG
            End If
        Loop
        Close #FHndl
        Close #THndl
        If Dir(fname) <> "" Then Kill fname
        THndl = FreeFile()
        Open tname For Binary Access Read As #THndl
        FHndl = FreeFile()
        Open fname For Binary Access Write As #FHndl
            Do While Not EOF(THndl)
                Get #THndl, , LSTAG
                If LSTAG.LocalLanguageStageName(0) <> &H0 Then
                    Put #FHndl, , LSTAG
                End If
            Loop
            Close #THndl
            
            PixelToFill = 0
            If NofPixel < 166 And NofPixel <> 0 Then
                PixelToFill = ((172 - NofPixel) / 6) - 1
                If PixelToFill > 16 Then
                    PixelToFill = 16
                End If
                If PixelToFill Mod 2 <> 0 Then PixelToFill = PixelToFill + 1
            End If
            
            j = 0
            For I = 0 To 22
                If PixelToFill <> 0 Then
                    For k = 0 To PixelToFill / 2
                        LSTAG.LocalLanguageStageName(I) = &H20
                        I = I + 1
                    Next
                    SpaceCount = PixelToFill / 2
                    PixelToFill = 0
                End If
                
                If FByte = &H20 Then
                    LSTAG.LocalLanguageStageName(I) = &H20
                Else
                    If LanguageStg(j) = &H0 Then
                        For k = 0 To SpaceCount
                        
                        ''''''''''' if added by syam
                        ''''reason :  sometimes i=24 causes out of bounds exception
                        
                        If I < 24 Then ''''''
                            LSTAG.LocalLanguageStageName(I) = &H20
                            I = I + 1
                            If I = 22 Then Exit For
                        End If '''''''''
                        Next
                        
                        Exit For
                    End If
                    LSTAG.LocalLanguageStageName(I) = LanguageStg(j)
                End If
                j = j + 1
            Next
            LSTAG.stagecode = StgCount
            LSTAG.RouteCode = Trim(FareTableFrm.txtrutcode) & Chr(0)
            Put #FHndl, , LSTAG
        Close #FHndl
    End If
    SAVE_FLAG = True
    SaveLanguageStg = True
    Exit Function
errLn:
    MsgBox "Error :" & err.Number & vbCrLf & err.Description, vbCritical, "ERROR"
    Close #FHndl
    Close #THndl
    SaveLanguageStg = False

End Function
Public Function LoadLanguageStage(ByVal bNos As Byte)
Dim sTempStr As String
Dim SavePath As String
Dim strSql As String
Dim I As Byte

        
    strSql = "SELECT * FROM LANGUAGE_ENABLED"
    Set LanguageRes = LanguageDB.OpenRecordset(strSql, dbOpenDynaset)
    
    
    '''''''''''''''''FLAGSET
    If LanguageRes!MALAYALAM = True Then
        LocalLanguage = MALAYALAM
        strLocalLanguage = "MALAYALAM"
        NoOfLocalLanguageChar = 43
    ElseIf LanguageRes!TAMIL = True Then
        LocalLanguage = TAMIL
        strLocalLanguage = "TAMIL"
        NoOfLocalLanguageChar = 43
    ElseIf LanguageRes!HINDI = True Then
        LocalLanguage = HINDI
        strLocalLanguage = "HINDI"
        NoOfLocalLanguageChar = 57
    ElseIf LanguageRes!SINHALA = True Then
        LocalLanguage = SINHALA
        strLocalLanguage = "SINHALA"
        NoOfLocalLanguageChar = 61
    Else
        LocalLanguage = 0
        strLocalLanguage = ""
        NoOfLocalLanguageChar = 0
        Exit Function
    End If
    
    If strLanguageStage = "" Then Exit Function
    sTempStr = strLanguageStage
    
    
    If InsertPhoto.State = adStateOpen Then
            InsertPhoto.Close
    End If
    
    
    
    With InsertPhoto
        
        For I = 1 To bNos
            .ActiveConnection = dbcon
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Source = "Select Iscii, Width from " & strLocalLanguage & " where MapId=" & Index + 1
            .Open
            Text7.Text = Trim(Text7.Text) & Format(.Fields(1), "00")
            SavePath = App.Path & "\Data\" & Index + 1 & ".bmp"
            Image2(IDxi).Picture = LoadPicture(SavePath)
            IDxi = IDxi + 1
            .Close
        Next
    End With
End Function

