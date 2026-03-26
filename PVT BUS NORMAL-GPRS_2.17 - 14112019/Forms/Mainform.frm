VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmMainform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Amphibia Bus Ticketing 2.17   Date : 03/09/2019"
   ClientHeight    =   10305
   ClientLeft      =   570
   ClientTop       =   630
   ClientWidth     =   17940
   Icon            =   "Mainform.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Mainform.frx":11F84
   ScaleHeight     =   10719.75
   ScaleMode       =   0  'User
   ScaleWidth      =   17940
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   6360
      Top             =   600
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   4560
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7320
      Index           =   3
      Left            =   6270
      ScaleHeight     =   7320
      ScaleWidth      =   1725
      TabIndex        =   4
      Top             =   525
      Visible         =   0   'False
      Width           =   1725
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date &  Time"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   4320
         Width           =   1590
      End
      Begin JeweledBut.JeweledButton Command4 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   39
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Entry"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFD07
         BC              =   12632256
         FC              =   0
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Transfer Data"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   3240
         Width           =   1590
      End
      Begin JeweledBut.JeweledButton Command4 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   40
         Top             =   1710
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Edit Data"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFD23
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command4 
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   41
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Settings"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFD3F
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command4 
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   42
         Top             =   2250
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Transfer"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFD5B
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command4 
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   43
         Top             =   5714
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFD77
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command4 
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   44
         Top             =   6270
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFD93
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdexit4 
         Height          =   480
         Left            =   0
         TabIndex        =   75
         Top             =   6845
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   847
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFDAF
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdLogo3 
         Height          =   495
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Logo Setup"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFDCB
         BC              =   12632256
         FC              =   0
      End
      Begin VB.Label lbltime 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   1860
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7320
      Index           =   2
      Left            =   4380
      ScaleHeight     =   7320
      ScaleWidth      =   1725
      TabIndex        =   3
      Top             =   525
      Visible         =   0   'False
      Width           =   1725
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bmp Settings"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   75
         TabIndex        =   63
         Top             =   5295
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Path Settings"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   75
         TabIndex        =   57
         Top             =   3180
         Width           =   1590
      End
      Begin JeweledBut.JeweledButton Command3 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   33
         Top             =   4020
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Entry"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFDE7
         BC              =   12632256
         FC              =   0
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Palmtec Setup"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   75
         TabIndex        =   17
         Top             =   2700
         Width           =   1590
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Port Setup"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   75
         TabIndex        =   16
         Top             =   4980
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Change Pswd"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   75
         TabIndex        =   15
         Top             =   2220
         Width           =   1590
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Delete User"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   1755
         Width           =   1590
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Add User"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   13
         Top             =   1335
         Width           =   1590
      End
      Begin JeweledBut.JeweledButton Command3 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   34
         Top             =   4590
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Edit Data"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFE03
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command3 
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   35
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Settings"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFE1F
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command3 
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   36
         Top             =   5164
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Transfer"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFE3B
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command3 
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   37
         Top             =   5714
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFE57
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command3 
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   38
         Top             =   6270
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFE73
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdexit3 
         Height          =   480
         Left            =   0
         TabIndex        =   74
         Top             =   6845
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   847
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFE8F
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdLogo2 
         Height          =   495
         Left            =   0
         TabIndex        =   84
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Logo Setup"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFEAB
         BC              =   12632256
         FC              =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7354
      Index           =   1
      Left            =   2520
      ScaleHeight     =   7350
      ScaleWidth      =   1725
      TabIndex        =   2
      Top             =   525
      Visible         =   0   'False
      Width           =   1725
      Begin JeweledBut.JeweledButton Command2 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Entry"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFEC7
         BC              =   12632256
         FC              =   0
      End
      Begin VB.CheckBox chkFareTable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fare Table"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   21
         Top             =   3255
         Width           =   1590
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stage"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   135
         TabIndex        =   12
         Top             =   4485
         Width           =   1590
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Route"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   3840
         Width           =   1590
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Graph Fare"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   2595
         Width           =   1590
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   28
         Top             =   1710
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Edit Data"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFEE3
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   29
         Top             =   525
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Settings"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFEFF
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   30
         Top             =   5164
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Transfer"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFF1B
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   31
         Top             =   5714
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFF37
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command2 
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   32
         Top             =   6270
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFF53
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdext1 
         Height          =   480
         Left            =   0
         TabIndex        =   73
         Top             =   6846
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   847
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFF6F
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdLogo1 
         Height          =   495
         Left            =   0
         TabIndex        =   83
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Logo Setup"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFF8B
         BC              =   12632256
         FC              =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7354
      Index           =   0
      Left            =   675
      ScaleHeight     =   7291.667
      ScaleMode       =   0  'User
      ScaleWidth      =   1725
      TabIndex        =   1
      Top             =   525
      Visible         =   0   'False
      Width           =   1725
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Waybill"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   8
         Left            =   75
         TabIndex        =   82
         Top             =   6240
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fare"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   75
         TabIndex        =   81
         Top             =   6000
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Vehicle"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   75
         TabIndex        =   67
         Top             =   4035
         Width           =   1470
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Expense"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   75
         TabIndex        =   64
         Top             =   3630
         Width           =   1590
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bus Type"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   75
         TabIndex        =   59
         Top             =   2025
         Width           =   1470
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Currency"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   75
         TabIndex        =   58
         Top             =   6435
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Crew Details"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   75
         TabIndex        =   9
         Top             =   3240
         Width           =   1590
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Delete Route"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   2850
         Width           =   1590
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New Route"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   2445
         Width           =   1590
      End
      Begin JeweledBut.JeweledButton Command1 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   22
         Top             =   4560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Edit Data"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFFA7
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command1 
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   23
         Top             =   555
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Settings"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFFC3
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command1 
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   24
         Top             =   5141
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Transfer"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFFDF
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command1 
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   25
         Top             =   5695
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":BFFFB
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command1 
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   26
         Top             =   6250
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C0017
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command1 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   68
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Entry"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C0033
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command7 
         Height          =   480
         Left            =   0
         TabIndex        =   71
         Top             =   6854
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   847
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C004F
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdLogo 
         Height          =   495
         Left            =   0
         TabIndex        =   72
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Logo Setup"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C006B
         BC              =   12632256
         FC              =   0
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   2790
      Top             =   645
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7320
      Index           =   4
      Left            =   8220
      ScaleHeight     =   7320
      ScaleWidth      =   1725
      TabIndex        =   5
      Top             =   525
      Visible         =   0   'False
      Width           =   1725
      Begin VB.CheckBox Chkrefund 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Refund"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   8
         Left            =   80
         TabIndex        =   88
         Top             =   6000
         Width           =   1590
      End
      Begin VB.CheckBox Chkfare 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fare wise"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   75
         TabIndex        =   80
         Top             =   5640
         Width           =   1590
      End
      Begin VB.CheckBox ChkStage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stage wise Consolidate"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   6
         Left            =   75
         TabIndex        =   79
         Top             =   5040
         Width           =   1590
      End
      Begin VB.CheckBox ChkTicket 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " TicketNo"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   75
         TabIndex        =   78
         Top             =   4800
         Width           =   1590
      End
      Begin VB.CheckBox ChkSch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Schedule"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   75
         TabIndex        =   70
         Top             =   4440
         Width           =   1590
      End
      Begin VB.CheckBox ChkBussmry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Buswise"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   75
         TabIndex        =   69
         Top             =   4080
         Width           =   1830
      End
      Begin VB.CheckBox ChkSchsmry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Trip"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   75
         TabIndex        =   66
         Top             =   3840
         Width           =   1470
      End
      Begin VB.CheckBox ChkExpense 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Expense "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   75
         TabIndex        =   65
         Top             =   3600
         Width           =   1590
      End
      Begin JeweledBut.JeweledButton Command5 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   45
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Entry"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C0087
         BC              =   12632256
         FC              =   0
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "View Report"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   19
         Top             =   3360
         Width           =   1590
      End
      Begin JeweledBut.JeweledButton Command5 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   46
         Top             =   1710
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Edit Data"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C00A3
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command5 
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   47
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Settings"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C00BF
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command5 
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   48
         Top             =   2250
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Transfer"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C00DB
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command5 
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   49
         Top             =   2820
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C00F7
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command5 
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   50
         Top             =   6270
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C0113
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdexit5 
         Height          =   480
         Left            =   0
         TabIndex        =   76
         Top             =   6840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   847
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C012F
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdLogo4 
         Height          =   495
         Left            =   0
         TabIndex        =   86
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Logo Setup"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C014B
         BC              =   12632256
         FC              =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7320
      Index           =   5
      Left            =   10080
      ScaleHeight     =   7320
      ScaleWidth      =   1725
      TabIndex        =   6
      Top             =   525
      Visible         =   0   'False
      Width           =   1725
      Begin JeweledBut.JeweledButton Command6 
         Height          =   495
         Index           =   0
         Left            =   30
         TabIndex        =   51
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Entry"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C0167
         BC              =   12632256
         FC              =   0
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "About Softland India Ltd."
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   825
         Index           =   0
         Left            =   75
         TabIndex        =   20
         Top             =   4575
         Width           =   1590
      End
      Begin JeweledBut.JeweledButton Command6 
         Height          =   495
         Index           =   1
         Left            =   30
         TabIndex        =   52
         Top             =   1710
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Edit Data"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C0183
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command6 
         Height          =   495
         Index           =   2
         Left            =   30
         TabIndex        =   53
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Settings"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C019F
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command6 
         Height          =   495
         Index           =   3
         Left            =   30
         TabIndex        =   54
         Top             =   2250
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Data Transfer"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C01BB
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command6 
         Height          =   495
         Index           =   4
         Left            =   30
         TabIndex        =   55
         Top             =   2820
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C01D7
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton Command6 
         Height          =   495
         Index           =   5
         Left            =   30
         TabIndex        =   56
         Top             =   3390
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C01F3
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton JeweledButton1 
         Height          =   480
         Left            =   0
         TabIndex        =   77
         Top             =   6840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   847
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C020F
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdLog5 
         Height          =   495
         Left            =   0
         TabIndex        =   87
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Logo Setup"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Mainform.frx":C022B
         BC              =   12632256
         FC              =   0
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FOR TESTING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   5040
      TabIndex        =   62
      Top             =   1440
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   6225
      TabIndex        =   0
      Top             =   4605
      Width           =   1860
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   11430
      Picture         =   "Mainform.frx":C0247
      Stretch         =   -1  'True
      Top             =   5550
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuNewRuote 
         Caption         =   "&New Route"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Route"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuCrew 
         Caption         =   "&Crew Details"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuGraph 
         Caption         =   "&Graph"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuRoute 
         Caption         =   "&Route"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuStage 
         Caption         =   "&Stage"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "T&ools"
      Visible         =   0   'False
      Begin VB.Menu mnuCalculator 
         Caption         =   "Ca&lculator"
      End
      Begin VB.Menu mnuAddUser 
         Caption         =   "&Add User"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuRemoveUser 
         Caption         =   "&Remove User"
      End
      Begin VB.Menu mnuPasword 
         Caption         =   "&Change Password"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPort 
         Caption         =   "&Port Setup"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Palmtec &Setup"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuTransfer 
      Caption         =   "&Transfer"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   
Option Explicit
Dim hWnd1 As Long
Dim Fhandle As Long
Private m_BGnd As CMdiBackground
Dim Result As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                   (ByVal hWnd As Long, ByVal lpszOp As String, _
                    ByVal lpszFile As String, ByVal lpszParams As String, _
                    ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                    As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Declare Function GetWindowText Lib "user32" _
Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As _
String, ByVal cch As Long) As Long
 
Private Declare Function GetWindowTextLength Lib "user32" _
Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
 
Private Declare Function GetNextWindow Lib "user32" _
Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) _
As Long
         
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" _
(ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal process As Long, lpExitCode As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
  (ByVal process As Long, ByVal uExitCode As Long) As Long
Public Function EndApplication(ByRef caption As String, ByRef frm As Form) As Boolean
    Dim hWnd As Long
    Dim appInstance As Long
    Dim process As Long
    Dim processID
    Dim Result As Boolean
    Dim exitCode As Long
    Dim returnValue As Long
    
    On Error GoTo Error
    
    If Trim(caption) = "" Then Exit Function
    Do
        hWnd = FindWindowByTitle(caption, frm)
        If hWnd = 0 Then Exit Do
        appInstance = GetWindowThreadProcessId(hWnd, processID)
        'Get a handle for the process we're looking for
        process = OpenProcess(PROCESS_ALL_ACCESS, 0&, processID)
        If process <> 0 Then
            'Next get our exit code (for use later)
            GetExitCodeProcess process, exitCode
                'Check for an exit code of 9 (zero)
                If exitCode <> 0 Then
                     'It's not zero so close the window
                    returnValue = TerminateProcess(process, exitCode)
                    If Result = False Then Result = returnValue > 0
                End If
        End If
    Loop
    EndApplication = Result
Exit Function
Error:
    MsgBox (err.Number & ": " & err.Description)
End Function

Private Function FindWindowByTitle(ByRef str As String, ByRef frm As Form) As Long
    Dim Handle As Long
    Dim caption As String
    Dim sTitle As String
    
    Handle = frm.hWnd
    sTitle = LCase(str)
    Do
       DoEvents
          If Handle = 0 Then Exit Do
            caption = LCase$(GetWindowCaption(Handle))
    
           If InStr(caption, sTitle) Then
              FindWindowByTitle = Handle
              Exit Do
           Else
             FindWindowByTitle = 0
           End If
           Handle = GetNextWindow(Handle, 2)
    Loop
End Function

Private Function GetWindowCaption(ByRef Handle As Long) As String
    Dim str As String
    Dim Length As Long
    
    Length& = GetWindowTextLength(Handle)
    str = String(Length, 0)
    Call GetWindowText(Handle, str, Length + 1)
    GetWindowCaption = str
End Function



    

Private Sub Check1_Click(Index As Integer)

'''''''''''''''
''''''''syam addeds
Editing = False

  
''''''''''''''

If Check1(Index).Value = vbUnchecked Then Exit Sub
Frmflag = False
 Select Case Index
 Case 0:
    If Not SAVEFLAG = False Then
        If SUPERUSER = True Then
            Load FareTableFrm
            FareTableFrm.Show vbModal
        Else
            RouteAdd = True
            Load SuperLogin
            SuperLogin.Show vbModal
        End If
    Else
        Load FareTableFrm
        FareTableFrm.Show vbModal
    End If
 Case 1:
    If SUPERUSER = True Then
        Load frmRouteDelete
        frmRouteDelete.Show vbModal
    Else
        DeleteRoute = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
 Case 2:
    'Load frmCrewDetails
    'frmCrewDetails.Show vbModal
    Load Crew
    Crew.Show vbModal
Case 3:
    Load frmAddCurrency
    frmAddCurrency.Show vbModal
Case 4:
    Load frmAddBusType
    frmAddBusType.Show vbModal
Case 5:
    Load frmExpense
    frmExpense.Show vbModal
Case 6:
    Load frmVehicle
    frmVehicle.Show vbModal
Case 7:
    Load Tarif_frm
    Tarif_frm.Show vbModal
Case 8:
    Load waybill_frm
    waybill_frm.Show vbModal
 End Select
    Check1(Index).Value = 0
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
 Select Case Index
 Case 0:
    If Not SAVEFLAG = False Then
        If SUPERUSER = True Then
            Load FareTableFrm
            FareTableFrm.Show vbModal
        Else
            RouteAdd = True
            Load SuperLogin
            SuperLogin.Show vbModal
        End If
    Else
        Load FareTableFrm
        FareTableFrm.Show vbModal
    End If
 Case 1:
    If SUPERUSER = True Then
        Load frmRouteDelete
        frmRouteDelete.Show vbModal
    Else
        DeleteRoute = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
 Case 2:
    'Load frmCrewDetails
    'frmCrewDetails.Show vbModal
    Load Crew
    Crew.Show vbModal
Case 3:
    Load frmAddCurrency
    frmAddCurrency.Show vbModal
 End Select
    Check1(Index).Value = 0
 End If
End Sub

Private Sub Check2_Click(Index As Integer)
 
 
 '''''''''''''''
'''''''syam added
Editing = True



''''''''''''''
 
 
 
 
 If Check2(Index).Value = 0 Then Exit Sub
 Select Case Index
 Case 0:
    If SUPERUSER = True Then
        Load frmgraphedit
        frmgraphedit.Show vbModal
    Else
        GraphEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 1:
    If SUPERUSER = True Then
        Load frmRoute
        frmRoute.Show vbModal
    Else
        RouteEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 2:
    If SUPERUSER = True Then
        RouteID = ""
        Load frmStage
        frmStage.Show vbModal
    Else
        StageEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
 End Select
    Check2(Index).Value = 0
End Sub

Private Sub Check2_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
 Select Case Index
 Case 0:
    If SUPERUSER = True Then
        Load frmgraphedit
        frmgraphedit.Show vbModal
    Else
        GraphEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 1:
    If SUPERUSER = True Then
        Load frmRoute
        frmRoute.Show vbModal
    Else
        RouteEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 2:
    If SUPERUSER = True Then
        RouteID = ""
        Load frmStage
        frmStage.Show vbModal
    Else
        StageEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
 End Select
 End If
    Check2(Index).Value = 0
End Sub
Private Sub ChkBussmry_Click(Index As Integer)
If ChkBussmry(Index).Value = 0 Then Exit Sub
Select Case Index
Case 3:
    frmBus.Show vbModal
 End Select
    ChkBussmry(Index).Value = 0
End Sub

Private Sub Chkfare_Click(Index As Integer)
  If Chkfare(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 7:
    Frmfarewise.Show vbModal
 End Select
    Chkfare(Index).Value = 0
    Command1_Click (4)
End Sub

Private Sub chkFareTable_Click(Index As Integer)
 If chkFareTable(Index).Value = 0 Then Exit Sub
 Select Case Index
 Case 3:
    If SUPERUSER = True Then
        Load frmFareTableEdit
        frmFareTableEdit.Show vbModal
    Else
        FareTableEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 1:
    If SUPERUSER = True Then
        Load frmRoute
        frmRoute.Show vbModal
    Else
        RouteEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 2:
    If SUPERUSER = True Then
        RouteID = ""
        Load frmStage
        frmStage.Show vbModal
    Else
        StageEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
 End Select
    chkFareTable(Index).Value = 0
End Sub

Private Sub chkFareTable_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
 Select Case Index
 Case 0:
    If SUPERUSER = True Then
        Load frmgraphedit
        frmFareTableEdit.Show vbModal
    Else
        GraphEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 1:
    If SUPERUSER = True Then
        Load frmRoute
        frmRoute.Show vbModal
    Else
        RouteEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 2:
    If SUPERUSER = True Then
        RouteID = ""
        Load frmStage
        frmStage.Show vbModal
    Else
        StageEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
 End Select
 End If
    Check2(Index).Value = 0
End Sub



Private Sub Check3_Click(Index As Integer)
 If Check3(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 0:
    Load frmAddUser
    frmAddUser.Show vbModal
  Case 1:
    frmRemoveUser.Show vbModal
  Case 2:
    frmChangePassword.Show vbModal
  Case 3:
    frmPortSetup.Show vbModal
  Case 4:
    If SUPERUSER = True Then
        Load frmSettings
        frmSettings.Show vbModal
    Else
        Settings = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 5:
    If SUPERUSER = True Then
        Load frmPCSettings
        frmPCSettings.Show vbModal
    Else
        ''Settings = True
        ''changed  by syam
        Settings = False
        PCSettings = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
 Case 6:
        frmBmpSettings.Show vbModal
 End Select
    Check3(Index).Value = 0
End Sub

Private Sub Check3_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
 Select Case Index
  Case 0:
    Load frmAddUser
    frmAddUser.Show vbModal
  Case 1:
    frmRemoveUser.Show vbModal
  Case 2:
    frmChangePassword.Show vbModal
  Case 3:
    frmPortSetup.Show vbModal
  Case 4:
    If SUPERUSER = True Then
        Load frmSettings
        frmSettings.Show vbModal
    Else
        Settings = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
  Case 5:
 End Select
 End If
    Check3(Index).Value = 0
End Sub

Private Sub Check4_Click(Index As Integer)
 If Check4(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 0:
     
    Load Transfer
    Transfer.Show vbModal
  Case 1:
    Load frmSetDateTime
    frmSetDateTime.Show vbModal
 End Select
   Check4(Index).Value = 0
    
End Sub

Private Sub Check4_KeyPress(Index As Integer, KeyAscii As Integer)
' If KeyAscii = 13 Then
' Select Case Index
'  Case 0:
'    Load FrmTransfer
'    FrmTransfer.Show vbModal
' End Select
' End If
'    Check4(Index).Value = 0



If KeyAscii = 13 Then
 Select Case Index
  Case 0:
    Load Transfer
    Transfer.Show vbModal
  Case 1:
    Load frmSetDateTime
    frmSetDateTime.Show vbModal
 End Select
   Check4(Index).Value = 0
End If
End Sub

Private Sub Check5_Click(Index As Integer)
 If Check5(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 0:
    ChDrive App.Path
    frmReport.Show vbModal
 End Select
    Check5(Index).Value = 0
    Command1_Click (4)
End Sub

Private Sub Check5_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
 Select Case Index
  Case 0:
    ChDrive App.Path
    frmReport.Show vbModal
 End Select
 End If
    Check5(Index).Value = 0
End Sub

Private Sub Check6_Click(Index As Integer)
If Check6(Index).Value = 0 Then Command1_Click (5): Command6(5).SetFocus: Exit Sub
Select Case Index
    Case 0:
        AboutUs_Frm.Show vbModal
End Select

Check6(Index).Value = 0
End Sub

Private Sub Check6_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
 Select Case Index
  Case 0:
    AboutUs_Frm.Show vbModal
 End Select
 End If
    Check6(Index).Value = 0
    Command1_Click (5)
End Sub
'''''''''''''''''''''''''''''RNC
Private Sub ChkExpense_Click(Index As Integer)
 If ChkExpense(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 1:
    frmExpenseRpt.Show vbModal
 End Select
    ChkExpense(Index).Value = 0
    Command1_Click (4)
End Sub

Private Sub ChkExpense_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
 Select Case Index
  Case 1:
    frmExpenseRpt.Show vbModal
 End Select
 End If
    ChkExpense(Index).Value = 0
End Sub

Private Sub Chkrefund_Click(Index As Integer)
If Chkrefund(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 8:
    frmRefundRpt.Show vbModal
 End Select
    Chkrefund(Index).Value = 0
    Command1_Click (4)
End Sub

Private Sub ChkSch_Click(Index As Integer)
    
    If ChkSch(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 4:
    frmSchpID.Show vbModal
 End Select
    ChkSch(Index).Value = 0
End Sub

Private Sub ChkSchsmry_Click(Index As Integer)
 If ChkSchsmry(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 2:
    frmSchedule.Show vbModal
 End Select
    ChkSchsmry(Index).Value = 0
End Sub

Private Sub ChkSchsmry_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
 Select Case Index
  Case 2:
    frmSchedule.Show vbModal
 End Select
 End If
    ChkSchsmry(Index).Value = 0
End Sub
''''''''''''''''''''''''''''''''''''''''''''

Private Sub ChkStage_Click(Index As Integer)
   If ChkStage(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 6:
    FrmConsole.Show vbModal
 End Select
    ChkStage(Index).Value = 0
    Command1_Click (4)
End Sub

Private Sub ChkTicket_Click(Index As Integer)
  
    If ChkTicket(Index).Value = 0 Then Exit Sub
 Select Case Index
  Case 5:
    FrmTicket.Show vbModal
 End Select
    ChkTicket(Index).Value = 0
End Sub

Private Sub cmdexit3_Click()
  Unload Me
  '  End
End Sub

Private Sub cmdexit4_Click()
 Unload Me
 '   End
End Sub


Private Sub cmdexit5_Click()
 Unload Me
'    End
End Sub


Private Sub cmdext1_Click()
  Unload Me
   ' End
End Sub

Private Sub cmdLog5_Click()
    frmLogo.Show
End Sub

Private Sub cmdLogo_Click()
    frmLogo.Show
End Sub

Private Sub cmdLogo1_Click()
    frmLogo.Show
End Sub

Private Sub cmdLogo2_Click()
    frmLogo.Show
End Sub

Private Sub cmdLogo3_Click()
frmLogo.Show
End Sub

Private Sub cmdLogo4_Click()
frmLogo.Show
End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
    Dim iCnt  As Integer
    For iCnt = 0 To 5
        Picture1(iCnt).Visible = False
    Next iCnt
      Picture1(Index).Move Picture1(0).Left, Picture1(Index).Top, Picture1(Index).Width, Picture1(Index).Height
    Select Case Index
        Case 0:
            Picture1(Index).Visible = True
            Command1(Index).SetFocus
        Case 1:
            Picture1(Index).Visible = True
            Command2(Index).SetFocus
        Case 2:
            Picture1(Index).Visible = True
            Command3(Index).SetFocus
        Case 3:
            Picture1(Index).Visible = True
            Command4(Index).SetFocus
        Case 4:
            Picture1(Index).Visible = True
            Command5(Index).SetFocus
        Case 5:
            Picture1(Index).Visible = True
            Command6(Index).SetFocus

    End Select
End Sub

Private Sub Command2_Click(Index As Integer)
On Error Resume Next
DoEvents
    Dim iCnt  As Integer
    For iCnt = 0 To 5
        Picture1(iCnt).Visible = False
    Next iCnt
      Picture1(Index).Move Picture1(0).Left, Picture1(Index).Top, Picture1(Index).Width, Picture1(Index).Height
    Select Case Index
        Case 0:
            Picture1(Index).Visible = True
            Command1(Index).SetFocus
        Case 1:
            Picture1(Index).Visible = True
            Command2(Index).SetFocus
        Case 2:
            Picture1(Index).Visible = True
            Command3(Index).SetFocus
        Case 3:
            Picture1(Index).Visible = True
            Command4(Index).SetFocus
        Case 4:
            Picture1(Index).Visible = True
            Command5(Index).SetFocus
        Case 5:
            Picture1(Index).Visible = True
            Command6(Index).SetFocus
    End Select
    
End Sub

Private Sub Command3_Click(Index As Integer)
On Error Resume Next
DoEvents
    Dim iCnt  As Integer
    For iCnt = 0 To 5
        Picture1(iCnt).Visible = False
    Next iCnt
      Picture1(Index).Move Picture1(0).Left, Picture1(Index).Top, Picture1(Index).Width, Picture1(Index).Height
    
    Select Case Index
        Case 0:
            Picture1(Index).Visible = True
            Command1(Index).SetFocus
        Case 1:
            Picture1(Index).Visible = True
            Command2(Index).SetFocus
        Case 2:
            Picture1(Index).Visible = True
            Command3(Index).SetFocus
        Case 3:
            Picture1(Index).Visible = True
            Command4(Index).SetFocus
        Case 4:
            Picture1(Index).Visible = True
            Command5(Index).SetFocus
        Case 5:
            Picture1(Index).Visible = True
            Command6(Index).SetFocus
    End Select

End Sub

Private Sub Command4_Click(Index As Integer)
On Error Resume Next
    Dim iCnt  As Integer
    For iCnt = 0 To 5
        Picture1(iCnt).Visible = False
    Next iCnt
      Picture1(Index).Move Picture1(0).Left, Picture1(Index).Top, Picture1(Index).Width, Picture1(Index).Height
    
    Select Case Index
        Case 0:
            Picture1(Index).Visible = True
            Command1(Index).SetFocus
        Case 1:
            Picture1(Index).Visible = True
            Command2(Index).SetFocus
        Case 2:
            Picture1(Index).Visible = True
            Command3(Index).SetFocus
        Case 3:
            Picture1(Index).Visible = True
            Command4(Index).SetFocus
        Case 4:
            Picture1(Index).Visible = True
            Command5(Index).SetFocus
        Case 5:
            Picture1(Index).Visible = True
            Command6(Index).SetFocus
    End Select
DoEvents
End Sub

Private Sub Command5_Click(Index As Integer)
On Error Resume Next
DoEvents
    Dim iCnt  As Integer
    For iCnt = 0 To 5
        Picture1(iCnt).Visible = False
    Next iCnt
      Picture1(Index).Move Picture1(0).Left, Picture1(Index).Top, Picture1(Index).Width, Picture1(Index).Height
    
    Select Case Index
        Case 0:
            Picture1(Index).Visible = True
            Command1(Index).SetFocus
        Case 1:
            Picture1(Index).Visible = True
            Command2(Index).SetFocus
        Case 2:
            Picture1(Index).Visible = True
            Command3(Index).SetFocus
        Case 3:
            Picture1(Index).Visible = True
            Command4(Index).SetFocus
        Case 4:
            Picture1(Index).Visible = True
            Command5(Index).SetFocus
        Case 5:
            Picture1(Index).Visible = True
            Command6(Index).SetFocus
    End Select
End Sub

Private Sub Command6_Click(Index As Integer)
On Error Resume Next
    Dim iCnt  As Integer
    For iCnt = 0 To 5
        Picture1(iCnt).Visible = False
    Next iCnt
      Picture1(Index).Move Picture1(0).Left, Picture1(Index).Top, Picture1(Index).Width, Picture1(Index).Height
    Select Case Index
        Case 0:
            Picture1(Index).Visible = True
            Command1(Index).SetFocus
        Case 1:
            Picture1(Index).Visible = True
            Command2(Index).SetFocus
        Case 2:
            Picture1(Index).Visible = True
            Command3(Index).SetFocus
        Case 3:
            Picture1(Index).Visible = True
            Command4(Index).SetFocus
        Case 4:
            Picture1(Index).Visible = True
            Command5(Index).SetFocus
        Case 5:
            Picture1(Index).Visible = True
            Command6(Index).SetFocus
    End Select

End Sub

Private Sub Command7_Click()
    Unload Me
   ' End
End Sub



Private Sub Form_Activate()
    Frmflag = False
    Picture1(0).Visible = True
  '  Command1(0).SetFocus = True
    If Command1(0).Enabled = True Then
    'Command7.Enabled = False
        '''''''''DEBUG
        ''frmMainform.Command1(0).SetFocus
                                                 

    End If
  '  frmMainform.Command1(0).SetFocus
    'Command1(0).SetFocus = True
    lblDate.Top = 200
    lbltime.Top = 500
    lblDate.Left = Screen.Width - lblDate.Width
    lbltime.Left = Screen.Width - lbltime.Width
                                                                                                                                                                             
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim sRoute As String
Dim Msg As VbMsgBoxResult

   ' MsgBox "1"
    Me.Picture = LoadPicture(App.Path & "\bus.jpg")
    'MsgBox "2"
    Set m_BGnd = New CMdiBackground
   With m_BGnd
      Set .Client = Me
      .Color = vbBlue
      .Color(mdiColorBottom) = vbBlack
      .BackStyle = mdiGradient
      .GraphicPosition = mdiNone
      .AutoRefresh = True
   End With
  ' Set m_BGnd.Graphic = Image1
    Result = ShellExecute(Me.hWnd, "Open", "intermediate.exe", vbNullString, vbNullString, 0)
 
If loginform.Visible = True Then loginform.Visible = False
Timer1.Enabled = False
    FindSysDir
    CONNECTDB
       
     '   CONNECTDB1

  
    'RES.Close
Timer1.Enabled = True
    TSQL = "SELECT * FROM PCSETUP"
    Set RES = CNN.OpenRecordset(TSQL, dbOpenDynaset)
    If RES.RecordCount > 0 Then
        RES.Edit
        If RES!TRANSFER_PATH = "" Then
            RES!TRANSFER_PATH = App.Path
        End If
        If RES!TICKET_PATH = "" Then
            RES!TICKET_PATH = App.Path & "\TICKET"
        End If
        RES.Update
    Else
        RES.AddNew
        RES!TRANSFER_PATH = App.Path
        RES!TICKET_PATH = App.Path & "\TICKET"
        RES.Update
    End If
    RES.Close
    
    
    USB_FLAG = False
'    res.Close
    PC_to_PMTC_Cntr = 0
       
    Timer2.Enabled = False
    SAVEFLAG = True
    FARESAVEFLAG = True
    If SUPERUSER = False Then
        mnuAddUser.Enabled = False
        frmChangePassword.chkAdmin.Enabled = False
    End If
    RSql = "SELECT * FROM ROUTE WHERE SAVEFLAG = FALSE"
    Set TDB = DAO.OpenDatabase(App.Path & "\GBackUp.mdb")
    Set RES = TDB.OpenRecordset(RSql, dbOpenDynaset)
    If RES.RecordCount > 0 Then
        
        Msg = MsgBox("Route details not Saved " & vbCrLf & "Do you want to continue ? " & vbCrLf & "Select YES to continue or select NO to remove this details ", vbYesNoCancel)
        If Msg = vbNo Then
        
            TSQL = "DELETE * FROM ROUTE "
            TDB.Execute (TSQL)
        
            TSQL = "DELETE * FROM TMPGRAPH "
            TDB.Execute (TSQL)
        
            TSQL = "DELETE * FROM STAGE "
            TDB.Execute (TSQL)
        
            TSQL = "DELETE * FROM TMPFARE"
            TDB.Execute (TSQL)
            
            TSQL = "DELETE * FROM STATUS"
            TDB.Execute (TSQL)
            
  
        ElseIf Msg = vbYes Then
            RSql = "SELECT * FROM ROUTE"
            Dim res1 As DAO.Recordset
            Set res1 = TDB.OpenRecordset(RSql, dbOpenDynaset)
         
            If res1.RecordCount > 0 Then
                'If NORECORDFLAG = True Then
                SAVEFLAG = False
                'End If
            Else
                SAVEFLAG = True
            End If
            res1.Close
            NOSTGS = RES!nostage
            FrType = RES!FareType
            
            If RES.RecordCount > 0 Then
                sRoute = RES!RUTCODE
                'res1.Close
                TSQL = "SELECT * FROM STAGE WHERE ROUTE = '" & sRoute & "'"
                Set res1 = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                If res1.RecordCount = 0 Then
                    RemoveLanguageStage (sRoute)
                End If
            End If
            'RES.Close
            Load FareTableFrm
            FareTableFrm.Show vbModal
            Exit Sub
        ElseIf Msg = vbCancel Then
            Unload Me
            Exit Sub
        End If
        RES.Close
        TSQL = "SELECT RUTCODE FROM ROUTE"
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
        
        If RES.RecordCount > 0 Then
            sRoute = RES!RUTCODE
            RemoveLanguageStage (sRoute)
        End If
        RES.Close
    End If
    LocalLanguage = GetLocalLanguage
   
End Sub


Private Sub Form_Terminate()
Unload Me
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (MsgBox("Do you want to exit from " & App.ProductName & " ?", vbQuestion + vbYesNo, App.ProductName)) = vbYes Then
   ' Result = ShellExecute(Me.hwnd, "end", "intermediate.exe", vbNullString, vbNullString, 0)
    'Call Shell(App.Path & "/intermediate.exe /CLOSE FILENAME.EXTENSION")
'call TerminateProcess(

 EndApplication "intermediate", Me
 
 
 
 Dim frm As Form

' Loop through all open forms and close them

For Each frm In Forms

    If frm.Name = "frmThisForm" Then GoTo SkipLoop

    Unload frm

SkipLoop:

Next frm

 
 
   'Unload Me
    End
    Else
        Cancel = True
    End If
End Sub

Private Sub JeweledButton1_Click()
Unload Me
'    End
End Sub

Private Sub mnuAbout_Click()
    AboutUs_Frm.Show vbModal
End Sub

Private Sub mnuAddUser_Click()
    Load frmAddUser
    frmAddUser.Show vbModal
End Sub

Private Sub mnuCalculator_Click()
    hWnd1 = FindWindow(vbNullString, "Calculator")
    If hWnd1 < 33 Then
        If SYSDIR <> "" Then
            Fhandle = Shell(SYSDIR, vbMaximizedFocus)
        Else
            Fhandle = WinExec(WINPATH, vbMaximizedFocus)
        End If
    End If

End Sub

Private Sub mnuCrew_Click()
    'Load frmCrewDetails
    'frmCrewDetails.Show vbModal
    Load Crew
    Crew.Show vbModal
End Sub

Private Sub mnuDelete_Click()
    If SUPERUSER = True Then
        Load frmRouteDelete
        frmRouteDelete.Show vbModal
    Else
        DeleteRoute = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuGraph_Click()
    If SUPERUSER = True Then
        Load frmgraphedit
        frmgraphedit.Show vbModal
    Else
        GraphEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
End Sub

Private Sub mnulogoff_Click()
    loginsucceed = False
    loginform.Show vbModal
End Sub

Private Sub mnuNewRuote_Click()
    If Not SAVEFLAG = False Then
        If SUPERUSER = True Then
            Load FareTableFrm
            FareTableFrm.Show vbModal
        Else
            RouteAdd = True
            Load SuperLogin
            SuperLogin.Show vbModal
        End If
    Else
        Load FareTableFrm
        FareTableFrm.Show vbModal
    End If
End Sub

Private Sub mnuPasword_Click()
    frmChangePassword.Show vbModal
End Sub

Private Sub mnuPort_Click()
    frmPortSetup.Show vbModal
End Sub

Private Sub mnuRemoveUser_Click()
    frmRemoveUser.Show vbModal
End Sub

Private Sub mnuReport_Click()
    ChDrive App.Path
    frmReport.Show vbModal
End Sub

Private Sub mnuRoute_Click()
    If SUPERUSER = True Then
        Load frmRoute
        frmRoute.Show vbModal
    Else
        RouteEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If

End Sub

Private Sub mnuSettings_Click()
    If SUPERUSER = True Then
        Load frmSettings
        frmSettings.Show vbModal
    Else
        Settings = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
End Sub

Private Sub mnuStage_Click()
    If SUPERUSER = True Then
        RouteID = ""
        Load frmStage
        frmStage.Show vbModal
    Else
        StageEdit = True
        Load SuperLogin
        SuperLogin.Show vbModal
    End If
End Sub

Private Sub mnuTransfer_Click()
    Load FrmTransfer
    FrmTransfer.Show vbModal
End Sub



Private Sub Timer1_Timer()
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
Dim str As String
Dim exdesc, expname As String
 lblDate.caption = Format(Date, "DD MMM YYYY")
 lbltime.caption = Format(Time, " HH:NN:SS AM/PM")
'   If ConnectMysqlDatabase = True Then
'     Timer3.Enabled = False
' sql = "select * from ticketdetails"
' If rslogin1.State = 1 Then rslogin1.Close
'   rslogin1.Open sql, adoDBCon, adOpenDynamic, adLockOptimistic
'sql = ""
'sql = "select * from TKTS"
'   rs.Open sql, condb, adOpenDynamic, adLockOptimistic
'   Do While Not rslogin1.EOF
'   If rs1.State = 1 Then rs1.Close
'     str = "select * from TKTS where Date='" & Format(rslogin1!TDate, "dd/mm/yyyy") & "' and Time='" & rslogin1!TTime & "' and PalmId='" & rslogin1!PalmID & "' and Schdule=" & rslogin1!ScheduleNo & " and  TripNo=" & rslogin1!TripNo & " and TicketNo=" & rslogin1!TicketNo & " and Amount=" & rslogin1!Amount & ""
'      rs1.Open str, condb, adOpenDynamic, adLockOptimistic
'      If rs1.EOF = True Then
'        rs.AddNew
'        rs!Date = Format(rslogin1!TDate, "dd/mm/yyyy")
'        rs!Time = rslogin1!TTime
'        rs!PalmID = rslogin1!PalmID
'        rs!Schdule = rslogin1!ScheduleNo
'        rs!TripNo = rslogin1!TripNo
'        rs!TicketNo = rslogin1!TicketNo
'        rs!Amount = rslogin1!Amount
'        rs!Luggage = rslogin1!Luggcnt
'        rs!FromStage = rslogin1!FromStage
'        rs!ToStage = rslogin1!ToStage
'        rs!Full = rslogin1!Fullcnt
'        rs!Half = rslogin1!Halfcnt
'        rs!st = rslogin1!Stcnt
'        rs!Phy = rslogin1!phycnt
'        rs!PassNo = rslogin1!PassNo
'       rs.Update
'
'    End If
'    rs1.Close
'       rslogin1.MoveNext
' '  If rs.State = 1 Then rs.Close
'   Loop
'   rslogin1.Close
'    If rs.State = 1 Then rs.Close
'    If rs1.State = 1 Then rs1.Close
'
'    If rslogin1.State = 1 Then rslogin1.Close
'    sql = "select * from rpt"
'   rslogin1.Open sql, adoDBCon, adOpenDynamic, adLockOptimistic
'sql = ""
'sql = "select * from rpt"
'   rs.Open sql, condb, adOpenDynamic, adLockOptimistic
'   Do While Not rslogin1.EOF
'   If rs1.State = 1 Then rs1.Close
'     str = "select * from rpt where StartDate='" & Format(rslogin1!SDate, "dd/mm/yyyy") & "' and StartTime='" & rslogin1!STime & "' and PalmId='" & rslogin1!PalmID & "' and Schedule=" & rslogin1!SCHEDULE & " and  TripNo=" & rslogin1!TripNo & " and STicket='" & rslogin1!Starttkt & "' and ETicketNo='" & rslogin1!Endtkt & "' and TotalColl=" & rslogin1!NetCol & ""
'      rs1.Open str, condb, adOpenDynamic, adLockOptimistic
'      If rs1.EOF = True Then
'        rs.AddNew
'        rs!Date = Format(Now, "dd/mm/yyyy")
'        rs!STicket = rslogin1!Starttkt
'        rs!ETicketNo = rslogin1!Endtkt
'        rs!TotalColl = rslogin1!NetCol
'        rs!FullColl = rslogin1!FullColl
'        rs!HalfColl = rslogin1!HalfColl
'        rs!PhyColl = rslogin1!PhyColl
'        rs!LuggageColl = rslogin1!LugColl
'        rs!STColl = rslogin1!StuColl
'        rs!AdjustColl = rslogin1!AdjColl
'        rs!Fulls = rslogin1!Full
'        rs!Half = rslogin1!Half
'        rs!Phy = rslogin1!Phy
'        rs!Luggage = rslogin1!Lug
'        rs!st = rslogin1!stu
'        rs!Adjust = rslogin1!Adj
'        rs!pass = rslogin1!pass
'        rs!SCHEDULE = rslogin1!SCHEDULE
'        rs!TripNo = rslogin1!TripNo
'        rs!RouteCode = rslogin1!RoutNo
'        rs!StartDate = Format(rslogin1!SDate, "dd/mm/yyyy")
'        rs!StartTime = rslogin1!STime
'        rs!EndDate = Format(rslogin1!EDate, "dd/mm/yyyy")
'        rs!EndTime = rslogin1!ETime
'        rs!expense = rslogin1!Expns
'        rs!PalmID = rslogin1!PalmID
'        rs!Busno = rslogin1!Busno
'        rs!Driver = rslogin1!Driver
'        rs!Conductor = rslogin1!Conductor
'        rs!Cleaner = rslogin1!Cleaner
'        rs!NoOfMisBill = rslogin1!NoOfMisBill
'        rs!InHandAmount = rslogin1!InHandAmt
'        rs!Free = rslogin1!Free
'        rs!Conc = rslogin1!Conc
'        rs!UpDownTrip = rslogin1!UpDownTrip
'       rs.Update
'
'    End If
'     rslogin1.MoveNext
'     ' If rs.State = 1 Then rs.Close
'   Loop
'   If rs.State = 1 Then rs.Close
'   rslogin1.Close
'
'
'    If rslogin1.State = 1 Then rslogin1.Close
'    sql = "select * from expense"
'   rslogin1.Open sql, adoDBCon, adOpenDynamic, adLockOptimistic
'        Do While Not rslogin1.EOF
'
'    If (getvalueQuery1("select Count(*) from EXPENSE where PALMID = '" & rslogin1!PalmID & "' AND rcpt_No= " & rslogin1!rcpt_No & " AND SCHEDULENO = " & rslogin1!ScheduleNo & " AND EXPCODE = '" & rslogin1!expcode & "' AND DATE = '" & Format(rslogin1!ExpDate, "dd/mm/yyyy") & "'")) = 0 Then
'
'
'        If Trim(rslogin1!expname) <> "1" Then
'            expname = exdesc
'        Else
'            expname = "Diesel Entry"
'        End If
'        'sql = "insert into EXPENSE (TripMasterReferenceId,ExpCode ,ExpAmt, ExpName, Date, Time, PalmID,ScheduleNo,BusNo,DriverName,rcpt_No) values('"
'        sql = "insert into EXPENSE values('" _
'        & getvalueQuery1("Select Trip_Master_ID FROM RPT WHERE PALMID='" & rslogin1!PalmID & "' AND SCHEDULE= " & rslogin1!ScheduleNo & " AND TRIPNO= " & rslogin1!TripNo & " AND DATE = DateValue('" & Format(rslogin1!ExpDate, "dd/mm/yyyy") & "')") & "','" _
'        & rslogin1!expcode & "'," _
'        & rslogin1!ExpAmt & ",'" _
'        & expname & "','" _
'        & Format(rslogin1!ExpDate, "dd/mm/yyyy") & "','" _
'        & rslogin1!ExpTime & "','" _
'        & rslogin1!PalmID & "'," _
'        & rslogin1!ScheduleNo & ",'" _
'        & rslogin1!Busno & "','" _
'        & rslogin1!DriverName & "'," _
'        & rslogin1!rcpt_No & ")"
'        condb.Execute sql
'  End If
'  rslogin1.MoveNext
'   Loop
'
'   If rslogin1.State = 1 Then rslogin1.Close
'
'   Timer3.Enabled = True
''    MsgBox "Mysql Connection Failed. Going To Abort"
''        End
''        Exit Sub
'    End If
End Sub

'Private Sub Timer2_Timer()
'    If GetDevices = True Then
'        If USB_FLAG = True Then
'            Disconnect_USB
'        End If
'    Else
'        USB_FLAG = False
'    End If
'End Sub


Private Sub Timer3_Timer()
    lblDate.caption = Format(Date, "DD MMM YYYY")
    lbltime.caption = Format(Time, " HH:NN:SS AM/PM")
End Sub

