VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FareTableFrm 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8460
   ClientLeft      =   5040
   ClientTop       =   2040
   ClientWidth     =   13260
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   8460
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FareFrame 
      BackColor       =   &H00FFC0C0&
      Height          =   8715
      Left            =   9960
      TabIndex        =   18
      Top             =   600
      Width           =   14640
      Begin VB.Frame StgNameFrame 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Stage Name Entries"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   3825
         Left            =   240
         TabIndex        =   19
         Top             =   4725
         Width           =   4815
         Begin VB.TextBox txtlanstage 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1410
            MaxLength       =   23
            TabIndex        =   65
            Top             =   840
            Visible         =   0   'False
            Width           =   3360
         End
         Begin JeweledBut.JeweledButton cmdStgSave 
            Height          =   435
            Left            =   1770
            TabIndex        =   60
            Top             =   3180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   767
            TX              =   "Sa&ve"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "Form1.frx":849DE
            BC              =   12632256
            FC              =   0
         End
         Begin VB.ListBox stagelist 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            ItemData        =   "Form1.frx":849FA
            Left            =   45
            List            =   "Form1.frx":849FC
            TabIndex        =   56
            Top             =   3540
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtDistance 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1410
            MaxLength       =   7
            TabIndex        =   55
            Top             =   1320
            Width           =   1920
         End
         Begin VB.TextBox txtStgName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1410
            MaxLength       =   11
            TabIndex        =   20
            Top             =   360
            Width           =   3360
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Local          "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   66
            Top             =   795
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label Label19 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Distance      "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   54
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label lblnoofstage 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "entry"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1290
            TabIndex        =   33
            Top             =   1890
            Width           =   1485
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Entries :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   405
            TabIndex        =   32
            Top             =   1890
            Width           =   945
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Stage Name "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   90
            TabIndex        =   21
            Top             =   390
            Width           =   1260
         End
      End
      Begin VB.Frame innerframe 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fare List Entries"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   4035
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   13755
         Begin JeweledBut.JeweledButton cmdFareEntry 
            Height          =   435
            Left            =   12180
            TabIndex        =   61
            Top             =   3540
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   767
            TX              =   "Save Fa&re"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "Form1.frx":849FE
            BC              =   12632256
            FC              =   0
         End
         Begin VB.TextBox txtFare1 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   270
            MaxLength       =   6
            TabIndex        =   51
            Top             =   2040
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox txtFare2 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   50
            Top             =   1950
            Visible         =   0   'False
            Width           =   1320
         End
         Begin MSFlexGridLib.MSFlexGrid FareTypeGrid 
            Height          =   2970
            Left            =   165
            TabIndex        =   23
            Top             =   555
            Width           =   13395
            _ExtentX        =   23627
            _ExtentY        =   5239
            _Version        =   393216
            ForeColor       =   255
            WordWrap        =   -1  'True
            ScrollTrack     =   -1  'True
            TextStyle       =   1
            HighLight       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblEntered 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "entry"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1935
            TabIndex        =   49
            Top             =   285
            Width           =   495
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Entries :"
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
            Left            =   195
            TabIndex        =   25
            Top             =   285
            Width           =   765
         End
         Begin VB.Label lblnoofentry 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "entry"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1035
            TabIndex        =   24
            Top             =   285
            Width           =   975
         End
      End
      Begin JeweledBut.JeweledButton cmdCreate 
         Height          =   500
         Left            =   13300
         TabIndex        =   63
         Top             =   7700
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         TX              =   "&Finish"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Form1.frx":84A1A
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdQuit 
         Height          =   500
         Left            =   13300
         TabIndex        =   62
         Top             =   7000
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         TX              =   "Cance&l"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Form1.frx":84A36
         BC              =   12632256
         FC              =   0
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5220
         MaxLength       =   11
         TabIndex        =   58
         Top             =   4800
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
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
         Left            =   5190
         MaxLength       =   68
         TabIndex        =   57
         Top             =   6375
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid stageGrid 
         Height          =   3450
         Left            =   5160
         TabIndex        =   59
         Top             =   4800
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   6085
         _Version        =   393216
         Cols            =   5
         BackColor       =   16777215
         ForeColor       =   255
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483646
         ScrollTrack     =   -1  'True
         Enabled         =   -1  'True
         TextStyle       =   1
         GridLines       =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "MinFare : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7035
         TabIndex        =   48
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label lblMinFareshow 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "MinFare"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   8415
         TabIndex        =   47
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Route Name: "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3045
         TabIndex        =   46
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "RouteCode : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   405
         TabIndex        =   45
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblRouteName 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "RouteName"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4530
         TabIndex        =   44
         Top             =   270
         Width           =   2445
      End
      Begin VB.Label lblShowRoute 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Route"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   43
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label lblfrtype 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "frtype"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   11415
         TabIndex        =   27
         Top             =   285
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "FARE TYPE : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10035
         TabIndex        =   26
         Top             =   285
         Width           =   1395
      End
   End
   Begin VB.Frame RouteListFrame 
      BackColor       =   &H00FFC0C0&
      Height          =   5415
      Left            =   240
      TabIndex        =   28
      Top             =   720
      Width           =   7665
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Height          =   4035
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   3495
         Begin VB.ComboBox cmbBustype 
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
            ItemData        =   "Form1.frx":84A52
            Left            =   1560
            List            =   "Form1.frx":84A54
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   3360
            Width           =   1815
         End
         Begin VB.TextBox txtminfare 
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
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   3
            Top             =   2250
            Width           =   975
         End
         Begin VB.ComboBox cbofaretype 
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
            Height          =   375
            ItemData        =   "Form1.frx":84A56
            Left            =   1560
            List            =   "Form1.frx":84A60
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2805
            Width           =   1815
         End
         Begin VB.TextBox txtnostage 
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
            Left            =   1575
            MaxLength       =   3
            TabIndex        =   2
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtrutcode 
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
            Height          =   420
            Left            =   1590
            MaxLength       =   4
            TabIndex        =   0
            Top             =   555
            Width           =   1455
         End
         Begin VB.TextBox txtnostop 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4800
            MaxLength       =   2
            TabIndex        =   15
            Top             =   2202
            Width           =   975
         End
         Begin VB.ComboBox cbousestop 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Form1.frx":84A72
            Left            =   4800
            List            =   "Form1.frx":84A7C
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   3864
            Width           =   1335
         End
         Begin VB.TextBox txtstartfrom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4800
            TabIndex        =   17
            Top             =   4440
            Width           =   975
         End
         Begin VB.TextBox txtrutname 
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
            Left            =   1560
            MaxLength       =   16
            TabIndex        =   1
            Top             =   1125
            Width           =   1515
         End
         Begin VB.Label lblBustype 
            BackStyle       =   0  'Transparent
            Caption         =   "Bus Type"
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
            Left            =   240
            TabIndex        =   64
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fare Type"
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
            Left            =   105
            TabIndex        =   42
            Top             =   2910
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Min Fare"
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
            Left            =   -15
            TabIndex        =   41
            Top             =   2340
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Height          =   375
            Left            =   225
            TabIndex        =   40
            Top             =   1740
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Route Name"
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
            Left            =   105
            TabIndex        =   39
            Top             =   1185
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Route Code"
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
            Left            =   85
            TabIndex        =   38
            Top             =   615
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "No of Stops"
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
            Left            =   3585
            TabIndex        =   36
            Top             =   2265
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Use Stop"
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
            Left            =   3585
            TabIndex        =   35
            Top             =   3960
            Width           =   1095
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Start From"
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
            Left            =   3600
            TabIndex        =   34
            Top             =   4500
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select Allowables"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   4035
         Left            =   3840
         TabIndex        =   30
         Top             =   360
         Width           =   3495
         Begin VB.CheckBox chkselect 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkAllowPass 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Allow Pass"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   2910
            Width           =   1410
         End
         Begin VB.CheckBox chkAdjust 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Adjust"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   2160
            Width           =   1410
         End
         Begin VB.CheckBox chkConc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Concession"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   2535
            Width           =   1410
         End
         Begin VB.CheckBox chkHalf 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Half"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   1125
            Width           =   1695
         End
         Begin VB.CheckBox chkph 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "PH"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1830
            Width           =   1095
         End
         Begin VB.CheckBox chkstudent 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Student"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   3270
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkluggage 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Luggage"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   1485
            Width           =   1695
         End
      End
      Begin JeweledBut.JeweledButton cmdExit 
         Height          =   495
         Left            =   4680
         TabIndex        =   67
         Top             =   4620
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   873
         TX              =   "&Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Form1.frx":84A89
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdCancel 
         Height          =   495
         Left            =   2880
         TabIndex        =   68
         Top             =   4620
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   873
         TX              =   "&Clear"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Form1.frx":84AA5
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdsave 
         Height          =   495
         Left            =   1080
         TabIndex        =   69
         Top             =   4620
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   873
         TX              =   "&Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Form1.frx":84AC1
         BC              =   12632256
         FC              =   0
      End
   End
   Begin VB.TextBox txtTemp 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   31
      Top             =   2280
      Width           =   945
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   4635
      Pattern         =   "*.txt;*.dat"
      TabIndex        =   37
      Top             =   4095
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stage Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1920
      TabIndex        =   53
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fare Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   3225
      TabIndex        =   52
      Top             =   105
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "BUS CONVERSION UTILITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   11340
      TabIndex        =   14
      Top             =   -360
      Width           =   3240
   End
End
Attribute VB_Name = "FareTableFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TestingBuffer As String
Dim Buffer As String
Dim flxgrdi As Integer
Dim glxgrdj As Integer
Dim FilePut As Integer
Dim CretFle As Integer
Dim n As Integer
Dim MaxSize
Dim StrCNTST(255) As String
Dim StGridEdit As Boolean
Dim sgRow As Integer
Dim cn As DAO.Database
Dim rs As DAO.Recordset
Dim rrow1 As Long, CCol1 As Long
Dim tstruct As table
Dim hndl As Integer
Public R As Integer
Public c As Integer
Public NOS As Integer
Public NoOfEntries As Long
Public EntryCount As Integer
Public EnteredCount As Integer
Public First As Boolean
Public FareSavedFlag As Boolean
Public ClearFlag As Boolean
'''''''''''FAIR STAGE MODIFICATION---75
Dim I As Long
Private Type ITEM
STAGE_MAX As Long
STAGE_BUFF As Long ''DEFAULT 333
End Type
Public BMPWidth As Long
Public BMPHeight As Long
Function getSTAGEmax() As Long
On Error GoTo err:
''SYAM ADDED
''DEFAULT SET TO 75 IF THE FAIR.TXT FILE IS AVAILABLE AND ITS CONTENTS ARE OK THEN READ THE DS FILE AND SET THE CONDITION
Dim lngRecordCount As Long
Dim lngFilPointer As Long
Dim SILSTAGE As String
Dim intMASTERHandle
Dim IT As ITEM
Dim LIMIT As Long
    SILSTAGE = App.Path & "/SILSTAGE.DAT"
    intMASTERHandle = FreeFile
    If Dir$(SILSTAGE) <> "" Then
        Open SILSTAGE For Binary Access Read As #intMASTERHandle
            Get #intMASTERHandle, , IT
            LIMIT = IT.STAGE_MAX
        Close #intMASTERHandle
    Else
        LIMIT = 128 '90 'SANGEETHA 21-08-13
        'LIMIT = 75
    End If
    If (IT.STAGE_BUFF <> 333) Then
        LIMIT = 128 '90 'SANGEETHA 21-08-13
        'LIMIT = 75
    End If
    getSTAGEmax = LIMIT
Exit Function
err:
    getSTAGEmax = 90
End Function
Private Sub cbofaretype_Change()
    If cbofaretype.ListIndex = 1 And val(txtnostage) > getSTAGEmax Then
        MsgBox "Maximum Allowed stages in Graph should be less than " & getSTAGEmax + 1 & " Stages.", vbInformation
        txtnostage.SetFocus
        txtnostage.Text = getSTAGEmax
        txtnostage.SelStart = 0
        txtnostage.SelLength = Len(txtnostage)
    End If
    If UCase(cbofaretype.List(cbofaretype.List(cbofaretype.ListIndex))) = "GRAPH" And val(txtnostage.Text) < 3 Then
        MsgBox "No of stages should be greater than 2", vbInformation, gblstrPrjTitle
        txtnostage.Text = ""
        txtnostage.SetFocus
        Exit Sub
    End If
End Sub
Private Sub cbofaretype_Click()
    If cbofaretype.ListIndex = 1 And val(txtnostage) > getSTAGEmax Then
        MsgBox "Maximum Allowed stages in Graph should be less than " & getSTAGEmax + 1 & " Stages.", vbInformation
        txtnostage.SetFocus
        txtnostage.Text = getSTAGEmax
        txtnostage.SelStart = 0
        txtnostage.SelLength = Len(txtnostage)
    End If
End Sub
Private Sub cbofaretype_LostFocus()
    If cbofaretype.ListIndex = 1 And val(txtnostage) > getSTAGEmax Then
        MsgBox "Maximum Allowed stages in Graph should be less than " & getSTAGEmax + 1 & " Stages.", vbInformation
        txtnostage.SetFocus
        txtnostage.Text = getSTAGEmax
        txtnostage.SelStart = 0
        txtnostage.SelLength = Len(txtnostage)
    End If
    If UCase(cbofaretype.List(cbofaretype.ListIndex)) = "GRAPH" And val(txtnostage.Text) < 3 Then
        MsgBox "No of stages should be greater than 2 in Graph fare", vbInformation, gblstrPrjTitle
        txtnostage.Text = ""
        txtnostage.SetFocus
        Exit Sub
    End If
    If val(txtnostage.Text) < 2 Then
        MsgBox "No of stages should be greater than 2", vbInformation, gblstrPrjTitle
        txtnostage.Text = ""
        txtnostage.SetFocus
        Exit Sub
    End If
End Sub
Private Sub chkAdjust_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chkAllowPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chkConc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chkHalf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chkluggage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chkph_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chkselect_Click()
    If ChkSelect.Value = 1 Then
        chkHalf.Value = 1
        chkAdjust.Value = 1
        chkAllowPass.Value = 1
        chkluggage.Value = 1
        chkConc.Value = 1
        chkph.Value = 1
        chkstudent.Value = 1
    ElseIf ChkSelect.Value = 0 Then
        chkHalf.Value = 0
        chkAdjust.Value = 0
        chkAllowPass.Value = 0
        chkluggage.Value = 0
        chkConc.Value = 0
        chkph.Value = 0
        chkstudent.Value = 0
    End If
End Sub
Private Sub ChkSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ChkSelect.Value = 1 Then
            cmdSave.SetFocus
        Else
            SendKeys "{TAB}"
        End If
    End If
End Sub
Private Sub chkstudent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cmbBustype_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        ChkSelect.SetFocus
    End Select
End Sub
Private Sub cmdCancel_Click()
    RSql = "DELETE * FROM ROUTE WHERE RUTCODE = '" & txtrutcode & "'"
    TDB.Execute (RSql)
    RSql = "DELETE * FROM TMPGRAPH WHERE ROUTE = '" & txtrutcode & "'"
    TDB.Execute (RSql)
    RSql = "DELETE * FROM TMPFARE WHERE ROUTE = '" & txtrutcode & "'"
    TDB.Execute (RSql)
    RSql = "DELETE * FROM STAGE WHERE ROUTE = '" & txtrutcode & "'"
    TDB.Execute (RSql)
    CLEARFIELDS
    txtrutcode.Locked = False
    cmbBusType.Locked = False
    cbofaretype.Locked = False
    txtrutname.Locked = False
    txtnostage.Locked = False
    txtminfare.Locked = False
    SAVEFLAG = True
End Sub
Private Sub cmdCreate_Click()
Dim Msg As VbMsgBoxResult
Dim RouteCode As String
Dim StageId As Integer
On Error GoTo err
    RouteCode = Mid$(DatFileName, 1, InStr(1, DatFileName, ".") - 1)
    Msg = MsgBox("Going to Create Route" & vbCrLf & "Do you want to continue ?", vbYesNo)
    If Msg = vbYes Then
        RSql = "SELECT * FROM ROUTE"
        TSQL = "SELECT * FROM ROUTE WHERE RUTCODE = '" & RouteCode & "'"
        
        Set RES = cn.OpenRecordset(RSql, dbOpenDynaset)
        
        Set rs = TDB.OpenRecordset(TSQL, dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveLast
        RES.AddNew
        RES!Id = rs!Id
        RES!RUTCODE = rs!RUTCODE
        RES!rutname = rs!rutname
        RES!nostage = rs!nostage
        RES!nostop = rs!nostop
        RES!MinFare = rs!MinFare

        RES!FareType = rs!FareType
        RES!BusType = rs!BusType
        RES!UseStop = rs!UseStop
        RES!Half = rs!Half
        RES!Luggage = rs!Luggage
        RES!student = rs!student
        RES!Adjust = rs!Adjust
        RES!PASSALLOW = rs!ALLOW '04/01/2010
        RES!Conc = rs!Conc
        RES!ph = rs!ph
        RES!StartFrom = rs!StartFrom
        
        rs.Edit
        rs!SAVEFLAG = True
        SAVEFLAG = True
        rs.Update
        RES.Update
        RES.Close
        rs.Close
        
        RSql = "DELETE * FROM ROUTE WHERE RUTCODE = '" & RouteCode & "'"
        TDB.Execute (RSql)
        
        RSql = "SELECT * FROM FARE"
        
        Set RES = cn.OpenRecordset(RSql, dbOpenDynaset)
        TSQL = "SELECT * FROM TMPGRAPH WHERE ROUTE = '" & RouteCode & "'" & " ORDER BY NUMBER"
        
        Set rs = TDB.OpenRecordset(TSQL, dbOpenDynaset)
        rs.MoveFirst
        If RES.RecordCount > 0 Then RES.MoveLast
        Do While Not rs.EOF
            RES.AddNew
            RES!row = rs!row
            RES!Col = rs!Col
            RES!FARE = rs!FARE
            RES!Route = rs!Route
            RES.Update
            rs.MoveNext
        Loop
        RSql = "DELETE * FROM TMPGRAPH WHERE ROUTE = '" & RouteCode & "'"
        TDB.Execute (RSql)
        RSql = "DELETE * FROM TMPFARE WHERE ROUTE = '" & RouteCode & "'"
        TDB.Execute (RSql)
        RSql = "SELECT * FROM STAGE"
        
        Set RES = cn.OpenRecordset(RSql, dbOpenDynaset)
        RSql = "SELECT * FROM STAGE WHERE ROUTE = '" & RouteCode & "'" & "ORDER BY ID"
        
        Set rs = TDB.OpenRecordset(RSql, dbOpenDynaset)
        rs.MoveFirst
        If RES.RecordCount > 0 Then
            RES.MoveLast
            StageId = RES!Id
            StageId = StageId + 1
        Else
            StageId = 0
        End If
        Do While Not rs.EOF
            RES.AddNew
            RES!StageName = rs!StageName
            If LocalLanguage > 0 Then
                RES!STG_LOCAL_LANGUAGE = rs!STG_LOCAL_LANGUAGE  'AJS
            ElseIf LocalLanguage = 0 Then
                RES!STG_LOCAL_LANGUAGE = "20-20-20-20"
            End If
            RES!Distance = rs!Distance
            RES!Route = rs!Route
            RES!Id = StageId
            RES.Update
            rs.MoveNext
            StageId = StageId + 1
        Loop
        RSql = "DELETE * FROM STAGE WHERE ROUTE = '" & RouteCode & "'"
        TDB.Execute (RSql)
    Else
        RSql = "DELETE * FROM ROUTE WHERE RUTCODE = '" & RouteCode & "'"
        TDB.Execute (RSql)
        RSql = "DELETE * FROM TMPGRAPH WHERE ROUTE = '" & RouteCode & "'"
        TDB.Execute (RSql)
        RSql = "DELETE * FROM TMPFARE WHERE ROUTE = '" & RouteCode & "'"
        TDB.Execute (RSql)
        RSql = "DELETE * FROM STAGE WHERE ROUTE = '" & RouteCode & "'"
        TDB.Execute (RSql)
        SAVEFLAG = True
    End If
    Unload Me
Exit Sub
err:
    Select Case err.Number
    Case Else
        MsgBox "Error No : " & err.Number & vbCrLf & err.Description, vbInformation, "Route"
        Unload Me
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExit_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub
Private Sub cmdQuit_Click()
Dim I As Integer
Dim RouteCode As String
Dim Msg As VbMsgBoxResult
Dim temp As Integer
    With FareTypeGrid
        temp = .Rows
        If .Cols < 3 And .TextMatrix(1, 1) = "" Then
            FARESAVEFLAG = True
            SAVEFLAG = True
            NORECORDFLAG = True
            ClearFlag = False
        ElseIf .TextMatrix(1, 1) = "" And FrType = 2 Then
            FARESAVEFLAG = True
            SAVEFLAG = True
            NORECORDFLAG = True
            ClearFlag = False
        ElseIf FareSavedFlag = False Then
            FARESAVEFLAG = True
            SAVEFLAG = True
            NORECORDFLAG = True
            ClearFlag = False
        Else
            ClearFlag = True
        End If
        tstruct.Col = .Col
        tstruct.row = .row
        hndl = FreeFile()
        Open App.Path & "\temp.dat" For Binary Access Write As hndl
            Put #hndl, , tstruct
        Close hndl
    End With
    If ClearFlag = True Then
        Msg = MsgBox("Route not created !" & vbCrLf & _
            "YES to save" & vbCrLf & _
            "NO to remove", vbYesNoCancel)
    Else
        Msg = vbNo
    End If
    If Msg = vbNo Then
        RouteCode = Trim(lblShowRoute.caption)
        
        Set DB = DAO.OpenDatabase(App.Path & "\GBackUp.mdb", dbOpenDynaset)
        RemoveLanguageStage (RouteCode)
        
        RSql = "DELETE * FROM ROUTE WHERE rutcode='" & RouteCode & "'"
        DB.Execute (RSql)
    
        RSql = "DELETE * FROM TMPGRAPH WHERE route='" & RouteCode & "'"
        DB.Execute (RSql)
    
        RSql = "DELETE * FROM STAGE WHERE route='" & RouteCode & "'"
        DB.Execute (RSql)
    
        TSQL = "DELETE * FROM TMPFARE"
        TDB.Execute (TSQL)
        
        TSQL = "DELETE * FROM STATUS"
        TDB.Execute (TSQL)
        
        TSQL = "DELETE * FROM TMPGRAPH"
        TDB.Execute (TSQL)
        
        If Dir(App.Path & "\" & RouteCode & ".dat") <> "" Then Kill App.Path & "\" & RouteCode & ".dat"
        RSql = "SELECT ID FROM STAGE"
        
        Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
        I = 0
        If RES.RecordCount > 0 Then
            RES.MoveFirst
            Do While Not RES.EOF
                RES.Edit
                RES!Id = I
                I = I + 1
                RES.Update
                RES.MoveNext
            Loop
            RES.Close
        End If
        SAVEFLAG = True
        FARESAVEFLAG = True
        FRAMEFLAG = 0
        Unload Me
    ElseIf Msg = vbYes Then
        If NORECORDFLAG <> True Then
            SAVEFLAG = False
        End If
        FRAMEFLAG = 0
        Unload Me
    ElseIf Msg = vbCancel Then
        FRAMEFLAG = 3
        Exit Sub
    End If
End Sub
Private Sub cmdSave_Click()   'SAVE THE ROUTELIST
Dim strCnt As String
Dim sql As String
Dim RouteID As Integer
Dim rtnam As String
Dim I As Integer
Dim tcount As Integer
Dim Strstg As String
Dim Gpos As Integer
Dim tmpbuf As String
Dim Rcount As Integer
Dim CCount As Integer
Dim NSatge As Long
Dim Ctrl As Control
Dim vbres As VbMsgBoxResult
    If cbofaretype.ListIndex = 0 And val(txtnostage) > 252 Then
        MsgBox "Number of stage allowed up to only 252 stages", vbInformation, "Route"
            txtnostage = ""
            txtnostage.SetFocus
            txtnostage.SelStart = 0
            txtnostage.SelLength = Len(txtnostage)
        Exit Sub
    ElseIf val(txtnostage) < 2 Then
        txtnostage = "2"
        Exit Sub
    End If
    For Each Ctrl In Me.Controls
        If TypeOf Ctrl Is TextBox Then
            If Ctrl.Name = "txtrutcode" Or Ctrl.Name = "txtrutname" Or Ctrl.Name = "txtnostage" Or Ctrl.Name = "txtminfare" Then
                If Ctrl.Text = "" Then
                    MsgBox "Some field missing!", vbInformation, gblstrPrjTitle
                    Exit Sub
                End If
            End If
        End If
    Next
    If cbofaretype.ListIndex = 1 And val(txtnostage) > getSTAGEmax Then
        MsgBox "Maximum Allowed stages in Graph should be less than " & getSTAGEmax + 1 & " Stages.", vbInformation
        txtnostage.SetFocus
        txtnostage.Text = getSTAGEmax
        txtnostage.SelStart = 0
        txtnostage.SelLength = Len(txtnostage)
        Exit Sub
    ElseIf cbofaretype.ListIndex = 1 And val(txtnostage) > 91 Then
        If MsgBox("Fare report will not be available in Machine if number of stage is above 91." & vbCrLf & "Do you wish to continue ?", vbYesNo) = vbNo Then
            txtnostage.SetFocus
            txtnostage.Text = "91"
            txtnostage.SelStart = 0
            txtnostage.SelLength = Len(txtnostage)
            Exit Sub
        End If
    End If
    If UCase(cbofaretype.Text) = "GRAPH" And val(txtnostage.Text) < 3 Then
        MsgBox "No of stages should be greater than 2", vbInformation, gblstrPrjTitle
        txtnostage.Text = ""
        txtnostage.SetFocus
        Exit Sub
    End If
    Rutecode_new = Trim((txtrutcode.Text))
    BustypeID = getbustype(cmbBusType.Text)
    min_fare = Round(val(txtminfare), 2)
    Me.Left = 200
    Me.Top = 500
    Me.Width = 15000
    Me.Height = 9500
    FareFrame.Left = 100
    Label10.Left = (Me.Width / 2) - (Label10.Width / 2)
    cmdFareEntry.Enabled = False
    TSQL = "SELECT * FROM STATUS"
    Set TRES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
    If SAVEFLAG = False And txtnostage.Text <> "" Then
        TSQL = "SELECT * FROM ROUTE"
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
        rtnam = Trim((txtrutcode.Text))
        RES.Edit
        RES!RUTCODE = Trim((txtrutcode.Text))
        RES!rutname = txtrutname
        RES!nostage = txtnostage
        RES!MinFare = Round(val(txtminfare), 2)
        RES!BusType = getbustype(cmbBusType.Text)
        RES!BusTypeName = cmbBusType.Text
        If chkHalf.Value = 1 Then RES!Half = 1
        If chkluggage.Value = 1 Then RES!Luggage = 1
        If chkstudent.Value = 1 Then RES!student = 1
        If chkph.Value = 1 Then RES!ph = 1
        If chkConc.Value = 1 Then RES!Conc = 1
        If chkAdjust.Value = 1 Then RES!Adjust = 1
        RES!ALLOW = IIf(chkAllowPass.Value = 1, 1, 0) ' 04/01/2010
        DatFileName = rtnam & ".DAT"
        RES.Update
        RES.Close
        TSQL = "SELECT * FROM TMPGRAPH"
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
        If RES.RecordCount = 0 Then
            RES.Close
            If TRES.RecordCount <> 0 Then
                If TRES!Save = False Then
                    FARESAVEFLAG = False
                    TSQL = "SELECT * FROM TMPFARE"
                    Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                    RES.MoveFirst
                    Do While Not RES.EOF
                        With FareTypeGrid
                            .TextMatrix(CLng(RES!row), CLng(RES!Col)) = Round((RES!FARE), 2)
                            RES.MoveNext
                            If Not RES.EOF Then
                                If .Cols < CLng(RES!Col) + 1 Then
                                    .Cols = CLng(RES!Col) + 1
                                End If
                                If .Rows < CLng(RES!row) + 1 Then .Rows = CLng(RES!row) + 1
                            End If
                        End With
                    Loop
                    RES.MoveLast
                    EnteredCount = TRES!NoOfEntries
                    If FrType = 2 Then
                        NoOfEntries = (NOSTGS * (NOSTGS - 1)) / 2
                    ElseIf FrType = 1 Then
                        NoOfEntries = NOSTGS
                    End If
                    EntryCount = EnteredCount
                    If FrType = 2 Then
                        If EnteredCount <= NoOfEntries Then EnteredCount = EnteredCount + 1
                    End If
                    If RES!row < NOSTGS - 1 Then
                        Rcount = RES!row
                        CCount = RES!Col
                        If SAVEFLAG = False Then
                            hndl = FreeFile()
                            Open App.Path & "\temp.dat" For Binary Access Read As hndl
                            Get #hndl, , tstruct
                            If tstruct.Col > 1 Then
                                NOS = (NOSTGS - Rcount) + 1
                            Else
                                NOS = (NOSTGS - Rcount)
                            End If
                        Else
                            NOS = (NOSTGS - Rcount)
                        End If
                        If FrType = 2 Then
                            If CCount = NOSTGS - Rcount Then
                                FareTypeGrid.Rows = FareTypeGrid.Rows + 1
                                FareTypeGrid.row = Rcount + 1
                                R = Rcount + 1
                                FareTypeGrid.Col = 1
                                c = 1
                            Else
                                FareTypeGrid.row = Rcount
                                R = Rcount
                                If FareTypeGrid.Cols <= NOSTGS - 1 Then FareTypeGrid.Cols = FareTypeGrid.Cols + 1
                                FareTypeGrid.Col = CCount + 1
                                c = CCount + 1
                            End If
                            txtFare2.Left = FareTypeGrid.CellLeft + FareTypeGrid.Left
                            txtFare2.Top = FareTypeGrid.CellTop + FareTypeGrid.Top
                        ElseIf FrType = 1 Then
                            If CCount < NOSTGS - 1 Then
                                FareTypeGrid.Cols = FareTypeGrid.Cols + 1
                                c = CCount + 1
                                FareTypeGrid.Col = c
                                txtFare2.Left = FareTypeGrid.CellLeft + FareTypeGrid.Left
                                txtFare2.Top = FareTypeGrid.CellTop + FareTypeGrid.Top
                            Else
                            End If
                        End If
                    End If
                End If
            End If
        Else
            FARESAVEFLAG = True
            If FrType = 1 Then
                Do While Not RES.EOF
                    With FareTypeGrid
                        .TextMatrix(CLng(RES!row), CLng(RES!Col)) = Round(RES!FARE, 2)
                        RES.MoveNext
                        If Not RES.EOF Then
                            If .Cols < CLng(RES!Col) + 1 Then
                                .Cols = CLng(RES!Col) + 1
                            End If
                            .Rows = CLng(RES!row) + 1
                        End If
                    End With
                Loop
            ElseIf FrType = 2 Then
                rrow1 = 1
                CCol1 = 1
                Do While Not RES.EOF
                    With FareTypeGrid
                        .TextMatrix(CLng(RES!Col), CLng(RES!row)) = Round(RES!FARE, 2)
                        RES.MoveNext
                        If Not RES.EOF Then
                            If .Cols < CLng(RES!row) + 1 Then
                                .Cols = CLng(RES!row) + 1
                            ElseIf .Rows < CLng(RES!Col) + 1 Then
                                .Rows = CLng(RES!Col) + 1
                            End If
                        End If
                    End With
                Loop
            End If
        End If
        FareFrame.Visible = True
        If FARESAVEFLAG = True Then
            If FrType = 2 Then
                cmdFareEntry.Enabled = True
                lblEntered = (NOSTGS * (NOSTGS - 1)) / 2
            ElseIf FrType = 1 Then
                cmdFareEntry.Enabled = True
                lblEntered = NOSTGS
            End If
        Else
            If FrType = 2 Then
                If TRES!NoOfEntries < NoOfEntries Then
                    txtFare2.Visible = True
                    txtFare2.SetFocus
                    cmdFareEntry.Enabled = False
                    lblEntered = EnteredCount - 1
                Else
                    cmdFareEntry.Enabled = True
                    lblEntered = EnteredCount - 1
                End If
            ElseIf FrType = 1 Then  '6 stgs load
                If TRES!NoOfEntries < NOSTGS - 1 Then
                    txtFare2.Visible = True
                    txtFare2.SetFocus
                    cmdFareEntry.Enabled = False
                    lblEntered = EnteredCount - 1
                Else
                    cmdFareEntry.Enabled = True
                    lblEntered = EnteredCount - 1
                End If
            End If
        End If
        lblShowRoute.caption = txtrutcode.Text
        lblRouteName.caption = txtrutname.Text
        GetLocalLanguage
        If LocalLanguage > 0 Then 'LANG
            Label20.Visible = True
            txtlanstage.Visible = True
            If LocalLanguage = 1 Then txtlanstage.FontName = "senthamil"
            If LocalLanguage = 2 Then txtlanstage.FontName = "ML-TTKarthika"
        Else
            Label20.Visible = False
            txtlanstage.Visible = False
        End If
        lblMinFareshow.caption = Round(txtminfare.Text, 2)
        FAREFRAMEFUNC
        FRAMEFLAG = 2
        Exit Sub
    End If
    sql = "select count(*) from route"
    Set rs = cn.OpenRecordset(sql, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
        RouteID = val(rs.Fields(0)) + 1
    Else
        RouteID = 1
    End If
    rs.Close
    Set rs = TDB.OpenRecordset("SELECT * FROM ROUTE", dbOpenDynaset)
    If txtrutcode.Text <> "" And txtrutname.Text <> "" And txtnostage.Text <> "" And txtminfare.Text <> "" And cbofaretype.Text <> "" And txtstartfrom.Text <> "" Then
        If txtnostage < 2 And cbofaretype.ListIndex <> 0 Then
            MsgBox "Number of stage Must be Greater than 1", vbInformation, "BusTrans"
                txtnostage.SetFocus
                txtnostage.SelStart = 0
                txtnostage.SelLength = Len(txtnostage)
                txtnostage.SetFocus
            Exit Sub
        End If
        If FilePut <> -1 Then FilePut = -1
        rtnam = Trim((txtrutcode.Text))
        txtnostop.Text = txtnostage.Text
        NOSTGS = val(txtnostage.Text)
        TEST = NOSTGS
        NOS = NOSTGS
        NSatge = NOS
        NoOfEntries = (NSatge * (NSatge - 1)) / 2
        EntryCount = 1
        rs.AddNew
        rs!Id = RouteID
        rs!RUTCODE = rtnam
        rs!rutname = txtrutname.Text
        rs!nostage = txtnostage.Text
        rs!nostop = txtnostop.Text
        rs!MinFare = Round(val(txtminfare.Text), 2)
        If cbofaretype.Text = "TABLE" Then
            txtFare1.Text = ""
            FareTypeGrid.Cols = 2  'to set the rows to flexgrid
            FrType = 1
        ElseIf cbofaretype.Text = "GRAPH" Then
            FareTypeGrid.Cols = 2  'to set the rows to flexgrid
            FrType = 2
        End If
        rs!FareType = FrType
        rs!BusType = getbustype(cmbBusType.Text)
        rs!UseStop = 0
        rs!Half = chkHalf.Value
        rs!Luggage = chkluggage.Value
        rs!student = chkstudent.Value
        rs!ph = chkph.Value
        rs!Conc = chkConc.Value
        rs!Adjust = chkAdjust.Value
        rs!ALLOW = IIf(chkAllowPass.Value = 1, 1, 0) '04/01/2009
        rs!StartFrom = 0
        If strCnt = "" Then strCnt = "0"
            tcount = CInt(strCnt)
        For I = 0 To txtnostage - 1
            StrCNTST(I) = tcount
            tcount = tcount + 1
        Next
        For I = 0 To txtnostage - 1
            If I <> txtnostage - 1 Then
                Strstg = Strstg + StrCNTST(I) + ","
            ElseIf I = txtnostage - 1 Then
                Strstg = Strstg + StrCNTST(I)
            End If
        Next
        Gpos = InStr(Strstg, ",")
        If Gpos > 0 Then
            tmpbuf = Mid(Strstg, 1, Gpos - 1) & "-"
        End If
        Gpos = 0
        Gpos = InStrRev(Strstg, ",")
        If Gpos > 0 Then
            tmpbuf = tmpbuf & Mid(Strstg, Gpos + 1, Len(Strstg) - Gpos)
        End If
        rs!StageCount = tmpbuf 'Strstg
        rs.Update
        rs.Close
        If UCase(cbofaretype.Text) = "GRAPH" Then
            lblShowRoute.caption = txtrutcode.Text
            lblRouteName.caption = txtrutname.Text
            lblMinFareshow.caption = Round(txtminfare.Text, 2)
            lblfrtype.caption = "GRAPH"
            vbres = MsgBox("Do you want to import?", vbYesNo, gblstrPrjTitle)
            If vbres = vbYes Then
                lblShowRoute.caption = txtrutcode.Text
                lblRouteName.caption = txtrutname.Text
                lblMinFareshow.caption = Round(txtminfare.Text, 2)
                lblfrtype.caption = "GRAPH"
                
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
                Set rs = cn.OpenRecordset("SELECT * FROM ROUTE", dbOpenDynaset)
                If txtrutcode.Text <> "" And txtrutname.Text <> "" And txtnostage.Text <> "" And txtminfare.Text <> "" And cbofaretype.Text <> "" And txtstartfrom.Text <> "" Then
                    If txtnostage < 2 And cbofaretype.ListIndex <> 0 Then
                        MsgBox "Number of stage Must be Greater than 1", vbInformation, "BusTrans"
                        txtnostage.SetFocus
                        txtnostage.SelStart = 0
                        txtnostage.SelLength = Len(txtnostage)
                        txtnostage.SetFocus
                        Exit Sub
                    End If
                    rs.AddNew
                    rs!Id = RouteID
                    rs!RUTCODE = rtnam
                    rs!rutname = txtrutname.Text
                    rs!nostage = txtnostage.Text
                    rs!nostop = txtnostop.Text
                    rs!MinFare = Round(val(txtminfare.Text), 2)
                If cbofaretype.Text = "TABLE" Then
                    FrType = 1
                ElseIf cbofaretype.Text = "GRAPH" Then
                    FrType = 2
                End If
                rs!FareType = FrType
                rs!BusType = getbustype(cmbBusType.Text)
                rs!UseStop = 0
                rs!Half = chkHalf.Value
                rs!Luggage = chkluggage.Value
                rs!student = chkstudent.Value
                rs!ph = chkph.Value
                rs!Conc = chkConc.Value
                rs!Adjust = chkAdjust.Value
                rs!PASSALLOW = IIf(chkAllowPass.Value = 1, 1, 0) '04/01/2009
                rs!StartFrom = 0
                If strCnt = "" Then strCnt = "0"
                    tcount = CInt(strCnt)
                    For I = 0 To txtnostage - 1
                        StrCNTST(I) = tcount
                        tcount = tcount + 1
                    Next
                    For I = 0 To txtnostage - 1
                        If I <> txtnostage - 1 Then
                            Strstg = Strstg + StrCNTST(I) + ","
                        ElseIf I = txtnostage - 1 Then
                            Strstg = Strstg + StrCNTST(I)
                        End If
                    Next
                    Gpos = InStr(Strstg, ",")
                    If Gpos > 0 Then
                        tmpbuf = Mid(Strstg, 1, Gpos - 1) & "-"
                    End If
                    Gpos = 0
                    Gpos = InStrRev(Strstg, ",")
                    If Gpos > 0 Then
                        tmpbuf = tmpbuf & Mid(Strstg, Gpos + 1, Len(Strstg) - Gpos)
                    End If
                    rs!StageCount = tmpbuf 'Strstg
                    rs.Update
                    rs.Close
                    Unload Me
                    Frm_importexport.Show
                    Exit Sub
                End If
            End If
        End If
        Buffer = ""
        DatFileName = Trim(txtrutcode.Text) & ".TXT"
        FareFrame.Visible = True
        lblShowRoute.caption = txtrutcode.Text
        lblRouteName.caption = txtrutname.Text
        lblMinFareshow.caption = Round(txtminfare.Text, 2)
        FAREFRAMEFUNC
        FARELISTENTRY
        TSQL = "SELECT * FROM TMPFARE"
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
    Else
        MsgBox "Some Fields missing!", vbInformation, "Bus Info.."
        txtrutcode.SetFocus
    End If
    FRAMEFLAG = 2
End Sub
Private Sub cmdFareEntry_Click() 'Saving from Flexgrid
On Error GoTo err
Dim tabletypeNo As Integer
Dim tabFlag As Integer
Dim flxgrdj As Integer
Dim FareGrid As Long
Dim Result As VbMsgBoxResult
Dim FARE As Single
Dim R1 As Integer
Dim c1 As Integer
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
    If Not SAVEFLAG = False Or FARESAVEFLAG = False Then
        Result = MsgBox("Do You want to Save Fare ?" & vbCrLf & vbCrLf & _
            "NO to edit fare " & vbCrLf & _
            "YES to continue Stage Entry ", vbYesNo)
        If Result = vbNo Then Exit Sub
        If MinimumFareCheck = False Then Exit Sub
        FareTypeGrid.Enabled = False
        txtTemp.Visible = False
        tabletypeNo = 1
        tabFlag = 0
        buff = ""
        
        Set rs = TDB.OpenRecordset("TMPGRAPH", dbOpenDynaset)
        If FrType = 2 Then          ' for GRAPH
            RES.MoveFirst
            Do While Not RES.EOF
                rs.AddNew
                rs!row = RES!Col    '''vaisakh 30.03.11
                rs!Col = RES!row
                FARE = RES!FARE
                rs!FARE = FARE
                rs!Route = lblShowRoute.caption  'Format(lblShowRoute, "000")
                rs.Update
                RES.MoveNext
            Loop
            rs.Close
            RES.Close
            TSQL = "DELETE * FROM TMPFARE"
            TDB.Execute (TSQL)
            TSQL = "DELETE * FROM STATUS"
            TDB.Execute (TSQL)
        ElseIf FrType = 1 Then      'for TABLE
            RES.MoveFirst
            Do While Not RES.EOF
                rs.AddNew
                rs!row = RES!row
                rs!Col = RES!Col
                FARE = RES!FARE
                rs!FARE = FARE
                rs!Route = Mid$(DatFileName, 1, InStr(1, DatFileName, ".") - 1)
                rs.Update
                RES.MoveNext
            Loop
            TSQL = "DELETE * FROM TMPFARE"
            TDB.Execute (TSQL)
            TSQL = "DELETE * FROM STATUS"
            TDB.Execute (TSQL)
            FARESAVEFLAG = True
        End If
        GetLocalLanguage
        txtStgName.Enabled = True
        If LocalLanguage > 0 Then 'LANG
            Label20.Visible = True
            txtlanstage.Enabled = False
            txtlanstage.Visible = True
            If LocalLanguage = 1 Then txtlanstage.FontName = "senthamil"
            If LocalLanguage = 2 Then txtlanstage.FontName = "ML-TTKarthika"
        Else
            Label20.Visible = False
            txtlanstage.Visible = False
        End If
        txtStgName.SetFocus
        txtStgName.BackColor = &HC0EF00
    Else
        If MinimumFareCheck = False Then Exit Sub
        TSQL = "SELECT * FROM TMPGRAPH WHERE ROUTE = '" & txtrutcode & "'"
        Set rs = TDB.OpenRecordset(TSQL, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MsgBox "No Records Found!", vbInformation
            Exit Sub
        End If
        rs.MoveFirst
        c1 = NOSTGS - 1
        j = 1
        k = 1
        l = 1
        With FareTypeGrid
            If FrType = 2 Then
                For I = 1 To NOSTGS - 1
                    For k = j To c1
                        rs.Edit
                        rs!row = I
                        rs!Col = j
                        rs!FARE = (.TextMatrix(j, I)) '(j,i)
                        rs.Update
                        rs.MoveNext
                        j = j + 1
                    Next
                    j = l + 1
                    l = j
                Next
            ElseIf FrType = 1 Then
                I = 1
                j = 1
                Do While Not rs.EOF
                    rs.Edit
                    rs!row = I
                    rs!Col = j
                    rs!FARE = (.TextMatrix(I, j))
                    rs.Update
                    rs.MoveNext
                    j = j + 1
                Loop
            End If
        End With
        FareTypeGrid.Enabled = False
        rs.Close
        I = 1
        TSQL = "SELECT * FROM STAGE WHERE ROUTE = '" & txtrutcode & "'"
        Set rs = TDB.OpenRecordset(TSQL, dbOpenDynaset)
        If rs.RecordCount > 0 Then
            rs.MoveLast
            rs.MoveFirst
            For I = 1 To rs.RecordCount
                stageGrid.TextMatrix(I, 1) = rs!StageName
                stageGrid.TextMatrix(I, 2) = rs!Distance
                If LocalLanguage > 0 Then
                    stageGrid.TextMatrix(I, 3) = rs!STG_LOCAL_LANGUAGE
                End If
                rs.MoveNext
                If Not rs.EOF Then stageGrid.Rows = stageGrid.Rows + 1
            Next
            cmdStgSave.Enabled = True
            txtStgName.Enabled = False
        Else
            If LocalLanguage = 1 Then 'LANG
                txtlanstage.Font = "senthamil Plain"
            ElseIf LocalLanguage = 2 Then
                txtlanstage.Font = "ML-TTKarthika"
            End If
            txtStgName.Enabled = True
            txtStgName.SetFocus
            txtStgName.BackColor = &HC0EF00
        End If
    End If
    FareSavedFlag = True
    cmdFareEntry.Enabled = False
Exit Sub
err:
    MsgBox "Error !" & vbCrLf & "Err No: " & err.Number & vbCrLf & err.Description, vbCritical
End Sub
Private Sub cmdStgSave_Click()    ''14/01/2011
On Error Resume Next
Dim I As Integer
Dim StageCount As Integer
Dim Stage(255) As STAGEDETAILS
    If Not SAVEFLAG = False Then
        StageNameBuff = ""
        For I = 0 To NOSTGS - 1
            StageNameBuff = StageNameBuff & StgName(I)
            Stage(I).StageName = StgName(I)
            Stage(I).Distance = val(StageDistance(I))
        Next I
        TSQL = "SELECT ID FROM STAGE"
        Set rs = cn.OpenRecordset(TSQL, dbOpenDynaset)
        If rs.RecordCount > 0 Then
            rs.MoveLast
            StageCount = rs!Id
            StageCount = StageCount + 1
            rs.Close
        Else
            StageCount = 0
        End If
        Set rs = TDB.OpenRecordset("STAGE", dbOpenDynaset)
        For I = 0 To NOSTGS - 1
            If rs.RecordCount > 0 Then rs.MoveLast
            rs.AddNew
            rs!StageName = ParseTreeData(StgName(I))
            If LocalLanguage > 0 Then
                rs!STG_LOCAL_LANGUAGE = stageGrid.TextMatrix(I + 1, 3) 'AJS '14/01/2011
            End If
            rs!Distance = val(StageDistance(I))
            rs!Route = Mid$(DatFileName, 1, InStr(1, DatFileName, ".") - 1)
            rs!Id = StageCount
            rs.Update
            StageCount = StageCount + 1
        Next I
    Else
        I = 1
        TSQL = "SELECT * FROM STAGE WHERE ROUTE = '" & txtrutcode & "'"
        Set rs = TDB.OpenRecordset(TSQL, dbOpenDynaset)
        If rs.RecordCount > 0 Then
            rs.MoveLast
            rs.MoveFirst
            For I = 1 To rs.RecordCount
                rs.Edit
                rs!StageName = stageGrid.TextMatrix(I, 1)
                If LocalLanguage > 0 Then
                    rs!STG_LOCAL_LANGUAGE = stageGrid.TextMatrix(I, 3)  '14/01/2011
                End If
                rs!Distance = val(stageGrid.TextMatrix(I, 2))
                rs!Route = txtrutcode
                rs!Id = I - 1
                rs.Update
                rs.MoveNext
            Next
        Else
            rs.Close
            TSQL = "SELECT * FROM STAGE"
            Set rs = TDB.OpenRecordset(TSQL, dbOpenDynaset)
            For I = 1 To NOSTGS
                rs.AddNew
                rs!StageName = stageGrid.TextMatrix(I, 1)
                If LocalLanguage > 0 Then
                    rs!STG_LOCAL_LANGUAGE = stageGrid.TextMatrix(I, 3)  '14/01/2011
                End If
                rs!Distance = val(StageDistance(I))
                rs!Route = txtrutcode
                rs!Id = I - 1
                rs.Update
            Next
        End If
    End If
    rs.Close
    Text1.Visible = False
    cmdStgSave.Enabled = False
    cmdCreate.Enabled = True
    cmdCreate.SetFocus
End Sub
Private Sub FareTypeGrid_Click()
Dim MaxRow As Integer
Dim MaxCol As Integer
Dim CRow As Integer
Dim CCol As Integer
Dim ACol As Integer
    txtFare1.Visible = True
    MaxRow = NOSTGS - 1
    MaxCol = NOSTGS - 1
    With FareTypeGrid
        CRow = .row
        CCol = .Col
        ACol = NOSTGS - CRow
        If FrType = 2 Then
            If txtFare1.Visible = True Then
                txtFare1 = ""
                txtFare1.Visible = False
            End If
            If CCol > CRow Then Exit Sub
        End If
        If FrType = 1 Then
            If .Col = 1 Then Exit Sub
        End If
        txtFare1.Visible = True
        txtFare1.Top = .CellTop + .Top
        txtFare1.Left = .CellLeft + .Left
        txtFare1.Width = .CellWidth
        txtFare1.Height = .CellHeight
        txtFare1 = Round(val(.TextMatrix(.row, .Col)), 2)
        txtFare1.SelStart = 0
        txtFare1.SelLength = Len(txtFare1)
        txtFare1.SetFocus
        txtFare2.Visible = False
        txtFare2.Width = .CellWidth
        txtFare2.Height = .CellHeight
        txtFare2 = ""
    End With
End Sub
Private Sub Form_Load()
    chkstudent.Value = 0
    File1.Path = App.Path
    MOVE_ALL_FILES
    TestingBuffer = ""
    flxgrdCnt = 1
    COUNTER = 1
    StgNameCnt = 1
    toListBox = 0
    n = 1
    R = 1
    c = 1
    First = False
    EnteredCount = 1
    FARESAVEFLAG = True
    FRAMEFLAG = 1
    cmdCreate.Enabled = False
    txtStgName.Enabled = False
    cmdStgSave.Enabled = False
    txtDistance.Enabled = False
    
    txtnostop.Enabled = False
    cbousestop.Enabled = False
    txtstartfrom.Text = "0"
    cmdFareEntry.Enabled = False
    txtTemp.Visible = False
    buff = ""
    StageNameBuff = ""
    If LocalLanguage > 0 Then 'LANG
        Label20.Visible = True
        txtlanstage.Visible = True
        If LocalLanguage = 1 Then txtlanstage.FontName = "senthamil"
        If LocalLanguage = 2 Then txtlanstage.FontName = "ML-TTKarthika"
    Else
        Label20.Visible = False
        txtlanstage.Visible = False
    End If
    Dim BTYPE As DAO.Database
    Dim BTYPEREC As DAO.Recordset
    Set BTYPE = DAO.OpenDatabase(App.Path & "\pvt.mdb")
    TSQL = "SELECT * FROM BUSTYPE"
    Set BTYPEREC = BTYPE.OpenRecordset(TSQL, dbOpenDynaset)
    If BTYPEREC.RecordCount > 0 Then
        BTYPEREC.MoveFirst
        cmbBusType.Clear
        Do While Not BTYPEREC.EOF
            cmbBusType.AddItem (BTYPEREC!Name)
            BTYPEREC.MoveNext
        Loop
    Else
        MsgBox "Please create Bustype", vbInformation
        frmMainform.Check1(0).Value = vbUnchecked
        Unload Me
        Exit Sub
    End If
    Set cn = DAO.OpenDatabase(App.Path & "\PVT.Mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    Set TDB = DAO.OpenDatabase(App.Path & "\GBackUp.mdb")
    TSQL = "SELECT * FROM ROUTE"
    If SAVEFLAG = False Then
        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
        If RES.RecordCount > 0 Then
            txtrutcode = RES!RUTCODE
            txtrutname = RES!rutname
            txtnostage = RES!nostage
            txtminfare = (RES!MinFare)
            If RES!FareType = 2 Then
                cbofaretype.Text = cbofaretype.List(1)
                cbofaretype.Locked = True
            ElseIf RES!FareType = 1 Then
                cbofaretype.Text = cbofaretype.List(0)
                cbofaretype.Locked = True
            End If
            cmbBusType.Text = getbustypecmbid(RES!BusType)
            cmbBusType.Locked = True
            If RES!Half = 1 Then chkHalf.Value = 1
            If RES!Luggage = 1 Then chkluggage.Value = 1
            If RES!student = 1 Then chkstudent.Value = 1
            If RES!Conc = 1 Then chkConc.Value = 1
            If RES!ph = 1 Then chkph.Value = 1
            If RES!Adjust = 1 Then chkAdjust.Value = 1
            chkAllowPass = IIf(RES!ALLOW = 1, 1, 0) ' 04/01/2010
            cmdSave.Enabled = True
            txtrutcode.Locked = True
            txtrutname.Locked = True
            txtnostage.Locked = True
            txtminfare.Locked = True
            RES.Close
        End If
    Else
        cbofaretype.Text = cbofaretype.List(0)
        cmbBusType.Text = cmbBusType.List(0)
    End If
    Me.Width = 8500
    Me.Height = 6500
    RouteListFrame.Left = (Me.Width / 2) - (RouteListFrame.Width / 2)
    Label10.Left = (Me.Width / 2) - (Label10.Width / 2)
    With stageGrid  'To fix the size of the Columns initially
        .TextMatrix(0, 0) = "S.NO"
        .ColWidth(0) = 600
        .TextMatrix(0, 1) = "STAGE NAME"
        .ColWidth(1) = 2500
        .TextMatrix(0, 2) = "KM"
        .ColWidth(2) = 1200
        .TextMatrix(0, 3) = " STG"
        .ColWidth(3) = 4600
        .ColWidth(4) = 0
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload FareTableFrm
End Sub
Private Sub stageGrid_Click()
    If Not FareSavedFlag = True Then Exit Sub
    If cmdCreate.Visible = True And cmdStgSave.Visible = False Then
    Else
        ShowCols
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        If Text1.Text <> "" Then
            If stageGrid.Col = 1 Then
                StgName(stageGrid.row - 1) = UCase(Trim(Text1.Text)) & ","
                stageGrid.TextMatrix(stageGrid.row, 1) = UCase(Trim(Text1.Text))
            ElseIf stageGrid.Col = 2 Then
                stageGrid.TextMatrix(stageGrid.row, stageGrid.Col) = val(Text1.Text)
            ElseIf stageGrid.Col = 3 Then
                If LocalLanguage = 1 Then Text1.FontName = "senthamil"  'LANG
                If LocalLanguage = 2 Then Text1.FontName = "ML-TTKarthika"
                stageGrid.TextMatrix(stageGrid.row, stageGrid.Col) = Text1.Text
                If LocalLanguage = 1 Then stageGrid.CellFontName = "senthamil"
                If LocalLanguage = 2 Then stageGrid.CellFontName = "ML-TTKarthika"
            End If
            Text1.Visible = False
            If Not SAVEFLAG = False Then
                If (txtStgName.Enabled = True) Then
                    txtStgName.SetFocus
                    txtStgName.BackColor = &HC0EF00
                Else
                    If cmdCreate.Visible = False Then
                        cmdStgSave.SetFocus
                    Else
                        If cmdStgSave.Enabled = False Then
                            txtStgName.Enabled = True
                            txtStgName.SetFocus
                            txtStgName.BackColor = &HC0EF00
                        End If
                    End If
                End If
            Else
                If txtStgName.Enabled = True Then
                    txtStgName.SetFocus
                    txtStgName.BackColor = &HC0EF00
                Else
                    If cmdStgSave.Enabled = True Then
                        cmdStgSave.SetFocus
                    End If
                End If
            End If
            Exit Sub
        End If
        MsgBox "Enter Stage Name ", vbCritical, "BUSTrans"
        Text1.SetFocus
    End Select
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        If Text2.Text <> "" Then
            If stageGrid.Col = 3 Then
                stageGrid.TextMatrix(stageGrid.row, stageGrid.Col) = Text2.Text
            End If
            Text2.Visible = False
            If Not SAVEFLAG = False Then
                If (txtStgName.Enabled = True) Then
                    txtStgName.SetFocus
                    txtStgName.BackColor = &HC0EF00
                Else
                    If cmdCreate.Visible = False Then
                        cmdStgSave.SetFocus
                    Else
                        If cmdStgSave.Enabled = False Then
                            txtStgName.Enabled = True
                            txtStgName.SetFocus
                            txtStgName.BackColor = &HC0EF00
                        End If
                    End If
                End If
            Else
                If txtStgName.Enabled = True Then
                    txtStgName.SetFocus
                    txtStgName.BackColor = &HC0EF00
                Else
                    If cmdStgSave.Enabled = True Then
                        cmdStgSave.SetFocus
                    End If
                End If
            End If
            Exit Sub
        End If
        MsgBox "Enter Stage Name ", vbCritical, "BUS"
        Text2.SetFocus
    End Select
End Sub
Private Sub txtDistance_KeyPress(KeyAscii As Integer)
    Static LastText As String
    Static SecondTime As Boolean
    Const MaxDecimal As Integer = 1
    Const MaxWhole As Integer = 4
    If TextBoxValidityNumeric(KeyAscii) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    With txtDistance
        If Not SecondTime Then
            If .Text Like "*[!0-9.]*" _
                Or .Text Like "*.*.*" _
                Or .Text Like "*." & String$(1 + MaxDecimal, "#") _
                Or .Text Like String$(MaxWhole, "#") & "[!.]" Then
                SecondTime = True
                .SelStart = Len(.Text)
            Else
                LastText = .Text
            End If
        End If
    End With
    SecondTime = False
    Select Case KeyAscii
    Case 48 To 57, 8, 46
        KeyAscii = KeyAscii
    Case 13
        stageGrid.TextMatrix(StgNameCnt - 1, 2) = val(txtDistance.Text)
        StageDistance(StgNameCnt - 2) = val(txtDistance)
        txtDistance = ""
        If StgNameCnt <= val(txtnostage) Then
            txtStgName.Enabled = True
            txtStgName.SetFocus
            txtStgName.BackColor = &HC0EF00
        End If
        txtDistance.Enabled = False
        lblnoofstage.caption = StgNameCnt - 1 & "/" & NOSTGS
        If StgNameCnt > NOSTGS Then
            cmdStgSave.Enabled = True
            cmdStgSave.SetFocus
        End If
        stageGrid.CellBackColor = &H80000005
        If LocalLanguage > 0 Then
            stageGrid.Col = 2
            stageGrid.CellBackColor = &H80000005
        End If
        If StgNameCnt <= NOSTGS Then
            stageGrid.row = StgNameCnt
            stageGrid.Col = 1
            stageGrid.CellBackColor = &H900FFF
        End If
        txtDistance.BackColor = &H80000005
        stageGrid.row = stageGrid.Rows - 1
        stageGrid.TopRow = stageGrid.Rows - 1
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtFare1_LostFocus()
    txtFare1 = ""
    txtFare1.Visible = False 'SANGEETHA
End Sub

Private Sub txtFare2_LostFocus()
    txtFare2 = ""
    txtFare2.Visible = False
End Sub
Private Sub txtlanstage_KeyPress(KeyAscii As Integer)
    If LocalLanguage = 1 Then 'LANG
        txtlanstage.FontName = "senthamil"
    ElseIf LocalLanguage = 2 Then
        txtlanstage.FontName = "ML-TTKarthika"
    End If
    strBmpName = ""
    strBmpName = Trim(txtlanstage.Text)
    If KeyAscii = 13 Then
        LocalSTAGENamesOneByOne
        stageGrid.TextMatrix(0, 0) = "S.No"
        stageGrid.TextMatrix(0, 1) = "STAGE NAME"
        stageGrid.TextMatrix(0, 2) = "KM"
        If LocalLanguage > 0 Then
            stageGrid.TextMatrix(0, 3) = " STG"
        End If
        txtlanstage.BackColor = &H80000005
    End If
End Sub
Private Sub txtminfare_Click()
On Error Resume Next
    Call SendKeys("{HOME}+{END}") '05/01/2010
End Sub

Private Sub txtminfare_KeyPress(KeyAscii As Integer)
    Static LastText As String
    Static SecondTime As Boolean
    Const MaxDecimal As Integer = 1
    Const MaxWhole As Integer = 4
    If TextBoxValidityNumeric(KeyAscii) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    SecondTime = False
    If KeyAscii = 13 Then
        If PROJECT <> 1 Then
            If val(txtminfare.Text) <= 0 Then
                MsgBox "Minimum Fare Allowed Above 0", vbOKOnly
                txtminfare = ""
                txtminfare.SetFocus
                Exit Sub
            End If
            cbofaretype.SetFocus
        Else
            If val(txtminfare.Text) < 1 Then
                MsgBox "Minimum Fare Allowed Above 0", vbOKOnly
                txtminfare = ""
                txtminfare.SetFocus
                Exit Sub
            End If
            cbofaretype.SetFocus
        End If
    End If
End Sub
Private Sub cbofaretype_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        cmbBusType.SetFocus
        If cbofaretype.ListIndex = 1 And val(txtnostage) > getSTAGEmax Then
            MsgBox "Maximum Allowed stages in Graph should be less than " & getSTAGEmax + 1 & " Stages.", vbInformation
            txtnostage.Text = getSTAGEmax
            txtnostage.SetFocus
            txtnostage.SelStart = 0
            txtnostage.SelLength = Len(txtnostage)
        End If
    Case Else
    End Select
End Sub
Private Sub txtnostage_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 8
        KeyAscii = KeyAscii
    Case 13
        If val(txtnostage) > 252 Then
            MsgBox "Number of stage allowed up to only 252 stages", vbInformation, "Route"
            txtnostage = ""
            txtnostage.SetFocus
            Exit Sub
        ElseIf val(txtnostage) < 2 Then
            MsgBox "Number of stage must be greater than 1", vbInformation, "Route"
            txtnostage = ""
            txtnostage.SetFocus
            Exit Sub
        ElseIf val(txtnostage) < 3 And UCase(cbofaretype.List(cbofaretype.ListIndex)) = "GRAPH" Then
            MsgBox "Number of stage must be greater than 2 in Graph fare", vbInformation, "Route"
            txtnostage = ""
            txtnostage.SetFocus
            Exit Sub
        End If
        If cbofaretype.ListIndex = 1 And val(txtnostage) > getSTAGEmax Then
            MsgBox "Maximum Allowed stages in Graph should be less than " & getSTAGEmax + 1 & " Stages.", vbInformation
            txtnostage.Text = 50
            txtnostage.SetFocus
            txtnostage.SelStart = 0
            txtnostage.SelLength = Len(txtnostage)
        Else
            txtminfare.SetFocus
        End If
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtnostop_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 13, 8
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub
Private Sub txtrutcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 65 To 92, 97 To 122, 8, 32
        KeyAscii = KeyAscii
    Case 13
        If RouteExists(txtrutcode) = True Then
            MsgBox "Route already created" & vbCrLf & "Please give another code", vbInformation, gblstrPrjTitle
            txtrutcode.SetFocus
        Else
            txtrutname.SetFocus
        End If
    Case Else
        KeyAscii = 0
    End Select
End Sub
Private Sub txtrutcode_Validate(Cancel As Boolean)
On Error GoTo err
Dim cn As DAO.Database
Dim RES As DAO.Recordset
    Set cn = DAO.OpenDatabase(App.Path & "\PVT.Mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    Set RES = cn.OpenRecordset("select * from route where rutcode='" & Trim(txtrutcode) & "'", dbOpenDynaset)
    If RES.RecordCount > 0 Then
        MsgBox "Route already created" & vbCrLf & "Please give another code", vbInformation, gblstrPrjTitle
        txtrutcode.Text = ""
        txtrutcode.SetFocus
        Cancel = True
        Exit Sub
    End If
Exit Sub
err:
    MsgBox err.Number & " , " & err.Description, vbInformation, gblstrPrjTitle
    Exit Sub
End Sub

Private Sub txtrutname_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 65 To 92, 97 To 122, 8, 32, 45
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 13
        txtnostage.SetFocus
    Case Else
        KeyAscii = 0
    End Select
End Sub
Private Sub txtstartfrom_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 13, 8
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub
Private Sub txtStgName_KeyPress(KeyAscii As Integer)
    strBmpName = ""
    strBmpName = Trim(txtStgName.Text)
    Select Case KeyAscii
    Case 65 To 92, 97 To 122, 48 To 57, 13, 8, 32
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case Else
        KeyAscii = 0
    End Select
    If KeyAscii = 13 Then
        STAGENamesOneByOne
        stageGrid.TextMatrix(0, 0) = "S.No"
        stageGrid.TextMatrix(0, 1) = "STAGE NAME"
        stageGrid.TextMatrix(0, 2) = "KM"
        If LocalLanguage > 0 Then
            stageGrid.TextMatrix(0, 3) = " STG"
        End If
        txtStgName.BackColor = &H80000005
    End If
End Sub
Public Sub CLEARFIELDS()
    txtrutname = ""
    txtrutcode = ""
    txtnostage = ""
    txtnostop = ""
    txtminfare = ""
    cmbBusType.ListIndex = 0
    cbofaretype.ListIndex = 0
    chkHalf.Value = False
    chkluggage.Value = False
    chkstudent.Value = False
    chkph.Value = False
    chkConc.Value = False
    chkAdjust.Value = False
    chkAllowPass.Value = False '04/01/2010
    ChkSelect.Value = False
End Sub
Public Function FAREFRAMEFUNC()
    If FrType = 1 Then
        lblfrtype.caption = "TABLE"
        TotFare = NOSTGS
        lblnoofentry.caption = TotFare
        lblnoofstage.caption = "0/" & NOSTGS
    ElseIf FrType = 2 Then
        lblfrtype.caption = "GRAPH"
        TotFare = NOSTGS * (NOSTGS - 1) / 2
        lblnoofentry.caption = TotFare
        lblnoofstage.caption = "0/" & NOSTGS
    End If
    FieldCount = 1
End Function
Public Function FARELISTENTRY()
Dim temp As Integer
    If FrType = 2 Then
        FareTypeGrid.Cols = NOSTGS       ''''vaisakh 31.03.11
        FareTypeGrid.Rows = NOSTGS
        For temp = 1 To NOSTGS
            With FareTypeGrid
                .TextMatrix(0, temp - 1) = temp - 1
                .TextMatrix(temp - 1, 0) = temp
            End With
        Next temp
    End If
    txtFare1.Visible = True
    txtFare1.Top = FareTypeGrid.CellTop + FareTypeGrid.Top
    txtFare1.Left = FareTypeGrid.CellLeft + FareTypeGrid.Left
    txtFare1.SetFocus
    txtFare1_KeyPress (0)
End Function
Public Function STAGENamesOneByOne()   ' To Store the Datas by pressing ENTER key   ''' 14/01/2011
    If Trim(txtStgName.Text) <> "" Then
        If StgNameCnt < NOSTGS Then
            StageNameBuff = StageNameBuff + Trim(txtStgName.Text) + ","
            With stageGrid
                If .TextMatrix(StgNameCnt, 0) = "" Then
                    .AddItem ""
                End If
                .TextMatrix(StgNameCnt, 0) = StgNameCnt
                .TextMatrix(StgNameCnt, 1) = Trim(txtStgName.Text)
            End With
            If StgNameCnt > 9 Then
                stageGrid.TopRow = StgNameCnt - 1
            End If
            StgName(StgNameCnt - 1) = Trim(txtStgName.Text) & ","
            txtStgName.Text = ""
            txtStgName.SetFocus
            txtStgName.BackColor = &HC0EF00
            toListBox = toListBox + 1
        End If
        If StgNameCnt < stageGrid.Rows Then
            stageGrid.row = StgNameCnt
            Debug.Print StgNameCnt
        End If
        If LocalLanguage > 0 Then
            stageGrid.Col = 1
            stageGrid.CellBackColor = &H80000005
            stageGrid.Col = 2
            stageGrid.CellBackColor = &H80000005
            stageGrid.Col = 3
            stageGrid.CellBackColor = &H900FFF
        Else
            stageGrid.Col = 1
            stageGrid.CellBackColor = &H80000005
            stageGrid.Col = 2
            stageGrid.CellBackColor = &H900FFF
        End If
        If LocalLanguage > 0 Then
            txtlanstage.Enabled = True
            txtlanstage.SetFocus
            txtlanstage.BackColor = &HC0EF00
        Else
            txtDistance.Enabled = True
            txtDistance.SetFocus
            txtDistance.BackColor = &HC0EF00
            StgNameCnt = StgNameCnt + 1
        End If
    Else
        MsgBox "Enter the Stage Name", vbInformation
        txtStgName.SetFocus
        txtStgName.BackColor = &HC0EF00
    End If
    If StgNameCnt = NOSTGS And COUNTER >= 1 And LocalLanguage > 0 Then StgNameCnt = StgNameCnt + 1
    If StgNameCnt > NOSTGS And COUNTER >= 1 Then
        StageNameBuff = StageNameBuff + Trim(txtStgName.Text)
        With stageGrid
            .AddItem ""
            .TextMatrix(StgNameCnt - 1, 0) = StgNameCnt - 1
            .TextMatrix(StgNameCnt - 1, 1) = Trim(txtStgName.Text)
            .TextMatrix(StgNameCnt - 1, 3) = strLanguageStage
        End With
        StgName(StgNameCnt - 2) = Trim(txtStgName.Text)
        txtStgName.Text = ""
        toListBox = 0
        txtStgName.Enabled = False
        stageGrid.Rows = stageGrid.Rows - 1
        If StgNameCnt > NOSTGS And COUNTER >= 1 And LocalLanguage > 0 Then StgNameCnt = StgNameCnt - 1
        Exit Function
    End If
End Function
Public Function LocalSTAGENamesOneByOne()   ' To Store the Datas by pressing ENTER key   ''' 14/01/2011
    If Trim(txtlanstage.Text) <> "" Then
        If StgNameCnt < NOSTGS Then
            With stageGrid
                .TextMatrix(StgNameCnt, 0) = StgNameCnt
                .TextMatrix(StgNameCnt, 3) = Trim(txtlanstage.Text)
                If LocalLanguage > 0 Then 'LANG
                    .Col = 3
                    .row = StgNameCnt
                    If LocalLanguage = 1 Then .CellFontName = "senthamil"
                    If LocalLanguage = 2 Then .CellFontName = "ML-TTKarthika"
                End If
            End With
            If StgNameCnt > 9 Then
                stageGrid.TopRow = StgNameCnt - 1
            End If
            txtlanstage.Text = ""
            txtlanstage.SetFocus
            txtlanstage.BackColor = &HC0EF00
            toListBox = toListBox + 1
        End If
        If StgNameCnt < stageGrid.Rows Then
            stageGrid.row = StgNameCnt
        End If
        stageGrid.Col = 1
        stageGrid.CellBackColor = &H80000005
        stageGrid.Col = 2
        stageGrid.CellBackColor = &H900FFF
        stageGrid.Col = 3
        stageGrid.CellBackColor = &H80000005
        txtlanstage.Enabled = False
        txtDistance.Enabled = True
        
        txtDistance.SetFocus
        txtDistance.BackColor = &HC0EF00
        StgNameCnt = StgNameCnt + 1
    Else
        MsgBox "Enter the Stage Name", vbInformation
        txtlanstage.SetFocus
        txtlanstage.BackColor = &HC0EF00
    End If
    If StgNameCnt > NOSTGS And COUNTER >= 1 Then
        With stageGrid
            .AddItem ""
            .TextMatrix(StgNameCnt - 1, 0) = StgNameCnt - 1
            .TextMatrix(StgNameCnt - 1, 3) = Trim(txtlanstage.Text)
            If LocalLanguage > 0 Then 'LANG
                .Col = 3
                .row = StgNameCnt - 1
                If LocalLanguage = 1 Then .CellFontName = "senthamil"
                If LocalLanguage = 2 Then .CellFontName = "ML-TTKarthika"
            End If
        End With
        txtlanstage.Text = ""
        toListBox = 0
        txtlanstage.Enabled = False
        stageGrid.Rows = stageGrid.Rows - 1
        Exit Function
    End If
End Function
Private Sub ShowCols()
    If cmdCreate.Enabled <> True Then
        Text1.Width = stageGrid.CellWidth
        Text1.Height = stageGrid.CellHeight
        Text1.Left = stageGrid.Left + stageGrid.CellLeft
        Text1.Top = stageGrid.Top + stageGrid.CellTop
        Text1.Visible = True
        Text1.Text = stageGrid.TextMatrix(stageGrid.row, stageGrid.Col)
        Text1.SetFocus   ''Text1.Locked = True
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
        If LocalLanguage = 1 And stageGrid.Col = 3 Then 'LANG
            Text1.FontName = "senthamil"
        ElseIf LocalLanguage = 2 And stageGrid.Col = 3 Then
            Text1.FontName = "ML-TTKarthika"
        Else
            Text1.FontName = "MS Sans Serif"
        End If
    End If
 End Sub
Private Sub txtFare1_KeyPress(KeyAscii As Integer)
    Static LastText As String
    Static SecondTime As Boolean
    Const MaxDecimal As Integer = 1
    Const MaxWhole As Integer = 6
    If TextBoxValidityNumeric(KeyAscii) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    With txtFare1
        If Not SecondTime Then
            If .Text Like "*[!0-9.]*" _
                Or .Text Like "*.*.*" _
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
    If KeyAscii = 13 Then
        If txtFare1 = "." Then
            MsgBox "Please enter valid fare.", vbInformation, "Route"
            txtFare1.SelStart = 0
            txtFare1.SelLength = Len(txtFare1)
            txtFare1.SetFocus
            Exit Sub
        End If
        If val(txtFare1) < val(txtminfare) Then
            MsgBox "Minimum Fare is " & txtminfare, vbInformation, "Route"
            txtFare1.SelStart = 0
            txtFare1.SelLength = Len(txtFare1)
            txtFare1.SetFocus
            Exit Sub
        End If
        With FareTypeGrid
            If Not SAVEFLAG = False Or FARESAVEFLAG = False Then
                If val(txtFare1) < val(txtminfare) Then
                    MsgBox "Minimum fare is  " & txtminfare, vbExclamation
                    txtFare1 = ""
                    Exit Sub
                End If
                If FrType = 2 Then
                    .Cols = NOSTGS
                    .Rows = NOSTGS
                    .TextMatrix(.row, .Col) = Round(Trim(txtFare1), 2)
                    If .row <> 1 And First = False Then
                        Exit Sub
                    ElseIf First = False And FARESAVEFLAG = True Then
                        RES.AddNew
                        RES!row = .row
                        RES!Col = .Col
                        RES!FARE = Round((.TextMatrix(.row, .Col)), 2)
                        RES!Route = lblShowRoute.caption
                        RES.Update
                        TRES.AddNew
                        TRES!Save = False
                        TRES!NoOfEntries = EnteredCount
                        TRES.Update
                        lblEntered.caption = EnteredCount
                        First = True
                        .row = .row + 1
                        R = R + 1
                        EnteredCount = EnteredCount + 1
                        EntryCount = EntryCount + 1
                        txtFare2.Top = .CellTop + .Top
                        txtFare2.Left = .CellLeft + .Left
                        txtFare2.Text = .TextMatrix(R, c) ''' vaisakh 30.03.11
                        txtFare1.Width = .CellWidth
                        txtFare1.Height = .CellHeight
                    Else
                        RES.Close
                        TSQL = "SELECT * FROM TMPFARE WHERE Row = " & .row & " AND Col = " & .Col & " AND route= '" & lblShowRoute & "'"
                        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                        If RES.RecordCount <> 0 Then
                            RES.Edit
                            RES!FARE = Round((.TextMatrix(.row, .Col)), 2)
                            RES.Update
                        End If
                        RES.Close
                        .row = R
                        .Col = c
                        TSQL = "SELECT * FROM TMPFARE"
                        Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                        RES.MoveLast
                    End If
                    If EntryCount > NoOfEntries Or EnteredCount > NoOfEntries Then
                        txtFare1 = ""
                        txtFare1.Visible = False
                        Exit Sub
                    End If
                    txtFare1 = ""
                    txtFare1.Visible = False
                    txtFare2.Visible = True
                    txtFare2.Text = .TextMatrix(R, c) ''' vaisakh 30.03.11
                    txtFare2.SelStart = 0
                    txtFare2.SelLength = Len(txtFare2)
                    txtFare2.SetFocus
                ElseIf FrType = 1 Then
                    If First = False And FARESAVEFLAG = True Then
                        First = True
                        c = c + 1
                        .TextMatrix(.row, .Col) = 0
                        RES.AddNew
                        RES!row = .row
                        RES!Col = .Col
                        RES!FARE = Round((.TextMatrix(.row, .Col)), 2)
                        RES!Route = lblShowRoute.caption
                        RES.Update
                        TRES.AddNew
                        TRES!Save = False
                        .Cols = .Cols + 2
                        .Col = c
                        
                        .TextMatrix(.row, c) = val(txtFare1)
                        TRES!NoOfEntries = EnteredCount '2 changed by SAN to avoid blank text box while entering st>2
                        TRES.Update
                        RES.AddNew
                        RES!row = .row
                        RES!Col = .Col
                        RES!FARE = Round((.TextMatrix(.row, c)), 2)
                        RES!Route = lblShowRoute.caption
                        RES.Update
                        c = c + 1
                        .Col = c
                        EnteredCount = EnteredCount + 1
                        lblEntered.caption = EnteredCount
                        EntryCount = EntryCount + 1
                        If NOSTGS = 2 Then
                            txtFare1 = ""
                            txtFare1.Visible = False
                            cmdFareEntry.Enabled = True
                            cmdFareEntry.SetFocus
                            c = c - 1
                            .Cols = .Cols - 1
                            Exit Sub
                        End If
                        txtFare1.Visible = False
                        txtFare2.Visible = True
                        txtFare2.Top = .CellTop + .Top
                        txtFare2.Left = .CellLeft + .Left
                        txtFare2.SetFocus
                    Else
                        If EntryCount < NOS Then
                            .TextMatrix(.row, .Col) = Round(val(txtFare1), 2)
                            lblEntered.caption = EnteredCount
                            TSQL = "SELECT * FROM TMPFARE WHERE Row = " & .row & " AND Col = " & .Col & " AND route= '" & lblShowRoute & "'"
                            Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                            If RES.RecordCount > 0 Then
                                RES.Edit
                            Else
                                RES.AddNew
                            End If
                            RES!FARE = Round((.TextMatrix(.row, .Col)), 1)
                            RES.Update
                            RES.Close
                            TSQL = "SELECT * FROM TMPFARE"
                            Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                            .Col = c
                            .row = R
                            txtFare2.Visible = True
                            txtFare1.Visible = False
                            txtFare2.Top = .CellTop + .Top
                            txtFare2.Left = .CellLeft + .Left
                            txtFare2.SetFocus
                        Else
                            .TextMatrix(.row, .Col) = Round(val(txtFare1), 2)
                            RES.Close
                            TSQL = "SELECT * FROM TMPFARE WHERE Row = " & .row & " AND Col = " & .Col & " AND Route= '" & lblShowRoute & "'"
                            Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                            RES.Edit
                            RES!FARE = Round((.TextMatrix(.row, .Col)), 2)
                            RES.Update
                            RES.Close
                            TSQL = "SELECT * FROM TMPFARE"
                            Set RES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                            txtFare1 = ""
                            txtFare1.Visible = False
                            Exit Sub
                        End If
                    End If
                    txtFare1.Visible = False
                End If
            Else
                .TextMatrix(.row, .Col) = Round(val(txtFare1), 2)
                txtFare1.Visible = False
                TSQL = "SELECT * FROM TMPGRAPH WHERE ROW= " & .row & "AND COL = " & .Col
                Set TRES = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                If TRES.RecordCount > 0 Then
                    TRES.Edit
                    TRES!FARE = Round((val(.TextMatrix(.row, .Col))), 2)
                Else
                    TRES.AddNew
                    TRES!FARE = Round((val(.TextMatrix(.row, .Col))), 2)
                End If
                TRES.Update
            End If
        End With
    End If
End Sub
Private Sub txtFare2_KeyPress(KeyAscii As Integer)
Static LastText As String
Static SecondTime As Boolean
Const MaxDecimal As Integer = 1
Const MaxWhole As Integer = 5
    If TextBoxValidityNumeric(KeyAscii) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    With txtFare2
        If Not SecondTime Then
            If .Text Like "*[!0-9.]*" _
                Or .Text Like "*.*.*" _
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
    If KeyAscii = 13 Then
        With FareTypeGrid
            If txtFare2 = "." Then
                MsgBox "Please enter valid fare.", vbInformation, "Route"
                txtFare2 = ""
                Exit Sub
            End If
            If val(txtFare2) < val(txtminfare) Then
                MsgBox "Minimum fare is  " & txtminfare, vbExclamation
                txtFare2 = ""
                Exit Sub
            End If
            If FrType = 2 Then
                If val(.TextMatrix(R - 1, c)) > val(Trim(txtFare2)) Then
                    MsgBox "Fare must be greater than previous fare! ", vbExclamation
                    txtFare2 = ""
                    Exit Sub
                End If
            ElseIf val(.TextMatrix(.row, c - 1)) > val(Trim(txtFare2)) Then
                MsgBox "Fare must be greater than previous fare! ", vbExclamation
                txtFare2 = ""
                Exit Sub
            End If
            If FrType = 2 Then
                .TextMatrix(R, c) = Round(val(Trim(txtFare2)), 2)    '''vaisakh 30.03.11
                TSQL = "SELECT * FROM TMPFARE WHERE Row = " & .row & " AND Col = " & .Col & " AND route= '" & lblShowRoute & "'"
                
                Set RES3 = TDB.OpenRecordset(TSQL, dbOpenDynaset)
                If RES3.RecordCount <> 0 Then
                    RES3.Edit
                    RES3!FARE = Round((.TextMatrix(.row, .Col)), 2)
                    RES3.Update
                    RES3.Close
                Else
                    If RES.RecordCount > 0 Then RES.MoveLast
                    RES.AddNew
                    RES!row = .row
                    RES!Col = .Col
                    RES!FARE = Round((.TextMatrix(R, c)), 2)
                    RES!Route = lblShowRoute.caption
                    RES.Update
                    TRES.MoveFirst
                    TRES.Edit
                    TRES!Save = False
                    TRES!NoOfEntries = EnteredCount
                    TRES.Update
                    lblEntered.caption = EnteredCount
                End If
                If EntryCount < NoOfEntries Then
                    EntryCount = EntryCount + 1
                    If R < NOSTGS - 1 Then
                        R = R + 1
                        .row = R
                        .Col = c
                        EnteredCount = EnteredCount + 1
                    Else
                        If c < NOSTGS - 1 Then
                            c = .Col + 1
                            R = c
                            NOS = NOS - 1
                            EnteredCount = EnteredCount + 1
                        Else
                            txtFare2 = ""
                            txtFare2.Visible = False
                        End If
                        .Col = c
                        .row = R
                    End If
                    txtFare2.Top = .CellTop + .Top
                    txtFare2.Left = .CellLeft + .Left
                    txtFare2 = ""
                    txtFare1.Visible = False
                    txtFare2.Text = .TextMatrix(R, c)
                    txtFare2.SelStart = 0
                    txtFare2.SelLength = Len(txtFare2)
                    txtFare2.SetFocus
                Else
                    NOS = NOS - 1
                End If
                If NOS < 2 Then
                    txtFare2.Visible = False
                    cmdFareEntry.Enabled = True
                    cmdFareEntry.SetFocus
                End If
            ElseIf FrType = 1 Then
                .TextMatrix(.row, c) = Round(val(txtFare2), 2)
                RES.AddNew
                RES!row = .row
                RES!Col = .Col
                RES!FARE = Round((.TextMatrix(R, c)), 2)
                RES!Route = lblShowRoute.caption
                RES.Update
                TRES.MoveFirst
                TRES.Edit
                TRES!Save = False
                TRES!NoOfEntries = EnteredCount
                TRES.Update
                lblEntered.caption = EnteredCount + 1
                EntryCount = EntryCount + 1
                If EntryCount = NOS Then
                    txtFare2 = ""
                    txtFare2.Visible = False
                    cmdFareEntry.Enabled = True
                    cmdFareEntry.SetFocus
                    Exit Sub
                Else
                    c = c + 1
                    .Cols = .Cols + 1
                    .Col = c
                    EnteredCount = EnteredCount + 1
                    txtFare2 = ""
                    txtFare2.Top = .CellTop + .Top
                    txtFare2.Left = .CellLeft + .Left
                    txtFare2.SelStart = 0
                    txtFare2.SelLength = Len(txtFare2)
                    txtFare2.SetFocus
                End If
            End If
        End With
    End If
End Sub
Public Function MinimumFareCheck() As Boolean
Dim I As Integer
Dim j As Integer
Dim k As Integer
    With FareTypeGrid
        k = NOSTGS - 1
        For I = 1 To NOSTGS - 1
            For j = 1 To I
                If FrType = 1 And I > 1 Then MinimumFareCheck = True: Exit Function
                If val(txtminfare) > val(.TextMatrix(I, j)) Then
                    If FrType = 1 And j = 1 Then Exit For
                    MsgBox "Fare in cell (" & I & "," & j & ") is less than minimum fare", vbExclamation, "Route"
                    .row = I
                    .Col = j
                    .SetFocus
                    txtFare1.Visible = True
                    txtFare1.Left = .CellLeft + .Left
                    txtFare1.Top = .CellTop + .Top
                    txtFare1 = Round(.TextMatrix(.row, .Col), 2)
                    txtFare1.SelStart = 0
                    txtFare1.SelLength = Len(txtFare1)
                    txtFare1.SetFocus
                    MinimumFareCheck = False
                    Exit Function
                End If
            Next
        Next
        MinimumFareCheck = True
    End With
End Function
Public Function AddLanguageStageName()
On Error GoTo err
Dim sql As String
    sql = "SELECT STG_LOCAL_LANGUAGE FROM STAGE"
    Set RES = cn.OpenRecordset(sql, dbOpenDynaset)
    If RES.RecordCount > 0 Then
        RES.MoveLast
    End If
    RES.Edit
    RES!STG_LOCAL_LANGUAGE = strLanguageStage
    RES.Update
    RES.Close
Exit Function
err:
    MsgBox "Language Stage Name Save Error!", vbInformation, "BUS"
End Function
Public Function getbustype(bustype_name As String) As Integer
On Error GoTo lblErr
Dim sql As String
    sql = "SELECT id FROM bustype where name='" & bustype_name & "'"
    Set RESSAN = cn.OpenRecordset(sql, dbOpenDynaset)
    getbustype = RESSAN!Id
    RESSAN.Close
Exit Function
lblErr:
End Function
Public Function getbustypecmbid(bustype_id As Integer) As String
On Error GoTo lblErr
Dim sql As String
    sql = "SELECT name FROM bustype where id=" & bustype_id
    Set res1 = cn.OpenRecordset(sql, dbOpenDynaset)
    getbustypecmbid = res1!Name
    res1.Close
    Set res1 = Nothing
Exit Function
lblErr:
End Function
