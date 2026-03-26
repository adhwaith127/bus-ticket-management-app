VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form frmExpenseDetails 
   BackColor       =   &H80000009&
   Caption         =   "Expense  Details"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.TextBox txtExptime 
         Height          =   285
         Left            =   4920
         TabIndex        =   13
         Top             =   2640
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPExpense 
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Format          =   50790401
         CurrentDate     =   40760
      End
      Begin VB.TextBox txtExpamount 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtExpname 
         Height          =   285
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   8
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtRcptNo 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtExpcode 
         Height          =   285
         Left            =   4920
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin JeweledBut.JeweledButton cmdCancel 
         Height          =   495
         Left            =   5280
         TabIndex        =   14
         Top             =   3360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         TX              =   "&Cancel"
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
         MICON           =   "frmExpenseDetails.frx":0000
         BC              =   12632256
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cmdSave 
         Height          =   495
         Left            =   3960
         TabIndex        =   15
         Top             =   3360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         TX              =   "&Save"
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
         MICON           =   "frmExpenseDetails.frx":001C
         BC              =   12632256
         FC              =   0
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time"
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
         TabIndex        =   10
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
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
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exp.Code"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Caption         =   "Amount"
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
         TabIndex        =   4
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Name"
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
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "EXPENSE DETAILS"
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
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Receipt No:"
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
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmExpenseDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
