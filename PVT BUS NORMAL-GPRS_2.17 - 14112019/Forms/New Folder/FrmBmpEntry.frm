VERSION 5.00
Begin VB.Form frmBmpEntry 
   BorderStyle     =   0  'None
   Caption         =   "BMP Stage Entry"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleMode       =   0  'User
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileName 
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cmbFontSize 
      Height          =   315
      Left            =   9840
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
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
      Left            =   6120
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbFont 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Size"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHeight 
      Height          =   375
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   6120
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create BMP"
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtG 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmBmpEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias _
         "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, ByVal lParam As Long) As Long


Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long



Private Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type


Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10


Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4&    ' Draw the window's client area
Private Const PRF_CHILDREN = &H10& ' Draw all visible child
Private Const PRF_OWNED = &H20&    ' Draw all owned windows

Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Const BI_RGB = 0&
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Dim rv As Long

Public BmpWidth1 As Long
Public BmpHeight1 As Long


Private Sub pDrawOnPicbox(ByRef oPicBox As PictureBox, ByVal sText As String)
Dim r As RECT
    oPicBox.Cls

    With r
        .Left = 0
        .Top = 0
        .Right = oPicBox.ScaleWidth
        .Bottom = oPicBox.ScaleHeight
    End With
    SetRect r, 0, 0, oPicBox.ScaleWidth, oPicBox.ScaleHeight
    DrawText oPicBox.hdc, sText, Len(sText), r, DT_LEFT Or DT_WORDBREAK
End Sub

Private Sub Command1_Click()
    Dim str As String
    Dim bError As Byte
    pic.Cls
    
'    bError = TextToPicture(pic, txtG.Text, 0, , 4, vbButtonFace)
'    If Not bError Then
'      MsgBox "Unable to find suitable font size for text"
'    End If
    
    Call pDrawOnPicbox(pic, txtG)
'    SavePicture pic.Picture, App.Path & "\1.bmp"
End Sub


Private Sub chkBold_Click()
    If chkBold.Value = 1 Then
        txtG.FontBold = True
        pic.FontBold = True
    Else
        txtG.FontBold = False
        pic.FontBold = False
    End If
End Sub

Private Sub cmbFont_Click()
    txtG.Font = cmbFont.Text
    pic.Font = txtG.Font
End Sub


Private Sub cmbFontSize_Change()
    If val(cmbFontSize.Text) > 10 Then
        txtG.FontSize = val(cmbFontSize.Text)
        pic.FontSize = val(cmbFontSize.Text)
    End If
End Sub

Private Sub cmbFontSize_Click()
    If val(cmbFontSize.Text) > 10 Then
        txtG.FontSize = val(cmbFontSize.Text)
        pic.FontSize = val(cmbFontSize.Text)
    End If
End Sub

Private Sub cmdCreate_Click()
On Error GoTo err:
    Dim Fso As New FileSystemObject
    txtFileName = strBmpName
    pic.Visible = True
    Call pDrawOnPicbox(pic, txtG)
    If Fso.FolderExists(App.Path & "\pic") = False Then Fso.CreateFolder App.Path & "\pic"
    gblBMPName = ""
    gblBMPName = txtFileName & ".bmp"
    SavePictureBW pic, App.Path & "\pic\" & txtFileName & ".bmp"
    pic.Visible = False
    gblBMPName = txtFileName & ".bmp"
'    MsgBox "Successfully Bitmap file created successfully" & vbCrLf & vbCrLf & vbCrLf, vbInformation
    Unload Me
    Exit Sub
err:
    MsgBox "Error!" & vbCrLf & "Error No:" & err.Number & vbCrLf & err.Description, vbCritical
End Sub



Private Sub SavePictureBW(ByVal Ctrl As PictureBox, ByVal DestFile As String)
    Dim hdcMono As Long, hbmpMono As Long, hBmpOld As Long, dxBlt As Long, dyBlt As Long, success As Long
    Dim numscans As Long, byteswide As Long, totalbytes As Long, lfilesize As Long
    Dim bmpsrc As BITMAP, bmpdst As BITMAP
    Dim bInfo As BITMAPINFO
    Dim bitmaparray() As Byte, fileheader() As Byte
    Dim ff As Integer

    'Object's scalemode must be Pixel.
    dxBlt = Ctrl.ScaleWidth
    dyBlt = Ctrl.ScaleHeight

    'Create monochrome bitmap from control.
    hdcMono = CreateCompatibleDC(0)
    hbmpMono = CreateCompatibleBitmap(hdcMono, dxBlt, dyBlt)
    success = GetObject(hbmpMono, Len(bmpsrc), bmpsrc)
    hBmpOld = SelectObject(hdcMono, hbmpMono)
    success = BitBlt(hdcMono, 0, 0, dxBlt, dyBlt, Ctrl.hdc, 0, 0, SRCCOPY)

    'Calculate array size needed for bitmap bits (dword aligned)
    numscans = dyBlt
    by8 = dxBlt / 8
    If (dxBlt Mod 8) = 0 And (by8 Mod 4) = 0 Then
       byteswide = by8
    Else
       byteswide = (Int(by8) + 4) - (Int(by8) Mod 4)
    End If
    totalbytes = numscans * byteswide
    ReDim bitmaparray(1 To totalbytes)

    'Set BITMAPINFO values to pass to GetDIBits function.
    With bInfo
       .bmiHeader.biSize = Len(.bmiHeader)
       .bmiHeader.biWidth = bmpsrc.bmWidth
       .bmiHeader.biHeight = bmpsrc.bmHeight
       .bmiHeader.biPlanes = bmpsrc.bmPlanes
       .bmiHeader.biBitCount = bmpsrc.bmBitsPixel
       .bmiHeader.biCompression = BI_RGB
    End With

    success = GetDIBits(hdcMono, hbmpMono, 0, numscans, bitmaparray(1), bInfo, DIB_RGB_COLORS)
    'success = StretchBlt(hdcMono, 0, 0, dxBlt, dyBlt, pbSrc.hdc, 0, dyBlt - 1, dxBlt, -dyBlt, SRCCOPY)
    'bitmaparray should now contain bitmap bit data. Now create bitmap file header.
    ReDim fileheader(1 To &H3E)
    fileheader(1) = &H42 'B
    fileheader(2) = &H4D 'M
    lfilesize = UBound(fileheader) + UBound(bitmaparray)
    fileheader(3) = lfilesize And 255
    fileheader(4) = (lfilesize \ 256) And 255
    fileheader(5) = (lfilesize \ 65536) And 255
    fileheader(6) = (lfilesize \ 16777216) And 255
    fileheader(11) = &H3E 'offset
    fileheader(15) = &H28 'size of bitmapinfoheader
    fileheader(19) = dxBlt And 255
    fileheader(20) = (dxBlt \ 256) And 255
    fileheader(21) = (dxBlt \ 65536) And 255
    fileheader(22) = (dxBlt \ 16777216) And 255
    fileheader(23) = dyBlt And 255
    fileheader(24) = (dyBlt \ 256) And 255
    fileheader(25) = (dyBlt \ 65536) And 255
    fileheader(26) = (dyBlt \ 16777216) And 255
    fileheader(27) = 1
    fileheader(29) = 1
    fileheader(35) = UBound(bitmaparray) And 255
    fileheader(36) = (UBound(bitmaparray) \ 256) And 255
    fileheader(37) = (UBound(bitmaparray) \ 65536) And 255
    fileheader(38) = (UBound(bitmaparray) \ 16777216) And 255
    fileheader(47) = 2
    fileheader(51) = 2
    fileheader(59) = &HFF
    fileheader(60) = &HFF
    fileheader(61) = &HFF

    ff = FreeFile
    Open DestFile For Binary Access Write As #ff
       Put #ff, , fileheader
       Put #ff, , bitmaparray
    Close #ff

    ' Clean up
    Call SelectObject(hdcMono, hBmpOld)
    Call DeleteDC(hdcMono)
    Call DeleteObject(hbmpMono)
End Sub



Private Sub Command3_Click()
    If val(txtWidth) Mod 8 > 0 Then MsgBox "Width should be multiple of 8": txtWidth.SetFocus
    
    pic.ScaleWidth = val(txtWidth)
    pic.ScaleHeight = val(txtHeight)
End Sub

Private Sub Form_Activate()
    
    txtWidth = "384"
    txtHeight = "64"
    pic.Width = Me.Width - 500
    pic.Height = Me.Height - 500
    
    pic.ScaleWidth = val(txtWidth)
    pic.ScaleHeight = val(txtHeight)
    txtG.SetFocus
''''    LoadFont
    
End Sub


Public Sub LoadFont()
    Dim iVal As Integer
    
    For iVal = 1 To Screen.FontCount - 1
        cmbFont.AddItem Screen.Fonts(iVal)
    Next
    
    If cmbFont.ListCount > 0 Then
        cmbFont.Text = cmbFont.List(0)
    End If
    
    
    
    For iVal = 12 To 30
        cmbFontSize.AddItem iVal
    Next
    
    If cmbFontSize.ListCount > 0 Then
        cmbFontSize.Text = cmbFontSize.List(0)
    End If
    
End Sub

Private Sub Form_Load()
Dim dbRS As dao.Recordset

    LoadFont
    
    CONNECT_DB

    BmpWidth1 = 0: BmpHeight1 = 0

    TSQL = "SELECT * FROM BMP_Settings"
    Set dbRS = DB.OpenRecordset(TSQL, dbOpenDynaset)
    If dbRS.RecordCount > 0 Then
         txtG.Font = dbRS!Font_Name
         pic.Font = txtG.Font
         txtG.FontSize = dbRS!Font_Size
         pic.FontSize = dbRS!Font_Size
         cmbFontSize.Text = dbRS!Font_Size
         cmbFont.Text = dbRS!Font_Name
         
         If dbRS!Bold_EnableOrDisable = 1 Then
            txtG.FontBold = True
            pic.FontBold = True
            chkBold.Value = 1
         Else
            txtG.FontBold = False
            pic.FontBold = False
            chkBold.Value = 0
         End If
         BmpWidth1 = dbRS!Bmp_Width
         BmpHeight1 = dbRS!Bmp_Height
         pic.ScaleWidth = dbRS!Bmp_Width
         pic.ScaleHeight = dbRS!Bmp_Height
         txtWidth.Text = dbRS!Bmp_Width
         txtHeight.Text = dbRS!Bmp_Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gblBMPName = txtFileName & ".bmp"
End Sub



Private Sub txtG_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCreate_Click
    End If
End Sub

Private Sub txtHeight_Change()
    If val(txtHeight) > 480 Then txtHeight = "480"
End Sub

Private Sub txtWidth_Change()
    If val(txtWidth) > 384 Then txtWidth = "384"
End Sub
