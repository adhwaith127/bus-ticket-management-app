VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMdiBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©2001 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within Your own applications, but You
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

' Win32 APIs
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function Iswindow Lib "user32" (ByVal hWnd As Long) As Long

' GetSystemMetrics constants
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3

' SetWindowPos Flags
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_SHOWwindow = &H40
Private Const SWP_HIDEwindow = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200      '  Youn't You owner Z ordering

' Member variables
Private WithEvents m_Client As MDIForm
Attribute m_Client.VB_VarHelpID = -1
Private m_Canvas As PictureBox
Private m_Graphic As StdPicture
Private m_AutoRefresh As Boolean
Private m_ColorTop As OLE_COLOR
Private m_ColorBottom As OLE_COLOR
Private m_BackStyle As mdiBackStyles
Private m_GfxPos As mdiGraphicPositions
Private m_hWnd As Long

' Constants
Private Const pbID As String = "mdiBackPicture"

' Background styles
Public Enum mdiBackStyles
   mdiSolid = 0
   mdiGradient = 1
End Enum

Public Enum mdiColors
   mdiColorTop = 0
   mdiColorBottom = 1
End Enum

' Background graphic positions
Public Enum mdiGraphicPositions
   mdiNone = 0
   mdiCentered = 1
   mdiTiled = 2
   mdiStretched = 3
End Enum

' Default member values
Private Const defAutoRefresh As Boolean = False
Private Const defBackStyle As Long = mdiSolid
Private Const defColorTop As Long = vbApplicationWorkspace
Private Const defColorBottom As Long = vbApplicationWorkspace
Private Const defGfxPos As Long = mdiNone

' ********************************************
'  Class Events
' ********************************************
Private Sub Class_Initialize()
   ' Set default values for class
   m_AutoRefresh = defAutoRefresh
   m_BackStyle = defBackStyle
   m_ColorTop = defColorTop
   m_ColorBottom = defColorBottom
   m_GfxPos = defGfxPos
End Sub

Private Sub Class_Terminate()
   ' Clean up
   Set m_Graphic = Nothing
   Call ClientTearDown
End Sub

' ********************************************
'  Client Events
' ********************************************
Private Sub m_Client_Resize()
   ' Adjust canvas to fit
   If m_Client.WindowState <> vbMinimized Then
      Call CanvasResize
   End If
End Sub

' ********************************************
'  Public Properties
' ********************************************
Public Property Let AutoRefresh(ByVal NewVal As Boolean)
   m_AutoRefresh = NewVal
   If m_AutoRefresh Then Call CanvasRefresh
End Property

Public Property Get AutoRefresh() As Boolean
   AutoRefresh = m_AutoRefresh
End Property

Public Property Let BackStyle(ByVal NewVal As mdiBackStyles)
   m_BackStyle = NewVal
   If m_AutoRefresh Then Call CanvasRefresh
End Property

Public Property Get BackStyle() As mdiBackStyles
   BackStyle = m_BackStyle
End Property

Public Property Set Client(ByVal NewClient As MDIForm)
   ' Clean up old client, if need be
   Call ClientTearDown
   ' Set up new client
   Set m_Client = NewClient
   Call ClientSetup
End Property

Public Property Get Client() As MDIForm
   Set Client = m_Client
End Property

Public Property Let Color(Optional ByVal Which As mdiColors = mdiColorTop, ByVal NewVal As OLE_COLOR)
   Select Case Which
      Case mdiColorTop
         m_ColorTop = NewVal
      Case mdiColorBottom
         m_ColorBottom = NewVal
   End Select
   If m_AutoRefresh Then Call CanvasRefresh
End Property

Public Property Get Color(Optional ByVal Which As mdiColors = mdiColorTop) As OLE_COLOR
   Select Case Which
      Case mdiColorTop
         Color = m_ColorTop
      Case mdiColorBottom
         Color = m_ColorBottom
   End Select
End Property

Public Property Set Graphic(ByVal NewPict As StdPicture)
   Set m_Graphic = NewPict
   If m_AutoRefresh Then Call CanvasRefresh
End Property

Public Property Get Graphic() As StdPicture
   Set Graphic = m_Graphic
End Property

Public Property Let GraphicPosition(ByVal NewVal As mdiGraphicPositions)
   m_GfxPos = NewVal
   If m_AutoRefresh Then Call CanvasRefresh
End Property

Public Property Get GraphicPosition() As mdiGraphicPositions
   GraphicPosition = m_GfxPos
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Sub Refresh()
   Call CanvasRefresh
End Sub

' ********************************************
'  Private Methods
' ********************************************
Private Sub CanvasRefresh()
   Dim x As Long, y As Long
   Dim w As Long, h As Long
   Const swpFlags As Long = _
      SWP_FRAMECHANGED Or SWP_NOMOVE Or _
      SWP_NOZORDER Or SWP_NOSIZE
   
   ' Bail if no canvas established
   If m_Canvas Is Nothing Then Exit Sub
   
   ' Paint pretty picture :-)
   If m_BackStyle = mdiSolid Then
      m_Canvas.BackColor = m_ColorTop
   Else
      Call PaintGradient(m_Canvas, m_ColorTop, m_ColorBottom)
   End If
   
   ' Check for special graphic
   If Not (m_Graphic Is Nothing) Then
      With m_Canvas
         w = .ScaleX(m_Graphic.Width, vbHimetric, .ScaleMode)
         h = .ScaleY(m_Graphic.Height, vbHimetric, .ScaleMode)
         Select Case m_GfxPos
            Case mdiNone
            Case mdiCentered
               x = (.ScaleWidth - w) \ 2
               y = (.ScaleHeight - h) \ 2
               .PaintPicture m_Graphic, x, y
            Case mdiTiled
               For x = 0 To .ScaleWidth \ w
                  For y = 0 To .ScaleHeight \ h
                     .PaintPicture m_Graphic, x * w, y * h
                  Next y
               Next x
            Case mdiStretched
               .PaintPicture m_Graphic, 0, 0, .ScaleWidth, .ScaleHeight
         End Select
      End With
   End If
   
   ' Force client to repaint canvas
   Set m_Client.Picture = m_Canvas.Image
   Call SetWindowPos(m_Client.hWnd, 0, 0, 0, 0, 0, swpFlags)
End Sub

Private Sub CanvasResize()
   Dim w As Long, h As Long
   ' Bail if no canvas established
   If m_Canvas Is Nothing Then Exit Sub
   ' Adjust canvas to fit
   w = m_Client.ScaleWidth + (GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX)
   h = m_Client.ScaleHeight + (GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY)
   Call m_Canvas.Move(0, 0, w, h)
   ' Always update graphics on resize!
   Call CanvasRefresh
End Sub

Private Sub ClientSetup()
   ' Bail if no client established
   If m_Client Is Nothing Then Exit Sub
   ' Create background canvas
   Set m_Canvas = m_Client.Controls.Add("VB.PictureBox", pbID)
   With m_Canvas
      ' Set appropriate properties
      .AutoRedraw = True
      .BorderStyle = 0
      .ClipControls = False
      ' Cache window handle
      m_hWnd = .hWnd
   End With
   Call CanvasResize
End Sub

Private Sub ClientTearDown()
   ' Bail if no client established
   If m_Client Is Nothing Then Exit Sub
   ' Remove background canvas
   If Iswindow(m_hWnd) Then
      ' Testing for the window is required because
      ' the client may be in an indeterminate
      ' empcat and its control collection could be
      ' hosed if in the middle of unloading.
      On Error Resume Next
      Call m_Client.Controls.Remove(pbID)
      Set m_Canvas = Nothing
   End If
   ' Release reference to client
   Set m_Client = Nothing
End Sub

' ********************************************
'  Paint Methods
' ********************************************
Private Sub PaintGradient(pic As PictureBox, ByVal color1 As OLE_COLOR, ByVal color2 As OLE_COLOR)
   Dim oldScaleMode As ScaleModeConstants
   Dim I As Long
   Const Shades As Long = 64
   Dim rgbs() As Long
      
   ' Get array of colors to paint with
   rgbs = GetGradients(Shades, color1, color2)
   
   ' Rescale picture and paint
   With pic
      oldScaleMode = .ScaleMode
      .ScaleMode = vbUser
      .ScaleTop = 0
      .ScaleHeight = Shades
      .ScaleLeft = 0
      .ScaleWidth = 1
      For I = 1 To Shades
         pic.Line (0, I - 1)-(1, I), rgbs(I), BF
      Next I
      .ScaleMode = oldScaleMode
   End With
End Sub

Private Function GetGradients(ByVal Shades As Long, ByVal color1 As OLE_COLOR, ByVal color2 As OLE_COLOR) As Long()
   Dim I As Long
   Dim rShift As Integer
   Dim gShift As Integer
   Dim bShift As Integer
   Dim rNew As Integer
   Dim gNew As Integer
   Dim bNew As Integer
   Dim nRet() As Long
      
   ' Convert system to RGB colors
   color1 = CheckSysColor(color1)
   color2 = CheckSysColor(color2)
   
   ' Calc shift values for each channel
   rNew = GetRValue(color1)
   gNew = GetGValue(color1)
   bNew = GetBValue(color1)
   rShift = (rNew - GetRValue(color2)) \ Shades
   gShift = (gNew - GetGValue(color2)) \ Shades
   bShift = (bNew - GetBValue(color2)) \ Shades
   
   ' Create new array of color values
   ReDim nRet(1 To Shades) As Long
   For I = 1 To Shades
      nRet(I) = RGB(rNew, gNew, bNew)
      rNew = rNew - rShift
      gNew = gNew - gShift
      bNew = bNew - bShift
   Next I
   GetGradients = nRet
End Function

Private Function CheckSysColor(ByVal Color As Long) As Long
   Const HighBit = &H80000000
   
   ' If high bit set, strip, and get system color.
   If Color And HighBit Then
      CheckSysColor = GetSysColor(Color And Not HighBit)
   Else
      CheckSysColor = Color
   End If
End Function

Private Function GetRValue(ByVal Value As Long) As Byte
  ' #define GetRValue(rgb)   ((BYTE) (rgb))
  GetRValue = (Value And &HFF)
End Function

Private Function GetGValue(ByVal Value As Long) As Byte
  ' #define GetGValue(rgb)   ((BYTE) (((WORD) (rgb)) >> 8))
  GetGValue = (Value And &HFF00&) \ &H100
End Function

Private Function GetBValue(ByVal Value As Long) As Byte
  ' #define GetBValue(rgb)   ((BYTE) ((rgb) >> 16))
  GetBValue = (Value And &HFF0000) \ &H10000
End Function

