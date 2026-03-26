Attribute VB_Name = "modLogoBMP"
Option Explicit

Private Const NOOF_BYTES_PER_LINE = 48

Public Type BMP_HEAD
    MagicNo As Integer
    FileSize As Long
    BMPReserved1 As Integer
    BMPReserved2 As Integer
    DataOffset As Long
End Type

Public Type BMP_INFO
   HeadSize As Long
   BMPWidth As Long
   BMPHeight As Long
   ColorPlane As Integer
   BitsperPixel As Integer
   CompressionMethod As Long
   ImageSize As Long
   HorizontalResolution As Long
   VerticalResolution As Long
   ColorsinPalette As Long
   ImportantColors As Long
End Type


Public InputFile As String
Public SourcePath As String

Dim strQuery As String

Public Function CheckNumeric(ByVal KeyAscii As Integer, ByVal Expression As String) As Boolean
Dim strCheckString As String
On Error Resume Next
    strCheckString = "0123456789."
    If KeyAscii = 8 Then
        CheckNumeric = True
        Exit Function
    End If
    If KeyAscii = 46 Then
        If InStr(1, Expression, Chr(KeyAscii)) > 0 Then
            CheckNumeric = False
            Exit Function
        End If
    Else
        If InStr(1, strCheckString, Chr(KeyAscii)) <= 0 Then
            CheckNumeric = False
            Exit Function
        End If
    End If
    CheckNumeric = True
End Function

Public Function AppendCharacter(Source As String, Length As Integer, Character As String) As String
On Error Resume Next
    AppendCharacter = Mid(Trim(Source), 1, Length - 1) & Character
End Function

Public Function RemoveCharacter(Source As String, Character As String) As String
On Error Resume Next
    RemoveCharacter = Replace(Source, Character, "")
End Function

'Sub CreateTransferLog(WriteString As String)
'On Error Resume Next
'    FrmTransfer.txtDetails.Text = FrmTransfer.txtDetails.Text & WriteString
'End Sub

Public Function GetFileSize(file As String) As String
Dim fsoFile As New FileSystemObject
Dim fleFile  As file
On Error Resume Next
    Set fleFile = fsoFile.GetFile(file)
    GetFileSize = fleFile.Size & " bytes"
End Function

Public Function ConvertBMP(SourceFile As String, DestFile As String, Append As Boolean) As Boolean
Dim intReadHandle As Integer, intWriteHandle As Integer, lngFilePointer As Long, lngPrevPointer As Long, lngWritePointer As Long
Dim udtHead As BMP_HEAD, udtInfo As BMP_INFO
Dim intAvailableDataByte As Long
Dim lngTempCount As Long
Dim bytVar As Byte
Dim gblstrPrjTitle As String
gblstrPrjTitle = "Pvt Bus"
Dim lngRowCount As Long, lngColCount As Long, lngPixelWidth As Long, lngLoopCount As Long
On Error GoTo CatchError
    If Dir(SourceFile, vbNormal) = "" Then
        MsgBox "Logo file not found", vbInformation, gblstrPrjTitle
        Exit Function
    End If
    If Dir(SourceFile, vbNormal) <> "" Then
        intReadHandle = FreeFile
        Open SourceFile For Binary Access Read As #intReadHandle
        Get #intReadHandle, , udtHead
        If EOF(intReadHandle) Then
            MsgBox "Error While reading from logo file", vbInformation, gblstrPrjTitle
            Close #intReadHandle
            Exit Function
        End If
        Get #intReadHandle, , udtInfo
        If EOF(intReadHandle) Then
            MsgBox "Error While reading from logo file", vbInformation, gblstrPrjTitle
            Close #intReadHandle
            Exit Function
        End If
        intWriteHandle = FreeFile
        If Append = True Then
            If Dir(DestFile, vbNormal) <> "" Then
                lngWritePointer = FileLen(DestFile) + 1
            Else
                lngWritePointer = 1
            End If
        Else
            If Dir(DestFile, vbNormal) <> "" Then Call Kill(DestFile)
            lngWritePointer = 1
        End If
        Open DestFile For Binary Access Write As #intWriteHandle
        intAvailableDataByte = FileLen(SourceFile) - udtHead.DataOffset
        Put #intWriteHandle, lngWritePointer, CInt(udtInfo.BMPHeight)
        lngWritePointer = lngWritePointer + 2
        lngLoopCount = 1
        For lngTempCount = 3 To NOOF_BYTES_PER_LINE
            Put #intWriteHandle, lngWritePointer, bytVar
            lngWritePointer = lngWritePointer + 1
        Next lngTempCount
        lngPixelWidth = Int(udtInfo.BMPWidth) / 8
        If lngPixelWidth < ((udtInfo.BMPWidth) / 8) Then lngPixelWidth = lngPixelWidth + 1
        lngFilePointer = (FileLen(SourceFile) - lngPixelWidth) + 1
        lngLoopCount = 1
         Do While (intAvailableDataByte)
            For lngColCount = lngLoopCount To lngPixelWidth
                Get #intReadHandle, lngFilePointer, bytVar
                If intAvailableDataByte < 1 Then Exit Do
                If EOF(intReadHandle) Then Exit Do
                Put #intWriteHandle, lngWritePointer, Not bytVar
                lngWritePointer = lngWritePointer + 1
                lngFilePointer = lngFilePointer + 1
                intAvailableDataByte = intAvailableDataByte - 1
                lngRowCount = lngRowCount + 1
                bytVar = 0
            Next lngColCount
            bytVar = 0
            For lngTempCount = lngColCount To NOOF_BYTES_PER_LINE
                Put #intWriteHandle, lngWritePointer, bytVar
                lngWritePointer = lngWritePointer + 1
            Next lngTempCount
            lngFilePointer = FileLen(SourceFile) - (lngRowCount + lngPixelWidth) + 1
            If lngFilePointer <= udtHead.DataOffset Then
                lngLoopCount = udtHead.DataOffset + 1
                lngPixelWidth = lngPrevPointer
            End If
            lngPrevPointer = lngFilePointer
        Loop
    Else
        intWriteHandle = FreeFile
        If Append = True Then
            If Dir(DestFile, vbNormal) <> "" Then
                lngWritePointer = FileLen(DestFile) + 1
            Else
                lngWritePointer = 1
            End If
        Else
            If Dir(DestFile, vbNormal) <> "" Then Call Kill(DestFile)
            lngWritePointer = 1
        End If
        Open DestFile For Binary Access Write As #intWriteHandle
        intAvailableDataByte = NOOF_BYTES_PER_LINE
        Put #intWriteHandle, lngWritePointer, CInt(1)
        lngWritePointer = lngWritePointer + 2
        bytVar = 0
        For lngTempCount = 3 To NOOF_BYTES_PER_LINE * 2
            Put #intWriteHandle, lngWritePointer, bytVar
            lngWritePointer = lngWritePointer + 1
        Next lngTempCount
    End If
    Close #intReadHandle
    Close #intWriteHandle
    ConvertBMP = True
    Exit Function
CatchError:
    If intReadHandle > 0 Then Close #intReadHandle
    If intWriteHandle > 0 Then Close #intWriteHandle
End Function
Public Function ConvertLanBMP(SourceFile As String, DestFile As String, Append As Boolean) As Boolean

Dim intReadHandle As Integer, intWriteHandle As Integer, lngFilePointer As Long, lngPrevPointer As Long, lngWritePointer As Long
Dim udtHead As BMP_HEAD, udtInfo As BMP_INFO
Dim intAvailableDataByte As Long
Dim lngTempCount As Long
Dim bytVar As Byte
Dim gblstrPrjTitle As String
gblstrPrjTitle = "CONVET bMP TO LOGO"
Dim lngRowCount As Long, lngColCount As Long, lngPixelWidth As Long, lngLoopCount As Long
On Error GoTo CatchError
    If SourceFile <> "" Then
'        MsgBox "Logo file not found", vbInformation, gblstrPrjTitle
'        Exit Function
'    End If
    If Dir(SourceFile, vbNormal) <> "" Then
        intWriteHandle = FreeFile
        If Append = True Then
            If Dir(DestFile, vbNormal) <> "" Then
                lngWritePointer = FileLen(DestFile)
                If FileLen(DestFile) = 0 Then lngWritePointer = 1
                If FileLen(DestFile) = 1 Then lngWritePointer = 2
            Else
                lngWritePointer = 1
            End If
        Else
            If Dir(DestFile, vbNormal) <> "" Then Call Kill(DestFile)
            lngWritePointer = 1
        End If
        Open DestFile For Binary Access Write As #intWriteHandle
        intReadHandle = FreeFile
        Open SourceFile For Binary Access Read As #intReadHandle
        While EOF(intReadHandle) <> True
            Get #intReadHandle, , bytVar
            
            Put #intWriteHandle, lngWritePointer, bytVar
            lngWritePointer = lngWritePointer + 1
        Wend
     End If
    Else
         intWriteHandle = FreeFile
        If Append = True Then
            If Dir(DestFile, vbNormal) <> "" Then
                lngWritePointer = FileLen(DestFile)
              '  If FileLen(DestFile) <> 0 Then lngWritePointer = lngWritePointer + 1
            Else
                lngWritePointer = 1
            End If
        Else
            If Dir(DestFile, vbNormal) <> "" Then Call Kill(DestFile)
            lngWritePointer = 1
        End If
        Open DestFile For Binary Access Write As #intWriteHandle
        intReadHandle = FreeFile
        
            Put #intWriteHandle, lngWritePointer, "|"
            lngWritePointer = lngWritePointer + 1
            Put #intWriteHandle, lngWritePointer, bytVar
'
'            lngWritePointer = lngWritePointer + 1
            'Put #intWriteHandle, lngWritePointer, 0
        'Wend
'''        Next lngTempCount
   
End If
    Close #intReadHandle
    Close #intWriteHandle
    ConvertLanBMP = True
    Exit Function
CatchError:
    If intReadHandle > 0 Then Close #intReadHandle
    If intWriteHandle > 0 Then Close #intWriteHandle
End Function



