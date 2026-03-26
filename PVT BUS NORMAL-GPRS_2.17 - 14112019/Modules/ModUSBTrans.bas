Attribute VB_Name = "ModUSBTrans"

Public Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwmilliseconds As Long)

Public WriteBuffer As String, ReadBuffer As String
Public USB_FLAG As Boolean

 Dim VendorID As Long
 Dim ProductID As Long
 Dim Manufacturer As String
 Dim SerialNumber As String
 Dim DeviceName As String
 Dim OutputReportID As String





'''''''''''''''''''''form1''''''''''''''
Private Type mdeviceList
    DeviceName As Long      ' Device name
    Manufacturer As Long    ' Manufacturer
    SerialNumber As Long    ' Serial number
    VendorID As Long        ' Vendor ID
    ProductID As Long       ' Product ID
    InputReportLen As Long  ' Length of HID input report (bytes)
    OutputReportLen As Long ' Length of HID output report (bytes)
    Interface As Long       ' Interface (added 3/12/07)
    Collection As Long      ' Collection (added 3/12/07)
End Type

' This constant defines the maximum number of USB devices we will
' attempt to retreive using the GetList() function.
Private Const MAX_DEVICES = 10

Private Declare Function SetInstance Lib "UsbHidApi" (ByVal val As Long) As Long
Private Declare Function GetLibVersion Lib "UsbHidApi.dll" (ByVal pBuf As String) As Long
Private Declare Sub CloseWrite Lib "UsbHidApi.dll" ()
Private Declare Sub CloseRead Lib "UsbHidApi.dll" ()
Private Declare Function Read Lib "UsbHidApi.dll" (ByRef pBuf As Byte) As Long
Private Declare Function Write2 Lib "UsbHidApi.dll" Alias "Write" (ByRef pBuf As Byte) As Long
Private Declare Sub ShowVersion Lib "UsbHidApi.dll" ()
Private Declare Sub GetReportLengths Lib "UsbHidApi.dll" (ByRef inlen As Long, ByRef outlen As Long)
Private Declare Function Open2 Lib "UsbHidApi.dll" Alias "Open" (ByVal VendorID As Long, ByVal ProductID As Long, ByVal Manufacturer As String, ByVal SerialNumber As String, ByVal DeviceName As String, ByVal bAsync As Long) As Long
Private Declare Function GetList Lib "UsbHidApi.dll" (ByVal VendorID As Long, ByVal ProductID As Long, ByVal Manufacturer As String, ByVal SerialNumber As String, ByVal DeviceName As String, ByRef DeviceList As mdeviceList, ByVal MaxDevices As Long) As Long
Private Declare Sub SetCollection Lib "UsbHidApi.dll" (ByVal Col)
Private Declare Function GetCollection Lib "UsbHidApi.dll" () As Long
Private Declare Sub SetInterface Lib "UsbHidApi.dll" (ByVal iface)
Private Declare Function GetInterface Lib "UsbHidApi.dll" () As Long

Dim StrDeviceName(MAX_DEVICES) As String
Dim StrManufacturer(MAX_DEVICES) As String
Dim StrSerialNumber(MAX_DEVICES) As String
Dim StrVendorID(MAX_DEVICES) As String
Dim StrProductID(MAX_DEVICES) As String
Dim StrInReportLen(MAX_DEVICES) As String
Dim StrOutReportLen(MAX_DEVICES) As String
Dim Interfaces(MAX_DEVICES) As Long
Dim Collections(MAX_DEVICES) As Long



Private Const PACKET_SIZE = 64
Private Const DATA_SIZE = 59
Private Const DATA_START = PACKET_SIZE - DATA_SIZE
Private Const USB_TimeOut = 300

'Public VendorID As Long
'Public ProductID As Long
'Public Manufacturer As String
'Public SerialNumber As String
'Public DeviceName As String
'Public InputReportLen As Long
'Public OutputReportLen As Long
'Public bAsync As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''Dialog.frm'''''''''''''''''''''''''''
Public k As Integer
'Public VendorID As Long
'Public ProductID As Long
'Public DeviceName As String
'Public Manufacturer As String
'Public SerialNumber As String
Public InputReportLen As Long
Public OutputReportLen As Long
Public Interface As Long
Public Collection As Long
Public Buf As String, filename As String
Public TPacket As Integer
Dim sPacket() As String
Dim Fso As New FileSystemObject, file As TextStream, log As TextStream
Dim t As Long, PBINc As Integer, Lost As Integer
'Dim USB_TimeOut As Long
Public USB_Path As String
Public unvFlg As Boolean
Public Mode As String
Public TimeOut As Double

''''''''fro file details'''''''''
Public fname() As String, Org_Size() As String, Size() As String, S() As String, no_of_files  As Integer
Public Battery_Level As Integer, AMP_DTE As String, AMP_TME As String


Private Type PacketData
    PStatus As Byte
    DataCount As Long
    DATA As String * 59
End Type

'''''''''''''''''''''''''''''''''



'Private Declare Function Open2 Lib "UsbHidApi.dll" Alias "Open" (ByVal VendorID As Long, ByVal ProductID As Long, ByVal Manufacturer As String, ByVal SerialNumber As String, ByVal DeviceName As String, ByVal bAsync As Long) As Long
'Private Declare Sub CloseWrite Lib "UsbHidApi.dll" ()
'Private Declare Sub CloseRead Lib "UsbHidApi.dll" ()
'Private Declare Function Read Lib "UsbHidApi.dll" (ByRef pBuf As Byte) As Long
'Private Declare Function Write2 Lib "UsbHidApi.dll" Alias "Write" (ByRef pBuf As Byte) As Long
'Private Declare Sub SetCollection Lib "UsbHidApi.dll" (ByVal col As Long)
'Private Declare Function GetCollection Lib "UsbHidApi.dll" () As Long
'Private Declare Sub SetInterface Lib "UsbHidApi.dll" (ByVal iface As Long)
'Private Declare Function GetInterface Lib "UsbHidApi.dll" () As Long
'Private Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwmilliseconds As Long)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub WriteData()
   Dim iDx As Integer
   Dim wresult As Integer
   Dim rdresult As Integer
   Dim retry As Integer
   Dim send_buf(65) As Byte
   
'   TimeOut = 5
   ' Clear the buffers
   For iDx = 0 To 64
      send_buf(iDx) = 0
   Next
'   Status = ""
'   Refresh
'   USB_TimeOut = Timer
   send_buf(0) = val("&H" & OutputReportID)
   Dim a() As String, I As Integer, Buf As String
   For iDx = 1 To 64 'Len(WriteBuffer(1))
         If iDx < Len(WriteBuffer) Then
            send_buf(iDx) = Asc(Mid(Trim(WriteBuffer), iDx, 1))
         Else
            send_buf(iDx) = Asc(Chr(0))
         End If
   Next
   ' Clear the fields
   WriteBuffer = ""
   retry = 0
   wresult = Write2(send_buf(0))
'   Debug.Print send_buf(0)
    retry = retry + 1
    Call Check_USBConnection
End Sub
Private Sub ReadData()
  Dim iDx As Integer
   Dim wresult As Integer
   Dim rdresult As Long
   Dim retry As Integer
   'Dim buf As String
'''''''''''''''''''''''''''''''''''''''''''''''Do While Not UCase(Chr(k)) = "Q"

   DoEvents
   Dim recv_buf(64) As Byte  'changed from 1000 to 64'
'Do
   ' Clear the buffers
'''''   For idx = 0 To 63  'changed from 99 to 63'
'''''      recv_buf(idx) = 0
'''''   Next
   ' Clear the status box
'   Status = ""
'''''   Refresh
   ' Clear the fields
'''''   ReadBuffer(0) = ""
'''''   ReadBuffer(0).Refresh
'''''   InputReportID(1) = ""
'''''   InputReportID(1).Refresh
   retry = 1
   ' Attempt to read the response
   rdresult = Read(recv_buf(0))
   ReDim a(64) As String  'changed from 1000 to 64'
   For iDx = 1 To 64
       Buf = Buf & Chr(recv_buf(iDx))
   Next
End Sub

Public Function processFile(filename As String) As Boolean
On Error GoTo err
Dim hndlFile As Integer, byt As Byte, hndlPackets As Integer
Dim FSize As Long, I As Integer, c As Integer
hndlFile = FreeFile
Open USB_Path & "\" & filename For Binary Access Read As hndlFile
FSize = LOF(hndlFile)
'''''''''''''''for 0 size files'''''''
''''''If FSize = 0 Then
''''''  ReDim Spacket(1) As String, PSize(1) As Integer
''''''  Spacket(1) = Chr(0)
''''''  PSize(1) = 1
''''''  Close hndlFile
''''''  processFile = True
''''''  TPacket = 1
''''''  Exit Function
''''''End If
'''''''''''''''''''''''''''''''''''''''''''''''''
TPacket = Round(FSize / 63, 0)
If TPacket < (FSize / 63) Then TPacket = TPacket + 1
'showlog "Total no.of Packets = " & TPacket
''frmGPRS.lstStatus.ListIndex = frmGPRS.lstStatus.ListCount - 1
ReDim sPacket(0 To TPacket + 1) As String ', PSize(0 To TPacket + 1) As Integer
I = 1: c = 0

Do While Not EOF(hndlFile)
   Get #hndlFile, , byt
   If EOF(hndlFile) = True Then Exit Do
'   If hndlPackets = 0 Then
'      hndlPackets = FreeFile
'      Open USB_Path & "\send" & i & ".snd" For Binary Access Write As hndlPackets
'      Put #hndlPackets, , byt
'   End If
   If c = 63 Then
     ' Debug.Print Spacket(i)
'      PSize(i) = c
      c = 0
      Close hndlPackets: hndlPackets = 0
      I = I + 1
   End If
   sPacket(I) = sPacket(I) & Chr(byt)
   c = c + 1
Loop
'PSize(i) = c
'Close hndlPackets: hndlPackets = 0
Close hndlFile
processFile = True
Exit Function
err:
     processFile = False
End Function

Public Function GetDevices() As Boolean

    GetDevices = True
   Dim dcnt As Integer

   Dim Buf As String * 20  ' Without the specifier, the string len will be 0!
   Dim val As Long
   Dim inlen As Long
   Dim outlen As Long
   
   
   ' Get list of available USB devices.  (This call will only
   ' retrieve HID-class devices.)
   Dim Devices(MAX_DEVICES) As mdeviceList
   
   ' Get and display the DLL version
   'Val = GetLibVersion(buf)
   'Call MsgBox("GetLibVersion:  " & buf)

   ' Alternative version display
   'ShowVersion

   ' Set the object instance
   ' (This is optional since '0' is the default.)
   val = SetInstance(0)

   DeviceName = vbNullString    ' NULL => Any device
   SerialNumber = vbNullString ' NULL => Any serial number
   Manufacturer = vbNullString ' NULL => Any manufacturer
   VendorID = 65535     ' 0xffff => Any vendor ID
   ProductID = 65535    ' 0xffff => Any product ID

   ' Prepare for the ensuing call to retrieve devices
   Dim cnt As Integer
   For cnt = 0 To MAX_DEVICES

      ' Allocate space to each string.
      '(I'm sure there's a better way to do this.)
      StrDeviceName(cnt) = "                                                  "
      StrManufacturer(cnt) = "                                                  "
      StrSerialNumber(cnt) = "                                                  "

      ' Convert from Unicode to ANSI
      StrDeviceName(cnt) = StrConv(StrDeviceName(cnt), vbFromUnicode)
      StrManufacturer(cnt) = StrConv(StrManufacturer(cnt), vbFromUnicode)
      StrSerialNumber(cnt) = StrConv(StrSerialNumber(cnt), vbFromUnicode)

      ' Assign string addresses to structure fields
      Devices(cnt).DeviceName = StrPtr(StrDeviceName(cnt))
      Devices(cnt).Manufacturer = StrPtr(StrManufacturer(cnt))
      Devices(cnt).SerialNumber = StrPtr(StrSerialNumber(cnt))

   Next

   '
   ' Call the DLL GetList() routine to get a list of available
   ' USB HID devices.
   '
  
   dcnt = GetList(VendorID, ProductID, Manufacturer, SerialNumber, DeviceName, Devices(0), MAX_DEVICES)
   'Call MsgBox("GetList returned " & Val)
   If (dcnt <= 0) Then
      Mode = "COM"
      'Call MsgBox("No USB HID devices found - terminating!")
   End If

   ' Process the list of device(s) we found
   For cnt = 0 To (dcnt - 1)
  
      ' String fields must be converted back to Unicode
'      StrDeviceName(cnt) = TrimChr(StrConv(StrDeviceName(cnt), vbUnicode))
'      StrManufacturer(cnt) = TrimChr(StrConv(StrManufacturer(cnt), vbUnicode))
'      StrSerialNumber(cnt) = TrimChr(StrConv(StrSerialNumber(cnt), vbUnicode))
'      StrVendorID(cnt) = TrimChr(Str(Devices(cnt).VendorID))
'      StrProductID(cnt) = TrimChr(Str(Devices(cnt).ProductID))
'      StrInReportLen(cnt) = TrimChr(Str(Devices(cnt).InputReportLen))
'      StrOutReportLen(cnt) = TrimChr(Str(Devices(cnt).OutputReportLen))
      DeviceName = TrimChr(StrConv(StrDeviceName(cnt), vbUnicode))
      Manufacturer = TrimChr(StrConv(StrManufacturer(cnt), vbUnicode))
      SerialNumber = TrimChr(StrConv(StrSerialNumber(cnt), vbUnicode))
      VendorID = TrimChr(str(Devices(cnt).VendorID))
      ProductID = TrimChr(str(Devices(cnt).ProductID))
      InReportLen = TrimChr(str(Devices(cnt).InputReportLen))
      OutReportLen = TrimChr(str(Devices(cnt).OutputReportLen))
      ' Store interface(s) and/or collection(s)
      ' added 3/12/07
      Interfaces(cnt) = Devices(cnt).Interface
      Collections(cnt) = Devices(cnt).Collection
      If UCase(DeviceName) = "PALMTEC" And UCase(Manufacturer) = "SOFTLAND" Then
          val = Open2(VendorID, ProductID, Manufacturer, SerialNumber, DeviceName, True)
          Exit For
      End If
   Next

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
'      val = Open2(49745, 5889, "Softland", "SIL123456789", "Palmtec", True)

If val <> 1 Then
   GetDevices = False
   
Else
   GetDevices = True
   WriteBuffer = "ttttttttttttttttttttttttttt" ''''setting transfer mode in amphibia'''


   Call WriteData
'   gblblnUSBInTransfer = True
'    With frmMainDispaly
'        .lblMessage.Caption = "Device in Transfer Mode "
'        .Picture1.Visible = True
'        .fraMessage.Top = .Height - (.StatusBar1.Height + .fraMessage.Height)
'        .fraMessage.Left = .Width - .fraMessage.Width
'        .fraMessage.Visible = True
'        .Timer1.Enabled = True
'    End With
End If
Exit Function
err:
   GetDevices = False
End Function

Public Function Refresh_Device() As Boolean
   Dim Device As ListItem
   Dim nDevices As Integer
   Dim dev As Integer
   ' Disable the "Open" button
'   OpenDev(0).Enabled = False
   
   ' Clear the listview contents
'   DeviceList.ListItems.Clear
   
   ' Retrieve list of USB HID devices
   nDevices = GetDevices()
   ' Display the retreived device(s)
   If (nDevices > 0) Then
      For dev = 0 To (nDevices - 1)
         ' Populate the listview control
         Set Device = DeviceList.ListItems.Add
         If UCase(Trim(Mid(StrDeviceName(dev), 1, InStr(StrDeviceName(dev), Chr(0)) - 1))) = "PALMTEC" Then
            Device.Text = StrDeviceName(dev)
            Device.SubItems(1) = StrManufacturer(dev)
            Device.SubItems(2) = StrVendorID(dev)
            Device.SubItems(3) = StrProductID(dev)
            Device.SubItems(4) = StrSerialNumber(dev)
            Device.SubItems(5) = StrInReportLen(dev)
            Device.SubItems(6) = StrOutReportLen(dev)
         End If
      Next
   End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Write_USB(filename As String) As Boolean
On Error GoTo err
Dim strYear As String
    Write_USB = True
    delayTr 1
    TransferTimeOut = 10
    TimeOut = TransferTimeOut
   
    If filename = "SHUTDOWN" Then
        WriteBuffer = "s@"
        Call WriteData
        Write_USB = True
        Exit Function
    End If
    
    If Mid$(filename, 1, 11) = "DATEANDTIME" Then
        filename = Mid$(filename, InStr(1, filename, "-") + 1)
'        WriteBuffer = "t@" & Trim(Filename) & Chr(0)
        WriteBuffer = "D@" & Trim(filename) & Chr(0)
        Call WriteData
        TimeOut = Timer
        Do While Not InStr(Buf, "t") > 0
            If IsUSBTimeOut(TimeOut, USB_TimeOut) Then
                Write_USB = True
                Exit Function
            End If
            Buf = ""
            Call ReadData
            DoEvents
        Loop
        If Mid$(Buf, 2, 7) = "Success" Then
            Write_USB = True
        Else
            Write_USB = False
        End If
        Exit Function
    End If
    
    If filename = "DELETETKTS" Then
        WriteBuffer = "r@" & filename
        Call WriteData
        Write_USB = True
        Exit Function
    End If
    If filename = "VERISON_SUCESS" Then
        WriteBuffer = "v@"
        Call WriteData
        Write_USB = True
        Exit Function
    End If
    If filename = "FORMATMAC" Then
        WriteBuffer = "f@"
        Call WriteData
        Write_USB = True
        Exit Function
    End If
    
    If filename = "UPLOADSUCCESS" Then  'SANGEETHA
        WriteBuffer = "o@"
        Call WriteData
        Write_USB = True
        Exit Function
    End If
    
    Transfer.pBar.Value = 0
    If Dir(USB_Path & "\" & filename) = "" Then
        MsgBox "File Not found", vbInformation, gblstrPrjTitle
        Exit Function
    End If
    processFile (filename)
    Dim hndl As Integer
    hndl = FreeFile
    Open USB_Path & "\" & filename For Binary Access Read As hndl
    FSize = LOF(hndl)
    Close #hndl
    If FSize = 0 Then
        Write_USB = True
        Exit Function
    End If
           
    WriteBuffer = "b" & filename & "#" & FSize & "#"
    Call WriteData
        
    '''''''''waiting for conformation  'k'''''''
    TimeOut = Timer
    Do While Not InStr(Buf, "k") > 0
        If IsUSBTimeOut(TimeOut, USB_TimeOut) Then
            Write_USB = False
            Exit Function
        End If
        Buf = ""
        Call ReadData
        DoEvents
    Loop
    TimeOut = TransferTimeOut
    '''''''''''''''''''''''''''''''''''
    ''''''''sending file'''''''
    Dim I As Integer
    t = Timer
    PBINc = 100 / TPacket
    Transfer.txtTotalPacket = TPacket
    Transfer.pBar.MinValue = 0
    Transfer.pBar.MaxValue = TPacket
    For I = 1 To TPacket
        Buf = ""
        WriteBuffer = "d" & sPacket(I) & "V"
        Call WriteData
        Transfer.txtPacketNo = I
        Transfer.pBar.Value = I
        DoEvents
        If I Mod 260 = 0 Then
            TimeOut = Timer
            Do While Not InStr(Buf, "k") > 0
                If I = TPacket Then Exit Do
                If IsUSBTimeOut(TimeOut, USB_TimeOut) Then
                    Write_USB = False
                    Exit Function
                End If
                Buf = ""
                Call ReadData
                DoEvents
            Loop
        End If
'        If (PB1.Value + PBINc) < 100 Then
'            PB1.Value = PB1.Value + PBINc
'        Else
'            PBINc = (PB1.Value + PBINc) - 100
'            PB1.Value = PB1.Value + PBINc
'        End If
    Next I
        ''''''''''''''''''''''''''''
'    file.Close
    t = Timer - t
'    MsgBox "Sending complete  in " & T & " sec"
    Exit Function
err:
   Write_USB = False
End Function

Public Function Read_USB(filename As String) As Boolean
Dim Datasize As Integer, BufferSize As Integer, BufferStartPos As Byte, DataStartPos As Byte, Status As Long
Dim intFP As Long
On Error GoTo err
'    Read_USB = True
    delayTr 1
    
    WriteBuffer = "a" & filename & "#"
            
    Call WriteData
            
    ''''for file size'''
'    BufferStartPos = 1
'    BufferSize = PACKET_SIZE
'    DataStartPos = DATA_START + 1
'    Datasize = DATA_SIZE
    
    BufferStartPos = 1
    BufferSize = 64
    DataStartPos = ((BufferStartPos + Len(Status)) + 1)
    Datasize = (BufferSize - (BufferStartPos + Len(Status)))
    TimeOut = Timer
    Dim FSize As Long, p As Integer
    Do While Not p > 0
        If IsUSBTimeOut(TimeOut, USB_TimeOut) Then
            Read_USB = False
            Exit Function
        End If
        Buf = ""
        
        Call ReadData
'        Debug.Print "Buffer" & "-- " & Buf
        If Buf <> "" Then
            p = InStr(Buf, "s")
            If p > 0 Then
                FSize = val(Mid(Buf, p + 1, InStr(Buf, "#") - 1))
                If FSize Mod Datasize = 0 Then
                    Transfer.txtTotalPacket = FSize / Datasize
'                    FSize = FSize / Datasize
                Else
                    Transfer.txtTotalPacket = Fix((FSize / Datasize)) + 1
'                    FSize = Fix((FSize / Datasize)) + 1
                End If
            End If
        End If
    Loop
        
    '''''''''''sending confirmation''''
    
    WriteBuffer = "kkkkkkkkkkkkkkkkkkkkkkkkkkkkkk"
    Call WriteData
             
    ''''''''''''''''''''''''''''''''''
            
    If FSize = 0 Then
        Read_USB = True
        Exit Function
    End If
    Transfer.pBar.MinValue = 0
    Transfer.pBar.Value = 0
    Transfer.pBar.MaxValue = Transfer.txtTotalPacket
    Transfer.pBar.Visible = True
    
    ''''opening file;'''''
    Set file = Fso.OpenTextFile(USB_Path & "\" & filename, ForWriting, True)
    
    ''''''''''''''''''''''''''''
           
    '''''''''''''getting file'''''''''''''
    t = Timer
    Dim PSize As Long, PrePkt As Long, PckRecv As String
    Dim PckLost(50) As Integer, bufWrite As String, j As Integer
    PrePkt = 0
    Lost = 1
    bufWrite = ""
    Transfer.txtPacketNo = 0
    TimeOut = Timer
    Dim inthan As Integer, inthan1 As Integer
    Dim pData As PacketData
    
'    If Dir(App.Path & "\t1") <> "" Then Call Kill(App.Path & "\t1")
'    inthan = FreeFile
'    Open App.Path & "\t1" For Binary Access Write As #inthan
'    Close #inthan
    
'    Do While Not PSize > FSize
    Do While FSize > PSize
        If IsUSBTimeOut(TimeOut, USB_TimeOut) Then
            Read_USB = False
            Close #inthan
            file.Close
            Exit Function
        End If
        Buf = ""
        
        Call ReadData
'        If Dir(App.Path & "\t1") <> "" Then Call Kill(App.Path & "\t1")
''        intFP = FileLen(App.Path & "\t1")
'        inthan = FreeFile
'        Open App.Path & "\t1" For Binary Access Write As #inthan
'        Put #inthan, , Buf
'        Close #inthan
'        inthan = FreeFile
'        Open App.Path & "\t1" For Binary Access Read As #inthan
'        Get #inthan, , pData
'        Close #inthan
'
        Status = Str2Long(Mid(Buf, 2, 4))
        
'        Status = CLng(Asc(Mid(Buf, 2, 5)))
'        Print #inthan, Status
'        Status = val(Mid(Buf, 2, 4))
'        Debug.Print Status
        PckRecv = PckRecv & Format(Status, "0000")
'        Debug.Print "Data Buffer --" & Buf
'        If Mid(pData.DATA, 1, 1) = "d" Then

        If Mid(Buf, 1, 1) = "d" Then

            If Status <> PrePkt And Status <> 0 Then
                Transfer.pBar.Value = val(Transfer.txtPacketNo) + 1
                Transfer.txtPacketNo = Status
                PSize = PSize + Datasize ' (Len(buf) - 2)
                Transfer.txtTotalPacket.Text = PSize
'                j = j + 1
'                If Status > 0 Then
'                    Debug.Print Status
'                End If
                '''''''''checking for missing packet'''
                       
'                Dim k As Integer
'                For k = PrePkt + 1 To Status - 1
'                    PckLost(Lost) = Lost
'                    Lost = Lost + 1
'                Next k
                       
                ''''''''''''''''''''''''''''''''''''''''''''
'                Debug.Print Buf
                
                If PSize > FSize Then
                    Debug.Print "Last Pack"
                    Dim e As Integer
'                    e = PSize - FSize '- (PSize)
                    e = FSize - (PSize - Datasize)
'                    file.Write Mid(pData.DATA, DataStartPos, e)
                    file.Write Mid(Buf, DataStartPos, e)
                Else
                    file.Write Mid(Buf, DataStartPos)
                End If
                PrePkt = Status
                TimeOut = Timer
            End If
        End If
    Loop
    file.Close
    Close #inthan
    If Lost > 1 Then
        MsgBox "Packet Lost" & "  " & Lost
    End If
    t = Timer - t
    Read_USB = True
    Exit Function
err:
    Read_USB = False
    Close #inthan
    file.Close
            ''''''''''''''''''''''''''''''''''''''
End Function

Public Function Initiate_USB() As Boolean
  If Refresh_Device = False Then
     MsgBox "Error in detedcting Amphibia", vbInformation, gblprjTitle
  End If
   
End Function

Public Function Get_File_Details() As Boolean
On Error GoTo err
Dim buf_len As Integer
Dim buf_temp As String, PacketNo As Integer, PrePacket As Integer
Get_File_Details = True
'Set File = fso.OpenTextFile(App.Path & "\temp.dat", ForWriting, True)

WriteBuffer = ""
WriteBuffer = "cccccccccccc"
Call WriteData
'Do While Not Right(buf, 1) = "&"
Do While Not Left(Buf, 1) = "s"
   Buf = ""
   Call ReadData
Loop
'   If InStr(buf, Chr(0)) > 1 Then
      If Left(Buf, 1) = "s" Then
            PacketNo = Asc(Mid(Buf, 2, 1))
            TotalPacket = Asc(Mid(Buf, 3, 1))
            buf_temp = buf_temp & Mid(Buf, 3)
            For I = 2 To TotalPacket
                Buf = ""
                Do While Not Left(Buf, 1) = "s"
                    Buf = ""
                    Call ReadData
                Loop
                PacketNo = Asc(Mid(Buf, 2, 1))
                Debug.Print "PNo=" & PacketNo & "   :  PrePck=" & PrePacket
                If PacketNo <> PrePacket Then
                    
                   If PacketNo = TotalPacket Then
                        buf_temp = buf_temp & Mid(Buf, 3, InStr(Buf, "&") - 2)
                   Else
                        buf_temp = buf_temp & Mid(Buf, 3)
                   End If
                   PrePacket = PacketNo
                End If
            Next I
            'If Mid(buf, 3, 1) = "&" Then
'                 If Right(buf, 1) = "&" Then
'                    buf_temp = buf_temp & Mid(buf, 3, Len(buf) - 1)
'                 Else
'                    buf_temp = buf_temp & Mid(buf, 3)
'                 End If
'            'End If
'            If Right(buf, 1) = "&" And Left(buf, 1) <> "&" Then
'                 buf_temp = buf_temp & Mid(buf, 2, Len(buf) - 1)
'            End If
      End If
'   End If
'Loop
'Dim FileName() As String, Org_Size() As String, Size() As String, S() As String
Dim no As Integer
S() = Split(buf_temp, "$")
no_of_files = UBound(S)
ReDim fname(0 To no_of_files) As String, Org_Size(0 To no_of_files) As String, Size(0 To no_of_files) As String
cnt = 0
buf_temp = Mid(buf_temp, 2)
Do While Not buf_temp = ""
   fname(cnt) = Mid(buf_temp, InStr(buf_temp, "*") + 1, InStr(buf_temp, "$") - 2)
   buf_temp = Mid(buf_temp, InStr(buf_temp, "$") + 1)
   Size(cnt) = Mid(buf_temp, 1, InStr(buf_temp, "#") - 1)
   buf_temp = Mid(buf_temp, InStr(buf_temp, "#") + 1)
   Org_Size(cnt) = Mid(buf_temp, 1, InStr(buf_temp, "*") - 1)
   buf_temp = Mid(buf_temp, InStr(buf_temp, "*"))
   If buf_temp = "*&" Then buf_temp = ""
   cnt = cnt + 1
Loop


''''''''''''''''''''''''''''''
Buf = buf_temp
Exit Function
err:
MsgBox err.Description
     Get_File_Details = False
    '''''''''''''''''''''''''
End Function

Public Function Check_USBConnection()

End Function

Public Function Disconnect_USB() As Boolean
 On Error Resume Next
    WriteBuffer = "xxxxxxxxxxxxx"
    Call WriteData
'    With frmMainDispaly
'        .lblMessage.Caption = "Device exit from Transfer Mode"
'        .fraMessage.Visible = True
'        .Timer1.Enabled = True
'    End With
'    gblblnUSBInTransfer = False
End Function

Public Function Get_Date_Time() As Boolean
On Error GoTo err
Get_Date_Time = True
WriteBuffer = ""
WriteBuffer = "eeeeeeeeeeeeeeeeee"
Call WriteData
Buf = ""
Do While TrimChr(Buf) = ""
   Buf = ""
   Call ReadData
Loop
'If Mid(buf, 1, 1) = "^" Then
'   Battery_Level = Mid(buf, 2, InStr(2, buf, "^"))
   Buf = Mid(Buf, InStr(5, Buf, "^") + 1)
'End If
If Mid(Buf, 1, 1) = "@" Then
   AMP_TME = Mid(Buf, 2, InStr(2, Buf, "@") - 2)
   Buf = Mid(Buf, InStr(2, Buf, "@") + 1)
End If
If Mid(Buf, 1, 1) = "!" Then
   AMP_DTE = Mid(Buf, 2, InStr(2, Buf, "!") - 2)
   Buf = Mid(Buf, InStr(2, Buf, "!") + 1)
End If
Exit Function
err:
   Get_Date_Time = False
End Function

Public Function Get_Battery() As Boolean
On Error GoTo err
Get_Battery = True
WriteBuffer = ""
WriteBuffer = "eeeeeeeeeeeeeeeeee"
Call WriteData
Buf = ""
Do While TrimChr(Buf) = ""
   Buf = ""
   Call ReadData
Loop
If Mid(Buf, 1, 1) = "^" Then
   Battery_Level = Mid(Buf, 2, InStr(2, Buf, "^"))
'   buf = Mid(buf, InStr(2, buf, "^") + 1)
End If
Exit Function
err:
   Get_Battery = False
End Function
'Public Function TrimChr(ByVal Buf As String) As String
'    Dim tmp As Integer
'    Dim Buf1 As String
'    tmp = InStr(Buf, Chr$(0))
'    If tmp > 0 Then
'        TrimChr = Trim(Mid(Buf, 1, tmp - 1))
'    Else
'        TrimChr = Trim(Buf)
'    End If
'    'TrimChr = left(Buf, InStr(1, strString, Chr$(0)) - 1) 'Mid(Buf, 1, (InStr(Buf, Chr(0)) - 1))
'End Function



Public Function Set_Date_Time(DTE As String, TME As String) As Boolean
On Error GoTo err
Set_Date_Time = True
Dim Hr As String, Min As String, Sec As String
Dim Dat As String, mon As String, Yr As String
TME = Format(TME, "HHMMSS")
Hr = Mid(TME, 1, 2)
Min = Mid(TME, 3, 2)
Sec = Mid(TME, 5, 2)
TME = Chr(Hr) & Chr(Min) & Chr(Sec)
Dat = Format(DTE, "DD")
mon = Format(DTE, "MM")
Yr = Format(DTE, "YYYY")
DTE = Chr(Dat) & Chr(mon) & CLng(Yr)
WriteBuffer = "f@" & TME & DTE & "@"
Call WriteData
Exit Function
err:
   Set_Date_Time = False
End Function

Public Function delayTr(x As Double)
    Dim Tin As Double
    Dim Tout As Double
    Tin = Timer
    Do While (Tout - Tin < x)
        Tout = Timer
    Loop
End Function


Public Function DeviceConnected() As Boolean

   Dim dcnt As Integer

   Dim Buf As String * 20  ' Without the specifier, the string len will be 0!
   Dim val As Long
   Dim inlen As Long
   Dim outlen As Long
   
   
   ' Get list of available USB devices.  (This call will only
   ' retrieve HID-class devices.)
   Dim Devices(MAX_DEVICES) As mdeviceList
   
   ' Get and display the DLL version
   'Val = GetLibVersion(buf)
   'Call MsgBox("GetLibVersion:  " & buf)

   ' Alternative version display
   'ShowVersion

   ' Set the object instance
   ' (This is optional since '0' is the default.)
   val = SetInstance(0)

   DeviceName = vbNullString   ' NULL => Any device
   SerialNumber = vbNullString ' NULL => Any serial number
   Manufacturer = vbNullString ' NULL => Any manufacturer
   VendorID = 65535     ' 0xffff => Any vendor ID
   ProductID = 65535    ' 0xffff => Any product ID

   ' Prepare for the ensuing call to retrieve devices
   Dim cnt As Integer
   For cnt = 0 To MAX_DEVICES

      ' Allocate space to each string.
      '(I'm sure there's a better way to do this.)
      StrDeviceName(cnt) = "                                                  "
      StrManufacturer(cnt) = "                                                  "
      StrSerialNumber(cnt) = "                                                  "

      ' Convert from Unicode to ANSI
      StrDeviceName(cnt) = StrConv(StrDeviceName(cnt), vbFromUnicode)
      StrManufacturer(cnt) = StrConv(StrManufacturer(cnt), vbFromUnicode)
      StrSerialNumber(cnt) = StrConv(StrSerialNumber(cnt), vbFromUnicode)

      ' Assign string addresses to structure fields
      Devices(cnt).DeviceName = StrPtr(StrDeviceName(cnt))
      Devices(cnt).Manufacturer = StrPtr(StrManufacturer(cnt))
      Devices(cnt).SerialNumber = StrPtr(StrSerialNumber(cnt))

   Next

   '
   ' Call the DLL GetList() routine to get a list of available
   ' USB HID devices.
   '
  
   dcnt = GetList(VendorID, ProductID, Manufacturer, SerialNumber, DeviceName, Devices(0), MAX_DEVICES)
   'Call MsgBox("GetList returned " & Val)
   If (dcnt <= 0) Then
      Mode = "COM"
      'Call MsgBox("No USB HID devices found - terminating!")
   End If

   ' Process the list of device(s) we found
   For cnt = 0 To (dcnt - 1)
  
      ' String fields must be converted back to Unicode
'      StrDeviceName(cnt) = TrimChr(StrConv(StrDeviceName(cnt), vbUnicode))
'      StrManufacturer(cnt) = TrimChr(StrConv(StrManufacturer(cnt), vbUnicode))
'      StrSerialNumber(cnt) = TrimChr(StrConv(StrSerialNumber(cnt), vbUnicode))
'      StrVendorID(cnt) = TrimChr(Str(Devices(cnt).VendorID))
'      StrProductID(cnt) = TrimChr(Str(Devices(cnt).ProductID))
'      StrInReportLen(cnt) = TrimChr(Str(Devices(cnt).InputReportLen))
'      StrOutReportLen(cnt) = TrimChr(Str(Devices(cnt).OutputReportLen))
      DeviceName = TrimChr(StrConv(StrDeviceName(cnt), vbUnicode))
      Manufacturer = TrimChr(StrConv(StrManufacturer(cnt), vbUnicode))
      SerialNumber = TrimChr(StrConv(StrSerialNumber(cnt), vbUnicode))
      VendorID = TrimChr(str(Devices(cnt).VendorID))
      ProductID = TrimChr(str(Devices(cnt).ProductID))
      InReportLen = TrimChr(str(Devices(cnt).InputReportLen))
      OutReportLen = TrimChr(str(Devices(cnt).OutputReportLen))
      ' Store interface(s) and/or collection(s)
      ' added 3/12/07
      Interfaces(cnt) = Devices(cnt).Interface
      Collections(cnt) = Devices(cnt).Collection
      If UCase(DeviceName) = "PALMTEC" And UCase(Manufacturer) = "SOFTLAND" Then
          val = Open2(VendorID, ProductID, Manufacturer, SerialNumber, DeviceName, True)
          Exit For
      End If
   Next

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
'      val = Open2(49745, 5889, "Softland", "SIL123456789", "Palmtec", True)

If val <> 1 Then
   DeviceConnected = False
   
Else
   DeviceConnected = True
'   WriteBuffer = "ttttttttttttttttttttttttttt" ''''setting transfer mode in amphibia'''
'
'
'   Call WriteData
'  frmmain.cmdFileDets.Value = True
End If
Exit Function
err:
   DeviceConnected = False
End Function

Function IsUSBTimeOut(ByVal StartTime As Double, TimeOutPeriod As Long) As Boolean
On Error Resume Next
    If (Timer - StartTime) > TimeOutPeriod Then
        IsUSBTimeOut = True
    Else
        IsUSBTimeOut = False
    End If
End Function
