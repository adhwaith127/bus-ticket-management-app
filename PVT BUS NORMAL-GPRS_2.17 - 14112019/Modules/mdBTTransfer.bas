Attribute VB_Name = "mdBTTransfer"
Option Explicit

Public Const BT_TIMEOUT = 20
Public Const BT_DATA_SIZE = 2048


Public Const CMD_SYNC = &H1
Public Const CMD_UPLOAD = &H2
Public Const CMD_DOWNLOAD = &H3
Public Const CMD_DELETE_DATA = &H4
Public Const CMD_FORMAT = &H5
Public Const CMD_SET_DATE_TIME = &H6
Public Const CMD_SHUTDOWN = &H7
Public Const CMD_STATUS = &H8


' Error / Status

Public Const C_OK = &H0                   '  Operation  succeed
Public Const C_CMDTIMEOUT = &H1           '  Instruction  overtime
Public Const C_CMDERR = &H3               '  Wrong  Commands
Public Const C_CMDINVALIDFILE = &H4       '  Invalid FileName
Public Const C_CMDFILE_NOT_FOUND = &H5    '  File Not Found
Public Const C_CMD_FILE_ERROR = &H6       '  File creation error in master
Public Const C_CMD_EXIT_TRN = &H7         '  Exiting from transfer

Public Const C_EOFFILE = &H10             '  End of ile
Public Const C_READYTOSEND = &H11         '  Device ready to send file details
Public Const C_NXT_PKT_PENDING = &H12     '  Next packet is available for sending
Public Const C_RECV_READY = &H13          '  Device is ready to recieve data from master


Public Const C_CRCERROR = &H82            '  Checksum  error
Public Const C_DEVICENOTFOUND = &HA0      '  No device found
Public Const C_NOTOK = &HFF               '  General  error


Public Const START = &H2                  '  Start Flag




Public strMessage As String
Public bRcvdByte As Byte
Public bDataReadyTimer As Byte
Public bStatus As Byte
Public blnDataReady As Boolean
Public blnStatus As Boolean

Public lRcvdByteCount As Long
Public lBytesToBeRead As Long
Public lTimer As Long
Public lFileLen As Long
Public lCheckSum As Long

Public sRcvData As String
Public sStatus As String
Public sPacket As String

Public iHdl As Integer

Public strFileName As String
Public strFilePath As String




Public Function SendCommand(sComd As String)
Dim iVal As Integer
    sStatus = ""
    For iVal = 1 To Len(sComd)
        If SendByte(Asc(Mid$(sComd, iVal, 1))) = False Then Exit For
    Next
End Function

Public Function InitTransfer() As Boolean
On Error GoTo err

    blnStatus = False
    sPacket = ""
    sPacket = sPacket & Chr(&H2)
    sPacket = sPacket & Chr(&H7)
    sPacket = sPacket & Chr(&H0)
    sPacket = sPacket & Chr(CMD_SYNC)
    sPacket = sPacket & Chr(C_OK)
    sPacket = sPacket & Chr(&H0)
    sPacket = sPacket & Chr(&HA)
    sPacket = sPacket & Chr(&H0)
    sPacket = sPacket & Chr(&H0)
    sPacket = sPacket & Chr(&H0)
    SendCommand (sPacket)
    lTimer = 5
    Do
        DoEvents
        If blnStatus = True Then
            If Asc(Mid$(sStatus, 5, 1)) <> 0 Then
                strMessage = "Command Error!"
                InitTransfer = False
            Else
                InitTransfer = True
                strMessage = "Connected with palmtec device"
            End If
            Exit Do
        End If
        If lTimer = 0 Then
            strMessage = "Time Out"
            InitTransfer = False
            Exit Do
        End If
    Loop
    
    Exit Function
err:
    strMessage = "Error!" & vbCrLf & "InitTransfer : " & err.Description
End Function

Public Function CreatePacket(bCmd As Byte, bSts As Byte, strData As String) As Boolean
On Error GoTo err
    
    Dim iVal As Integer
    CreatePacket = False
    
    sPacket = ""
    sPacket = sPacket & Chr(&H2)
    
    strData = strData
    iVal = Len(strData) + 6
    sPacket = sPacket & Mid$(Int2Str(iVal), 1, 1)
    sPacket = sPacket & Mid$(Int2Str(iVal), 2, 1)
    
    
    sPacket = sPacket & Chr(bCmd)
    sPacket = sPacket & Chr(bSts)
    sPacket = sPacket & strData
    
    lCheckSum = 0
    For iVal = 1 To Len(sPacket)
        lCheckSum = lCheckSum + Asc(Mid$(sPacket, iVal, 1))
    Next
    
    For iVal = 1 To 4
        sPacket = sPacket & Mid$(Long2Str(lCheckSum), iVal, 1)
    Next

    CreatePacket = True
    Exit Function
err:
    strMessage = "Error!" & vbCrLf & "CreatePacket : " & err.Description
End Function

Public Function Long2Str(l As Long) As String
'On error goto ERR1
    Long2Str = Chr(l And &HFF)
    'Debug.Print "Byte1 = " & Hex(l And &HFF)
    Long2Str = Long2Str & Chr(((l And &HFF00) / &H100) And &HFF)
    'Debug.Print "Byte2 = " & Hex(((l And &HFF00) / &H100) And &HFF)
    Long2Str = Long2Str & Chr(((l And &HFF0000) / &H10000) And &HFF)
    'Debug.Print "Byte3 = " & Hex(((l And &HFF0000) / &H10000) And &HFF)
    Long2Str = Long2Str & Chr(((l And &HFF000000) / &H1000000) And &HFF)
    'Debug.Print "Byte4 = " & Hex(((l And &HFF000000) / &H1000000) And &HFF)
    Exit Function
End Function
'
'Public Function Str2Long(str As String) As Long
''On error goto ERR1
'    Dim lt As Long
'    If str = "" Then
'        Str2Long = -1
'        Exit Function
'    End If
'    Str2Long = Asc(Mid(str, 1, 1))
'    'Debug.Print "Byte1 = " & Hex(Asc(Mid(str, 1, 1)))
'    lt = Asc(Mid(str, 2, 1))
'    'Debug.Print "Byte2 = " & Hex(lt)
'    Str2Long = Str2Long Or (lt * &H100)
'    lt = Asc(Mid(str, 3, 1))
'    'Debug.Print "Byte3 = " & Hex(lt)
'    Str2Long = Str2Long Or (lt * &H10000)
'    lt = Asc(Mid(str, 4, 1))
'    'Debug.Print "Byte4 = " & Hex(lt)
'    Str2Long = Str2Long Or (lt * &H1000000)
'ERR1:
'    Exit Function
'End Function

Public Function SendPacket() As Boolean
On Error GoTo err
    
    sStatus = ""
    SendPacket = False
    blnStatus = False
    SendCommand (sPacket)
    lTimer = 10
    Do
        DoEvents
        If blnStatus = True Then
            bStatus = Asc(Mid$(sStatus, 5, 1))
            Select Case bStatus
                Case C_CMDFILE_NOT_FOUND
                Case C_CMDINVALIDFILE
                    strMessage = "Command Error!"
                    SendPacket = False
                Case C_EOFFILE
                    strMessage = "End of file"
                    SendPacket = True
                Case C_READYTOSEND
                    strMessage = "Transfering data ..."
                    SendPacket = True
                Case C_CMD_EXIT_TRN
                    strMessage = "Transfering Over ..."
                    SendPacket = True
                Case C_RECV_READY
                    strMessage = "Device ready to send file "
                    SendPacket = True
            End Select
            Exit Do
        End If
        If lTimer = 0 Then
            SendPacket = False
            strMessage = "Time Out"
            Exit Do
        End If
    Loop

    Exit Function
err:
    strMessage = "Error!" & vbCrLf & "SendPacket : " & err.Description
End Function


Public Function CreateFile(strpath As String) As Boolean
On Error GoTo err
    CreateFile = False
    
    If Dir(strpath) <> "" Then Kill strpath
    
    iHdl = FreeFile()
    Open strpath For Binary Access Write As #iHdl
    Close #iHdl
    CreateFile = True
    Exit Function
err:
    If iHdl > -1 Then Close #iHdl
    strMessage = "Error!" & vbCrLf & "CreateFile : " & err.Description
End Function

Public Function RecvPacket() As Byte
On Error GoTo err
    
    RecvPacket = True
    blnStatus = False
    lTimer = 10
    Do
        DoEvents
        If blnStatus = True Then
            bStatus = Asc(Mid$(sStatus, 5, 1))
            RecvPacket = bStatus
            Exit Do
        End If
        If lTimer = 0 Then
            RecvPacket = C_CMDTIMEOUT
            strMessage = "Time Out"
            Exit Do
        End If
    Loop

    Exit Function
err:
    strMessage = "Error!" & vbCrLf & "RecvPacket : " & err.Description
    RecvPacket = C_NOTOK
End Function

Public Function WriteToFile(strpath As String, strData As String) As Boolean
On Error GoTo err

    WriteToFile = False
    iHdl = FreeFile()

    Open strpath For Binary Access Read Write As #iHdl
        If LOF(iHdl) <> 0 Then Seek #iHdl, LOF(iHdl) + 1
        Put #iHdl, , strData
    Close #iHdl
    WriteToFile = True
    Exit Function
err:
    If iHdl > -1 Then Close #iHdl
    strMessage = "Error!" & vbCrLf & "WriteToFile : " & err.Description
End Function

Public Function UploadFile(strpath As String, strFlname As String) As Boolean
On Error GoTo err
    
    Dim strSendData As String
    Dim strFname As String * 16
    Dim iVal As Integer, iPacketCount As Integer
    Dim bRStatus As Byte
    Dim iDataLen As Integer
    
    UploadFile = False
    
    strFileName = Trim(UCase(strFlname))
    strFilePath = strpath & strFlname
    
    If InitTransfer = False Then
'        UploadFile = strMessage
        Exit Function
    End If
    
    Debug.Print vbCrLf & "1 Status : 0x" & Asc(Mid$(sStatus, 5, 1))
    
    strFname = strFileName & Chr(0)
    lFileLen = 0
    strSendData = strFname & Long2Str(lFileLen)
    If CreatePacket(CMD_UPLOAD, C_OK, strSendData) = False Then
'        UploadFile = strMessage
        Exit Function
    End If
    
    If SendPacket() = False Then
'        UploadFile = strMessage
        Exit Function
    End If
    
    Debug.Print vbCrLf & "2 Status : 0x" & Asc(Mid$(sStatus, 5, 1))
    
    
    If iPacketCount = 0 And bStatus = C_EOFFILE Then
'        UploadFile = "File size is zero byte" & vbCrLf & "Transfer Over!"
        Exit Function
    End If
    
    lFileLen = Str2Long(Mid$(sStatus, 22, 4))
    iPacketCount = lFileLen / BT_DATA_SIZE
    If lFileLen Mod BT_DATA_SIZE <> 0 Then iPacketCount = iPacketCount + 1
    
    If CreateFile(strFilePath) = False Then
        strSendData = Chr(0)
        If CreatePacket(CMD_UPLOAD, C_CMD_FILE_ERROR, strSendData) = True Then
            If SendPacket() = False Then
                Debug.Print vbCrLf & "3 Status : 0x" & Asc(Mid$(sStatus, 5, 1))
'                UploadFile = strMessage
                Exit Function
            End If
        Else
'            UploadFile = strMessage
            Exit Function
        End If
    End If
    
    strSendData = Chr(0)
    If CreatePacket(CMD_UPLOAD, C_OK, strSendData) = True Then
        SendCommand (sPacket)
        Debug.Print vbCrLf & "4 Status : 0x"
    Else
'        UploadFile = strMessage
        Exit Function
    End If
    iVal = 1
    For iVal = 1 To iPacketCount
        bRStatus = RecvPacket
        Debug.Print vbCrLf & "5 Status : 0x - RcvPacket - " & iPacketCount
        If bRStatus = C_CRCERROR Or bRStatus = C_NOTOK Or bRStatus = C_CMDTIMEOUT Then
            Exit For
        Else
            iDataLen = (Str2Int(Mid$(sStatus, 2, 2))) - 6
            sStatus = Mid$(sStatus, 6, iDataLen)
            If WriteToFile(strFilePath, sStatus) = False Then
'                UploadFile = strMessage
                Exit Function
            End If
        End If
    
        strSendData = Chr(0)
        If CreatePacket(CMD_UPLOAD, C_OK, strSendData) = True Then
            SendCommand (sPacket)
            If bRStatus = C_EOFFILE Then
                Exit For
            End If
        Else
'            UploadFile = strMessage
            Exit Function
        End If
    Next
    
    UploadFile = True
    Exit Function
err:
    strMessage = "Error!" & vbCrLf & "WriteToFile : " & err.Description
    UploadFile = strMessage
End Function

Public Function SendByte(bByte As Byte) As Boolean
On Error GoTo err
    SendByte = False
    SerialComm.Output = Chr(bByte)
    SendByte = True
   Exit Function
err:
    MsgBox "Error! " & vbCrLf & "SendByte() " & vbCrLf & "Error No: " & err.Number & vbCrLf & err.Description
End Function

Public Function Int2Str(ByVal l As Integer) As String
   
    Int2Str = Chr(l And &HFF)
    'Debug.Print "Byte1 = " & Hex(l And &HFF)
    Int2Str = Int2Str & Chr(((l And &HFF00) / &H100) And &HFF)
    'Debug.Print "Byte2 = " & Hex(((l And &HFF00) / &H100) And &HFF)
    
End Function

Public Function Str2Int(str As String) As Integer
    Dim lt As Integer
    Str2Int = Asc(Mid(str, 1, 1))
    'Debug.Print "Byte1 = " & Hex(Asc(Mid(str, 1, 1)))
    lt = Asc(Mid(str, 2, 1))
    'Debug.Print "Byte2 = " & Hex(lt)
    Str2Int = Str2Int Or (lt * &H100)
End Function

Public Function UpdateData(sData As String) As Boolean
On Error GoTo err
    Dim iHdl1 As Integer
    UpdateData = False
    If sData <> "" Then
        iHdl1 = FreeFile()
        Open App.Path & "\BT.HEX" For Append As #iHdl1
            Print #iHdl1, sData
        Close #iHdl1
    End If
    
    UpdateData = True
    Exit Function
    
err:
    MsgBox "Error! " & vbCrLf & "UpdateData() " & vbCrLf & "Error No: " & err.Number & vbCrLf & err.Description
End Function
