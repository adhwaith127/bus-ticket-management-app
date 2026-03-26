Attribute VB_Name = "ModTrans"
Option Explicit
'***************************************************************************'
'* Project      : Data Transfer protocol for Palmtec                       *'
'* Function     : Transfer functioning as slave/master                     *'
'* Date         : 01/06/2004                                               *'
'* Developed by : Maneish A.R.                                             *'
'* Place        : Softland India Ltd                                       *'
'* This is an unpublished work & is confidential property of softland      *'
'* Packet Framing :                                                        *'
'* Packet Struct ,Total size 1024 byte for data                            *'
'*<FrameFlag><Datalength><PacketNo><PacketType><Data><CheckSum><FrameFlag> *'
'*     1          2          2         1          x        4      1        *'
'*     1          2,3       4,5        6          7       7+x    11+x      *'
'* Check sum = sum of all Datalength, PacketNo, PacketType, Data           *'
'* ####################################################################### *'
'* Protocol                                                                *'
'* CheckConnection: Send RQTS should be send by MASTER when it is ready    *'
'*                  then SLAVE must send RTR then flush port & again send  *'
'*                  RTR.                                                   *'
'* ReadyToRecv    : Same as CheckConnection except for a time out          *'
'* ReadyToSend    : Wait for RQTS send by MASTER, then SLAVE must send RTR *'
'*                  then flush port & again send  RTR                      *'
'* Receive Packet : This must be done only after a Ready To Receive.       *'
'*                  Checksum must be calculated in side Receive packet     *'
'*                  Retransmission for 3 times must be taken care          *'
'*                  Must set the no of retransmissions.                    *'
'*                  Return ERRORS                                          *'
'* Send    Packet : This must be done only after a Ready To Send.          *'
'*                  Must set the no of retransmissions.                    *'
'*                  Return ERRORS                                          *'
'* MakePacket     : Make packet calculating & seting checksum              *'
'* RecvFile       : Must keep a packet no to check with the received no.   *'
'*                  Check for file existence and wait for packets          *'
'*                  On error unlink file if file is created'truncked       *'
'*                  on success packet no eqilance send the same.           *'
'*                  on EOF close file and send status.                     *'
'* SendFile       : Check for file existence and send status.              *'
'*                  On error return file error                             *'
'*                  on EOF close file and send status.                     *'
'*                                                                         *'
'* ERRORS         : TimeOut                                                *'
'*                  PacketErr                                              *'
'*                  NoConnection                                           *'
'***************************************************************************'

'#Const LCDDEV = LCDDEV

'*<FrameFlag><Datalength><PacketNo><PacketType><Data><CheckSum><FrameFlag> *'
'*     1          2          2         1          x        4      1        *'
'*     1          2,3       4,5        6          7       7+x    11+x      *'
Private Const DATA_LEN_LOC = 2
Private Const PACKET_NO_LOC = 4
Private Const PACKET_TYPE_LOC = 6
Private Const DATA_LOC = 7
Private Const DATA_SIZE = 896
'Private Const DATA_SIZE = 768
'Private Const DATA_SIZE = 672
'Private Const DATA_SIZE = 448
'Private Const DATA_SIZE = 224
'Private Const DATA_SIZE = 32
Private Const CHECK_SUM_LOC = DATA_SIZE + 7
Private Const PACKET_SIZE = DATA_SIZE + 11
Private Const WRITE_BUFF_SIZE = 4096
Private Const DATA_START_LOC = 7

'--------------------* Timeout in sec -----------*
' Note :TimeOut is not required or implemented in
' the PC Side because data is Rcv in Interrupt
'------------------------------------------------*

Private Const TIME_OUT = 5
Private Const TIMEOUT_BETWEEN_PKT = 1
Private Const PACKET_NO_SUCCESS = 15
Private Const NO_OF_PACKET_REPEATS = 3


'---*REQUEST TO SEND---------*
Private Const RQTS = &HF1

'----*READY TO RECEIVE-------*
Private Const RTR = &HF2

'----*Start FRAME FLAG-------*
Private Const SFlag = &HF3

'----*End FRAME FLAG---------*
Private Const EFlag = &HF4

'----*DELAY FOR SECTOR ERASE---*
Private Const FILE_CREAT_DELAY = &HF8

'----*SECTOR ERASE OK---*
Private Const FILE_CREAT_OK = &HF9

'----*FILE END FLAG----------*
Private Const EOFF = &HFF


'----*NOT FILE END FLAG------*
Private Const NEOFF = &H0

Private Const FILE_FOUND = &HF5
Private Const FILE_GOT = &HF6
Private Const PACKET_GOT = &HF7

'---------* ERRORS-----------------*
Private Const ERR_PACKET = &HE0
Private Const ERR_PACKET_ACK = &HE1
Private Const ERR_NOFILE = &HE2
Private Const ERR_TIMEOUT = &HE3
Private Const ERR_PACKET_NO = &HE4
Private Const ERR_FILE = &HE5
Private Const ERR_EOFF = &HE6
Private Const ERR_DATA_NOT_UPLOADED = &HE7
Private Const ERR_FS = &HE8

Private Type TIMETYPE
    Seconds As Byte
    Minutes As Byte
    Hour As Byte
End Type

Private Type DATETYPE
    Year As Byte
    Month As Byte
    Date As Byte
    Dow  As Byte
End Type

Private Type DATA_PACKET
    SFlag  As Byte      '* 0xF1 *
    Datalength(2) As Byte
    PacketNo(2) As Byte
    PacketType As Byte
    DATA(DATA_SIZE) As Byte
    CheckSum(4) As Byte
    EFlag As Byte       '* 0xF2 *
End Type

'--------------------*Commands list------------------------------------------*
'                   Trans Commands Byte
'----------------------------------------------------------------------------*
'* 1 FILENAME*
Private Const C_SEND_FILE = "1"
'* 2 FILENAME*
Private Const C_RECEIVE_FILE = "2"
'* 3 *
Private Const C_SEND_ALL_FILENAME = "3"
'* 4 DATETYPE TIMETYPE *
Private Const C_SEND_DATE_TIME = "4"
'* 5 DATETYPE TIMETYPE *
Private Const C_SET_DATE_TIME = "5"
'* 6 *
Private Const C_SEND_RTC_DATA = "6"
'* 7 FORMAT (for conformation)*
Private Const C_FORMAT = "7"
'* 8 CLEAR  (for conformation)*
Private Const C_CLEAR_RTC = "8"
'* 9 (rtc data,date time,filenames,status)*
Private Const C_SEND_ALL = "9"
'* 10 *
Private Const C_STATUS = "10"
'* 50 *

Private I As Integer
Private TransParm As String
Private PacketWriteBuff  As String
Private PacketBuff As String
Private WriteBuff(DATA_SIZE - 1) As Byte
Private PacketType As Byte
Private SendPacketBuff As String
Private SendDataBuff As String
Private PacketNo As Long
Private TimeOut As Integer
Private DataLen As Long
Private DataWriteLen As Long
Private ch As Byte
Private Recvchar As Byte
Private CheckSum As Long
Private PacketIndex As Long
Private NoOfTimes As Integer
Private Error As Integer
Private RptCount As Integer
Private Command As Integer
Private CommandStr As String
Private Port As String
Private DataType As Integer
Private ReceiveCount As Integer
Private ReadComCount As Integer
Dim Filehdl As Integer
Private StTime As Single
Dim tt As Single
Public Msg As String
Private PKTStartGot As Boolean
Private Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwmilliseconds As Long)

'Convertion Declerations******************************************
Private Type CONV_TIME
    Hour As Byte
    Minutes As Byte
    Seconds As Byte
End Type

Private Type CONV_DATE
    Date As Byte
    Month As Byte
    Year As Byte
End Type

'Convertion Declerations******************************************

Public SerialComm As MSComm

Public TransPath As String
Public TransStatusNo As Integer
Public TransMsg As String
Public filename As String

Public Function InitPort(Port As Integer, baud As String, Optional ByVal intRTThreshold As Integer = 0) As Boolean
On Error GoTo CommErr
With SerialComm
    'Buffer to hold input string
    If .PortOpen = True Then .PortOpen = False
    .CommPort = Port
    ' 128000 baud, no parity, 8 data, and 1 stop bit.
    .Settings = baud & ",N,8,1"
    '.Settings = "57600,N,8,1"
    '.Settings = "9600,N,8,1"
    ' Tell the control to read each byte on interrupt
    .InputLen = 1
    ' Open the port.
    ' Set InputMode to read binary data
    .InputMode = comInputModeBinary
    'RThreshold to 1, causes the MSComm control
    'to generate the OnComm event every time a
    'single character is placed in the receive buffer.
    .RThreshold = intRTThreshold
    .SThreshold = 0
'    .RTSEnable = False
'    .DTREnable = True
     .PortOpen = True
End With
InitPort = True
Exit Function
CommErr:
    'MsgBox Err.Description & vbCrLf & "Close any application using same port.", vbExclamation, "Trans - Com" & Port
    TransStatusNo = err.Number
    TransMsg = err.Description & vbCrLf & "Close any application using same port."
    'End
End Function

Private Sub SendByte(b As Byte)
    'SerialControl.Output = Chr(B)
    SerialComm.Output = Chr(b)
End Sub

Private Function RecvByte(ByRef b As Byte) As Boolean
Dim Rbyte As Byte
    tt = Timer
With SerialComm
RCV:
    While .InBufferCount < 1
        DoEvents
         If IsTimeOut() Then
            'RecvByte = ERR_TIMEOUT
            Exit Function
        End If
    Wend
    Rbyte = SerialComm.Input(0)
    If Rbyte = FILE_CREAT_DELAY Then
        Do
'            frmArm.lblStatus.Caption = "Creating File..."
            While .InBufferCount < 1
                DoEvents
            Wend
            Rbyte = SerialComm.Input(0)
            If Rbyte = FILE_CREAT_OK Then
'                frmArm.lblStatus.Caption = "File Created."
                tt = Timer
                GoTo RCV
            End If
        Loop
    End If
    'FrmTransfer.SerialCom
    b = Rbyte
    RecvByte = True
End With
End Function

Private Sub FlushPort()
With SerialComm
    .InBufferCount = 0
    .OutBufferCount = 0
End With
    PacketBuff = ""
End Sub

Private Function ReadyToRecv() As Integer
    tt = Timer
    For I = 0 To 16
        If RecvByte(ch) Then
            If ch = RQTS Then
                'SendByte RTR
                FlushPort
                'SendByte RTR
                ReadyToRecv = RQTS
                Exit Function
            End If
        End If
    Next I
    ReadyToRecv = ERR_TIMEOUT
End Function

Private Function ReadyToSend() As Integer
    For I = 0 To 3
        'SendByte RQTS
        If RecvByte(ch) Then
            If ch = RTR Then
                If RecvByte(ch) Then
                     If ch = RTR Then
                        ReadyToSend = RTR
                        Exit Function
                   End If
                End If
            End If
        End If
    Next I
    ReadyToSend = ERR_TIMEOUT
End Function

Private Function CalculateCheckSum() As Long
Dim ChkSum As Long
    ChkSum = 0
    If PacketBuff = "" Then Exit Function
    For I = 2 To PACKET_SIZE - 5
        ChkSum = ChkSum + Asc(Mid(PacketBuff, I, 1))
    Next I
    CalculateCheckSum = ChkSum
End Function

Private Function SendPacket() As Integer
Dim SndByte As Byte
Dim str As String
    
    For NoOfTimes = 1 To NO_OF_PACKET_REPEATS
        CheckSum = 0


        If ReadyToSend() <> RTR Then
            SendPacket = ERR_TIMEOUT
            TransMsg = "Erroneous Connection"
            Exit Function
'            Debug.Print "RTR Not Got ReadyToSend..."
'        Else
'            Debug.Print "RTR Got Ready to Send"
        End If
        '*<FrameFlag><Datalength><PacketNo><PacketType>
        '*     1          2          2         1
        '*     1          2,3       4,5        6
        SndByte = Asc(Mid(SendPacketBuff, 1, 1))
        'SendByte SndByte
        For PacketIndex = 2 To 6
            SndByte = Asc(Mid(SendPacketBuff, PacketIndex, 1))
            'SendByte SndByte
            CheckSum = CheckSum + SndByte
        Next PacketIndex
'        Debug.Print "Frame Data Send "
        '<Data>
        'x
        '7
        For PacketIndex = 1 To DATA_SIZE
            SndByte = Asc(Mid(SendDataBuff, PacketIndex, 1))
            'SendByte SndByte
            CheckSum = CheckSum + SndByte
        Next PacketIndex
        '<CheckSum><FrameFlag> *'
        '      4      1        *'
        '     7+x    11+x      *'
       'Sleep (10)
        str = Long2Str(CheckSum)
        'Debug.Print "Data Send CheckSum : " & CheckSum
        'SendByte Asc(Mid(str, 1, 1))
        'SendByte Asc(Mid(str, 2, 1))
        'SendByte Asc(Mid(str, 3, 1))
        'SendByte Asc(Mid(str, 4, 1))
        'SendByte EFlag
       
            '***Wait for Packet got***
        'Sleep (50)
        If RecvByte(Recvchar) Then
            If Recvchar = PACKET_GOT Then
                SendPacket = PACKET_GOT
                Exit Function
            ElseIf Recvchar = ERR_PACKET Then
                'Sleep (50)
                'SendByte ERR_PACKET_ACK
                'Debug.Print "Recv ERR_PACKET Sending ERR_PACKET_ACK"
            End If
        End If
    Next NoOfTimes
    SendPacket = ERR_PACKET
End Function


Private Sub DELAY(Sec As Long)
    Dim S As Long
    S = Timer
    Do
        DoEvents
        If (Timer - S) >= Sec Then Exit Sub
    Loop
End Sub

Private Function RecvPacket() As Integer
Dim PrevRecvCount As Integer, InBuffCnt As Integer
Dim DataPacket As String
Dim PktCount As Integer
    TimeOut = TIME_OUT
With SerialComm
    For NoOfTimes = 1 To NO_OF_PACKET_REPEATS
        CheckSum = 0
        'Debug.Print "Going To receive RTR before sending packet"
        TimeOut = TIME_OUT  '_BETWEEN_PKT
        tt = Timer
        For I = 0 To 3
            If RecvByte(ch) Then
                If ch = RQTS Then
                    'SendByte RTR
                    FlushPort
                    .InputLen = PACKET_SIZE
                    .InputMode = comInputModeText
                    PktCount = 0
                    TimeOut = TIME_OUT
                    'SendByte RTR
                    RecvPacket = RQTS
                    Exit For
                End If
            End If
        Next I
        
        If RecvPacket <> RQTS Then
            RecvPacket = ERR_TIMEOUT
            Exit Function
        End If
        'While ReceiveCount <> PACKET_SIZE
        While PktCount < PACKET_SIZE
            DoEvents
            PacketBuff = PacketBuff & .Input
            PktCount = Len(PacketBuff)
            'FrmTransfer.Text3.Text = PktCount 'ReceiveCount
            If IsTimeOut() Then
                RecvPacket = ERR_TIMEOUT
                Exit Function
            End If
            TimeOut = TIME_OUT  '_BETWEEN_PKT
        Wend
        .InputLen = 1
        .InputMode = comInputModeBinary
        
        TimeOut = TIME_OUT
        CheckSum = Str2Long(Mid(PacketBuff, DATA_SIZE + 7, 4))
        If CheckSum = -1 Then CheckSum = -2
        If (CheckSum - &HF3) <> CalculateCheckSum() Or RecvPacket = ERR_TIMEOUT Then
            'SendByte ERR_PACKET
            If RecvByte(ch) Then
                If ch <> ERR_PACKET_ACK Then
                    RecvPacket = ERR_PACKET
                    Exit Function
                End If
            Else
                RecvPacket = ERR_PACKET
                Exit Function
            End If
        Else
            'SendByte PACKET_GOT
            If RecvByte(ch) Then
                If ch = PACKET_GOT Then
                    RecvPacket = PACKET_GOT
                    Exit Function
                End If
            End If
        End If
    Next NoOfTimes
End With
    RecvPacket = ERR_PACKET
End Function

Private Sub SetPacketPrefix(DataLen As Long, PktNo As Long, PktType As Byte)
' Reset packet and add <FrameFlag><Datalength><PacketNo><PacketType>
'*<FrameFlag><Datalength><PacketNo><PacketType><CheckSum><Data><FrameFlag> *'
'*     1          2          2         1          4         x      1       *'
    SendPacketBuff = Chr(SFlag) & Int2Str(DataLen) & Int2Str(PktNo) & Chr(PktType)
End Sub

Private Function FileExists(CommandStr As String) As Boolean
'On error goto Errhdl
    Filehdl = FreeFile()
    Open TransPath & "\" & CommandStr For Input As #Filehdl
    Close #Filehdl  ' Close file.
    FileExists = True
    Exit Function
Errhdl:
    'Debug.Print "FileExists Err :" & Err.Description
    FileExists = False
    TransMsg = "FileExists Err :" & err.Description
    Exit Function
End Function

'' Sendfile is to calculate the check sum,put frame flag ,datalen &
' packet number then calls SendPacket() then waits for a packet got
' else returns packet err
Private Function SendFile() As Integer
  Dim FilePoint As Long
    
    PacketNo = 0: PacketType = 0
    UCase (CommandStr)
    'Debug.Print "File name : " & CommandStr
    TransMsg = "Sending File " & CommandStr
    TransStatusNo = 0
    If Dir$(TransPath & "\" & CommandStr) = "" Then
             'SendByte (ERR_NOFILE)
             SendFile = ERR_NOFILE
             'MsgBox TransPath & "\" & CommandStr & " not found", , "TRANSFER"
             Exit Function
    End If
         
    Filehdl = FreeFile()
    Open TransPath & "\" & CommandStr For Binary Access Read Write As #Filehdl Len = DATA_SIZE
    'SendByte (FILE_FOUND)
    'Debug.Print "File found, send FILEFOUND flag"
    FilePoint = 1
    Do
        PacketNo = PacketNo + 1
        SendDataBuff = String(DATA_SIZE, Chr(0))
        Get #Filehdl, FilePoint, SendDataBuff
        'Debug.Print "File Pointer " & FilePoint
        If EOF(Filehdl) Then
            ' Memsetted buffer so set it only on EOF
            'DataLen = InStrRev(SendDataBuff, Chr(0), 1, vbBinaryCompare)
            DataLen = LOF(Filehdl) - FilePoint
            DataLen = DataLen + 1
            If DataLen = 1 Then DataLen = 0
            PacketType = EOFF
            'Debug.Print "Last packet of the file"
        Else
            DataLen = DATA_SIZE
        End If
        FilePoint = FilePoint + DataLen
        If PacketType = EOFF Then
            SetPacketPrefix DataLen, PacketNo, PacketType
        Else
            SetPacketPrefix DataLen, PacketNo, PacketType
        End If
'        Debug.Print "Sending Packet no: " & PacketNo
'        Debug.Print "Send repeated    : " & NoOfTimes & " times"
'        Debug.Print "Data length      : " & DataLen
'        Debug.Print "Check sum        : " & CheckSum
'        Debug.Print "tt               : " & tt
'        Debug.Print "Going to send RTS & receive RTR"
        Error = SendPacket()
        If Error = ERR_PACKET Then
'            Debug.Print "PACKET_ERROR"
'            Debug.Print "Send repeated    : " & NoOfTimes & " times "
            TransMsg = "Err Transfer: Send repeated " & NoOfTimes & " times "
            TransStatusNo = -4
            SendFile = Error
            Close #Filehdl
            Exit Function
        ElseIf Error = ERR_TIMEOUT Then
'            Debug.Print "Time over"
 '           Debug.Print "Send repeated    : " & NoOfTimes & " times "
            SendFile = Error
            Close #Filehdl
            TransMsg = "Err Transfer: Send repeated " & NoOfTimes & " times "
            TransStatusNo = -100
            Exit Function
        End If
        
        
        
        If EOF(Filehdl) Then
            SendFile = IsNoOfPktsRcvd_AOE(PacketNo)
            Close #Filehdl
            Exit Function
        End If
        'FrmTransfer.StatusBar1.SimpleText = "Pkt No: " & PacketNo
        'Status
        TransMsg = "Send Packet No: " & PacketNo
        TransStatusNo = 0
    Loop
Close #Filehdl
Exit Function
FileErr:
End Function

Private Function RecvFile() As Integer
    Dim FileChanged As Integer
    If RecvByte(ch) Then
        If ch = ERR_NOFILE Then
            RecvFile = ERR_NOFILE
            Exit Function
        ElseIf ch = ERR_FS Then
            RecvFile = ERR_FS
            Exit Function
        End If
    Else
        RecvFile = ERR_NOFILE
        Exit Function
    End If
    
    UCase CommandStr
    PacketWriteBuff = ""
    DataLen = 0
    DataWriteLen = 0
    PacketNo = 0
    If Dir$(TransPath & "\TEMPFILE.BIN") <> "" Then
        Kill TransPath & "\TEMPFILE.BIN"
    End If
    Do
        ch = RecvPacket()
        If ch = PACKET_GOT Then
            PacketNo = PacketNo + 1
            TransMsg = "Received PacketNo : " & PacketNo
            If PacketNo = 1 Then
'                Debug.Print "Creating file on receiving first packet..."
                Filehdl = FreeFile()
                Open TransPath & "\TEMPFILE.BIN" For Binary Access Read Write As #Filehdl
            End If
'*<FrameFlag><Datalength><PacketNo><PacketType><Data><CheckSum><FrameFlag> *'
'*     1          2          2         1          x        4      1        *'
'*     1          2,3       4,5        6          7       7+x    11+x      *'
            If Str2Int(Mid(PacketBuff, 4, 2)) <> PacketNo Then
                RecvFile = ERR_PACKET
                Close #Filehdl
                If Dir$(TransPath & "\TEMPFILE.BIN") <> "" Then
                    Kill TransPath & "\TEMPFILE.BIN"
                End If
                Exit Function
            End If
'            Debug.Print "Total Bytes Written : " & PacketNo * DATA_SIZE
'            Debug.Print "NO of Packet repeats.." & NoOfTimes
'            Debug.Print "Going to write to file.. Packet NO : " & PacketNo
'            Debug.Print "Total Bytes Written : " & Val(PacketNo * DATA_SIZE)
            DataLen = Str2Int(Mid(PacketBuff, DATA_LEN_LOC, 2))
            If DataLen = DATA_SIZE Then
                For I = 0 To DATA_SIZE - 1
                    WriteBuff(I) = Asc(Mid(PacketBuff, DATA_START_LOC + I, 1))
                Next I
                Put #Filehdl, , WriteBuff
            Else
                For I = 0 To DataLen - 1
                    Put #Filehdl, , CByte(Asc(Mid(PacketBuff, DATA_START_LOC + I, 1)))
                Next I
            End If
            'Put #Filehdl, DataLen, Mid(PacketBuff, DATA_START_LOC, DataLen)
            'writing file after n = 5 packets
            TransStatusNo = 0
'            Debug.Print "DATALEN    : " & DataLen
'            Debug.Print "CHECK_SUM  : " & CheckSum
'            Debug.Print "Packet GOT SEND BACK after writing packet to file"
            If Asc(Mid(PacketBuff, PACKET_TYPE_LOC, 1)) = EOFF Then
                TransStatusNo = 0
'                Debug.Print "EOFF GOT EOFF GOT..."
                Close Filehdl
                RecvFile = IsNoOfPktsSend_BOE
                FileCopy TransPath & "\TEMPFILE.BIN", TransPath & "\" & CommandStr
                If Dir$(TransPath & "\TEMPFILE.BIN") <> "" Then
                    Kill TransPath & "\TEMPFILE.BIN"
                End If
                Exit Function
            End If
        Else
'            Debug.Print "IN REPEAT PACKET RPTCOUNT :" & NoOfTimes
            RecvFile = ch
            If Filehdl <> -1 Then Close #Filehdl
            If Dir$(TransPath & "\TEMPFILE.BIN") <> "" Then
                Kill TransPath & "\TEMPFILE.BIN"
            End If
            Exit Function
        End If
    Loop
End Function
Private Function IsTimeOut() As Boolean
    If Timer - tt > TimeOut Then
        IsTimeOut = True
    Else
        IsTimeOut = False
    End If
End Function

Private Function CheckConnection() As Boolean
'Dim i As Integer
'For i = 1 To 3
    'SendByte RQTS
    If RecvByte(ch) Then
        If ch = RTR Then
            If RecvByte(ch) Then
                If ch = RTR Then
                    'SendByte RQTS
                    CheckConnection = True
                    Exit Function
                End If
            End If
        End If
    End If
    CheckConnection = False
'Next i
End Function

Private Function Int2Str(ByVal l As Integer) As String
   
    Int2Str = Chr(l And &HFF)
    'Debug.Print "Byte1 = " & Hex(l And &HFF)
    Int2Str = Int2Str & Chr(((l And &HFF00) / &H100) And &HFF)
    'Debug.Print "Byte2 = " & Hex(((l And &HFF00) / &H100) And &HFF)
    
End Function

Private Function Str2Int(str As String) As Integer
    Dim lt As Integer
    Str2Int = Asc(Mid(str, 1, 1))
    'Debug.Print "Byte1 = " & Hex(Asc(Mid(str, 1, 1)))
    lt = Asc(Mid(str, 2, 1))
    'Debug.Print "Byte2 = " & Hex(lt)
    Str2Int = Str2Int Or (lt * &H100)
End Function


Public Function Trans(CmdStr As String) As Boolean
    Dim CmdBuff As String, j As Integer
    TimeOut = TIME_OUT
    SerialComm.InBufferCount = 0
    SerialComm.OutBufferCount = 0
    ReceiveCount = 0
    ReadComCount = 0
    PacketBuff = ""
    Filehdl = -1
    
    If Not CheckConnection Then
        TransStatusNo = -2
'        Debug.Print "Check Connection ! GOING TO ABORT..."
        TransMsg = "Check Connection ! GOING TO ABORT..."
        Exit Function
    End If
    PacketNo = 1
'    Debug.Print "Check Connection OK"
    I = Len(Trim(CmdStr))
    SendDataBuff = CmdStr & String(DATA_SIZE - I, Chr(0))
    DataLen = I
    SetPacketPrefix DataLen, 1, 0
'    Debug.Print "Send Command Packet"
    ch = SendPacket()
'    Debug.Print "Checked Connection OK ..."
    If ch <> PACKET_GOT Then
        TransMsg = "Erroneous Connection... "
        TransStatusNo = -3
'        Debug.Print "Packet not got after sendpacket \n just after check connection "
'        Debug.Print "Erroneous Connection : Cannot Send Packet,Press any key... "
        Exit Function
    End If
    CmdBuff = Trim(CmdStr)
    I = InStr(1, CmdBuff, " ")
    If I < 1 Then I = Len(CmdBuff) + 1
    Command = val(Mid(CmdBuff, 1, I - 1))
    I = I + 1
    j = InStr(I, CmdBuff, " ")
    If j <= I Then j = Len(CmdBuff) + 1
    If j > I Then
        CommandStr = Trim(Mid(CmdBuff, I, j - I))
        j = j + 1
        I = InStr(j, CmdBuff, " ")
        If I <= j Then I = Len(CmdBuff) + 1
        If I > j Then
            TimeOut = val(Mid(CmdBuff, j, I - j))
        End If
    End If
    
    Select Case Command
        Case 1:
            I = RecvFile()
            If I = EOFF Then
                TransMsg = "Success"
                'TransMsg = "File " & CommandStr & " Uploaded Successfully from Palmtec"
                TransStatusNo = 1
                Trans = True
            ElseIf I = ERR_NOFILE Then
                TransMsg = "File Not Found"
                TransStatusNo = -1
                Trans = False
            ElseIf I = ERR_FS Then
                TransMsg = "FS Error"
                TransStatusNo = -1
                Trans = False
            End If
        Case 2:
            TimeOut = 1
            TimeOut = TIME_OUT
            If SendFile = EOFF Then
                TransMsg = "Success"
                'TransMsg = "File " & CommandStr & " Downloaded Successfully to Palmtec"
                TransStatusNo = 2
                Trans = True
            End If
        Case 5, 7, 50:
            TransMsg = "Success"
            'TransMsg = "System time " & Format(Now, "dd/MM/yyyy hh:mm:ss") & " Set"
            'TransStatusNo = 2
            Trans = True
        Case 13:
            CommandStr = "I2C.DAT"
            I = RecvFile()
            If I = EOFF Then
                TransMsg = "Success"
                'TransMsg = "File " & CommandStr & " Uploaded Successfully from Palmtec"
                TransStatusNo = 1
                Trans = True
            ElseIf I = ERR_NOFILE Then
                TransMsg = "File Not Found"
                TransStatusNo = -1
                Trans = False
            End If
            
    End Select
    If Filehdl <> -1 Then Close (Filehdl)
End Function

Private Function IsNoOfPktsRcvd_AOE(ByVal Pkts As Integer) As Integer
  Dim GotPackets As Integer
  Dim b1 As Byte, b2 As Byte
    If ReadyToRecv() <> RQTS Then
        IsNoOfPktsRcvd_AOE = ERR_TIMEOUT
        Exit Function
    End If
    If RecvByte(ch) Then
    
        If ch = EOFF Then
            Error = 0
'* going to get the no of packets received at the other end*
            If RecvByte(b1) Then
                If Not RecvByte(b2) Then
                    Error = ERR_TIMEOUT
                End If
            Else
                Error = ERR_TIMEOUT
            End If
            GotPackets = (b2 * &H100) Or b1
            If GotPackets <> Pkts Then
                'SendByte ERR_PACKET_NO
            Else
                'SendByte PACKET_NO_SUCCESS
                IsNoOfPktsRcvd_AOE = EOFF '*Receive transaction file*
                Exit Function
            End If
        Else
             IsNoOfPktsRcvd_AOE = ERR_EOFF
             Exit Function
        End If
    End If
    IsNoOfPktsRcvd_AOE = ERR_TIMEOUT
End Function

Private Function IsNoOfPktsSend_BOE() As Integer

    If ReadyToSend() <> RTR Then
        IsNoOfPktsSend_BOE = ERR_TIMEOUT
        Exit Function
    End If
    'SendByte EOFF
    'SendByte (PacketNo And &HFF)
    'SendByte ((PacketNo And &HFF00) / &H100)
    If RecvByte(ch) Then
        If ch = ERR_PACKET_NO Then
            IsNoOfPktsSend_BOE = ERR_PACKET_NO
            Exit Function
        End If
            IsNoOfPktsSend_BOE = EOFF
    Else
        IsNoOfPktsSend_BOE = ERR_TIMEOUT
    End If

End Function



Private Function CStr2VBStr(b() As Byte) As String
    Dim I As Integer
'On error goto ERR1
    I = 1
    While b(I) <> 0
        CStr2VBStr = CStr2VBStr & Chr(b(I))
        I = I + 1
    Wend
ERR1:
    Exit Function
End Function

Public Function VBStr2CStr(ByVal str As String, b() As Byte, l As Integer) As Integer
'On error goto ERR1
    Dim I As Integer
    Dim slen As Integer
    slen = Len(str)
    I = 1
    While I <= l
        If I <= slen Then
            b(I) = Asc(Mid(str, I, 1))
        Else
            b(I) = 0
        End If
        I = I + 1
    Wend
    VBStr2CStr = I
    
ERR1:
    Exit Function
End Function

Private Function Byte2VbStr(b() As Byte, l As Integer) As String
'On error goto ERR1
    Dim I As Integer
    I = 1
    While I <= l
        Byte2VbStr = Byte2VbStr & Chr(b(I))
        I = I + 1
    Wend
ERR1:
    Exit Function
End Function

Private Sub SetBuff(b() As Byte, ByVal ch As Byte, ByVal sz As Integer)
'On error goto ERR1
    Dim I As Integer
    I = 1
    While I < sz
        b(I) = ch
        I = I + 1
    Wend
ERR1:
    Exit Sub
End Sub

Private Function Rs2Paise(ByVal rs As String) As Long
'On error goto ERR1
    Dim dot As Integer
    Dim str As String
    str = Trim(rs)
    dot = InStr(1, str, ".")
    If dot > 0 Then
        rs = Left(str, dot - 1)
        rs = rs & Right(str, Len(str) - dot)
    End If
    Rs2Paise = val(rs)
ERR1:
    Exit Function
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
ERR1:
    Exit Function
End Function

Public Function Str2Long(str As String) As Long
'On error goto ERR1
    Dim lt As Long
    If str = "" Then
        Str2Long = -1
        Exit Function
    End If
    Str2Long = Asc(Mid(str, 1, 1))
    'Debug.Print "Byte1 = " & Hex(Asc(Mid(str, 1, 1)))
    lt = Asc(Mid(str, 2, 1))
    'Debug.Print "Byte2 = " & Hex(lt)
    Str2Long = Str2Long Or (lt * &H100)
    lt = Asc(Mid(str, 3, 1))
    'Debug.Print "Byte3 = " & Hex(lt)
    Str2Long = Str2Long Or (lt * &H10000)
    lt = Asc(Mid(str, 4, 1))
    'Debug.Print "Byte4 = " & Hex(lt)
    Str2Long = Str2Long Or (lt * &H1000000)
ERR1:
    Exit Function
End Function
'Convertion functions******************************************
Public Function Transfer(ActionType As String, FileLocation As String, Message As String) As Boolean
    Dim t  As Date
    Dim str As String
    Dim stat As Integer
    Message = ""
    TransPath = FileLocation
    ActionType = UCase$(ActionType)
'    If ActionType = "DOWNLOAD" Then
'
'        If Trans("1 MASTER.TRN") = True Then
'            stat = 1
'            If Trans("1 SHIFT.RPT") = True Then
'                stat = 2
'                If DecryptData = True Then
'                    stat = 3
'                Else
'                    TransMsg = "Partial"
'                End If
'            End If
'        End If
'
'        If stat = 3 Then
'            Transfer = True
'        End If
'    ElseIf ActionType = "UPLOAD" Then
'
'        If EncryptMaster() < 0 Then
'            Transfer = False
'            Message = TransMsg
'            Exit Function
'        End If
'        If Trans("2 MASTER.DAT") Then
'            Transfer = True
'        End If
'
    If ActionType = "SETDATEANDTIME" Then

        t = Now
        str = "5 " _
            & Chr(DatePart("yyyy", Now) Mod 100) _
            & Chr(DatePart("m", Now)) _
            & Chr(DatePart("d", Now)) _
            & Chr(DatePart("w", Now)) _
            & Chr(DatePart("s", Now)) _
            & Chr(DatePart("n", Now)) _
            & Chr(DatePart("h", Now))
        If Trans(str) Then
            Transfer = True
        End If
    
'    ElseIf ActionType = "ENABLEMASTERDOWNLOAD" Then
'
'        If Trans("50") Then
'            Transfer = True
'        End If
    
    ElseIf ActionType = "RECOVERDATA" Then
        
        If Trans("13") Then
           'RecoverData
            Transfer = True
        End If
    
    ElseIf ActionType = "FORMAT" Then
        
        If Trans("7 FORMAT") Then
            Transfer = True
        End If
    
    End If
    Message = TransMsg

End Function

