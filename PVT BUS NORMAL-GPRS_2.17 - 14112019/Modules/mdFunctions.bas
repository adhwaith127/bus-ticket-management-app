Attribute VB_Name = "mdFunctions"
Option Explicit
Public Const T_CASH = 0
Public Const T_ADJUST = 9
Public Const T_PHY = 10
Public Const T_CONC = 20
Public Const T_CONC_E = 40

Public LoginStatus As Byte

Public filename As String
Public ChStr As String
Public DatFileName As String
Public FieldCount As Integer
Public str1 As String
Public MainBuffer As String
Public buff As String
Public OneLine As String
Public ret As Long
Public HNDWND As Long
Public PortNo As String

Public PC_to_PMTC_Cntr As Integer
Public GlobUsrName As String
Public flxgrdCnt As Integer
Public toListBox As Integer
Public COUNTER As Integer
Public StageNameBuff As String
Public StgNameCnt As Long     'No to store the stage name
Public FrType As Integer
Public NOSTGS As Long        'To Store No of Stages
Public TotFare As Integer       'Total no of Fares
Public TotBasicFare As Integer  'Total no of Basic Fares

Public CharWcnt As Integer      'Language
Public strLanguageStage As String
Public LanguageStage(23) As Byte
Public LocalLanguage As Byte
Public strLocalLanguage As String

Public StageFile(3134) As Byte
Public RouteID As String    ' for Language stage edit

Dim Fso1

Dim strftr As String
Public sDataBase As String
Public StrHeader As String
Public StgName(255) As String
Public StageDistance(255) As String
Public FileDateCret As String
Public pmtcID As String
Public scduID As String
Public ArryCount  As Integer
Public DataTrans As Boolean, RouteAdd As Boolean, GraphEdit As Boolean, FareTableEdit As Boolean
Public RouteEdit As Boolean, StageEdit As Boolean
Public DeleteRoute As Boolean
Public Settings As Boolean
Public PCSettings As Boolean
Public admin As Boolean
Public SAVEFLAG As Boolean
Public SUPERUSER As Boolean
Public FARESAVEFLAG As Boolean
Public NORECORDFLAG As Boolean
Public FRAMEFLAG As Integer
Public CREW_FLAG As Boolean
Public loginsucceed As Boolean
Public CancelFlag As Boolean

Public gblBMPName As String

Public Type table
Col As Integer
row As Integer
End Type

Public DB As DAO.Database
Public DB1 As DAO.Database ' 04/01/2010
Public RES, RESSAN, res1, RES3, RES2, RES14, RESROUTE, RESLaN, RES5, RES6, RES7, RES8, res9 As DAO.Recordset
Public RES4 As DAO.Recordset

Public CON, adoDBCon As New ADODB.Connection          'RNC
Public rs As ADODB.Recordset, rslogin1 As New ADODB.Recordset
Public rs1 As New ADODB.Recordset 'sangeetha
Public rs2 As New ADODB.Recordset 'sangeetha
Public rs3 As New ADODB.Recordset 'sangeetha
Public rs4 As New ADODB.Recordset 'sangeetha
Public RCD As ADODB.Recordset
Public SQLQuery As String
Public Query As String
 

Public RSql As String
Public sSQL As String
Public sql As String

Public TDB As DAO.Database
Public TRES As DAO.Recordset
Public TSQL As String
Public passTSQL As String
Public HStr As SETUP
Public hardwaresettings As HARDWARE_SETUP

Public CONS As CONCESSION
'Public Crew As CREWDETAILS
Public Crew As CREWDET
Public LSTAG As LANGUAGE_STAG
Public odmtr As ODOMETER
Public insptr As INSPECTORDET
Public exp As EXPENSEDET
Public expobj As EXPENSES
Public adoc As New ADODB.Connection
Public gbladoCon As New ADODB.Connection
Public DbZb As New ADODB.Connection
Public RsZb As New ADODB.Recordset
Public RSDT As New ADODB.Recordset
Public Dbs As New ADODB.Connection
Public rsado As New ADODB.Recordset

Public Type PROJECT_VALIDITY
    ProjectStartDate As Date
    LastUsedDate As Date
    MaximumTrialDays As Integer
    ValidityCount As Integer
    ExpiredFlag As Boolean
    TrialFlag As Boolean
End Type

Public Type PASSCONC      '56 Bytes
    
    PassNo As String * 11
    TicketNo As Long
    Name As String * 12 '24 asish(30082019)
    Amount As Single
    DateofBirth As String * 4
    Startingperiod As Byte
    EndPeriod As Byte
    StartStage As Byte
    EndStage As Byte
    ViaStage As Byte
    TransID As Byte
    PassType As Integer
    Reserved As String * 1
End Type

Public Type PReport
    STicketNo As Long
    ETicketNo As Long
    InHandAmount As Single
    TotalColl As Single
    FullColl As Single
    HalfColl As Single
    PhyColl As Single
    LuggageColl As Single
    STColl As Single
    AdjustColl As Single
    TotalKM As Single
    CessAmount As Single
    ExpenseAmount As Single
    WarrantAmount As Single
    DeafAndDumpAmount As Single
    OperatedKM As Single
    DieselFilled As Single
    CancelledKM As Single
     
    Full As Integer
    Half As Integer
    Phy As Integer
    st As Integer
    Free As Integer
    Adjust As Integer
    Conc As Integer
    Others As Integer
    Lugg As Integer
    gPass As PASSCONC
    sPass As Integer
    Warrant As Integer
    DeafAndDump As Integer
    CessCount As Integer
    TotalPassenger As Integer
    
    SCHEDULE As Byte
    TripNo As Byte
    UpDownTrip As Byte
    RouteCode As String * 5
    StartD As Byte
    StartMO As Byte
    StartY As Byte
    StartH As Byte
    StartM As Byte
    EndD As Byte
    EndMO As Byte
    EndY As Byte
    EndH As Byte
    EndM As Byte
    
    DepotCode As String * 6
    DepotName As String * 11
    Driver As String * 16
    Conductor As String * 16
    Cleaner As String * 16
    Busno As String * 16
    FareType As Byte
    
    ladies_count As Integer ' asish(30082019)
    seniar_count As Integer
    ladies_coll As Single
    seniar_coll As Single
End Type    'size 256
Public Type fare_type
    TktNo As Long
    FARE As Single
End Type
Public Type Route
    RouteCode As String * 5
    FareType As Byte
    NOS As Byte
    NoOfDupFare As Byte
End Type

Public Type STAGEDETAILS  '''14/01/2011
    StageName As String * 12
    Distance As Single
End Type


Public Type PTicket 'size '48'32
   TicketNo As Long
   From As Byte
   To As Byte
   Full As Byte
   Half As Byte
   st As Byte
   Phy As Byte
   Lugg As Byte
   Amount As Single
   Luggage As Single
   Hr As Byte
   Minut As Byte
   Typ As Byte
   Dy As Byte
   Mn As Byte
   ucTemp As String * 5
   cTripNo As Byte                        'added for kct
   bBusWarrent As Byte
   Refundsts As Byte '1 for refund
   RefundAmt As Single
   ladies_count As Byte 'asish(30082019)
   seniar_count As Byte 'asish(30082019)
   Reserved As String * 10 '12
End Type
''EDITED BY SYAM ON  17-04-2008 FOR BUSTYPE ADD--DATABSE TABLE ADDEDFOR
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Public Type RouteLST
'    Code As String * 5
'    Name As String * 25
'        NoOfStage As Byte
'        MinFare As Single
'        FareType As Byte
'    AllowHalf As Byte
'    AllowConc As Byte
'    AllowPh  As Byte
'    AllowLug As Byte
'    AllowAdjust As Byte
'    StartFrom  As Byte
'    BusType As Byte

'    InterState As Byte
'    FirstState As Byte
'    SecondState As Byte
'    FirstStateCross As Byte
'    SecondStateCross As Byte
'    ISEPoint As Byte

'    OptedKM As Single
'    cTemp As String * 11
' End Type
'nEW> follows
Public Type RouteLST
    Code As String * 5
    Name As String * 25
    NoOfStage As Byte
    MinFare As Single
    FareType As Byte
    AllowHalf As Byte
    AllowConc As Byte
    AllowPh  As Byte
    AllowLug As Byte
    AllowAdjust As Byte
    StartFrom  As Byte
    BusType As Byte
    BusTypeName As String * 16 'NEW FIELD
    OptedKM As Single
'    cTemp As Byte 'String * 1 ' 04/01/2010
    AllowPass As Byte ' 04/01/2010
 End Type
 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Type CONCESSION
    Name As String * 8
    PercentN As Byte
    PercentD As Byte
End Type
 
Public Type SETUP   'BUS.DAT

'''Modified on 14 -03 -2009y syamkrishna as per the code given by R&D
    MainDisp As String * 18
    MainDisp2 As String * 23
    bhl1 As String * 32
    bhl2 As String * 32
    bhl3 As String * 32
    bfl1 As String * 32
    bfl2 As String * 32
    PaperFeed As Byte
    PalmtecID As String * 6
    DefaultFull As Byte
    HalfPer As Byte
    ConPer As Byte
    STMaxAmt As Single
    STMinCon As Single
    PhyPer As Byte
    LuggageUnitRateEdit As Byte
    LuggageUnitRate As Single
    StageUpdation As Byte
    StageDisplayFont As Byte
    
    UseDuplicate As Byte
    UseDup1 As Byte
    Roundoff As Byte
    RoundUp As Byte
    Currency    As String * 8
    RoundAmt As Integer 'in paisa
    '''''''''''''''''''''''''added ON 14-3-2009
    ucbAdjust As Byte
    ucbReviewPasswd As Byte
    ucbReportPasswd As Byte
    ucbSTFromStage As Byte
    ucbSTFareEdit As Byte
    cMasterClearPassword As String * 11
    ReportFlag As Byte
    NextFareRound As Byte
    UpdateStageMsg As Byte '0 - Disable 1 - Enable
    EnableRemoveTicket As Byte '0 - Disable 1 - Enable
    EnableStageFont As Byte '0 - Disable 1 - Enable
    EnableStageDefault As Byte
    PrinterSel As Byte
    OdometerEntry As Byte
    TicketNoBigFont As Byte
    CrewCheck As Byte
    PhNo As String * 13
    TripSMS As Byte
    ScheduleSMS As Byte
   ''''*****added By Sangeetha On 13-08-2012******''''
    TicketRpt As Byte
    Busno As Byte
    Driver As Byte
    Conductor As Byte
    Inspectorreport As Byte
    RepeatST As Byte
    sendbillEnable As Byte
    TripsendEnable As Byte
    SchedulesendEnable As Byte
    Sendpend As Byte
    PhNo2 As String * 13
   'GPRSPACK
    AccessPoint As String * 24
    DestAdds As String * 32
   ' PrtNum As String * 8
    Username As String * 16
    PassWord As String * 16
    Uploadpath As String * 32
    Downloadpath As String * 32
    HttpUrl As String * 64
    GprsEnable As Byte
    MsgPrompt As Byte
    ExpEnable As Byte
    SmartCard As Byte  ''sarika on 05-10-2012
    Modomon As Byte
    FtpEnable As Byte
    RemovePswd As String * 11
    StageReport_E_D As Byte
    StRoundoff_E_D As Byte
    StRoundoff_Amt As Integer
    Simplereport As Byte
    ReportFONT As Byte
    MultiplePass As Byte
    InspectorSMS As Byte
    StageEntry As Byte
    PhNo3 As String * 13
    AutoShutdownEnable As Byte
    UserPasswordEnable As Byte
    DieselEntryEnable As Byte
    TripTimeEnable As Byte
    TripCloseReport As Byte
    ucPaperFeed As Byte
    refund As Byte
    shedule_close_rpt As Byte 'asish(30082019)
    ladis_per As Byte 'asish(30082019)
    seniar_per As Byte 'asish(30082019)
    ucTemp As String * 70 '73 '454 '470 '456
    
   ' ucTemp As String * 474 '489 ' //960+64-> 1024BYTE
''''********End**********''''
End Type

Public Type P_DATE
    da_day As Byte
    da_mon As Byte
    da_year As Integer
End Type
Public Type P_TIME
    ti_hour As Byte
    ti_min As Byte
    ti_sec As Byte
    ti_hund As Byte
End Type

''''''''''''''''''''''''''''*************************'''''''''''''''''''''''
'RNC 16082011
Public Type ODOMETER
    ucScheduleNo As Byte
    ucTripNo As Byte
    Driver As String * 16
    Busno As String * 16
    Startr As Single
    Endr As Single
    SHour As String * 1
    SMinutes As String * 1
    SDay As Byte
    SMonth As Byte
    SYear As Integer
    EHour As String * 1
    EMinutes As String * 1
    EDay As Byte
    EMonth As Byte
    EYear As Integer
   Reserved As String * 10
End Type


Public Type INSPECTORDET
    ucInspectorId As String * 8
    ucScheduleNo As Byte
    ucTripNo As Byte
    RouteCode As String * 5
    Stage As String * 12
    InsDay As Byte
    InsMonth As Byte
    InsYear As Integer
    InsHour As Byte
    InsMinutes As Byte
    Reserved As String * 31
End Type


Public Type EXPENSEDET
    ucType As String * 5
    expname As String * 16
    Reserved As String * 43
End Type

Public Type EXPENSES

    ucScheduleNo As Byte
    ucTripNo As Byte
    EName As String * 16
    Busno As String * 16
     fExpens As Single
    fDiesel As Single
    ucType As Byte
    Hour As Byte
    Minutes As Byte
    Day As Byte
    Month As Byte
    Year As Integer
'    Eid As String * 4
    RptNo As Long
    Reserved As String * 11
End Type




''''''''''''''''''''''''''''*************************'''''''''''''''''''''''



Public Type HARDWARE_SETUP
    Ptime As P_TIME
    Pdate As P_DATE
    MSR_PSWD As String * 11
    USR_PSWD As String * 11
    SPR_PSWD As String * 11
    val_contrast As Byte
    val_brightness As Byte
    screensaver_onoff As Byte
    backlit_timer As Byte
    keyhitdelay As Byte
    boarder_en As Byte
    dooropen_alert As Byte
    paperout_alert As Byte
    ucHalfPagePrinter As Byte
    buzz_onoff As Byte
    rs232_baud As Byte
    ir_baud As Byte
    rf_baud As Byte
    connecting_medium As Byte
    footer_stat As Byte
    select_language As Byte
    login_mode As Byte
    ucKPLight_opt As Byte
    usShuntdownTime As Integer
    LangNo   As Byte
    ucTemp As String * 2
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Type CREWDETAILS
'    Driver As String * 16
'    Conductor As String * 16
'    Cleaner As String * 16
'    BusNo As String * 16
'    ucType As Byte
'    DrvrID As String * 5
'    CndrID As String * 5
'    ClnrID As String * 5
'    BusId As String * 5   'BusTypeId
'    'BusType As String * 1
'    BusTypeName As String * 16
'    Reserved As String * 27
'End Type

Public Type CREWDET
    EmpName As String * 16
    EmpId As String * 8
    EmpType As Byte
    PassWord As String * 7
   
End Type

Public Type VEHICLE
    BUSID As Byte
    Busno As String * 16
    Reserved As String * 15
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type Stage
    StageName As String * 12
    Distance As Integer
    temp As String * 2
End Type

Public Type LANGUAGE_STAG
    RouteCode As String * 5
    LocalLanguageStageName(23) As Byte
    stagecode As Integer
    temp As Byte
End Type

Private Type Footer
    FooterString As String * 31
End Type
Public Enum ValidateOptions
    IntegerValue = 0
    FloatingPointValue = 1
    StrictAlphaNumeric = 2
    AlphaNumeric = 3
    Other = 4
End Enum
Public ValidationMode As ValidateOptions
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwmilliseconds As Long)
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public TEST As Integer
Dim Fso As New FileSystemObject

Public strBmpName As String

'Public Const gblstrPrjTitle = "PVTBUS Normal 2.17 190515"
Public Const gblstrPrjTitle = "Amphibia Bus Ticketing 2.17"

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const FLOAT = 1, SINK = 0


Public Function CreateTables()
Dim rs1 As ADODB.Recordset
Dim DB As New ADODB.Connection
Dim sDataBase As String
sDataBase = App.Path & "\PVT.MDB"
CreateTables = False
    
  
   
 If CheckTableExistsOrNot("FARERPT") = False Then
  Set rs1 = New ADODB.Recordset
                      If DB.State <> 0 Then DB.Close
                        DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source= " & sDataBase & "; Jet OLEDB:Database Password =silbus"

                    If rs1.State <> 0 Then rs1.Close
                    RSql = "CREATE TABLE FARERPT " & _
                            "(SLNO AUTOINCREMENT , " & _
                            "TKTNO NUMBER, " & _
                            "SCHNO NUMBER, " & _
                            "TRPNO NUMBER, " & _
                            "SCH_STDATE TEXT(15), " & _
                            "SCH_ENDDATE TEXT(15), " & _
                            "RTCD TEXT(50), " & _
                            "RTNME TEXT(50), " & _
                            "FARE NUMBER, " & _
                            "PID TEXT(10))"
                   rs1.Open RSql, DB, adOpenDynamic, adLockOptimistic
                 If rs1.State <> 0 Then rs1.Close
End If
If CheckTableExistsOrNot("TMPFARERPT") = False Then
  Set rs1 = New ADODB.Recordset
                      If DB.State <> 0 Then DB.Close
                        DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source= " & sDataBase & "; Jet OLEDB:Database Password =silbus"

                    If rs1.State <> 0 Then rs1.Close
                    RSql = "CREATE TABLE TMPFARERPT " & _
                            "(SLNO AUTOINCREMENT , " & _
                            "TKTNO NUMBER, " & _
                            "FARE NUMBER) "
                   rs1.Open RSql, DB, adOpenDynamic, adLockOptimistic
                 If rs1.State <> 0 Then rs1.Close
End If
 If CheckTableExistsOrNot("FARESLAB") = False Then
  Set rs1 = New ADODB.Recordset
                      If DB.State <> 0 Then DB.Close
                        DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source= " & sDataBase & "; Jet OLEDB:Database Password =silbus"

                    If rs1.State <> 0 Then rs1.Close
                    RSql = "CREATE TABLE FARESLAB " & _
                            "(SLNO AUTOINCREMENT , " & _
                            "BUSID NUMBER, " & _
                            "STARTKM NUMBER, " & _
                            "ENDKM NUMBER, " & _
                            "FARE NUMBER)"
                   rs1.Open RSql, DB, adOpenDynamic, adLockOptimistic
                 If rs1.State <> 0 Then rs1.Close
End If
 If CheckTableExistsOrNot("Waybill_tab") = False Then
  Set rs1 = New ADODB.Recordset
                      If DB.State <> 0 Then DB.Close
                        DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source= " & sDataBase & "; Jet OLEDB:Database Password =silbus"

                    If rs1.State <> 0 Then rs1.Close
                    RSql = "CREATE TABLE Waybill_tab " & _
                            "(waybill_ID AUTOINCREMENT , " & _
                            "waybill_number TEXT(50), " & _
                            "Duty_date TEXT(20), " & _
                            "Conductor TEXT(150), " & _
                            "Driver TEXT(150), " & _
                            "Schedule_KM NUMBER, " & _
                            "Schedule_Trip TEXT(150), " & _
                            "Bus_no TEXT(100), " & _
                            "Service_No TEXT(200))"
                   rs1.Open RSql, DB, adOpenDynamic, adLockOptimistic
                 If rs1.State <> 0 Then rs1.Close
End If
    
    Set rs1 = New ADODB.Recordset
    If DB.State <> 0 Then DB.Close
      DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source= " & sDataBase & "; Jet OLEDB:Database Password =silbus"
       If rs1.State <> 0 Then rs1.Close
       RSql = "select EmployeeTypeId from EMPLOYEETYPE where EmployeeTypeId=4"
       Set rs1 = New ADODB.Recordset
      Set rs1 = DB.Execute(RSql)
   
       If rs1.EOF = True Then
            RSql = "insert into EMPLOYEETYPE(EmployeeTypeId,TypeName) values('4','Inspector')"
            DB.Execute RSql
       End If
       
       
      
End Function

Private Function MOVE_DOWNLOAD_FILES()
 On Error GoTo err
 If Dir$(App.Path & "\DOWNLOAD", vbDirectory) = "" Then
    MkDir App.Path & "\DOWNLOAD"
 End If
 
 If Dir(App.Path & "\*.LST") <> "" Then
    ret = Shell("MOVE *.LST " & App.Path & "\DOWNLOAD", vbHide)
 End If
 If Dir(App.Path & "\*.DAT") <> "" Then
    ret = Shell("MOVE *.DAT " & App.Path & "\DOWNLOAD", vbHide)
 End If
err:
    MsgBox err.Number & " , " & err.Description
    
    Exit Function
End Function
Public Sub stayOnTop(frm As Form)
    Call SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 1)
    
End Sub

Public Function MOVE_ALL_FILES() 'Move *.dat, *.txt and *.lst files
  
  Dim fname As String
  Dim I As Integer
  fname = ""

End Function
Public Function STAGE_CONVERT() As Boolean
On Error GoTo err
Dim fname As String
Dim Buffer As String
Dim rea As String
Dim buff As String
Dim Thandle As Integer
Dim lHandle As Integer
Dim I As Integer
Dim Blen As Integer
Dim stg As Stage
Dim Distance As Integer
Distance = 0
I = 0
fname = App.Path & "\stage.txt"
Thandle = FreeFile()
If Dir(fname) <> "" Then
    Open fname For Input As #Thandle
        lHandle = FreeFile()
        Open App.Path & "\stage.lst" For Binary Access Write As #lHandle
        Do While Not (EOF(Thandle))
        Input #Thandle, buff
      '  Buffer = rea
        Blen = Len(buff)
        If Blen < 12 Then
            For I = 1 To 12 - Blen
                 buff = buff & Chr(0)
            Next
        End If
        stg.StageName = buff
        stg.Distance = Distance
        Distance = Distance + 5
        stg.temp = Chr(0)
'         If Blen > 12 Then
'            buff = Mid(buff, 1, 12) & Chr(0)
'         Else
'            buff = buff & Space(12 - Blen)
'            buff = Mid(buff, 1, 12) & Chr(0)
'         End If
         
          Put #lHandle, , stg
        Loop
    Close #Thandle
    Close #lHandle
    STAGE_CONVERT = True
Else
    MsgBox "STAGE.TXT Not found ", vbInformation, "BUS"
    
End If

Exit Function
err:
    MsgBox err.Number & " , " & err.Description
   
    
    Close #Thandle
    Close #lHandle
    STAGE_CONVERT = False
    Exit Function
End Function
'Public Sub main()
'Load loginform
'frmMainform.Enabled = False
'frmMainform.Show
'loginform.Show vbModal
'End Sub
Public Function TktsConvert(fname As String) As Boolean
'TKTS Schid.DAT - TKTS shid.csv
On Error GoTo err
Dim Handle As Integer
'Dim fname As String
Dim buff As String
Dim scid As String
Dim file As TextStream
Dim tk As PTicket
Dim st As String
Dim Mode As String
Dim ctype As Integer
Dim pos As Integer

    scid = Mid(fname, 5, 2)
    
    ReDim arty(0 To 6) As String
    
    TktsConvert = True
    ctype = 0
    arty(0) = "CASH"
    arty(1) = "ADJUST"
    arty(2) = "CNONC"
    arty(3) = "FREE"
    arty(4) = "PHY"
    arty(5) = "DUP1"
    arty(6) = "DUP2"
    
    st = App.Path & "\TKTS" & scid & ".DAT"
    
    Handle = FreeFile()
    Open st For Binary Access Read As #Handle
    
    Set file = Fso.OpenTextFile(App.Path & "\TKTS" & scid & ".CSV", ForWriting, True)
    file.Write "TKTNO,AMOUNT,LUGGAGE,TKT FROM,TKT TO,FULL,HALF,ST,PH,HR:MIN,TYPE" & vbCrLf
    
    Do While Not EOF(Handle)
        Get #Handle, , tk
        If EOF(Handle) Then Exit Do
        buff = ""
        buff = tk.TicketNo & ","
        buff = buff & tk.Amount & "," '
        buff = buff & tk.Luggage & "," '
        buff = buff & tk.From & ","
        buff = buff & tk.To & ","
        buff = buff & tk.Full & ","
        buff = buff & tk.Half & ","
        buff = buff & tk.st & ","
        buff = buff & tk.Phy & ","
        buff = buff & tk.Hr & ":" & tk.Minut & ","
        If (tk.Typ = T_CASH) Then
            If (tk.Phy) Then
                ctype = 4
            Else
                ctype = 0
            End If
        ElseIf (tk.Typ = T_ADJUST) Then
            ctype = 1
        ElseIf (tk.Typ = T_ADJUST + 1) Then
            ctype = 5
        ElseIf (tk.Typ = T_ADJUST + 2) Then
            ctype = 6
        ElseIf (tk.Typ >= T_CONC And tk.Typ <= T_CONC_E) Then
            If (tk.Amount) Then
                ctype = 2
            Else
                ctype = 3
            End If
        End If
        buff = buff & arty(ctype) & vbCrLf
        file.Write buff
    Loop
    file.Close
    Close #Handle
    DBTKTS (st)
    TktsConvert = True
    Exit Function
err:
    MsgBox err.Number & " , " & err.Description
    file.Close
    
    Close #Handle
    TktsConvert = False
    Exit Function
End Function
Public Function RptConvert(fname As String) As Boolean
'RPT SCHD.DAT - RPT schid.CSV & Ticket.TXT
On Error GoTo err
Dim Handle As Integer
Dim Thandle As Integer
'Dim fname As String
Dim buff As String
Dim Sch(100) As Long
Dim isch As Long
Dim file As TextStream
Dim rp As PReport
Dim st As String
Dim Mode As String
Dim scid As String
    
    RptConvert = True
    isch = 1
    Sch(1) = 0
    scid = Mid(fname, 4, 2)
    
'    RSQL = "SELECT * FROM RPT"
'    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB")
'    Set RES = DB.OpenRecordset(RSQL, dbOpenDynaset)
'    If RES.RecordCount > 0 Then RES.MoveLast
    Handle = FreeFile()
    Open App.Path & "\" & fname For Binary Access Read As #Handle
    
    Thandle = FreeFile()
    Open App.Path & "\Ticket.txt" For Output As #Thandle
    
    Set file = Fso.OpenTextFile(App.Path & "\Rpt" & scid & ".CSV", ForWriting, True)
    file.Write "TKTSNO,TKTENO,INHAND AMT,TOT COLL,FULL COLL,HALF COLL,PH COLL,LUG COLL,ST COLL,ADJ COLL,FULL,HALF,PH,ST,FREE,ADJUST,CONC,SHEDULE NO,TRIPNO,UP/DN,ROUTE CODE,START DATE,START TIME,END DATE,END TIME,DRIVER,CONDUCTOR,CLEANER,BUS NO,MISSED BILLS" & vbCrLf
    
    Do While Not EOF(Handle)
        Get #Handle, , rp
        If EOF(Handle) Then Exit Do
        buff = ""
        buff = rp.STicketNo & ","
        buff = buff & rp.ETicketNo & ","
        buff = buff & rp.InHandAmount & ","
        buff = buff & rp.TotalColl & ","
        buff = buff & rp.FullColl & ","
        buff = buff & rp.HalfColl & ","
        buff = buff & rp.PhyColl & ","
        buff = buff & rp.LuggageColl & ","
        buff = buff & rp.STColl & ","
        buff = buff & rp.AdjustColl & ","
        buff = buff & rp.Full & ","
        buff = buff & rp.Half & ","
        buff = buff & rp.Phy & ","
        buff = buff & rp.st & ","
        buff = buff & rp.Free & ","
        buff = buff & rp.Adjust & ","
        buff = buff & rp.Conc & ","
        buff = buff & rp.SCHEDULE & ","
        buff = buff & rp.TripNo & ","
        buff = buff & Chr(rp.UpDownTrip) & ","
        buff = buff & rp.RouteCode & ","
        buff = buff & rp.StartD & "/" & rp.StartMO & "/" & rp.StartY & ","
        buff = buff & rp.StartH & ":" & rp.StartM & ","
        buff = buff & rp.EndD & "/" & rp.EndMO & "/" & rp.EndY & ","
        
        
        'Ticket.TXT
        If (Sch(isch - 1) <> rp.SCHEDULE) Then 'Or (sch(isch - 1) <> isch)
            Sch(isch) = rp.SCHEDULE
            Print #Thandle, "TKTS" & Format(rp.SCHEDULE, "00") & ".DAT  - " & Format(rp.StartD, "00") & "/" & Format(rp.StartMO, "00") & "/" & Format(rp.StartY, "2000") & Chr(0)
            isch = isch + 1
        End If
        buff = buff & rp.EndH & ":" & rp.EndM & ","
        buff = buff & rp.Driver & ","
        buff = buff & rp.Conductor & ","
        buff = buff & rp.Cleaner & ","
        buff = buff & rp.Busno & ","
        buff = buff & rp.FareType & vbCrLf
        file.Write buff
    Loop
    file.Close
    Close #Thandle
    Close #Handle
    DBRpt (fname)
    Exit Function
err:
    MsgBox err.Number & " , " & err.Description
    file.Close
    
    Close #Thandle
    Close #Handle
    RptConvert = False
    Exit Function
End Function

Public Function ReportCnv(Pmid As String, scid As String) As Boolean
On Error GoTo err
Dim Handle As Integer
Dim file As TextStream
Dim TripFile As TextStream
Dim Mode As String 'cash cheqye
Dim tk As PTicket ' structure asignd
Dim rp As PReport ' structure asignd
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim TktNo(150) As Long
Dim Trip(100) As Long
Dim N_Trip(50) As Long
Dim TotalTkts As Integer
Dim TotalAmt As Double
Dim HeadRead As Integer
Dim TrpCount As Integer
Dim Busno(150) As String
Dim busdr(150) As String
Dim cond(150) As String
Dim rou(150) As String
Dim Dt(150) As String
Dim SCHEDULE(150) As String
Dim ctype As Integer
Dim trpid As Integer
Dim TrpTotTkts As Integer
Dim TrpTotAmt As Double
Dim PrnFlag As Integer
'----------------report header
Dim rptScID As String
Dim rptTripNo As String
Dim rptDate As String
Dim rptBuso As String
Dim rptBusDr As String
Dim rptBusCond As String
Dim rptRoute As String
Dim rptSDate As String
Dim rptEDate As String
Dim rptSTime As String
Dim rptETime As String
'----------------------------
    ReDim arty(0 To 6) As String * 6
    
    ReportCnv = True
    ctype = 0
    arty(0) = "CASH"
    arty(1) = "ADJUST"
    arty(2) = "CNONC"
    arty(3) = "FREE"
    arty(4) = "PHY"
    arty(5) = "DUP1"
    arty(6) = "DUP2"
    
    TrpCount = 1
    j = 0
    k = 0
    
    Handle = FreeFile()
    Open App.Path & "\Rpt" & Pmid & ".DAT" For Binary Access Read As #Handle
    Do While Not EOF(Handle)
        Get #Handle, , rp
        
        If EOF(Handle) Then Exit Do
        Busno(j) = rp.Busno
        busdr(j) = rp.Driver
        cond(j) = rp.Conductor
        rou(j) = rp.RouteCode
        Dt(j) = rp.StartD & "/" & rp.StartMO & "/" & rp.StartY
        TktNo(j) = rp.ETicketNo
        Trip(k) = rp.TripNo
        j = j + 1
        k = k + 1
    Loop
    
    Close #Handle
    
    j = 0
    k = 0
    
    While (Trip(j))
        If (Trip(j + 1) = 1 Or Trip(j + 1) = 0) Then
            N_Trip(k) = TrpCount
            TrpCount = 1
            j = j + 1
            k = k + 1
        Else
            TrpCount = TrpCount + 1
            j = j + 1
        End If
    Wend
    trpid = 1
    
    Handle = FreeFile()
    Open App.Path & "\Rpt" & Pmid & ".DAT" For Binary Access Read As #Handle
    Do While Not EOF(Handle)
        Get #Handle, , rp
        If EOF(Handle) Then Exit Do
        If rp.SCHEDULE = scid Then
            rptScID = scid
            rptTripNo = rp.TripNo
            rptSDate = Format(rp.StartD, "00") & "/" & Format(rp.StartMO, "00") & "/20" & Format(rp.StartY, "00")
            rptSTime = Format(rp.StartH, "00") & ":" & Format(rp.StartH, "00")
            rptEDate = Format(rp.EndD, "00") & "/" & Format(rp.EndMO, "00") & "/20" & Format(rp.EndY, "00")
            rptETime = Format(rp.EndH, "00") & ":" & Format(rp.EndM, "00")
            rptBuso = TrimChr(rp.Busno)
            rptBusDr = TrimChr(rp.Driver)
            rptBusCond = TrimChr(rp.Conductor)
            rptRoute = TrimChr(rp.RouteCode)
            Exit Do
        End If
    Loop
    Close #Handle
    
    Set file = Fso.OpenTextFile(App.Path & "\REPORT" & scid & ".TXT", ForWriting, True)
    file.WriteLine "-----------------------------------------------------------------------"
    HeadRead = FreeFile()
    Open App.Path & "\Bus.DAT" For Binary Access Read As #HeadRead
    Get #HeadRead, , HStr
    Close #HeadRead
    StrHeader = Space(24) & "SCHEDULE REPORT"
    file.WriteLine StrHeader
    
    StrHeader = ""
    StrHeader = Space(27) & strSpace(Mid$(HStr.bhl1, 1, InStr(1, HStr.bhl2, Chr(0)) - 1), 36)
    If Trim(HStr.bhl2) <> "" Then _
                StrHeader = StrHeader & vbCrLf & Space(27) & strSpace(Mid$(HStr.bhl2, 1, InStr(1, HStr.bhl2, Chr(0)) - 1), 36)
    If Trim(HStr.bhl3) <> "" Then _
                StrHeader = StrHeader & vbCrLf & Space(27) & strSpace(Mid$(HStr.bhl3, 1, InStr(1, HStr.bhl3, Chr(0)) - 1), 36)
     
    file.WriteLine StrHeader
    '-------------- newly added by deej on 11.05.07
    StrHeader = ""
    StrHeader = "SCHEDULE NO:" & Space(1) & Format(scid, "00") & Space(11) & "TRIP NO:" & Space(1) & Format(rptTripNo, "000") & Space(15) & "ROUTE NO :" & Format(rptRoute, "0000")
    file.WriteLine StrHeader
    file.WriteLine "-----------------------------------------------------------------------"
    StrHeader = ""
    StrHeader = "BUS No:" & Space(1) & Format(rptBuso, Space(16))
    file.WriteLine StrHeader
    file.WriteBlankLines (1)
    StrHeader = "BUS DRIVER :" & Format(rptBusDr, Space(16)) & Space(2) & "BUS CONDUCTOR :" & Format(rptBusCond, Space(16))
    file.WriteLine StrHeader
    file.WriteLine "-----------------------------------------------------------------------"
    file.WriteBlankLines (1)
    StrHeader = "START DATE:" & rptSDate & Space(2) & "START TIME :" & rptSTime
    file.WriteLine StrHeader
    StrHeader = "END DATE" & Space(2) & ":" & rptEDate & Space(2) & "END TIME" & Space(3) & ":" & rptETime
    file.WriteLine StrHeader
    
    'END DATE  :DD/MM/YYYY  END TIME   : HH:MM
    'START DATE:00/00/0000  START TIME : HH:MM
    '-------------------*******************
    file.WriteLine "-----------------------------------------------------------------------"
    file.WriteLine "NO OF TRIPS " & Space(3) & ":" & Space(2) & N_Trip(CLng(scid) - 1)
    file.WriteLine "-----------------------------------------------------------------------"
    file.WriteLine "TICKET NO | FROM |  TO  | FULL | HALF | LUGG | ST | PHY| TYPE |AMOUNT" & vbCrLf
    file.WriteLine "-----------------------------------------------------------------------"
    'j = 0
    'k = 0
    TotalTkts = 0
    TotalAmt = 0
    
    Handle = FreeFile()
    Open App.Path & "\TKTS" & scid & ".DAT" For Binary Access Read As #Handle
    Do While Not EOF(Handle)
        
        Get #Handle, , tk
        
        If EOF(Handle) Then Exit Do
        
        If (tk.Typ = T_CASH) Then
            If (tk.Phy) Then
            ctype = 4
            Else
            ctype = 0
            End If
        ElseIf (tk.Typ = T_ADJUST) Then
            ctype = 1
        ElseIf (tk.Typ = T_ADJUST + 1) Then
            ctype = 5
        ElseIf (tk.Typ = T_ADJUST + 2) Then
            ctype = 6
        ElseIf (tk.Typ >= T_CONC And tk.Typ <= T_CONC_E) Then
            If (tk.Amount) Then
                ctype = 2
            Else
                ctype = 3
            End If
        End If
        
        If PrnFlag = 0 Then
            If Dir(App.Path & "\" & FileDateCret & "\" & pmtcID & "\TripDetails\Trip" & scduID & trpid & ".TXT") <> "" Then
                Fso.DeleteFile App.Path & "\" & FileDateCret & "\" & pmtcID & "\TripDetails\Trip" & scduID & trpid & ".TXT"
            End If
            'fortesting in default folder
            Set TripFile = Fso.OpenTextFile(App.Path & "\" & FileDateCret & "\" & pmtcID & "\TripDetails\Trip" & scid & trpid & ".TXT", ForWriting, True)
            '   Set TripFile = Fso.OpenTextFile(App.Path & "\Trip" & scid & trpid & ".TXT", ForWriting, True) ' & FileDateCret & "\" & pmtcID & "\TripDetails\Trip"
'            TripFile.WriteLine Space(27) & strSpace(Trim(HStr.bhl1), 36)
'            TripFile.WriteLine Space(28) & strSpace(Trim(HStr.bhl2), 36)
'            TripFile.WriteLine Space(29) & strSpace(Trim(HStr.bhl3), 36)
'            TripFile.WriteLine Space(30) & strSpace(Trim(HStr.bhl4), 36)
            StrHeader = ""
            StrHeader = Space(27) & strSpace(Mid$(HStr.bhl1, 1, InStr(1, HStr.bhl1, Chr(0)) - 1), 36)
            If Trim(HStr.bhl2) <> "" Then _
                        StrHeader = StrHeader & vbCrLf & Space(27) & strSpace(Mid$(HStr.bhl2, 1, InStr(1, HStr.bhl2, Chr(0)) - 1), 36)
            If Trim(HStr.bhl3) <> "" Then _
                        StrHeader = StrHeader & vbCrLf & Space(27) & strSpace(Mid$(HStr.bhl3, 1, InStr(1, HStr.bhl3, Chr(0)) - 1), 36)
             
            TripFile.WriteLine StrHeader
            
            TripFile.WriteLine "-------------------------------------------------------------------------"
            TripFile.WriteLine "TRIP NO " & Space(8) & ":" & Space(2) & trpid & Space(10) & "SCHEDULE NO :" & rptScID
            TripFile.WriteLine "DRIVER NAME " & Space(3) & " :" & Space(2) & strSpace(Mid$(busdr(ArryCount), 1, InStr(1, busdr(ArryCount), Chr(0)) - 1), 16) & Space(15) & "DATE  :" & Format(Dt(ArryCount), "dd-MM-yy")
            TripFile.WriteLine "CONDUCTOR NAME  :" & Space(2) & strSpace(Mid$(cond(ArryCount), 1, InStr(1, cond(ArryCount), Chr(0)) - 1), 16) & Space(15) & "ROUTE :" & Mid$(rou(ArryCount), 1, InStr(1, rou(ArryCount), Chr(0)) - 1)
            TripFile.WriteLine "-------------------------------------------------------------------------"
            TripFile.WriteLine "TICKET NO | FROM |  TO  | FULL | HALF | LUGG | ST | PHY| TYPE | AMOUNT " & vbCrLf  '
            TripFile.WriteLine "-------------------------------------------------------------------------"
            PrnFlag = 1
        End If
        file.WriteLine "" & Space(2) & Format(tk.TicketNo, "000000") & Space(2) & "|" & Space(2) & strSpace(CStr(tk.From), 2) & Space(2) & "|" & Space(2) & strSpace(CStr(tk.To), 2) & Space(2) & "|" & Space(2) & strSpace(CStr(tk.Full), 2) & Space(2) & "|" & Space(2) & strSpace(CStr(tk.Half), 2) & Space(2) & "|" & Space(2) & strSpace(CStr(tk.Lugg), 2) & Space(2) & "|" & Space(1) & strSpace(CStr(tk.st), 2) & Space(1) & "|" & Space(1) & strSpace(CStr(tk.Phy), 2) & Space(1) & "|" & Space(0) & arty(ctype) & Space(0) & "|" & Space(1) & tk.Amount    '
        TripFile.WriteLine "" & Space(2) & Format(tk.TicketNo, "000000") & Space(2) & "|" & Space(2) & strSpace(CStr(tk.From), 2) & Space(2) & "|" & Space(2) & strSpace(CStr(tk.To), 2) & Space(2) & "|" & Space(2) & strSpace(CStr(tk.Full), 2) & Space(2) & "|" & Space(2) & strSpace(CStr(tk.Half), 2) & Space(2) & "|" & Space(2) & strSpace(CStr(tk.Lugg), 2) & Space(2) & "|" & Space(1) & strSpace(CStr(tk.st), 2) & Space(1) & "|" & Space(1) & strSpace(CStr(tk.Phy), 2) & Space(1) & "|" & Space(0) & arty(ctype) & Space(0) & "|" & Space(1) & tk.Amount
        TrpTotTkts = TrpTotTkts + 1
        TrpTotAmt = TrpTotAmt + tk.Amount
        
        Dim n As Integer
        n = 0
        Do While Not n = j
            If TktNo(n) = tk.TicketNo Then
                file.WriteLine "-----------------------------------------------------------------------"
                TripFile.WriteLine "-------------------------------------------------------------------------"
                TripFile.WriteLine "TOTAL NO.OF TICKETS :" & TrpTotTkts
                TripFile.WriteLine "TotalAmt " & Space(11) & ":" & TrpTotAmt
                TrpTotTkts = 0
                TrpTotAmt = 0
                ArryCount = ArryCount + 1
                trpid = trpid + 1
                TripFile.Close
                PrnFlag = 0
                Exit Do
            End If
            n = n + 1
        Loop
        
        TotalTkts = TotalTkts + 1
        TotalAmt = TotalAmt + tk.Amount
    
    Loop
    Close #Handle
    
    file.WriteLine "-----------------------------------------------------------------------"
    file.WriteLine "TOTAL NO.OF TICKETS :" & TotalTkts
    file.WriteLine "Total Amount " & Space(7) & ":" & TotalAmt
    If PrnFlag = 1 Then
        TripFile.WriteLine "-------------------------------------------------------------------------"
        TripFile.WriteLine "TOTAL NO.OF TICKETS :" & TrpTotTkts
        TripFile.WriteLine "TotalAmt " & Space(11) & ":" & TrpTotAmt
        TripFile.Close
    End If
    file.Close
    ReportCnv = True
    
    Exit Function
err:
    MsgBox err.Number & " , " & err.Description
    Close #Handle
    file.Close
    ReportCnv = False
    Exit Function
End Function


'Public Function RouteLST() As Boolean
'On Error GoTo err
'Dim handle As Integer
'Dim Buf As String
'Dim FBuf As String
'Dim Fhandle As Integer
'Dim Route As RouteLST
'handle = FreeFile()
'Open App.Path & "\RouteLst.TXT" For Input As #handle
'Fhandle = FreeFile()
' Open App.Path & "\RouteLst.LST" For Binary Access Write As #Fhandle
'   Do While Not EOF(handle)
'    Line Input #handle, Buf
'        Route.Code = ParseTreeData(Buf) & Chr(0)
'        Route.Name = ParseTreeData(Buf) & Chr(0)
'        Route.NoOfStop = CInt(Val(ParseTreeData(Buf)))
'        Route.NoOfStage = CInt(Val(ParseTreeData(Buf)))
'        Route.MinFare = CInt(Val((ParseTreeData(Buf))))
'        Route.FareType = CInt(Val(ParseTreeData(Buf)))
'        Route.UseStop = CByte(Val(ParseTreeData(Buf)))
'        Route.AllowHalf = CByte(Val(ParseTreeData(Buf)))
'        Route.AllowLug = CByte(Val(ParseTreeData(Buf)))
'        Route.AllowST = CByte(Val(ParseTreeData(Buf)))
'        Route.AllowAdjust = CByte(Val(ParseTreeData(Buf)))
'        Route.AllowConc = CByte(Val(ParseTreeData(Buf)))
'        Route.AllowPh = CByte(Val(ParseTreeData(Buf)))
'        Route.StartFrom = CByte(Val(Buf))
'
'        Put #Fhandle, , Route
'   Loop
'Close #Fhandle
'Close #handle
' RouteLST = True
'
'  Exit Function
'err:
'    MsgBox err.Number & " , " & err.Description
'    Close #Fhandle
'Close #handle
'    RouteLST = False
'    Exit Function
'End Function


'Public Function FareConv(fname As String) As Boolean
'On Error GoTo err
'Dim handle As Integer
'Dim Fhandle As Integer
'Dim strfn  As String
'Dim Buf As String
'Dim Longstr As Long
'Dim IntStr As Integer
'Dim chkFlag As Integer
'chkFlag = 0
'strfn = Mid(fname, 1, 3)
'
'handle = FreeFile()
' Open App.Path & "\" & fname For Input As #handle
'Fhandle = FreeFile()
'Open App.Path & "\" & strfn & ".DAT" For Binary Access Write As #Fhandle
'
'
'    Do While Not EOF(handle)
'        Line Input #handle, Buf
'       Do While Not Buf = ""
'            Longstr = CLng(Val(ParseTreeData(Buf)))
'            Put #Fhandle, , Longstr
'       Loop
'        Line Input #handle, Buf
'       Do While Not Buf = ""
'            IntStr = CInt(Val(ParseTreeData(Buf)))
'            Put #Fhandle, , IntStr
'       Loop
'    Loop
'
' Close #handle
'Close #Fhandle
'FareConv = True
'Exit Function
'err:
'    MsgBox err.Number & " , " & err.Description
'    Close #Fhandle
'Close #handle
'    FareConv = False
'    Exit Function
'End Function

Public Function ParseTreeData(pstrData As String)
Dim strWorking As String
Dim pos As Integer

 If InStr(1, pstrData, ",") = 0 Then
    strWorking = pstrData
    pstrData = vbNullString
Else
    strWorking = Mid(pstrData, 1, InStr(1, pstrData, ",") - 1)
    pstrData = Mid(pstrData, InStr(1, pstrData, ",") + 1)
End If
ParseTreeData = strWorking
End Function
Public Function strSpace(str As String, siz As Integer)
Dim pos As Integer
Dim S As Integer
    pos = Len(str)
    S = siz - pos
    If pos = 0 Or pos > 0 Then
       strSpace = str & Space(S)
    End If
End Function
Public Function DBRpt(fname As String)
On Error GoTo err
Dim Filehdl As Integer
Dim freefileInt As Integer
Dim rp As PReport
Dim fare_obj As fare_type
Dim gp As PASSCONC
Dim sql As String
Dim sHandle As Integer
Dim PalmID As String
Dim sqlRPT As String
Dim pos As Long
    pos = 0
    sHandle = FreeFile()
    Open App.Path & "\BUS.DAT" For Binary Access Read As #sHandle
        Get #sHandle, , HStr
    Close #sHandle
    PalmID = Mid$(HStr.PalmtecID, 1, InStr(1, HStr.PalmtecID, Chr(0)) - 1)
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    Filehdl = FreeFile()
    Open App.Path & "\" & fname For Binary Access Read As #Filehdl
        Do While Not EOF(Filehdl)
            Get #Filehdl, , rp
            If rp.SCHEDULE = 0 Then Exit Do
            Dim DA As String, DAA As String
            DA = rp.StartD & "/" & rp.StartMO & "/" & rp.StartY
            DAA = rp.EndD & "/" & rp.EndMO & "/" & rp.EndY
            sql = "DELETE * FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & rp.SCHEDULE & " AND TRIPNO= " & rp.TripNo & " AND DATE BETWEEN DateValue('" & Format(DA, "DD/MM/YYYY") & "') AND DateValue('" & Format(DAA, "DD/MM/YYYY") & "')"
            Set RES = DB.OpenRecordset("RPT", dbOpenDynaset)
            Dim f1 As Boolean
            f1 = False
            If RES.RecordCount > 0 Then RES.MoveLast
            RES.AddNew
            f1 = True
            If (getvalueQuery("select Count(*) from RPT where PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & rp.SCHEDULE & " AND TRIPNO= " & rp.TripNo & " AND DATE BETWEEN DateValue('" & Format(DA, "DD/MM/YYYY") & "') AND DateValue('" & Format(DAA, "DD/MM/YYYY") & "')")) = 0 Then
                With RES
                    DA = rp.StartD & "/" & rp.StartMO & "/" & rp.StartY
                    !Date = Format(DA, "DD/MM/YYYY")
                    !PalmID = TrimChr(HStr.PalmtecID) 'SANGEETHA
                    !STicket = rp.STicketNo
                    !ETicketNo = rp.ETicketNo
                    !InHandAmount = rp.InHandAmount
                    !TotalColl = rp.TotalColl
                    !FullColl = rp.FullColl
                    !HalfColl = rp.HalfColl
                    !PhyColl = rp.PhyColl
                    !LuggageColl = rp.LuggageColl
                    !STColl = rp.STColl
                    !AdjustColl = rp.AdjustColl
                    !Fulls = rp.Full
                    !Half = rp.Half
                    !Phy = rp.Phy
                    !Luggage = rp.Lugg  'SANGEETHA
                    !st = rp.st
                    !Free = rp.Free
                    !Adjust = rp.Adjust
                    !Conc = rp.Conc
                    !SCHEDULE = rp.SCHEDULE
                    !TripNo = rp.TripNo
                    !UpDownTrip = Chr(rp.UpDownTrip)
                    !RouteCode = TrimChr(rp.RouteCode)
                    If rp.StartD <> 0 And rp.StartMO <> 0 And rp.StartY <> 0 Then !StartDate = Format(rp.StartD & "/" & rp.StartMO & "/" & rp.StartY, "dd/mm/yyyy")
                    If rp.StartH <> 0 And rp.StartM <> 0 Then !StartTime = rp.StartH & ":" & rp.StartM
                    If rp.EndD <> 0 And rp.EndMO <> 0 And rp.EndY <> 0 Then !EndDate = Format(rp.EndD & "/" & rp.EndMO & "/" & rp.EndY, "dd/mm/yyyy")
                    If rp.EndH <> 0 And rp.EndM <> 0 Then !EndTime = rp.EndH & ":" & rp.EndM
                    !Busno = Trim(TrimChr(rp.Busno))
                    !Driver = rp.Driver
                    !Conductor = rp.Conductor
                    !Cleaner = rp.Cleaner
                    !NoOfMisBill = rp.FareType
                    
                    !ladies_count = IIf(IsNumeric(rp.ladies_count), rp.ladies_count, 0)
                    !ladies_coll = IIf(IsNumeric(rp.ladies_coll), rp.ladies_coll, 0)
                    !senior_count = IIf(IsNumeric(rp.seniar_count), rp.seniar_count, 0)
                    !senior_coll = IIf(IsNumeric(rp.seniar_coll), rp.seniar_coll, 0)
                    .Update
                End With
            End If
        Loop
    Close #Filehdl
    If f1 Then RES.Close
    gbladoCon.Execute ("delete * from TMPFARERPT")
    gbladoCon.Execute ("ALTER TABLE TMPFARERPT ALTER COLUMN SLNO COUNTER (1, 1)")
    freefileInt = FreeFile()
    Open App.Path & "\FAREWISE.DAT" For Binary Access Read As #freefileInt
        Do While Not EOF(freefileInt)
            Get #freefileInt, , fare_obj
            If EOF(freefileInt) Then Exit Do
            gbladoCon.Execute ("insert into TMPFARERPT (TKTNO,FARE) values(" & Trim(fare_obj.TktNo) & "," & Trim(fare_obj.FARE) & ")")
        Loop
    Close #freefileInt
    Exit Function
err:
    MsgBox "Error!" & vbCrLf & "Err No : " & err.Number & vbCrLf & err.Description, vbCritical
End Function
Public Function DBSet(fname As String)
Dim Filehdl As Integer
Dim sp As SETUP
Dim hsp As HARDWARE_SETUP
Dim sql As String

    Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    
    Set RES = DB.OpenRecordset("Settings", dbOpenDynaset)
    If RES.RecordCount > 0 Then
        Filehdl = FreeFile()
        Open App.Path & "\" & fname For Binary Access Read As #Filehdl
        
        Do While Not EOF(Filehdl)
            
            Get #Filehdl, , sp
            Get #Filehdl, , hsp
        If EOF(Filehdl) Then Exit Do
            
        If RES.RecordCount > 0 Then RES.MoveLast
            
            With RES
                .Edit
                !UserPWD = hsp.USR_PSWD
                !MasterPWD = hsp.MSR_PSWD
                !MainDisplay = Trim(sp.MainDisp)
                !MainDisplay2 = sp.MainDisp2
                !HEADER1 = sp.bhl1
                !HEADER2 = sp.bhl2
                !Header3 = sp.bhl3
                !Footer1 = sp.bfl1
                !Footer2 = sp.bfl2
                !PalmtecID = sp.PalmtecID
                !HalfPer = sp.HalfPer
                !ConPer = sp.ConPer
                !StFareEdit = sp.ucbSTFareEdit
                !STMaxAmt = sp.STMaxAmt
                !STMinCon = sp.STMinCon
                !PhyPer = sp.PhyPer
                !Roundoff = sp.Roundoff
                !RoundUp = sp.RoundUp
                !RoundAmt = sp.RoundAmt
                !LuggageUnitRate = sp.LuggageUnitRate
                !Currency = TrimChr(sp.Currency)
                !ReportFlag = sp.ReportFlag
                Call CreateFields
                !NextFareFlag = sp.NextFareRound
                !RemoveTicketFlag = sp.EnableRemoveTicket
                !StageFontFlag = sp.EnableStageFont
                !LANGUAGEOPTION = hsp.select_language
                !STAGEUPDATIONMSG = sp.UpdateStageMsg
                !DEFAULTSTAGE = sp.EnableStageDefault
                !OdometerEntry = sp.OdometerEntry
                !TicketNoBigFont = sp.TicketNoBigFont
                !CrewCheck = sp.CrewCheck
                !PhNo = sp.PhNo
                !TripSMS = sp.TripSMS
                !shSMS = sp.ScheduleSMS 'SANGEETHA
                !sendbillEnable = sp.sendbillEnable
        !TripsendEnable = sp.TripsendEnable
        !SchedulesendEnable = sp.SchedulesendEnable
        !Sendpend = sp.Sendpend
        !PhNo2 = sp.PhNo2
        !AccessPoint = sp.AccessPoint
        !DestAdds = sp.DestAdds
        !Username = sp.Username
        !PassWord = sp.PassWord
        !Uploadpath = sp.Uploadpath
        !Downloadpath = sp.Downloadpath
        !HttpUrl = sp.HttpUrl
        !GprsEnable = sp.GprsEnable
        !SmartCard = sp.SmartCard
        !ExpEnable = sp.ExpEnable
        !GprsEnableMessage = sp.MsgPrompt
        !FtpEnable = sp.FtpEnable
        !InspectRpt = sp.Inspectorreport
        !StRoundoffEnable = sp.StRoundoff_E_D
        !StRoundoffAmt = sp.StRoundoff_Amt
        !ReportFONT = sp.ReportFONT
        !MultiplePass = sp.MultiplePass
        !InspectorSMS = sp.InspectorSMS
        !Simplereport = sp.Simplereport
        !PhNo3 = sp.PhNo3
        !Autoshutdown = sp.AutoShutdownEnable
        !UserpswdEnable = sp.UserPasswordEnable
        !refund = sp.refund
        !ladies_ratio = IIf(IsNumeric(sp.ladis_per), sp.ladis_per, 0)
        !senior_ratio = IIf(IsNumeric(sp.seniar_per), sp.seniar_per, 0)
                .Update
             End With
            Loop
        Close #Filehdl
        
    Else
   
        
        Set RES = DB.OpenRecordset("Settings", dbOpenDynaset)
        If RES.RecordCount > 0 Then RES.MoveLast
        Filehdl = FreeFile()
        Open App.Path & "\" & fname For Binary Access Read As #Filehdl
        
        Do While Not EOF(Filehdl)
            
            Get #Filehdl, , sp
            Get #Filehdl, , hsp
            If EOF(Filehdl) Then Exit Do
            
            If RES.RecordCount > 0 Then RES.MoveLast
            
            With RES
                .AddNew
                !UserPWD = hsp.USR_PSWD
                !MasterPWD = hsp.MSR_PSWD
                !MainDisplay = Trim(sp.MainDisp)
                !MainDisplay2 = Trim(sp.MainDisp2)
                !HEADER1 = sp.bhl1
                !HEADER2 = sp.bhl2
                !Header3 = sp.bhl3
                !Footer1 = sp.bfl1
                !Footer2 = sp.bfl2
                !PalmtecID = sp.PalmtecID
                !HalfPer = sp.HalfPer
                !ConPer = sp.ConPer
                !STMaxAmt = sp.STMaxAmt
                !STMinCon = sp.STMinCon
                !PhyPer = sp.PhyPer
                !Roundoff = sp.Roundoff
                !RoundUp = sp.RoundUp
                !RoundAmt = sp.RoundAmt
                !LuggageUnitRate = sp.LuggageUnitRate
                !Currency = sp.Currency
                !ReportFlag = sp.ReportFlag
                If CreateFields() = True Then
                 !RemoveTicketFlag = sp.EnableRemoveTicket
                !StageFontFlag = sp.EnableStageFont
                End If
                !NextFareFlag = sp.NextFareRound
                !RemoveTicketFlag = sp.EnableRemoveTicket
                !StageFontFlag = sp.EnableStageFont
                !LANGUAGEOPTION = hsp.select_language
                !STAGEUPDATIONMSG = sp.UpdateStageMsg
                !DEFAULTSTAGE = sp.EnableStageDefault
                !OdometerEntry = sp.OdometerEntry
                !TicketNoBigFont = sp.TicketNoBigFont
                !CrewCheck = sp.CrewCheck
                !PhNo = sp.PhNo
                !TripSMS = sp.TripSMS
                !shSMS = sp.ScheduleSMS 'SANGEETHA
                  !sendbillEnable = sp.sendbillEnable
        !TripsendEnable = sp.TripsendEnable
        !SchedulesendEnable = sp.SchedulesendEnable
        !Sendpend = sp.Sendpend
        !PhNo2 = sp.PhNo2
        !AccessPoint = sp.AccessPoint
        !DestAdds = sp.DestAdds
        !Username = sp.Username
        !PassWord = sp.PassWord
        !Uploadpath = sp.Uploadpath
        !Downloadpath = sp.Downloadpath
        !HttpUrl = sp.HttpUrl
        !GprsEnable = sp.GprsEnable
        !SmartCard = sp.SmartCard
        !ExpEnable = sp.ExpEnable
        !GprsEnableMessage = sp.MsgPrompt
        !FtpEnable = sp.FtpEnable
        !InspectRpt = sp.Inspectorreport
        !StFareEdit = sp.ucbSTFareEdit
        !StRoundoffEnable = sp.StRoundoff_E_D
        !StRoundoffAmt = sp.StRoundoff_Amt
        
        !ReportFONT = sp.ReportFONT
        !MultiplePass = sp.MultiplePass
        !InspectorSMS = sp.InspectorSMS
        !Simplereport = sp.Simplereport
        !PhNo3 = sp.PhNo3
        !Autoshutdown = sp.AutoShutdownEnable
        !UserpswdEnable = sp.UserPasswordEnable
        !refund = sp.refund
        !ladies_ratio = IIf(IsNumeric(sp.ladis_per), sp.ladis_per, 0)
        !senior_ratio = IIf(IsNumeric(sp.seniar_per), sp.seniar_per, 0)
                        
                .Update
             End With
        Loop
        Close #Filehdl
        
    End If
    RES.Close
End Function
Public Function DBTKTS(fname As String)
Dim Filehdl As Integer
Dim tk As PTicket
Dim rp As PReport
Dim FHndl As Integer
Dim fShndl, TNO As Integer
Dim tktK As PTicket
Dim fBuff As String
Dim FARE As fare_type
Dim FnameUp, TMRId As String
Dim PFname, DA As String
Dim mystr As String
Dim pHandle, sHandle As Integer
Dim pass, PalmID As String
Dim gPass As PASSCONC
Dim Fileint As Long
Dim myres As New ADODB.Recordset
Dim S As Integer
Dim mybool As Boolean
mybool = False
Dim pos As Integer
'''gPassCount added by syam
Dim gPassCount As Long
    FnameUp = fname
    PFname = App.Path & "\PASS.PAS"
    sHandle = FreeFile()
    Open App.Path & "\BUS.DAT" For Binary Access Read As #sHandle
        Get #sHandle, , HStr
        Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
        Set RES = DB.OpenRecordset("TKTS", dbOpenDynaset)
        Filehdl = FreeFile()
        Open fname For Binary Access Read As #Filehdl
            Do While Not EOF(Filehdl)
                Get #Filehdl, , tk
                If EOF(Filehdl) Then Exit Do
                pHandle = FreeFile()
                Open PFname For Binary Access Read As #pHandle
                    Do While Not EOF(pHandle)
                        Get #pHandle, , gPass
                        If tk.TicketNo = gPass.TicketNo Then
                            pass = gPass.PassNo
                            Exit Do
                        Else
                            pass = Chr(0)
                        End If
                    Loop
                Close #pHandle
                PalmID = Replace(HStr.PalmtecID, Chr(0), "")
                DA = tk.Dy & "/" & tk.Mn & "/" & Format(Now, "YYYY")
                sql = "SELECT * FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE=" & val(Mid(fname, 5, 2)) & " AND TRIPNO = " & tk.cTripNo & " AND DATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "') "
                Set RES4 = DB.OpenRecordset(sql, dbOpenDynaset)
                If RES4.RecordCount > 0 Then
                    TMRId = RES4!Trip_Master_ID
                End If
                If (getvalueQuery("select Count(*) from TKTS where PALMID='" & PalmID & "' AND SCHDULE=" & val(Mid(fname, 5, 2)) & " AND TicketNo=" & tk.TicketNo & " AND DATE='" & Format(DA, "DD/MM/YYYY") & "'")) = 0 Then
                    With RES
                        If .RecordCount > 0 Then .MoveLast
                        .AddNew
                        !TripMasterId = TMRId
                        !Date = Format(DA, "DD/MM/YYYY")
                        !Time = Format(Time, "HH:MM:SS") 'Format(Time, "hh:mm:ssAM/PM")
                        !PalmID = HStr.PalmtecID
                        !SCHDULE = val(Mid(fname, 5, 2))
                        !TripNo = tk.cTripNo
                        !TicketNo = tk.TicketNo
                        !Amount = tk.Amount
                        !LuggAmont = tk.Lugg
                        !Luggage = tk.Luggage
                        !FromStage = tk.From
                        !ToStage = tk.To
                        !Full = tk.Full
                        !Half = tk.Half
                        !st = tk.st
                        !Phy = tk.Phy
                        !HourMint = tk.Hr & ":" & tk.Minut
                        !ctype = tk.Typ
                        !RefundAmt = tk.RefundAmt
                        !Refundsts = tk.Refundsts
                        !ladies_count = IIf(IsNumeric(tk.ladies_count), (tk.ladies_count), 0)
                        !senior_count = IIf(IsNumeric(tk.seniar_count), (tk.seniar_count), 0)
                        If Trim(pass) <> "" Then !PassNo = pass
                        .Update
                    End With
                    If val(getvalueQuery("select Count(*) from FARERPT where PID='" & TrimChr(PalmID) & "' and TKTNO =" & TrimChr(tk.TicketNo) & " AND SCHNO= " & val(Mid(fname, 5, 2)) & " AND TRPNO= " & tk.cTripNo & " AND SCH_STDATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "') ")) = 0 Then
                        mystr = "select TKTNO,FARE from TMPFARERPT where TKTNO=" & TrimChr(tk.TicketNo) & "  order by SLNO"
                        Set myres = New ADODB.Recordset
                        Set myres = gbladoCon.Execute(mystr)
                        Do While Not myres.EOF
                            If myres("TKTNO") = tk.TicketNo Then
                                If tk.Full > 0 Then
                                    For S = 1 To tk.Full
                                        sql = "insert into FARERPT (TKTNO,SCHNO,TRPNO,SCH_STDATE,SCH_ENDDATE,FARE,PID) values(" & myres("TKTNO") & "," & val(Mid(fname, 5, 2)) & " ," & tk.cTripNo & ",DateValue('" & Format(DA, "DD/MM/YYYY") & "') ,DateValue('" & Format(DA, "DD/MM/YYYY") & "') ," & TrimChr(myres("FARE")) / tk.Full & ",'" & TrimChr(HStr.PalmtecID) & "' )"
                                        gbladoCon.Execute sql
                                    Next S
                                    myres.MoveNext
                                End If
                                If tk.Half > 0 Then
                                    For S = 1 To tk.Half
                                        sql = "insert into FARERPT (TKTNO,SCHNO,TRPNO,SCH_STDATE,SCH_ENDDATE,FARE,PID) values(" & myres("TKTNO") & ", " & val(Mid(fname, 5, 2)) & " ," & tk.cTripNo & ",DateValue('" & Format(DA, "DD/MM/YYYY") & "') ,DateValue('" & Format(DA, "DD/MM/YYYY") & "') ," & TrimChr(myres("FARE")) / tk.Half & ",'" & TrimChr(HStr.PalmtecID) & "' )"
                                        gbladoCon.Execute sql
                                    Next S
                                    myres.MoveNext
                                End If
                                If tk.st > 0 Then
                                    For S = 1 To tk.st
                                        sql = "insert into FARERPT (TKTNO,SCHNO,TRPNO,SCH_STDATE,SCH_ENDDATE,FARE,PID) values(" & myres("TKTNO") & "," & val(Mid(fname, 5, 2)) & " ," & tk.cTripNo & ",DateValue('" & Format(DA, "DD/MM/YYYY") & "') ,DateValue('" & Format(DA, "DD/MM/YYYY") & "') ," & TrimChr(myres("FARE")) / tk.st & ",'" & TrimChr(HStr.PalmtecID) & "' )"
                                        gbladoCon.Execute sql
                                    Next S
                                    myres.MoveNext
                                End If
                                If tk.Phy > 0 Then
                                    For S = 1 To tk.Phy
                                        sql = "insert into FARERPT (TKTNO,SCHNO,TRPNO,SCH_STDATE,SCH_ENDDATE,FARE,PID) values(" & myres("TKTNO") & "," & val(Mid(fname, 5, 2)) & " ," & tk.cTripNo & ",DateValue('" & Format(DA, "DD/MM/YYYY") & "') ,DateValue('" & Format(DA, "DD/MM/YYYY") & "') ," & TrimChr(myres("FARE")) / tk.Phy & ",'" & TrimChr(HStr.PalmtecID) & "' )"
                                        gbladoCon.Execute sql
                                    Next S
                                    myres.MoveNext
                                End If
                                If tk.Lugg > 0 Then
                                    For S = 1 To tk.Lugg
                                        sql = "insert into FARERPT (TKTNO,SCHNO,TRPNO,SCH_STDATE,SCH_ENDDATE,FARE,PID) values(" & myres("TKTNO") & "," & val(Mid(fname, 5, 2)) & " ," & tk.cTripNo & ",DateValue('" & Format(DA, "DD/MM/YYYY") & "') ,DateValue('" & Format(DA, "DD/MM/YYYY") & "') ," & TrimChr(myres("FARE")) / tk.Lugg & ",'" & TrimChr(HStr.PalmtecID) & "' )"
                                        gbladoCon.Execute sql
                                    Next S
                                    myres.MoveNext
                                End If
                                'ladies and ss count
                                If tk.ladies_count > 0 Then
                                    For S = 1 To tk.ladies_count
                                        sql = "insert into FARERPT (TKTNO,SCHNO,TRPNO,SCH_STDATE,SCH_ENDDATE,FARE,PID) values(" & myres("TKTNO") & "," & val(Mid(fname, 5, 2)) & " ," & tk.cTripNo & ",DateValue('" & Format(DA, "DD/MM/YYYY") & "') ,DateValue('" & Format(DA, "DD/MM/YYYY") & "') ," & TrimChr(myres("FARE")) / tk.ladies_count & ",'" & TrimChr(HStr.PalmtecID) & "' )"
                                        gbladoCon.Execute sql
                                    Next S
                                    myres.MoveNext
                                End If
                                If tk.seniar_count > 0 Then
                                    For S = 1 To tk.seniar_count
                                        sql = "insert into FARERPT (TKTNO,SCHNO,TRPNO,SCH_STDATE,SCH_ENDDATE,FARE,PID) values(" & myres("TKTNO") & "," & val(Mid(fname, 5, 2)) & " ," & tk.cTripNo & ",DateValue('" & Format(DA, "DD/MM/YYYY") & "') ,DateValue('" & Format(DA, "DD/MM/YYYY") & "') ," & TrimChr(myres("FARE")) / tk.seniar_count & ",'" & TrimChr(HStr.PalmtecID) & "' )"
                                        gbladoCon.Execute sql
                                    Next S
                                    myres.MoveNext
                                End If
                            End If
                        Loop
                    End If
                Else
                    With RES
                        If .RecordCount > 0 Then .MoveLast
                        .Edit
                        !TripMasterId = TMRId
                        !Date = Format(DA, "DD/MM/YYYY")
                        !Time = Format(Time, "HH:MM:SS") 'Format(Time, "hh:mm:ssAM/PM")
                        !PalmID = HStr.PalmtecID
                        !SCHDULE = val(Mid(fname, 5, 2))
                        !TripNo = tk.cTripNo
                        !TicketNo = tk.TicketNo
                        !Amount = tk.Amount
                        !LuggAmont = tk.Lugg
                        !Luggage = tk.Luggage
                        !FromStage = tk.From
                        !ToStage = tk.To
                        !Full = tk.Full
                        !Half = tk.Half
                        !st = tk.st
                        !Phy = tk.Phy
                        !HourMint = tk.Hr & ":" & tk.Minut
                        !ctype = tk.Typ
                        !PassNo = pass
                        !RefundAmt = tk.RefundAmt
                        !Refundsts = tk.Refundsts
                        
                        !ladies_count = IIf(IsNumeric(tk.ladies_count), (tk.ladies_count), 0)
                        !senior_count = IIf(IsNumeric(tk.seniar_count), (tk.seniar_count), 0)
                        
                        .Update
                    End With
                End If
            Loop
        Close #sHandle
    Close #Filehdl
    RES.Close
End Function
Public Function DBODMTR(fname As String)      ''''''''''''''''RNC 16082011
Dim Fhdl, sHandle, pHandle As Integer
Dim odmtr As ODOMETER
Dim FnameUp As String
Dim TMRId As Integer
Dim PFname As String
Dim DA, sql As String
Dim PalmID As String

   pHandle = FreeFile()
   Open App.Path & "\BUS.DAT" For Binary Access Read As #pHandle
   Get #pHandle, , HStr
   Close #pHandle
        
    FnameUp = fname
    PFname = App.Path & "\ODOMETER.DAT"
    sHandle = FreeFile()
    Open App.Path & "\ODOMETER.DAT" For Binary Access Read As #sHandle
        
    PalmID = Mid$(HStr.PalmtecID, 1, InStr(1, HStr.PalmtecID, Chr(0)) - 1)

    'DA = odmtr.SDay & "/" & odmtr.SMonth & "/" & odmtr.SYear
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
'    Set RES = DB.OpenRecordset("ODOMETER", dbOpenDynaset)
      

    Do While Not EOF(sHandle)
        Get #sHandle, , odmtr
        If odmtr.ucScheduleNo = 0 Then Exit Do
        DA = odmtr.SDay & "/" & odmtr.SMonth & "/" & odmtr.SYear
       sql = "DELETE FROM ODOMETER WHERE PalmID= '" & Trim(PalmID) & "' AND ScheduleNo = " & odmtr.ucScheduleNo & " AND TRIPNO = " & odmtr.ucTripNo & " AND SDATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "') "
        DB.Execute (sql)
        'If odmtr.ucScheduleNo = 0 Then Exit Do
        
       sql = "Select * FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & odmtr.ucScheduleNo & " AND TRIPNO= " & odmtr.ucTripNo & " AND DATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "') "
        'AND TRIPNO= " & odmtr.ucTripNo & "
        Set RES4 = DB.OpenRecordset(sql, dbOpenDynaset)
        
        'Set rs = CON.Execute(sql)
        If RES4.RecordCount > 0 Then
            TMRId = RES4!Trip_Master_ID
        End If
        
        Set RES = DB.OpenRecordset("ODOMETER", dbOpenDynaset)
        With RES
                  If .RecordCount > 0 Then .MoveLast
                     .AddNew
                        !TirpMasterReferenceId = TMRId
                        !SDate = Format(odmtr.SDay, "00") & "/" & Format(odmtr.SMonth, "00") & "/" & Format(odmtr.SYear, "0000")
                        !EDate = Format(odmtr.EDay, "00") & "/" & Format(odmtr.EMonth, "00") & "/" & Format(odmtr.EYear, "0000")
                        !SOdometer = odmtr.Startr
                        !EOdometer = odmtr.Endr
                        !PalmID = HStr.PalmtecID
                        !TripNo = odmtr.ucTripNo
                        !ScheduleNo = odmtr.ucScheduleNo
                      .Update
         End With
      Loop
     Close #sHandle
    RES.Close
    RES4.Close
               
End Function

Public Function DBINSPR(fname As String)
Dim Fhdl, Shdl, pHandle As Integer
Dim insptr As INSPECTORDET
Dim FnameUp As String
Dim PFname As String
Dim TMRId As Integer
Dim PalmID As String
Dim DA As String

    pHandle = FreeFile()
    Open App.Path & "\BUS.DAT" For Binary Access Read As #pHandle
    Get #pHandle, , HStr
    Close #pHandle
     
    FnameUp = fname
    PFname = App.Path & "\INSPECTOR.DAT"
    Shdl = FreeFile()
    Open App.Path & "\INSPECTOR.DAT" For Binary Access Read As #Shdl
    
    PalmID = Mid$(HStr.PalmtecID, 1, InStr(1, HStr.PalmtecID, Chr(0)) - 1)
      
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
    Set RES = DB.OpenRecordset("INSPECTORDET", dbOpenDynaset)
'    DA = insptr.InsDay & "/" & insptr.InsMonth & "/" & insptr.InsYear
'    Fhdl = FreeFile()
'    Open fname For Binary Access Read As #Fhdl
'    sql = "DELETE FROM INSPECTORDET WHERE PalmID='" & PalmId & "' And SCHEDULE = " & insptr.ucScheduleNo & " And TripNo = " & insptr.ucTripNo & " AND DATE='" & Format(DA, "DD/MM/YYYY") & "'"
    'WHERE PALMID='" & PalmId & "' AND SCHDULE=" & val(Mid(fname, 5, 2)) & " AND TicketNo=" & tk.TicketNo & " AND DATE='" & Format(DA, "DD/MM/YYYY") & "'"
'    DB.Execute (sql)
    
    Do While Not EOF(Shdl)
        Get #Shdl, , insptr
        DA = insptr.InsDay & "/" & insptr.InsMonth & "/" & insptr.InsYear
        sql = "DELETE FROM INSPECTORDET WHERE PalmID='" & PalmID & "' And SCHEDULENO = " & insptr.ucScheduleNo & " And TripNo = " & insptr.ucTripNo & " AND DATE='" & Format(DA, "DD/MM/YYYY") & "'"
        DB.Execute (sql)
        'If EOF(Fhdl) Then Exit Do
        If insptr.ucScheduleNo = 0 Then Exit Do
        

        sql = "Select * FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & insptr.ucScheduleNo & " AND TRIPNO= " & insptr.ucTripNo & " AND DATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "') "
        Set RES4 = DB.OpenRecordset(sql, dbOpenDynaset)
        If RES4.RecordCount > 0 Then
            TMRId = RES4!Trip_Master_ID
        End If
        
       ' Set RES = DB.OpenRecordset("INSPECTORDET", dbOpenDynaset)
                With RES
                    If .RecordCount > 0 Then .MoveLast
                    
                    .AddNew
                    !TripMasterReferenceId = TMRId
                    !InspectorID = insptr.ucInspectorId
                    !StationNo = insptr.Stage
                    !Date = Format(insptr.InsDay, "00") & "/" & Format(insptr.InsMonth, "00") & "/" & Format(insptr.InsYear, "0000")
                    !Time = Format(insptr.InsHour, "00") & ":" & Format(insptr.InsMinutes, "00")
                    !PalmID = HStr.PalmtecID
                    !ScheduleNo = insptr.ucScheduleNo
                    !TripNo = insptr.ucTripNo
                    .Update
                End With
    Loop
    Close #Shdl
'    Close #Fhdl
    
    'RES.Close
    'res4.Close
End Function
Public Function DBEXPENSE1(fname As String)
Dim Fhdl, Shdl, pHandle As Integer
Dim expobj As EXPENSES
Dim TMRId As Integer
'Dim exp As EXPENSEDET
Dim FnameUp, exdesc As String
Dim PFname, PalmID, DA As String
Dim Sqlselect, sqlpass, sqlExp As String

    pHandle = FreeFile()
    Open App.Path & "\BUS.DAT" For Binary Access Read As #pHandle
        Get #pHandle, , HStr
        
    Fhdl = FreeFile()
    Open App.Path & "\EXPENSEDET.DAT" For Binary Access Read As #Fhdl
        Get #Fhdl, , exp
        
    FnameUp = fname
    PFname = App.Path & "\EXPENSE.DAT"
    Shdl = FreeFile()
    Open App.Path & "\EXPENSE.DAT" For Binary Access Read As #Shdl
        
      
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
   Set RES = DB.OpenRecordset("EXPENSE", dbOpenDynaset)
    Dim f1 As Boolean
        f1 = False
'    sql = "insert into EXPENSE (ERecieptNo,ExpCode ,ExpAmt, ExpName, Date, Time, PalmID) values('" _
'    & insptr.InspID & "','" _
'    & val & "','" _
'    & val & "','" _
'    & val & "','" _
'    & "')"
'    DB.Execute sql
    'Fhdl = FreeFile()
   ' Open fname For Binary Access Read As #Fhdl
   
   PalmID = Replace(HStr.PalmtecID, Chr(0), "")
'   DA = expobj.Day & "/" & expobj.Month & "/" & expobj.Year
'   sql = "DELETE FROM EXPENSE  WHERE PALMID = '" & PalmId & "' AND SCHEDULENO = " & val(Mid(fname, 5, 2)) & " AND EXPCODE = '" & expobj.ucType & "' AND DATE = '" & Format(DA, "DD/MM/YYYY") & "'"
'   DB.Execute (sql)
        sql = "DELETE FROM EXPENSE  WHERE PALMID = '" & PalmID & "' AND SCHEDULENO  = " & expobj.ucScheduleNo & " AND EXPCODE = '" & expobj.ucType & "' AND DATE = '" & Format(DA, "DD/MM/YYYY") & "' and   BusNo='" & TrimChr(expobj.Busno) & "' and Time='" & expobj.Hour & ":" & expobj.Minutes & "'"
        DB.Execute (sql)
    Do While Not EOF(Shdl)
        Get #Shdl, , expobj
        If EOF(Shdl) Then Exit Do
        DA = expobj.Day & "/" & expobj.Month & "/" & expobj.Year
        'sql = "DELETE FROM EXPENSE  WHERE PALMID = '" & PalmId & "' AND SCHEDULENO  = " & expobj.ucScheduleNo & " AND EXPCODE = '" & expobj.ucType & "' AND DATE = '" & Format(DA, "DD/MM/YYYY") & "' and   BusNo='" & TrimChr(expobj.BusNo) & "'"
        'sql = "DELETE FROM EXPENSE  WHERE PALMID = '" & PalmID & "' AND SCHEDULENO  = " & expobj.ucScheduleNo & " AND EXPCODE = '" & expobj.ucType & "' AND DATE = '" & Format(DA, "DD/MM/YYYY") & "' and   BusNo='" & TrimChr(expobj.BusNo) & "' and Time='" & expobj.Hour & ":" & expobj.Minutes & "'"
       
        
        
        
        sql = "SELECT * FROM EXPENSE ORDER BY DATE,TIME"
        Set RES = DB.OpenRecordset(sql, dbOpenDynaset)
        
        
        
        sql = "SELECT EXP_NAME FROM EXPMASTER WHERE EXP_CODE=" & expobj.ucType & " "
        Set RES4 = DB.OpenRecordset(sql, dbOpenDynaset)
        If RES4.RecordCount > 0 Then
            exdesc = RES4!exp_name
        End If
        
        sql = "Select * FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & expobj.ucScheduleNo & " AND TRIPNO= " & expobj.ucTripNo & " AND DATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "') "
        'AND TRIPNO= " & odmtr.ucTripNo & "
        Set RES4 = DB.OpenRecordset(sql, dbOpenDynaset)
        
        'Set rs = CON.Execute(sql)
        If RES4.RecordCount > 0 Then
            TMRId = RES4!Trip_Master_ID
        End If
                  
                With RES
                    If .RecordCount > 0 Then .MoveLast
                    
                    .AddNew
                    !expcode = expobj.ucType
                    If Trim(expobj.ucType) <> "1" Then
                        !expname = exdesc
                    Else
                        !expname = "Diesel Entry"
                    End If
                    !TripMasterReferenceId = TMRId
                    !ExpAmt = expobj.fExpens
                    !Date = Format(expobj.Day, "00") & "/" & Format(expobj.Month, "00") & "/" & Format(expobj.Year, "0000")
                    !Time = expobj.Hour & ":" & expobj.Minutes
                    !PalmID = TrimChr(HStr.PalmtecID)
                    !ScheduleNo = expobj.ucScheduleNo
                    !Busno = expobj.Busno
                    !DriverName = expobj.EName
                    !rcpt_No = expobj.RptNo
                    .Update
                
                End With
                
                sqlExp = ""
            sqlExp = "Select sum(ExpAmt)as EXP from EXPENSE where TripMasterReferenceId='" & RES4!Trip_Master_ID & "'"
            Set RES8 = CNN.OpenRecordset(sqlExp, dbOpenDynaset)
            'RES4.Edit
           sql = "UPDATE RPT SET Expense=" & RES8("EXP") & " where Trip_Master_ID=" & RES4!Trip_Master_ID
           'Set RES7 = CNN.OpenRecordset(sql, dbOpenDynaset)
           CNN.Execute (sql)
                
               
    Loop
    Close #Shdl
    Close #Fhdl
    Close #pHandle
   ' if res4.         res4.Close
    If f1 Then RES.Close
    'RES.Close
    

End Function

Public Function DBEXPENSE(fname As String)
Dim Fhdl, Shdl, pHandle As Integer
Dim expobj As EXPENSES
Dim TMRId As String
'Dim exp As EXPENSEDET
Dim FnameUp, exdesc, expname As String
Dim PFname, PalmID, DA As String
Dim Sqlselect, sqlpass, sqlExp As String
  Call ConnectDatabase(gbladoCon, App.Path & "\PVT.MDB", "silbus")
    pHandle = FreeFile()
    Open App.Path & "\BUS.DAT" For Binary Access Read As #pHandle
        Get #pHandle, , HStr
        
    Fhdl = FreeFile()
    Open App.Path & "\EXPENSEDET.DAT" For Binary Access Read As #Fhdl
        Get #Fhdl, , exp
        
    FnameUp = fname
    PFname = App.Path & "\EXPENSE.DAT"
    Shdl = FreeFile()
    Open App.Path & "\EXPENSE.DAT" For Binary Access Read As #Shdl
        
      
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.mdb", dbDriverComplete, False, ";UID=;PWD=silbus")
   Set RES = DB.OpenRecordset("EXPENSE", dbOpenDynaset)
    Dim f1 As Boolean
        f1 = False
'    sql = "insert into EXPENSE (ERecieptNo,ExpCode ,ExpAmt, ExpName, Date, Time, PalmID) values('" _
'    & insptr.InspID & "','" _
'    & val & "','" _
'    & val & "','" _
'    & val & "','" _
'    & "')"
'    DB.Execute sql
    'Fhdl = FreeFile()
   ' Open fname For Binary Access Read As #Fhdl
   
   PalmID = Replace(HStr.PalmtecID, Chr(0), "")
'   DA = expobj.Day & "/" & expobj.Month & "/" & expobj.Year
'   sql = "DELETE FROM EXPENSE  WHERE PALMID = '" & PalmId & "' AND SCHEDULENO = " & val(Mid(fname, 5, 2)) & " AND EXPCODE = '" & expobj.ucType & "' AND DATE = '" & Format(DA, "DD/MM/YYYY") & "'"
'   DB.Execute (sql)
        'sql = "DELETE FROM EXPENSE  WHERE PALMID = '" & PalmID & "' AND SCHEDULENO  = " & expobj.ucScheduleNo & " AND EXPCODE = '" & expobj.ucType & "' AND DATE = '" & Format(DA, "DD/MM/YYYY") & "' and   BusNo='" & TrimChr(expobj.BusNo) & "' and Time='" & expobj.Hour & ":" & expobj.Minutes & "'"
        'DB.Execute (sql)
    Do While Not EOF(Shdl)
        Get #Shdl, , expobj
        If EOF(Shdl) Then Exit Do
        DA = expobj.Day & "/" & expobj.Month & "/" & expobj.Year
        'sql = "DELETE FROM EXPENSE  WHERE PALMID = '" & PalmId & "' AND SCHEDULENO  = " & expobj.ucScheduleNo & " AND EXPCODE = '" & expobj.ucType & "' AND DATE = '" & Format(DA, "DD/MM/YYYY") & "' and   BusNo='" & TrimChr(expobj.BusNo) & "'"
        'sql = "DELETE FROM EXPENSE  WHERE PALMID = '" & PalmID & "' AND SCHEDULENO  = " & expobj.ucScheduleNo & " AND EXPCODE = '" & expobj.ucType & "' AND DATE = '" & Format(DA, "DD/MM/YYYY") & "' and   BusNo='" & TrimChr(expobj.BusNo) & "' and Time='" & expobj.Hour & ":" & expobj.Minutes & "'"
       
        
        
        
'        sql = "SELECT * FROM EXPENSE ORDER BY DATE,TIME"
'        Set RES = DB.OpenRecordset(sql, dbOpenDynaset)
'
'
'
        sql = "SELECT EXP_NAME FROM EXPMASTER WHERE EXP_CODE=" & expobj.ucType & " "
        Set RES4 = DB.OpenRecordset(sql, dbOpenDynaset)
        If RES4.RecordCount > 0 Then
            exdesc = RES4!exp_name
        End If
'
'        sql = "Select Trip_Master_ID FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & expobj.ucScheduleNo & " AND TRIPNO= " & expobj.ucTripNo & " AND DATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "') "
'        'AND TRIPNO= " & odmtr.ucTripNo & "
'        Set RES4 = DB.OpenRecordset(sql, dbOpenDynaset)
'
'        'Set rs = CON.Execute(sql)
'        If RES4.RecordCount > 0 Then
'            TMRId = RES4!Trip_Master_ID
'        End If
'
'                With RES
'                    If .RecordCount > 0 Then .MoveLast
'
'                    .AddNew
'                    !expcode = expobj.ucType
'                    If Trim(expobj.ucType) <> "1" Then
'                        !expname = exdesc
'                    Else
'                        !expname = "Diesel Entry"
'                    End If
'                    !TripMasterReferenceId = TMRId
'                    !ExpAmt = expobj.fExpens
'                    !Date = Format(expobj.Day, "00") & "/" & Format(expobj.Month, "00") & "/" & Format(expobj.Year, "0000")
'                    !Time = expobj.Hour & ":" & expobj.Minutes
'                    !PalmID = TrimChr(HStr.PalmtecID)
'                    !ScheduleNo = expobj.ucScheduleNo
'                    !BusNo = expobj.BusNo
'                    !DriverName = expobj.EName
'                    !rcpt_No = expobj.RptNo
'                    .Update
'
'                End With
'




    If (getvalueQuery("select Count(*) from EXPENSE where PALMID = '" & PalmID & "' AND rcpt_No= " & expobj.RptNo & " AND SCHEDULENO = " & expobj.ucScheduleNo & " AND EXPCODE = '" & expobj.ucType & "' AND DATE = '" & Format(DA, "DD/MM/YYYY") & "'")) = 0 Then
        
        
        If Trim(expobj.ucType) <> "1" Then
            expname = exdesc
        Else
            expname = "Diesel Entry"
        End If
        TMRId = ""
        sqlExp = "Select Trip_Master_ID  as TRM FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & expobj.ucScheduleNo & " AND TRIPNO= " & expobj.ucTripNo & " AND DateValue('" & Format(DA, "DD/MM/YYYY") & "') between StartDate and EndDate"
       ' sqlExp = "Select Trip_Master_ID  as TRM FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & expobj.ucScheduleNo & " AND TRIPNO= " & expobj.ucTripNo & " AND DATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "')"
        Set RES8 = DB.OpenRecordset(sqlExp, dbOpenDynaset)
        If (RES8("TRM")) > 0 Then
            TMRId = RES8("TRM")
        Else
            TMRId = ""
        End If
       
        'sleep (200)
        'sql = "insert into EXPENSE (TripMasterReferenceId,ExpCode ,ExpAmt, ExpName, Date, Time, PalmID,ScheduleNo,BusNo,DriverName,rcpt_No) values('"
        sql = "insert into EXPENSE values('" _
        & TrimChr(TMRId) & "','" _
        & expobj.ucType & "'," _
        & expobj.fExpens & ",'" _
        & expname & "','" _
        & Format(expobj.Day, "00") & "/" & Format(expobj.Month, "00") & "/" & Format(expobj.Year, "0000") & "','" _
        & expobj.Hour & ":" & expobj.Minutes & "','" _
        & TrimChr(HStr.PalmtecID) & "'," _
        & expobj.ucScheduleNo & ",'" _
        & TrimChr(expobj.Busno) & "','" _
        & TrimChr(expobj.EName) & "'," _
        & expobj.RptNo & ")"
     
        DB.Execute sql
        
'    Else
'        sql = "UPDATE EXPENSE SET ERecieptNo='" & AAAA & "',EXPCODE='' WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & expobj.ucScheduleNo & " AND TRIPNO= " & expobj.ucTripNo & " AND DATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "')"
'        DB.Execute sql
    End If

        sqlExp = "Select Trip_Master_ID  as TRM FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & expobj.ucScheduleNo & " AND TRIPNO= " & expobj.ucTripNo & " AND DateValue('" & Format(DA, "DD/MM/YYYY") & "') between StartDate and EndDate"
       ' sqlExp = "Select Trip_Master_ID  as TRM FROM RPT WHERE PALMID='" & Trim(PalmID) & "' AND SCHEDULE= " & expobj.ucScheduleNo & " AND TRIPNO= " & expobj.ucTripNo & " AND DATE = DateValue('" & Format(DA, "DD/MM/YYYY") & "')"
        Set RES8 = DB.OpenRecordset(sqlExp, dbOpenDynaset)
        If (RES8("TRM")) > 0 Then
            TMRId = RES8("TRM")
        Else
            TMRId = ""
        End If
            sqlExp = ""
            sqlExp = "Select sum(ExpAmt)as EXP1 from EXPENSE where TripMasterReferenceId='" & TMRId & "' "
            Set RES8 = DB.OpenRecordset(sqlExp, dbOpenDynaset)
            ' rs1 = CNN.Execute(sqlExp)
            'RES4.Edit
            If (RES8("EXP1")) > 0 Then
                sql = "UPDATE RPT SET Expense=" & RES8("EXP1") & " where Trip_Master_ID=" & TMRId & ""
           'Set RES7 = CNN.OpenRecordset(sql, dbOpenDynaset)
                DB.Execute (sql)
           End If
                
               
    Loop
    Close #Shdl
    Close #Fhdl
    Close #pHandle
   ' if res4.         res4.Close
    If f1 Then RES.Close
    'RES.Close
   

End Function


Public Function CreateNewtables() As Boolean
Dim tbls As TableDefs, tbl As TableDef
Dim fld As Field
Dim intCount As Integer
Dim TableExists As Boolean

On Error GoTo err
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
'''    Set tbls = DB.TableDefs
'''    For intCount = 0 To DB.TableDefs.Count - 1
'''        Set tbl = DB.TableDefs(intCount)
'''        If UCase(Trim(tbl.Name)) = "BMP_SETTINGS" Then
'''            TableExists = True
'''            Exit For
'''        End If
'''    Next intCount
'''
'''    If TableExists <> True Then
'''        Set tbl = DB.CreateTableDef("BMP_Settings", , "BMP_Settings")
'''        With tbl
'''            Set fld = .CreateField("Font_Name", dbText)
'''            fld.DefaultValue = "Arial"
'''            fld.AllowZeroLength = True
'''            .Fields.Append fld
'''
'''            Set fld = .CreateField("Font_Size", dbLong)
'''            fld.DefaultValue = 16
'''            .Fields.Append fld
'''
'''            Set fld = .CreateField("Bold_EnableOrDisable", dbLong)
'''            fld.DefaultValue = 0
'''            .Fields.Append fld
'''
'''            Set fld = .CreateField("Bmp_Width", dbLong)
'''            fld.DefaultValue = 192 '' 384
'''            .Fields.Append fld
'''
'''            Set fld = .CreateField("Bmp_Height", dbLong)
'''            fld.DefaultValue = 64
'''            .Fields.Append fld
'''
'''        End With
'''        tbls.Append tbl
'''    End If
    CreateNewtables = True
    
    Exit Function
err:
    MsgBox err.Number & "  " & err.Description & " in CreateNewtables"
End Function




Public Function CreateFields() As Boolean
Dim tbl As TableDef
Dim fld As Field
Dim intCount As Integer
Dim blnRFfieldExists As Boolean, blnSFfieldExists As Boolean, blnLanguagefieldExists As Boolean
Dim blnStgUpdtFieldExists As Boolean, blnDefaultStageFieldExists As Boolean
Dim blnNFfieldExists As Boolean, blnODEntryFieldExists As Boolean, blnTCKTFTFieldExists As Boolean
Dim blnCrewChkFieldExists As Boolean, blnAPfieldExists As Boolean, blnRefundfieldExists As Boolean, blnALPfieldExists As Boolean, blnSsendbillFieldExists As Boolean, blnStripenableFieldExists As Boolean, blnSSendpendFieldExists As Boolean, blnSTripSMSFieldExists As Boolean, blnSRechargeEnableFieldExists As Boolean, blnSDownloadEnableFieldExists As Boolean, blnSRechargeprEnableFieldExists As Boolean, blnSPhNoFieldExists As Boolean, blnSPhNo2FieldExists As Boolean, blnSPhNo3FieldExists As Boolean, blnSAccessPointFieldExists As Boolean, blnSDestAddsFieldExists As Boolean, blnSUsernameFieldExists As Boolean, blnSPasswordFieldExists As Boolean, blnSUploadpathFieldExists As Boolean ' 04/01/2010
Dim AutoshutdownExists As Boolean, UserpswdEnable, refundEnable As Boolean
Dim blnSDownloadpathFieldExists  As Boolean, blnSHttpUrlFieldExists As Boolean, blnSInspectrptFieldExists As Boolean, blnSGprsEnableFieldExists As Boolean, blnSGprsMessageEnableFieldExists As Boolean, blnSSTFareEnableFieldExists As Boolean, blnSSTRoundEnableFieldExists As Boolean, c As Boolean, blnSTRoundAmtFieldExists As Boolean, blnSMobileFieldExists As Boolean, blnSPinFieldExists As Boolean, blnSSchedulesendEnableFieldExists As Boolean, blnSScheduleSMSFieldExists As Boolean, SmartCardEnable As Boolean, ExpEnable As Boolean, FtpEnable As Boolean ' 04/01/2010
Dim blnBMPExists As Boolean    '' 14/01/2011
Dim blnphExists As Boolean
Dim blnTrExists As Boolean
Dim blnShExists As Boolean
Dim Simplereport As Boolean, ReportFONT As Boolean, MultiplePass As Boolean, InspectorSMS As Boolean
On Error GoTo err
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    
    blnRefundfieldExists = False
    Set tbl = DB.TableDefs("TKTS")  ' 04/01/2010
    For intCount = 0 To tbl.Fields.Count - 1
        If UCase(Trim(tbl.Fields(intCount).Name)) = "REFUNDSTS" Then
            blnRefundfieldExists = True
        End If
    Next intCount
    If blnRefundfieldExists <> True Then
        Set fld = tbl.CreateField("REFUNDSTS", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    blnRefundfieldExists = False
    For intCount = 0 To tbl.Fields.Count - 1
        If UCase(Trim(tbl.Fields(intCount).Name)) = "REFUNDAMT" Then
            blnRefundfieldExists = True
        End If
    Next intCount
    If blnRefundfieldExists <> True Then
        Set fld = tbl.CreateField("REFUNDAMT", dbDouble)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    
    blnAPfieldExists = False
    Set tbl = DB.TableDefs("RPT")  ' 04/01/2010
    For intCount = 0 To tbl.Fields.Count - 1
        If UCase(Trim(tbl.Fields(intCount).Name)) = "PENALITYCNT" Then
            blnAPfieldExists = True
        End If
    Next intCount
    If blnAPfieldExists <> True Then
        Set fld = tbl.CreateField("PENALITYCNT", dbByte)
        tbl.Fields.Append fld
   '     fld.DefaultValue
'        fld.DefaultValue = 0
        CreateFields = True
    End If
    
     blnAPfieldExists = False
    For intCount = 0 To tbl.Fields.Count - 1
        If UCase(Trim(tbl.Fields(intCount).Name)) = "PENALITYAMT" Then
            blnAPfieldExists = True
        End If
    Next intCount
    If blnAPfieldExists <> True Then
        Set fld = tbl.CreateField("PENALITYAMT", dbDouble)
        tbl.Fields.Append fld
   '     fld.DefaultValue
'        fld.DefaultValue = 0
        CreateFields = True
    End If
    
    
    blnAPfieldExists = False
    Set tbl = DB.TableDefs("CREW")  ' 04/01/2010
    For intCount = 0 To tbl.Fields.Count - 1
        If UCase(Trim(tbl.Fields(intCount).Name)) = "PSWD" Then
            blnAPfieldExists = True
        End If
    Next intCount
    If blnAPfieldExists <> True Then
        Set fld = tbl.CreateField("PSWD", dbText)
        tbl.Fields.Append fld
        fld.AllowZeroLength = True
   '     fld.DefaultValue
'        fld.DefaultValue = 0
        CreateFields = True
    End If
    Set tbl = DB.TableDefs("Settings")
    For intCount = 0 To tbl.Fields.Count - 1
        If UCase(Trim(tbl.Fields(intCount).Name)) = "REMOVETICKETFLAG" Then
            blnRFfieldExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "STAGEFONTFLAG" Then
            blnSFfieldExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "LANGUAGEOPTION" Then
            blnLanguagefieldExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "STAGEUPDATIONMSG" Then
            blnStgUpdtFieldExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "DEFAULTSTAGE" Then
            blnDefaultStageFieldExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "NEXTFAREFLAG" Then
            blnNFfieldExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "ODOMETERENTRY" Then
            blnODEntryFieldExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "TICKETNOBIGFONT" Then
            blnTCKTFTFieldExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "CREWCHECK" Then
            blnCrewChkFieldExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "PHNO" Then
            blnphExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "TRIPSMS" Then
            blnTrExists = True
        End If
        If UCase(Trim(tbl.Fields(intCount).Name)) = "SHSMS" Then
            blnShExists = True
        End If 'SANGEETHA
        
           If (Trim(tbl.Fields(intCount).Name)) = "sendbillEnable" Then   'SANGEETHA
            blnSsendbillFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "TripsendEnable" Then   'SANGEETHA
            blnStripenableFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "SchedulesendEnable" Then   'SANGEETHA
            blnSSchedulesendEnableFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "Sendpend" Then   'SANGEETHA
            blnSSendpendFieldExists = True
        End If
       
        If (Trim(tbl.Fields(intCount).Name)) = "GprsEnable" Then   'SANGEETHA
            blnSGprsEnableFieldExists = True
        End If
'        If (Trim(tbl.Fields(intCount).Name)) = "SmartCardEnable" Then   'SANGEETHA
'            SmartCardEnable = True
'        End If
        If (Trim(tbl.Fields(intCount).Name)) = "ExpEnable" Then   'SANGEETHA
            ExpEnable = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "GprsEnableMessage" Then
            blnSGprsMessageEnableFieldExists = True
        End If
        
        If (Trim(tbl.Fields(intCount).Name)) = "PhNo2" Then   'SANGEETHA
            blnSPhNo2FieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "AccessPoint" Then   'SANGEETHA
            blnSAccessPointFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "DestAdds" Then   'SANGEETHA
            blnSDestAddsFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "Username" Then   'SANGEETHA
            blnSUsernameFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "Password" Then   'SANGEETHA
            blnSPasswordFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "Uploadpath" Then   'SANGEETHA
            blnSUploadpathFieldExists = True
        End If
        
        If (Trim(tbl.Fields(intCount).Name)) = "Downloadpath" Then   'SANGEETHA
            blnSDownloadpathFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "HttpUrl" Then   'SANGEETHA
            blnSHttpUrlFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "InspectRpt" Then   'SANGEETHA
            blnSInspectrptFieldExists = True
        End If
         If (Trim(tbl.Fields(intCount).Name)) = "GprsEnable" Then   'sarika
            blnSGprsEnableFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "SmartCard" Then   'sarika
            SmartCardEnable = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "ExpEnable" Then
            ExpEnable = True
        End If
         If (Trim(tbl.Fields(intCount).Name)) = "FtpEnable" Then
            FtpEnable = True
        End If
         If (Trim(tbl.Fields(intCount).Name)) = "GprsEnableMessage" Then
            blnSGprsMessageEnableFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "StRoundoffEnable" Then
            blnSSTRoundEnableFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "StRoundoffAmt" Then
            blnSTRoundAmtFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "StFareEdit" Then
            blnSSTFareEnableFieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "SimpleReport" Then
            Simplereport = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "ReportFONT" Then
            ReportFONT = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "MultiplePass" Then
            MultiplePass = True
        End If
         If (Trim(tbl.Fields(intCount).Name)) = "InspectorSMS" Then
            InspectorSMS = True
        End If
         If (Trim(tbl.Fields(intCount).Name)) = "PhNo3" Then   'SANGEETHA
            blnSPhNo3FieldExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "AutoShutDown" Then   'SANGEETHA
            AutoshutdownExists = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "UserpswdEnable" Then   'SANGEETHA
            UserpswdEnable = True
        End If
        If (Trim(tbl.Fields(intCount).Name)) = "REFUND" Then   'SANGEETHA
            refundEnable = True
        End If
    Next intCount
    
    
    
    If blnRFfieldExists <> True Then
        Set fld = tbl.CreateField("RemoveTicketFlag", dbByte)
        tbl.Fields.Append fld
        fld.DefaultValue = 0
        CreateFields = True
    End If
    If blnSFfieldExists <> True Then
        Set fld = tbl.CreateField("StageFontFlag", dbByte)
        tbl.Fields.Append fld
         fld.DefaultValue = 0
        CreateFields = True
    End If
    If blnLanguagefieldExists <> True Then
        Set fld = tbl.CreateField("LANGUAGEOPTION", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnStgUpdtFieldExists <> True Then
        Set fld = tbl.CreateField("STAGEUPDATIONMSG", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnDefaultStageFieldExists <> True Then
        Set fld = tbl.CreateField("DEFAULTSTAGE", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnNFfieldExists <> True Then
        Set fld = tbl.CreateField("NextFareFlag", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnODEntryFieldExists <> True Then
        Set fld = tbl.CreateField("ODOMETERENTRY", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnTCKTFTFieldExists <> True Then
        Set fld = tbl.CreateField("TICKETNOBIGFONT", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnCrewChkFieldExists <> True Then
        Set fld = tbl.CreateField("CREWCHECK", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnphExists <> True Then
        Set fld = tbl.CreateField("PHNO", dbText)
        tbl.Fields.Append fld
        CreateFields = True
    End If
     If blnTrExists <> True Then
        Set fld = tbl.CreateField("TRIPSMS", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
      If blnShExists <> True Then
        Set fld = tbl.CreateField("SHSMS", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If 'sangeetha
    If blnStripenableFieldExists <> True Then
        Set fld = tbl.CreateField("TripsendEnable", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnSSchedulesendEnableFieldExists <> True Then
        Set fld = tbl.CreateField("SchedulesendEnable", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnSSendpendFieldExists <> True Then
        Set fld = tbl.CreateField("Sendpend", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
'    If blnSGprsEnableFieldExists <> True Then
'        Set fld = tbl.CreateField("GprsEnable", dbByte)
'        tbl.Fields.Append fld
'        CreateFields = True
'    End If
     If blnSInspectrptFieldExists <> True Then
        Set fld = tbl.CreateField("InspectRpt", dbByte)
        tbl.Fields.Append fld
        fld.DefaultValue = 0
        CreateFields = True
    End If
    
    
'    If SmartCardEnable <> True Then
'        Set fld = tbl.CreateField("SmartCard", dbByte)
'        tbl.Fields.Append fld
'        CreateFields = True
'    End If
'    If ExpEnable <> True Then
'        Set fld = tbl.CreateField("ExpEnable", dbByte)
'        tbl.Fields.Append fld
'        CreateFields = True
'    End If
'     If FtpEnable <> True Then
'        Set fld = tbl.CreateField("FtpEnable", dbByte)
'        tbl.Fields.Append fld
'        CreateFields = True
'    End If
'    If blnSGprsMessageEnableFieldExists <> True Then
'        Set fld = tbl.CreateField("GprsEnableMessage", dbByte)
'        tbl.Fields.Append fld
'        CreateFields = True
'    End If
    If blnSPhNo2FieldExists <> True Then
        Set fld = tbl.CreateField("PhNo2", dbText)
        tbl.Fields.Append fld
        fld.DefaultValue = 0
        CreateFields = True
    End If
    If blnSAccessPointFieldExists <> True Then
        Set fld = tbl.CreateField("AccessPoint", dbText)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnSDestAddsFieldExists <> True Then
        Set fld = tbl.CreateField("DestAdds", dbText)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnSUsernameFieldExists <> True Then
        Set fld = tbl.CreateField("Username", dbText)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnSPasswordFieldExists <> True Then
        Set fld = tbl.CreateField("Password", dbText)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnSUploadpathFieldExists <> True Then
        Set fld = tbl.CreateField("Uploadpath", dbText)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnSDownloadpathFieldExists <> True Then
        Set fld = tbl.CreateField("Downloadpath", dbText)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnSHttpUrlFieldExists <> True Then
        Set fld = tbl.CreateField("HttpUrl", dbText)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If blnSGprsEnableFieldExists <> True Then
        Set fld = tbl.CreateField("GprsEnable", dbText)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    If SmartCardEnable <> True Then
        Set fld = tbl.CreateField("SmartCard", dbText)
        tbl.Fields.Append fld
        fld.DefaultValue = 0
        CreateFields = True
    End If
    If ExpEnable <> True Then
        Set fld = tbl.CreateField("ExpEnable", dbText)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
     If FtpEnable <> True Then
        Set fld = tbl.CreateField("FtpEnable", dbText)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
    If blnSGprsMessageEnableFieldExists <> True Then
        Set fld = tbl.CreateField("GprsEnableMessage", dbText)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
    If blnSsendbillFieldExists <> True Then
        Set fld = tbl.CreateField("sendbillEnable", dbByte)
        tbl.Fields.Append fld
        CreateFields = True
    End If
    
     If blnSSTRoundEnableFieldExists <> True Then
        Set fld = tbl.CreateField("StRoundoffEnable", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
     If blnSTRoundAmtFieldExists <> True Then
        Set fld = tbl.CreateField("StRoundoffAmt", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
    If blnSSTFareEnableFieldExists <> True Then
        Set fld = tbl.CreateField("StFareEdit", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
    
  
     If ReportFONT <> True Then
        Set fld = tbl.CreateField("ReportFONT", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
     If MultiplePass <> True Then
        Set fld = tbl.CreateField("MultiplePass", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
     If Simplereport <> True Then
        Set fld = tbl.CreateField("SimpleReport", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
     If InspectorSMS <> True Then
        Set fld = tbl.CreateField("InspectorSMS", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
    If blnSPhNo3FieldExists <> True Then
        Set fld = tbl.CreateField("PhNo3", dbText)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
    If AutoshutdownExists <> True Then
        Set fld = tbl.CreateField("AutoShutDown", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
    If UserpswdEnable <> True Then
        Set fld = tbl.CreateField("UserpswdEnable", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
    If refundEnable <> True Then
        Set fld = tbl.CreateField("REFUND", dbByte)
        tbl.Fields.Append fld
          fld.DefaultValue = 0
        CreateFields = True
    End If
    
    Set tbl = DB.TableDefs("Route")  ' 04/01/2010
    For intCount = 0 To tbl.Fields.Count - 1
        If UCase(Trim(tbl.Fields(intCount).Name)) = "PASSALLOW" Then
            blnAPfieldExists = True
        End If
    Next intCount
    If blnAPfieldExists <> True Then
        Set fld = tbl.CreateField("PASSALLOW", dbByte)
        tbl.Fields.Append fld
   '     fld.DefaultValue
'        fld.DefaultValue = 0
        CreateFields = True
    End If
    
  
    '''--------------------------------------------------------------------
    
'''    blnBMPExists = False
'''    Set tbl = DB.TableDefs("Stage")  ' 04/01/2010
'''    For intCount = 0 To tbl.Fields.Count - 1
'''        If UCase(Trim(tbl.Fields(intCount).Name)) = "BMPFILE" Then
'''            blnBMPExists = True
'''        End If
'''    Next intCount
'''    If blnBMPExists <> True Then
'''        Set fld = tbl.CreateField("BMPFile", dbText)
'''        fld.AllowZeroLength = True
'''        tbl.Fields.Append fld
'''        CreateFields = True
'''    End If
    
    '''----------------------------------------------------------------------
    
    
    Set DB1 = DAO.OpenDatabase(App.Path & "\GBackUp.mdb", dbDriverComplete, False, ";UID=;PWD=")
    'App.Path & "\GBackUp.mdb"
    Set tbl = DB1.TableDefs("ROUTE")  ' 04/01/2010
    For intCount = 0 To tbl.Fields.Count - 1
        If UCase(Trim(tbl.Fields(intCount).Name)) = "ALLOW" Then
            blnALPfieldExists = True
        End If
    Next intCount
    If blnALPfieldExists <> True Then
        Set fld = tbl.CreateField("ALLOW", dbByte)
        tbl.Fields.Append fld
'        fld.DefaultValue = 0
        CreateFields = True
    End If
    
'''    blnBMPExists = False
'''    Set tbl = DB1.TableDefs("STAGE")  ' 04/01/2010
'''    For intCount = 0 To tbl.Fields.Count - 1
'''        If UCase(Trim(tbl.Fields(intCount).Name)) = "BMPFILE" Then
'''            blnBMPExists = True
'''        End If
'''    Next intCount
'''    If blnBMPExists <> True Then
'''        Set fld = tbl.CreateField("BMPFile", dbText)
'''        fld.AllowZeroLength = True
'''        tbl.Fields.Append fld
'''        CreateFields = True
'''    End If
    
    
    Exit Function
err:
    CreateFields = False
End Function

Public Function CreateSchedule() As Boolean

    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    If CreateRouteList = True Then
        CreateStageList
        CreateRoute
    End If
End Function
Public Function DownSetup() As Boolean
Dim Fhandle As Integer
Dim sql As String
Dim bytPrev As Byte, intLoopCount As Integer
On Error GoTo err
 Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
 Fhandle = FreeFile()
 ' If Dir(App.Path & "\BUS.DAT") <> "" Then Kill App.Path & "\BUS.DAT"
    
  Open App.Path & "\BUS.DAT" For Binary Access Read Write As #Fhandle
    Get #Fhandle, , HStr
    Get #Fhandle, , hardwaresettings
  Close #Fhandle
  
  Kill App.Path & "\BUS.DAT"
  Fhandle = FreeFile()
  Open App.Path & "\BUS.DAT" For Binary Access Write As #Fhandle

  sql = "select * from settings"
  Set RES = DB.OpenRecordset(sql, dbOpenDynaset)
  
  If RES.RecordCount <= 0 Then
      DownSetup = False
      Exit Function
  End If
  If RES!UserPWD = " " Or RES!MasterPWD = " " Then
      DownSetup = False
      Exit Function
  End If
 If RES.RecordCount > 0 Then
    RES.MoveFirst
    Do While Not RES.EOF
    
      If TrimChr(hardwaresettings.MSR_PSWD) <> TrimChr(RES!MasterPWD) Then
       hardwaresettings.MSR_PSWD = RES!MasterPWD & Chr(0)
       
      End If
      
      If TrimChr(hardwaresettings.USR_PSWD) <> TrimChr(RES!UserPWD) Then
       hardwaresettings.USR_PSWD = RES!UserPWD & Chr(0)
       
      End If
      If hardwaresettings.select_language <> IIf(IsNull(RES!LANGUAGEOPTION), 0, RES!LANGUAGEOPTION) Then
        hardwaresettings.select_language = IIf(IsNull(RES!LANGUAGEOPTION), 0, RES!LANGUAGEOPTION)
       
      End If
       If hardwaresettings.LangNo <> IIf(IsNull(RES!LANGUAGEOPTION), 0, RES!LANGUAGEOPTION) Then
        hardwaresettings.LangNo = IIf(IsNull(RES!LANGUAGEOPTION), 0, RES!LANGUAGEOPTION)
       
      End If
      hardwaresettings.ucTemp = 1
      If TrimChr(HStr.MainDisp) <> TrimChr(RES!MainDisplay) Then
       HStr.MainDisp = RES!MainDisplay & Chr(0)
       
      End If
      
      If TrimChr(HStr.MainDisp2) <> TrimChr(RES!MainDisplay2) Then
       HStr.MainDisp2 = RES!MainDisplay2 & Chr(0)
       
      End If
      
      If TrimChr(HStr.bhl1) <> TrimChr(RES!HEADER1) Then
       HStr.bhl1 = RES!HEADER1 & Chr(0)
       
      End If
      
      If TrimChr(HStr.bhl1) <> TrimChr(RES!HEADER1) Then
       HStr.bhl1 = RES!HEADER1 & Chr(0)
       
      End If
      
      If TrimChr(HStr.bhl2) <> TrimChr(RES!HEADER2) Then
       HStr.bhl2 = RES!HEADER2 & Chr(0)
       
      End If
      
      If TrimChr(HStr.bhl2) <> TrimChr(RES!HEADER2) Then
       HStr.bhl2 = RES!HEADER2 & Chr(0)
       
      End If
      
      If TrimChr(HStr.bhl3) <> TrimChr(RES!Header3) Then
       HStr.bhl3 = RES!Header3 & Chr(0)
       
      End If
      
      If TrimChr(HStr.bfl1) <> TrimChr(RES!Footer1) Then
       HStr.bfl1 = RES!Footer1 & Chr(0)
       
      End If
      
      If TrimChr(HStr.bfl2) <> TrimChr(RES!Footer2) Then
       HStr.bfl2 = RES!Footer2 & Chr(0)
       
      End If
        strftr = ftreditvalue
        HStr.bfl2 = IIf(strftr = "", HStr.bfl2, strftr & Chr(0))
        If TrimChr(HStr.PalmtecID) <> TrimChr(RES!PalmtecID) Then
            HStr.PalmtecID = RES!PalmtecID & Chr(0)
        End If
      
      If TrimChr(HStr.HalfPer) <> TrimChr(RES!HalfPer) Then
       HStr.HalfPer = CByte(val(RES!HalfPer))
       
      End If
      
      If TrimChr(HStr.HalfPer) <> TrimChr(RES!HalfPer) Then
       HStr.HalfPer = CByte(val(RES!HalfPer))
       
      End If
      
      If TrimChr(HStr.ConPer) <> TrimChr(RES!ConPer) Then
       HStr.ConPer = CByte(val(RES!ConPer))
       
      End If
      'chkstenableucbSTFareEdit
      If TrimChr(HStr.STMaxAmt) <> TrimChr(RES!STMaxAmt) Then
       HStr.STMaxAmt = CSng(val(RES!STMaxAmt))
       
      End If
      If TrimChr(HStr.STMinCon) <> TrimChr(RES!STMinCon) Then
       HStr.STMinCon = CSng(val(RES!STMinCon))
       
      End If
      
      If TrimChr(HStr.PhyPer) <> TrimChr(RES!PhyPer) Then
       HStr.PhyPer = CByte(val(RES!PhyPer))
       
      End If
      
      If TrimChr(HStr.Roundoff) <> TrimChr(RES!Roundoff) Then
       HStr.Roundoff = CByte(val(RES!Roundoff))
       
      End If
      If TrimChr(HStr.RoundUp) <> TrimChr(RES!RoundUp) Then
       HStr.RoundUp = CByte(val(RES!RoundUp))
       
      End If
      
      If TrimChr(HStr.RoundAmt) <> TrimChr(RES!RoundAmt) Then
       HStr.RoundAmt = CByte(val(RES!RoundAmt))
       
      End If
      
      If TrimChr(HStr.LuggageUnitRate) <> TrimChr(RES!LuggageUnitRate) Then
       HStr.LuggageUnitRate = CSng(val(RES!LuggageUnitRate))
       
      End If
     
      If TrimChr(HStr.Currency) <> TrimChr(RES!Currency) Then
       HStr.Currency = RES!Currency & Chr(0)
       End If
       
       '''''''''''''SYAM ADDED
       
       If (HStr.ReportFlag) <> (RES!ReportFlag) Then
            HStr.ReportFlag = CByte(val(RES!ReportFlag))
       End If
      bytPrev = 1
        For intLoopCount = 1 To 7
            bytPrev = (bytPrev * 2)
            Select Case bytPrev
''                Case 2:
''                    If (.ReportFlag And bytPrev) = bytPrev Then chkPrint.Value = 1
''                Case 4:
''                    If (.ReportFlag And bytPrev) = bytPrev Then chkAdvertise.Value = 1
''                Case 8:
''                    If (.ReportFlag And bytPrev) = bytPrev Then chklogenable.Value = 1
                Case 16:
                    If (HStr.ReportFlag And bytPrev) = bytPrev Then
                            HStr.Busno = 1
                            HStr.Conductor = 1
                            HStr.Driver = 1
                    Else
                            HStr.Busno = 0
                            HStr.Conductor = 0
                            HStr.Driver = 0
                    
                    End If
                Case 32:
                   ' If (.ReportFlag And bytPrev) = bytPrev Then chkbigfontenable.Value = 1
                Case 64:
                    
            End Select
        Next
       
       Call CreateFields
       
       If (HStr.EnableRemoveTicket) <> IIf(IsNull(RES!RemoveTicketFlag), 0, RES!RemoveTicketFlag) Then
        HStr.EnableRemoveTicket = CByte(IIf(IsNull(RES!RemoveTicketFlag), 0, RES!RemoveTicketFlag))
       End If
        
       If (HStr.UpdateStageMsg) <> IIf(IsNull(RES!STAGEUPDATIONMSG), 0, RES!STAGEUPDATIONMSG) Then
        HStr.UpdateStageMsg = CByte(IIf(IsNull(RES!STAGEUPDATIONMSG), 0, RES!STAGEUPDATIONMSG))
       End If
            
       If (HStr.EnableStageDefault) <> IIf(IsNull(RES!DEFAULTSTAGE), 0, RES!DEFAULTSTAGE) Then
       HStr.EnableStageDefault = CByte(IIf(IsNull(RES!DEFAULTSTAGE), 0, RES!DEFAULTSTAGE))
       End If
       
       If (HStr.EnableStageFont) <> IIf(IsNull(RES!StageFontFlag), 0, RES!StageFontFlag) Then
        HStr.EnableStageFont = CByte(IIf(IsNull(RES!StageFontFlag), 0, RES!StageFontFlag))
       End If
       
       If (HStr.NextFareRound) <> IIf(IsNull(RES!NextFareFlag), 0, RES!NextFareFlag) Then
        HStr.NextFareRound = CByte(IIf(IsNull(RES!NextFareFlag), 0, RES!NextFareFlag))
       End If
    
       '!NextFareFlag=sp.NextFareRound
       '''''''''''''''''''''''
       '''''''''''''''''''''''RNC
       If (HStr.OdometerEntry) <> IIf(IsNull(RES!OdometerEntry), 0, RES!OdometerEntry) Then
       HStr.OdometerEntry = CByte(IIf(IsNull(RES!OdometerEntry), 0, RES!OdometerEntry))
       End If
       
       If (HStr.TicketNoBigFont) <> IIf(IsNull(RES!TicketNoBigFont), 0, RES!TicketNoBigFont) Then
        HStr.TicketNoBigFont = CByte(IIf(IsNull(RES!TicketNoBigFont), 0, RES!TicketNoBigFont))
       End If
       
       If (HStr.CrewCheck) <> IIf(IsNull(RES!CrewCheck), 0, RES!CrewCheck) Then
        HStr.CrewCheck = CByte(IIf(IsNull(RES!CrewCheck), 0, RES!CrewCheck))
       End If
      
'       If (HStr.PhNo) <> IIf(IsNull(RES!PhNo), 0, RES!PhNo) Then
'        HStr.PhNo = CStr(IIf(IsNull(RES!PhNo), 0, RES!PhNo))
'       End If
       
        If TrimChr(HStr.PhNo) <> TrimChr(RES!PhNo) Then
         HStr.PhNo = RES!PhNo & Chr(0)
        End If
      
       If (HStr.TripSMS) <> IIf(IsNull(RES!TripSMS), 0, RES!TripSMS) Then
        HStr.TripSMS = CByte(IIf(IsNull(RES!TripSMS), 0, RES!TripSMS))
       End If
       If (HStr.ScheduleSMS) <> IIf(IsNull(RES!shSMS), 0, RES!shSMS) Then
        HStr.ScheduleSMS = CByte(IIf(IsNull(RES!shSMS), 0, RES!shSMS))
       End If 'SANGEETHA
      
       '''''''''''''''''''''''''''
        If (HStr.Sendpend) <> IIf(IsNull(RES!Sendpend), 0, RES!Sendpend) Then
        HStr.Sendpend = CByte(IIf(IsNull(RES!Sendpend), 0, RES!Sendpend))
       End If
'''       If (HStr.gprs) <> IIf(IsNull(RES!GprsEnable), 0, RES!GprsEnable) Then
'''        HStr.GprsEnable = CByte(IIf(IsNull(RES!GprsEnable), 0, RES!GprsEnable))
'''       End If
   
       If TrimChr(HStr.PhNo2) <> TrimChr(RES!PhNo2) Then
       HStr.PhNo2 = TrimChr(RES!PhNo2) & Chr(0)
       End If
       If TrimChr(HStr.AccessPoint) <> TrimChr(RES!AccessPoint) Then
       HStr.AccessPoint = TrimChr(RES!AccessPoint) & Chr(0)
       End If
       If TrimChr(HStr.DestAdds) <> TrimChr(RES!DestAdds) Then
       HStr.DestAdds = TrimChr(RES!DestAdds) & Chr(0)
       End If
       If TrimChr(HStr.Username) <> TrimChr(RES!Username) Then
       HStr.Username = TrimChr(RES!Username) & Chr(0)
       End If
       If TrimChr(HStr.PassWord) <> TrimChr(RES!PassWord) Then
       HStr.PassWord = TrimChr(RES!PassWord) & Chr(0)
       End If
       If TrimChr(HStr.Uploadpath) <> TrimChr(RES!Uploadpath) Then
       HStr.Uploadpath = TrimChr(RES!Uploadpath) & Chr(0)
       End If
       If TrimChr(HStr.Downloadpath) <> TrimChr(RES!Downloadpath) Then
       HStr.Downloadpath = TrimChr(RES!Downloadpath) & Chr(0)
       End If
'        MsgBox "2"
       If TrimChr(HStr.HttpUrl) <> TrimChr(RES!HttpUrl) Then
       HStr.HttpUrl = TrimChr(RES!HttpUrl) & Chr(0)
       End If
        
       If TrimChr(HStr.GprsEnable) <> TrimChr(RES!GprsEnable) Then
       HStr.GprsEnable = TrimChr(RES!GprsEnable) & Chr(0)
       End If
       
       If TrimChr(HStr.SmartCard) <> IIf(IsNull(RES!SmartCard), 0, RES!SmartCard) Then
       HStr.SmartCard = IIf(IsNull(RES!SmartCard), 0, RES!SmartCard) & Chr(0)
       End If
      
       If TrimChr(HStr.ExpEnable) <> IIf(IsNull(RES!ExpEnable), 0, RES!ExpEnable) Then
       HStr.ExpEnable = IIf(IsNull(RES!SmartCard), 0, RES!ExpEnable) & Chr(0)
       End If
      
       If TrimChr(HStr.MsgPrompt) <> IIf(IsNull(RES!GprsEnableMessage), 0, RES!GprsEnableMessage) Then
       HStr.MsgPrompt = IIf(IsNull(RES!GprsEnableMessage), 0, RES!GprsEnableMessage) & Chr(0)
       End If
       
        If (HStr.TripsendEnable) <> IIf(IsNull(RES!TripsendEnable), 0, RES!TripsendEnable) Then
        HStr.TripsendEnable = CByte(IIf(IsNull(RES!TripsendEnable), 0, RES!TripsendEnable))
       End If
       If (HStr.SchedulesendEnable) <> IIf(IsNull(RES!SchedulesendEnable), 0, RES!SchedulesendEnable) Then
        HStr.SchedulesendEnable = CByte(IIf(IsNull(RES!SchedulesendEnable), 0, RES!SchedulesendEnable))
       End If
       If (HStr.sendbillEnable) <> IIf(IsNull(RES!sendbillEnable), 0, RES!sendbillEnable) Then
        HStr.sendbillEnable = CByte(IIf(IsNull(RES!sendbillEnable), 0, RES!sendbillEnable))
       End If
       
        If (HStr.FtpEnable) <> IIf(IsNull(RES!FtpEnable), 0, RES!FtpEnable) Then
        HStr.FtpEnable = CByte(IIf(IsNull(RES!FtpEnable), 0, RES!FtpEnable))
       End If
       
       If (HStr.Inspectorreport) <> IIf(IsNull(RES!InspectRpt), 0, RES!InspectRpt) Then
            HStr.Inspectorreport = CByte(IIf(IsNull(RES!InspectRpt), 0, RES!InspectRpt))
       End If
       
        If (HStr.StRoundoff_E_D) <> IIf(IsNull(RES!StRoundoffEnable), 0, RES!StRoundoffEnable) Then
            HStr.StRoundoff_E_D = CByte(IIf(IsNull(RES!StRoundoffEnable), 0, RES!StRoundoffEnable))
       End If
       If TrimChr(HStr.StRoundoff_Amt) <> IIf(IsNull(RES!StRoundoffAmt), 0, RES!StRoundoffAmt) Then
       HStr.StRoundoff_Amt = CByte(IIf(IsNull(RES!StRoundoffAmt), 0, RES!StRoundoffAmt))
       
      End If
      If (HStr.ucbSTFareEdit) <> IIf(IsNull(RES!StFareEdit), 0, RES!StFareEdit) Then
            HStr.ucbSTFareEdit = CByte(IIf(IsNull(RES!StFareEdit), 0, RES!StFareEdit))
     End If
       
     If (HStr.ReportFONT) <> IIf(IsNull(RES!ReportFONT), 0, RES!ReportFONT) Then
            HStr.ReportFONT = CByte(IIf(IsNull(RES!ReportFONT), 0, RES!ReportFONT))
     End If
     If (HStr.MultiplePass) <> IIf(IsNull(RES!MultiplePass), 0, RES!MultiplePass) Then
            HStr.MultiplePass = CByte(IIf(IsNull(RES!MultiplePass), 0, RES!MultiplePass))
     End If
    
       If (HStr.Simplereport) <> IIf(IsNull(RES!Simplereport), 0, RES!Simplereport) Then
            HStr.Simplereport = CByte(IIf(IsNull(RES!Simplereport), 0, RES!Simplereport))
     End If
     If (HStr.InspectorSMS) <> IIf(IsNull(RES!InspectorSMS), 0, RES!InspectorSMS) Then
            HStr.InspectorSMS = CByte(IIf(IsNull(RES!InspectorSMS), 0, RES!InspectorSMS))
     End If
     If TrimChr(HStr.PhNo3) <> IIf(IsNull(RES!PhNo3), 0, RES!PhNo3) Then
        HStr.PhNo3 = TrimChr(IIf(IsNull(RES!PhNo3), 0, RES!PhNo3)) & Chr(0)
     End If
    If (HStr.AutoShutdownEnable) <> IIf(IsNull(RES!Autoshutdown), 0, RES!Autoshutdown) Then
            HStr.AutoShutdownEnable = CByte(IIf(IsNull(RES!Autoshutdown), 0, RES!Autoshutdown))
    End If
    If (HStr.UserPasswordEnable) <> IIf(IsNull(RES!UserpswdEnable), 0, RES!UserpswdEnable) Then
            HStr.UserPasswordEnable = CByte(IIf(IsNull(RES!UserpswdEnable), 0, RES!UserpswdEnable))
    End If
    If (HStr.refund) <> IIf(IsNull(RES!refund), 0, RES!refund) Then
        HStr.refund = CByte(IIf(IsNull(RES!refund), 0, RES!refund))
    End If
    HStr.ladis_per = IIf(IsNumeric(RES!ladies_ratio), RES!ladies_ratio, 0)
    HStr.seniar_per = IIf(IsNumeric(RES!senior_ratio), RES!senior_ratio, 0)

       ''''''''''''''''''''''''''''
      
      Put #Fhandle, , HStr
      Put #Fhandle, , hardwaresettings
      'MsgBox "1"
      RES.MoveNext
    Loop
   End If
  RES.Close
Close #Fhandle
DownSetup = True
    Exit Function
err:
    MsgBox err.Number & " , " & err.Description
    Close #Fhandle
    DownSetup = False
    RES.Close
    Exit Function

End Function
  

Public Function CreateRouteList() As Boolean
''On Error GoTo err
Dim Fhandle As Integer
Dim Route As RouteLST
Dim ubcount As Integer
Dim I As Integer

If SeletedRouteCount = 0 Then Exit Function

    If UBound(SeletedRoute) > 0 Then
    
    'ubcount = UBound(SeletedRoute) - 1
    For I = 0 To UBound(SeletedRoute) - 1
    If SeletedRoute(I) <> "" Then
    ubcount = ubcount + 1
    End If
    Next I
    RSql = "SELECT * FROM ROUTE where rutcode='"
    For I = 0 To UBound(SeletedRoute)
        If SeletedRoute(I) <> "" Then
            RSql = RSql & SeletedRoute(I)
            If I < ubcount - 1 Then
            RSql = RSql & "' OR rutcode='"
            End If
        End If
    Next I
    'RSql = RSql & "'"-----------------
    RSql = RSql & "' order by RUTCODE"
    Debug.Print RSql
    Else
    'RSql = "SELECT * FROM ROUTE order by RUTCODE"
    RSql = "SELECT * FROM ROUTE "
    'RSql = "SELECT * FROM ROUTE order by RUTCODE"
    End If
         Fhandle = FreeFile()
         If Dir(App.Path & "\RouteLst.LST") <> "" Then Kill (App.Path & "\RouteLst.LST")
         Open App.Path & "\RouteLst.LST" For Binary Access Write As #Fhandle
     For I = 0 To UBound(SeletedRoute)
         RSql = "SELECT * FROM ROUTE where rutcode='"

       If SeletedRoute(I) <> "" Then
           RSql = RSql & SeletedRoute(I) & "'"
            'RSql = RSql & SeletedRoute(I) & "' order by RUTCODE"
        'RSql = "SELECT * FROM ROUTE"
         Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
         
         Dim BusSql As String 'SYAM
         
         If RES.RecordCount > 0 Then
            RES.MoveFirst
            Do While Not RES.EOF
                BusSql = "SELECT NAME FROM BUSTYPE WHERE ID=" & val(RES!BusType) ' Done By Rensy
                
                Set res1 = DB.OpenRecordset(BusSql, dbOpenDynaset)
                Route.Code = RES!RUTCODE & Chr(0)
                Route.Name = RES!rutname & Chr(0)
                Route.NoOfStage = CByte(val(RES!nostage))
                Route.MinFare = CSng(val((RES!MinFare)))
                Route.FareType = CByte(val(RES!FareType))
                Route.BusType = CByte(val(RES!BusType))
                If res1.RecordCount > 0 Then
                    res1.MoveFirst
                    If Len(res1!Name) > 15 Then
                        Route.BusTypeName = Mid(Trim(res1!Name), 0, 15) & Chr(0)
                    Else
                        Route.BusTypeName = Trim(res1!Name) & Chr(0)
                    End If
                Else
                    Route.BusTypeName = "ORDINARY"
                End If
                Route.AllowHalf = CByte(val(RES!Half))
                Route.AllowLug = CByte(val(RES!Luggage))
                Route.AllowAdjust = CByte(val(RES!Adjust))
                If Not IsNull(RES!PASSALLOW) Then
                    Route.AllowPass = CByte(val(RES!PASSALLOW)) '04/01/2010
                Else
                    Route.AllowPass = 0
                End If
                Route.AllowConc = CByte(val(RES!Conc))
                Route.AllowPh = CByte(val(RES!ph))
                Route.StartFrom = CByte(val(RES!StartFrom))
'                Route.cTemp = 1 ''Chr(0) '04/01/2010
                res1.Close
                Put #Fhandle, , Route
                RES.MoveNext
            Loop
         End If
         RES.Close
      End If
    Next I
         Close #Fhandle
    CreateRouteList = True
    Exit Function
err:
    MsgBox err.Number & " , " & err.Description
    Close #Fhandle
    CreateRouteList = False
    RES.Close
    Exit Function

End Function


'''''''''''''' Original   20/01/2011
Public Function CreateStageList() As Boolean
Dim Fhandle As Integer
Dim sHandle As Integer
Dim lHandle As Integer
Dim Stage As STAGEDETAILS
Dim LanguageStageCode As String * 24

    RSql = "SELECT StageName,Distance,STG_LOCAL_LANGUAGE FROM STAGE ORDER BY ID"
    Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)


    If RES.RecordCount > 0 Then
        If Dir(App.Path & "\STAGE.LST") <> "" Then Kill App.Path & "\STAGE.LST"
        If Dir(App.Path & "\LANGUAGE.DAT") <> "" Then Kill App.Path & "\LANGUAGE.DAT"
        Fhandle = FreeFile()
        Open App.Path & "\STAGE.LST" For Binary Access Write As #Fhandle
        lHandle = FreeFile()
        Open App.Path & "\LANGUAGE.DAT" For Binary Access Write As #lHandle
        RES.MoveFirst
        Do While Not RES.EOF
'        If RES!StageName = "WEQ" Then
'            MsgBox "1"
'        End If
        
            Stage.StageName = Mid$(Trim(RES!StageName), 1, 11) & Chr(0)
            Debug.Print Stage.StageName
            Stage.Distance = Format(val(IIf(IsNull(RES!Distance), 0, RES!Distance)), "0.00")
            Put #Fhandle, , Stage
            If RES!STG_LOCAL_LANGUAGE <> "" Then  'LAN
                LanguageStageCode = RES!STG_LOCAL_LANGUAGE & Chr(0)
            Else
                LanguageStageCode = &H20
            End If
'            If strLocalLanguage <> "" Then
'                CovertLanguageStageName (LanguageStageCode)
'            End If
            
            Put #lHandle, , LanguageStageCode
            RES.MoveNext
        Loop
    End If
    Close #Fhandle
    Close #lHandle
    RES.Close
End Function
'
'''''''''''''''Dim Fhandle As Integer
'''''''''''''''Dim sHandle As Integer
'''''''''''''''Dim lHandle As Integer
'''''''''''''''Dim Stage As STAGEDETAILS
'''''''''''''''Dim LanguageStageCode As String
'''''''''''''''
'''''''''''''''
'''''''''''''''    RSql = "SELECT StageName,Distance,STG_LOCAL_LANGUAGE FROM STAGE ORDER BY ID"
'''''''''''''''    Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
'''''''''''''''Debug.Print
'''''''''''''''
'''''''''''''''    If RES.RecordCount > 0 Then
'''''''''''''''        If Dir(App.Path & "\STAGE.LST") <> "" Then Kill App.Path & "\STAGE.LST"
'''''''''''''''        If Dir(App.Path & "\LANGUAGE.DAT") <> "" Then Kill App.Path & "\LANGUAGE.DAT"
'''''''''''''''        Fhandle = FreeFile()
'''''''''''''''        Open App.Path & "\STAGE.LST" For Binary Access Write As #Fhandle
'''''''''''''''        lHandle = FreeFile()
'''''''''''''''        Open App.Path & "\LANGUAGE.DAT" For Binary Access Write As #lHandle
'''''''''''''''        RES.MoveFirst
'''''''''''''''        Do While Not RES.EOF
'''''''''''''''            Stage.StageName = Mid$(Trim(RES!StageName), 1, 11) & Chr(0)
'''''''''''''''            Stage.Distance = RES!Distance
'''''''''''''''            Put #Fhandle, , Stage
'''''''''''''''            If Not IsNull(RES!STG_LOCAL_LANGUAGE) Then '
'''''''''''''''                If RES!STG_LOCAL_LANGUAGE <> "" Then
'''''''''''''''                    LanguageStageCode = RES!STG_LOCAL_LANGUAGE
'''''''''''''''                    If LanguageStageCode = "-40" Then LanguageStageCode = "20-20-20"
'''''''''''''''                    ''CovertLanguageStageName (LanguageStageCode)  i iam removing this code to support native language
'''''''''''''''                    ''for english  ---------and adding the following
'''''''''''''''                           Dim i As Long
'''''''''''''''
'''''''''''''''                           For i = 0 To UBound(LanguageStage)
'''''''''''''''                           LanguageStage(i) = &H0
'''''''''''''''
'''''''''''''''                           Next
''''''''''''''''                Else
''''''''''''''''                    LanguageStageCode = &H20
'''''''''''''''                End If
'''''''''''''''            End If
'''''''''''''''            Put #lHandle, , LanguageStage
'''''''''''''''            RES.MoveNext
'''''''''''''''        Loop
'''''''''''''''    End If
'''''''''''''''    Close #Fhandle
'''''''''''''''    Close #lHandle
'''''''''''''''    RES.Close
'End Function


'''''''''''  Modified  On  : 20/01/2011
''Public Function CreateStageList() As Boolean                  '''''''' commented by vaisakh 30.06.11
''Dim Fhandle As Integer
''Dim sHandle As Integer
''Dim lHandle As Integer
''Dim Stage As STAGEDETAILS
''Dim LanguageStageCode As String
''Dim BMPStageFile As String
''Dim intBmpHandle As Integer
''Dim intReadBmpHdl As Integer
''Dim bytByte As Byte
''Dim i As Integer
''
''   RSql = "SELECT StageName,Distance,STG_LOCAL_LANGUAGE FROM STAGE ORDER BY ID" ''' 21/01/2011
''''    RSql = "SELECT StageName,Distance,STG_LOCAL_LANGUAGE,BmpFile FROM STAGE ORDER BY ID"
''    Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
''
''
''    If RES.RecordCount > 0 Then
''
''
''        If Dir(App.Path & "\STAGE.LST") <> "" Then Kill App.Path & "\STAGE.LST"
''
''        If Dir(App.Path & "\LANGUAGE.DAT") <> "" Then Kill App.Path & "\LANGUAGE.DAT"
'''''        If Dir(App.Path & "\GRSTAGE.LST") <> "" Then Kill App.Path & "\GRSTAGE.LST"    ''' 20/01/2011
''
''        Fhandle = FreeFile()
''        Open App.Path & "\STAGE.LST" For Binary Access Write As #Fhandle
''
''        lHandle = FreeFile()
''        Open App.Path & "\LANGUAGE.DAT" For Binary Access Write As #lHandle
''
'''''        intBmpHandle = FreeFile()    ''' 20/01/2011
'''''        Open App.Path & "\GRSTAGE.LST" For Binary Access Write As #intBmpHandle
''
''        RES.MoveFirst
''
''        Do While Not RES.EOF
''
''            Stage.StageName = Mid$(Trim(RES!StageName), 1, 11) & Chr(0)
''            Stage.Distance = Format(val(IIf(IsNull(RES!Distance), 0, RES!Distance)), "0.00")
''            Put #Fhandle, , Stage
''
''            If RES!STG_LOCAL_LANGUAGE <> "" Then
''                LanguageStageCode = RES!STG_LOCAL_LANGUAGE
''            Else
''                LanguageStageCode = &H20
''            End If
''            If strLocalLanguage <> "" Then
''                CovertLanguageStageName (LanguageStageCode)
''            End If
''            Put #lHandle, , LanguageStage
''            Debug.Print LanguageStage
            
            ''''''''''''''''  modified on : 20/01/2011
'''            If RES!BmpFile <> "" Or RES!BmpFile <> "" Then
'''                BMPStageFile = RES!BmpFile
'''            Else
'''                BMPStageFile = "DEFAULT.bmp"
'''
'''            End If
            
'''            If strLocalLanguage <> "" Then
'''
'''                If Dir(App.Path & "\pic\" & BMPStageFile) <> "" Then
'''                    intReadBmpHdl = FreeFile
'''                    Open App.Path & "\pic\" & BMPStageFile For Binary Access Read As #intReadBmpHdl
'''                        Do While Not EOF(intReadBmpHdl)
'''                            bytByte = 0
'''                            Get #intReadBmpHdl, , bytByte
'''                            If EOF(intReadBmpHdl) Then Exit Do
'''                            Put #intBmpHandle, , bytByte
'''                        Loop
'''                    Close #intReadBmpHdl
'''                Else
'''                    For i = 1 To 1598  ''   3134
'''                        bytByte = &HFF
'''                        Put #intBmpHandle, , bytByte
'''                    Next i
'''                End If
'''
'''            Else
'''
'''                For i = 1 To 1598  ''  3134
'''                    bytByte = &HFF
'''                    Put #intBmpHandle, , bytByte
'''                Next i
'''
'''            End If
            ''''''''''''''''''''''''''''''''''
            
'''            RES.MoveNext                     ''''''''' commented by vaisakh on 30.06.11
'''        Loop
'''        If Fhandle > 0 Then Close #Fhandle
'''        If lHandle > 0 Then Close #lHandle
''''''        If intBmpHandle > 0 Then Close #intBmpHandle
'''    End If
'''    RES.Close
'''End Function


Public Function CreateRoute() As Boolean
Dim RouteCode(500) As String
Dim I As Integer
Dim Handle As Integer
Dim filename As String
Dim FARE As Single
Dim Count As Integer
Dim NumOFStages As Integer
Dim NumOFStagesEntry As Integer
Dim j As Integer
Dim FareType(500) As Byte
Dim Nosg(500) As Byte
Dim NoOfDupFare(500) As Byte
Dim RTEST As Route
    'RSql = "SELECT rutcode,faretype, NOSTAGE FROM route"
    RSql = "SELECT rutcode,faretype, NOSTAGE FROM route order by rutcode"
    Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
    If RES.RecordCount > 0 Then
        I = 0
        RES.MoveFirst
        While Not RES.EOF
            RouteCode(I) = TrimChr(RES!RUTCODE)
            FareType(I) = CByte(RES!FareType)
            Nosg(I) = CByte(RES!nostage)
            RES.MoveNext
            I = I + 1
        Wend
        RES.Close
    End If
    I = 0
        filename = App.Path & "\RTE.dat"
        Handle = FreeFile()
            If Dir(filename) <> "" Then Kill filename
            Open filename For Binary Access Write As #Handle
    Do While RouteCode(I) <> ""
'        Filename = App.Path & "\" & Format(RouteCode(i), "000") & ".dat"
'        Handle = FreeFile()
             RTEST.RouteCode = Trim(RouteCode(I)) & Chr(0)
             RTEST.FareType = CByte(FareType(I))
             RTEST.NOS = CByte(Nosg(I))
             RTEST.NoOfDupFare = 0
                
                
        If FareType(I) = 2 Then
                
            RSql = "SELECT DISTINCT FARE FROM FARE WHERE ROUTE = '" & RouteCode(I) & "'" & "ORDER BY FARE"
            Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
            If RES.RecordCount > 0 Then
                RES.MoveFirst
                RES.MoveLast
                RTEST.NoOfDupFare = 0 'CByte(RES.RecordCount + 1)
            Else
                MsgBox "No Fare Details Found!" & vbCrLf & "Going to Exit Program...", vbCritical, "ROUTE"
                End
            End If
            RES.Close
        End If
             
             Put #Handle, , RTEST
    
        
        
        If FareType(I) = 2 Then
        
            
            RSql = "SELECT NOSTAGE FROM ROUTE WHERE RUTCODE = '" & RouteCode(I) & "'"
            Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
            
            NumOFStages = val(RES!nostage)
            RES.Close
            
            
            RSql = "SELECT FARE FROM FARE WHERE ROUTE = '" & RouteCode(I) & "' ORDER BY NUMBER ASC"
            Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
            
            NumOFStagesEntry = (NumOFStages * (NumOFStages - 1)) / 2
            FARE = 0
            For j = 1 To NumOFStagesEntry
                FARE = CSng(val(RES!FARE))
                Put #Handle, , FARE
                RES.MoveNext
            Next
            
            RES.Close
        
'            RSql = "SELECT DISTINCT FARE FROM FARE WHERE ROUTE = '" & RouteCode(I) & "'" & "ORDER BY FARE"
'            Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
'
'            RES.MoveFirst
'            Count = 0
'            FARE = 0
'            Put #Handle, , FARE
'            Do While Not RES.EOF
'                FARE = CSng(val(RES!FARE))
'                Put #Handle, , FARE
'                RES.MoveNext
'            Loop
'            RES.Close
            
            RSql = "SELECT ID FROM STAGE WHERE ROUTE = '" & RouteCode(I) & "'" & " ORDER BY ID"
            Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
            
            RES.MoveFirst
            
            Do While Not RES.EOF
                Count = CInt(RES!Id)
                Put #Handle, , Count
                RES.MoveNext
            Loop
            
            RES.Close
        
        Else
        
            RSql = "SELECT FARE FROM FARE WHERE ROUTE = '" & RouteCode(I) & "'" & " ORDER BY FARE"
            Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
            If RES.RecordCount > 0 Then
                RES.MoveFirst
                Do While Not RES.EOF
                    FARE = CSng(val(RES!FARE))
                    Put #Handle, , FARE
                    RES.MoveNext
                Loop
                
                RES.Close
                RSql = "SELECT ID FROM STAGE WHERE ROUTE = '" & RouteCode(I) & "'" & " ORDER BY ID"
                Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
                
                RES.MoveFirst
                
                Do While Not RES.EOF
                    Count = CInt(RES!Id)
                    Put #Handle, , Count
                    RES.MoveNext
                Loop
                
                RES.Close
            
            End If
        End If
        I = I + 1
    Loop
            Close #Handle
End Function

Public Function RouteExists(RouteCode As String) As Boolean
On Error GoTo err
    RouteExists = True
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    RSql = "SELECT RUTCODE FROM ROUTE WHERE RUTCODE = '" & RouteCode & "'"
    Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
    If RES.RecordCount = 0 Then RouteExists = False
    Exit Function
err:
    Select Case err.Number
        Case Else
            MsgBox "Error No : " & err.Number & vbCrLf & err.Description, vbInformation, "Route"
            Exit Function
    End Select
End Function
Public Function TextBoxValidity(Key As Integer) As Integer
Dim st As String
  Key = Asc(UCase(Chr(Key)))
    st = "`~!@#$%^&*()_-+=/?<>;', .][\\:{}|"
    TextBoxValidity = InStr(st, Chr(Key))
End Function

Public Function TextBoxPalmValidity(Key As Integer) As Integer
Dim st As String
    Key = Asc(UCase(Chr(Key)))
    st = "`~!@#$%^&*()_-+=/?<>;',. ][\\:{}|"""
    TextBoxPalmValidity = InStr(st, Chr(Key))
End Function

Public Function TextBoxValidityNumeric(Key As Integer) As Integer
Dim st As String
    Key = Asc(UCase(Chr(Key)))
    st = "ABCDEFGHIJKLMNOPQRSTUVWXYZ`~!@#$%^&*()_-+=/?\[{]}:<>;|', """
    TextBoxValidityNumeric = InStr(st, Chr(Key))
End Function
Public Function TextBoxValidityonlyNumeric(Key As Integer) As Integer
Dim st As String
    Key = Asc(UCase(Chr(Key)))
    st = "ABCDEFGHIJKLMNOPQRSTUVWXYZ`~!@#$%^&*()_-+=/?\[{]}:<>;|',. "
    TextBoxValidityonlyNumeric = InStr(st, Chr(Key))
End Function

Public Function CloseProgram(ByVal caption As String)
    Dim Handle As Long
    Handle = FindWindow(vbNullString, caption)
    If Handle = 0 Then Exit Function
    SendMessage Handle, WM_CLOSE, 0&, 0&
End Function

Public Function Change_System_Date()
    Dim dwLCID As Long
    dwLCID = GetSystemDefaultLCID()
    If SetLocaleInfo(dwLCID, LOCALE_SLONGDATE, "dd/MM/yyyy") _
       = False Then
       MsgBox "System Date  Formatting Changing Failed" & vbCrLf & "Do it manually..."
       Exit Function
    End If
    If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, "dd/MM/yyyy") _
       = False Then
       MsgBox "System Date  Formatting Changing Failed" & vbCrLf & "Do it manually..."
       Exit Function
    End If
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
 End Function


Public Function FindSysDir()
Dim Buf As String * 64
Dim Gpos As Integer
    
    GetSystemDirectory Buf, Len(Buf)
    Gpos = InStr(Buf, Chr(0))
    If Gpos > 0 Then Buf = Mid(Buf, 1, Gpos - 1)
    SYSDIR = Buf
    Buf = ""
    GetWindowsDirectory Buf, Len(Buf)
    Gpos = InStr(Buf, Chr(0))
    If Gpos > 0 Then Buf = Mid(Buf, 1, Gpos - 1)
    WINPATH = Buf
    If Dir(Trim(SYSDIR) & "\Calc.exe") <> "" Then
        SYSDIR = Trim(SYSDIR) & "\Calc.exe"
        WINPATH = ""
    End If
    If Dir(Trim(WINPATH) & "\calc.exe") <> "" Then
        WINPATH = Trim(WINPATH) & "\calc.exe"
        SYSDIR = ""
    End If
End Function


Public Function CONNECT_DB()
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
End Function

Public Function TrimChr(ByVal Buf As String) As String
    Dim tmp As Integer
    Dim Buf1 As String
    tmp = InStr(Buf, Chr$(0))
    If tmp > 0 Then
        TrimChr = Trim(Mid(Buf, 1, tmp - 1))
    Else
        TrimChr = Trim(Buf)
    End If
    'TrimChr = left(Buf, InStr(1, strString, Chr$(0)) - 1) 'Mid(Buf, 1, (InStr(Buf, Chr(0)) - 1))
End Function

Public Function LPad(ByVal strData As String, ByVal strFill As String, ln As Integer) As String
    Dim strTemp As String
    Dim I As Integer
    For I = 1 To Len(strData)
        If Mid$(strData, I, 1) <> vbNullChar Then
            strTemp = strTemp & Mid$(strData, I, 1)
        End If
    Next I
    strTemp = String$(ln, strFill) & Trim$(strTemp)
    strTemp = Right$(strTemp, ln)
    LPad = strTemp
End Function
Public Function RPad(ByVal strData As String, strFill As String, ln As Integer)
  Dim intLen As Integer
    intLen = Len(strData)
    If ln > intLen Then
        strData = strData & String$(ln, strFill)
        strData = Left$(strData, ln)
        RPad = strData
    Else
        RPad = Left$(strData, ln)
    End If
End Function

Public Function Convert_Report() As Boolean
Dim Handle As Integer
Dim Rept As PReport
Handle = FreeFile()
Open App.Path & "\RPT01.DAT" For Binary Access Read As #Handle
    Do
        Get #Handle, , Rept
        If EOF(Handle) Then Exit Do
    Loop
Close #Handle

End Function


Public Function ProjectValidity() As Boolean
On Error Resume Next
Dim sSytemDir As String * 64
Dim Gpos As Integer
Dim iHandle As Integer
Dim iMaximumTrial As Integer

    ProjectValidity = True
    GetSystemDirectory sSytemDir, Len(sSytemDir)
    Gpos = InStr(sSytemDir, Chr(0))
    If Gpos > 0 Then sSytemDir = Mid(sSytemDir, 1, Gpos - 1)
    SYSDIR = sSytemDir
    
    If Dir(App.Path & "\License.dat") <> "" Then
        iHandle = FreeFile()
        Open App.Path & "\License.dat" For Binary Access Read As #iHandle
            Get #iHandle, , PrjValidity
        Close #iHandle
        iMaximumTrial = PrjValidity.MaximumTrialDays
    Else
        iMaximumTrial = NOOFDAYS_VALIDITY
    End If
    
    iHandle = FreeFile()
    If Dir(Trim(SYSDIR) & "\slic.sys") <> "" Then
        Open Trim(SYSDIR) & "\slic.sys" For Binary Access Read As #iHandle
            Get #iHandle, , PrjValidity
        Close #iHandle
'        If iMaximumTrial > PrjValidity.MaximumTrialDays Then PrjValidity.ExpiredFlag = False
        
        If PrjValidity.TrialFlag = True Then
            If PrjValidity.ExpiredFlag = True Then
                MsgBox "Palmtec Amphibia Bus Ticketing Utility" & vbCrLf & _
                       "Trial Version has been Expired!" & vbCrLf & _
                       "For Future use please purchase our orginal version" & vbCrLf & vbCrLf & _
                       "Please contact us" & vbCrLf & vbCrLf & _
                       "SOFTLAND INDIA LTD" & vbCrLf & _
                       "KINFRA SMALL INDUSTRIES PARK" & vbCrLf & _
                       "MENAMKULAM, THUMBA" & vbCrLf & _
                       "TRIVANDRUM-695586" & vbCrLf & _
                       "PH: 91-471-2704090" & vbCrLf & _
                       "email: info@softlandindia.co.in", vbOKOnly, "Palmtec Amphibia Bus Ticketing"
                ProjectValidity = False
                Exit Function
            End If
            If ValidateDate = True Then
                PrjValidity.ValidityCount = PrjValidity.ValidityCount + 1
                PrjValidity.LastUsedDate = Date
            End If
            
            If PrjValidity.ValidityCount >= iMaximumTrial Then
                PrjValidity.ExpiredFlag = True
            End If
            
            iHandle = FreeFile()
            Open Trim(SYSDIR) & "\slic.sys" For Binary Access Write As #iHandle
                Put #iHandle, , PrjValidity
            Close #iHandle
        End If
    Else
        PrjValidity.ProjectStartDate = Date
        PrjValidity.LastUsedDate = Date
        PrjValidity.ValidityCount = 0
        PrjValidity.ExpiredFlag = False
        PrjValidity.TrialFlag = True
        Open Trim(SYSDIR) & "\slic.sys" For Binary Access Write As #iHandle
            Put #iHandle, , PrjValidity
        Close #iHandle
    End If

End Function

Public Function ValidateDate() As Boolean
On Error Resume Next
    
    ValidateDate = False
    If Day(PrjValidity.LastUsedDate) = Day(Date) Then
        If Month(PrjValidity.LastUsedDate) = Month(Date) Then
            If Year(PrjValidity.LastUsedDate) = Year(Date) Then
                If PrjValidity.ValidityCount = 0 Then ValidateDate = True
            Else
                ValidateDate = True
            End If
        Else
            ValidateDate = True
        End If
    Else
        ValidateDate = True
    End If

End Function

Public Function RemoveLanguageStage(Route As String)
On Error GoTo err
    Dim fname As String
    Dim tname As String
    Dim FHndl As Integer
    Dim THndl As Integer
    Dim StgCount As Integer
    Dim RouteCode As String
    Dim sSQL As String
    Reset
    fname = App.Path & "\LOCAL_LANGUAGE.DAT"
    tname = App.Path & "\TEMP.DAT"
    If Dir(tname, vbNormal) <> "" Then Call Kill(tname)
    If Dir(fname, vbNormal) <> "" Then
        FHndl = FreeFile()
        Open fname For Binary Access Read As #FHndl
        THndl = FreeFile()
        Open tname For Binary Access Write As #THndl
        Do While Not EOF(FHndl)
            Get #FHndl, , LSTAG
            RouteCode = Mid(LSTAG.RouteCode, 1, InStr(1, LSTAG.RouteCode, Chr(0)) - 1)
            
            sSQL = "SELECT count(*) FROM ROUTE WHERE rutcode='" & Route & "'"
            Set RES = TDB.OpenRecordset(sSQL, dbOpenDynaset)
            If RES.Fields(0) > 0 Then
                If RouteCode <> Route Then
                    If LSTAG.LocalLanguageStageName(0) <> &H0 Then
                        Put #THndl, , LSTAG
                    End If
                End If
            End If
            RES.Close
        Loop
        Close #FHndl
        Close #THndl
        Kill fname
        THndl = FreeFile()
        Open tname For Binary Access Read As #THndl
        FHndl = FreeFile()
        StgCount = 0
        Open fname For Binary Access Write As #FHndl
            Do While Not EOF(THndl)
                Get #THndl, , LSTAG
                LSTAG.stagecode = StgCount
                If LSTAG.LocalLanguageStageName(0) <> &H0 Then
                    Put #FHndl, , LSTAG
                    StgCount = StgCount + 1
                End If
            Loop
            Close #THndl
        Close #FHndl
    End If
    Exit Function
err:
    MsgBox "Language Stage Coneversion Error!", vbInformation, "BUS"
End Function

Public Function CovertLanguageStageName(ByVal sStageName As String)
On Error GoTo err
Dim TempStage(23) As Byte
Dim I As Byte, j As Byte
Dim str As String
Dim DB_SIN As DAO.Database
Dim REC_SIN As DAO.Recordset
Dim SQL_SIN As String
Dim ChrWdth As Byte
Dim PixelCount As Integer
Dim PixelToFill As Integer
Dim SpaceToFill As Byte
Dim ChrCount As Byte
Dim cbyte1 As Byte
    
    ChrWdth = 0
    
    PixelCount = 0
    ChrCount = 0
    Set DB_SIN = DAO.OpenDatabase(App.Path & "\HCHAR.mdb", dbDriverComplete, False, ";UID=;PWD=siljvvnl")
    For I = 0 To UBound(LanguageStage)
       LanguageStage(I) = &H0
       TempStage(I) = &H0
    Next
    If sStageName = "" Then sStageName = "20-20-20-20-20-20"
    sStageName = sStageName & "-"
    
    
    For I = 0 To 22
        If Len(Mid(sStageName, 1, InStr(1, sStageName, "-"))) = 0 Then Exit For
        str = Mid$(sStageName, 1, 2)
        SQL_SIN = "SELECT WIDTH FROM " & strLocalLanguage & " WHERE ISCII='" & str & "'"
        Set REC_SIN = DB_SIN.OpenRecordset(SQL_SIN, dbOpenDynaset)
        If REC_SIN.RecordCount > 0 Then
            ChrWdth = ChrWdth + val(REC_SIN!Width)
        ElseIf str = "20" Then '''''''''''''
            ChrWdth = ChrWdth + val(16) '''
        Else
            ChrWdth = ChrWdth + 6
        End If
        REC_SIN.Close
        ChrCount = ChrCount + 1
        sStageName = Mid$(sStageName, InStr(1, sStageName, "-") + 1)
        cbyte1 = CByte("&H" & Mid(str, 1, 1))
        TempStage(I) = ((cbyte1 * 16) + CByte("&H" & Mid(str, 2, 1)))
    Next

    PixelCount = ChrWdth
    If PixelCount < 166 And PixelCount <> 0 Then
        PixelToFill = ((172 - PixelCount) / 6) - 1
        If PixelToFill > 16 Then
            PixelToFill = 16
        End If
        If PixelToFill Mod 2 <> 0 Then PixelToFill = PixelToFill + 1
    End If
    SpaceToFill = PixelToFill / 2
    If SpaceToFill = 0 Then SpaceToFill = 1
    For I = 0 To SpaceToFill - 1
        LanguageStage(I) = &H20
    Next
'    For i = 0 To SpaceToFill - 1  'vaisakh 29.06.11
'        LanguageStage(i) = &H0
'    Next
    
    For j = 0 To ChrCount - 1
        LanguageStage(I) = TempStage(j)
        I = I + 1
        If I >= 23 Then Exit For
    Next
    If I < 22 Then LanguageStage(I) = &H20
    Exit Function
err:
    MsgBox "Error!" & vbCrLf & err.Number & " : " & err.Description, vbInformation, "BUS"
End Function


Public Function ShiftLeft(ByVal bByte As Byte)
Dim tByte As Byte
    
    tByte = bByte And &HF
    
End Function

Public Function SetLocalLanguage(ByVal LanguageIndex As Byte)
Dim FieldIndex As Byte
    Set DB = DAO.OpenDatabase(App.Path & "\HCHAR.MDB", dbDriverComplete, False, ";UID=;PWD=siljvvnl")
    Set RES = DB.OpenRecordset("SELECT * FROM LANGUAGE_ENABLED", dbOpenDynaset)
     
    If RES.RecordCount > 0 Then
        RES.Edit
        For FieldIndex = 1 To 4
            RES.Fields(FieldIndex) = False
        Next
        If LanguageIndex > 0 Then
            If LanguageIndex = 1 Then '
            RES.Fields(2) = True 'tamil
            ElseIf LanguageIndex = 2 Then
            RES.Fields(1) = True 'malyalam
            ElseIf LanguageIndex = 3 Then
            RES.Fields(3) = True 'sinhala
            ElseIf LanguageIndex = 4 Then
            RES.Fields(4) = True 'hindi
            End If
            ''REFER DATABSE STRUCTURE FOR CLARIFICATIONS
           '' RES.Fields(LanguageIndex) = True
        End If
        RES.Update
        LocalLanguage = LanguageIndex
    End If
    RES.Close
    DB.Close
End Function

Public Function GetLocalLanguage() As Byte
    Set DB = DAO.OpenDatabase(App.Path & "\HCHAR.MDB", dbDriverComplete, False, ";UID=;PWD=siljvvnl")
    Set RESLaN = DB.OpenRecordset("SELECT * FROM LANGUAGE_ENABLED", dbOpenDynaset)
    
    GetLocalLanguage = 0
     
    If RESLaN.RecordCount > 0 Then
        If RESLaN!MALAYALAM = True Then
            LocalLanguage = 2
            strLocalLanguage = "MALAYALAM"
        ElseIf RESLaN!TAMIL = True Then
            LocalLanguage = 1
            strLocalLanguage = "TAMIL"
        ElseIf RESLaN!HINDI = True Then
            LocalLanguage = 3
            strLocalLanguage = "HINDI"
        ElseIf RESLaN!SINHALA = True Then
            LocalLanguage = 2
            strLocalLanguage = "SINHALA"
        Else
            LocalLanguage = 0
            strLocalLanguage = ""
        End If
        GetLocalLanguage = LocalLanguage
    End If
    RESLaN.Close
    DB.Close
End Function

Public Function CreateFolder(Path As String) As String
Dim fsoFolder As Object, fldFolder As Object
Dim strTemp As String, strErrTemp As String
Dim intCount As Integer
On Error GoTo CatchError
    strTemp = Path
    Set fsoFolder = CreateObject("Scripting.FileSystemObject")
    If fsoFolder.FolderExists(Path) = False Then
        strErrTemp = Path
        strTemp = strTemp & "\"
        For intCount = 1 To Len(strTemp)
            Path = Mid(strTemp, 1, InStr(intCount, strTemp, "\"))
            If fsoFolder.FolderExists(Path) = False Then
                Set fldFolder = fsoFolder.CreateFolder(Path)
            Else
                Set fldFolder = fsoFolder.GetFolder(Path)
            End If
            strErrTemp = fldFolder.Path
            intCount = InStr(intCount, strTemp, "\")
        Next intCount
    End If
    Set fldFolder = fsoFolder.GetFolder(Path)
    CreateFolder = fldFolder.Path
    Exit Function
CatchError:
    If strErrTemp = "" Then
       CreateFolder = App.Path
    Else
        CreateFolder = strErrTemp
    End If
End Function


Function ConnectDatabase(ByRef ConnectionObject As ADODB.Connection, _
                         ByVal Database As String, _
                         Optional ByVal DatabasePassword As String) As Boolean
On Error Resume Next
    If ConnectionObject.State = adStateOpen Then ConnectionObject.Close
    ConnectionObject.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database & ";Persist Security Info=False;Jet OLEDB:Database Password=" & DatabasePassword
    ConnectionObject.Open
    If ConnectionObject.State = adStateOpen Then ConnectDatabase = True
    Exit Function
CatchError:
End Function

Public Function checkfooter() As String
Dim strpath As String
Dim strBuffer As String * 250
    Call GetSystemDirectory(strBuffer, Len(strBuffer))
    strBuffer = Mid(strBuffer, 1, InStr(1, strBuffer, Chr(0)) - 1)
    strpath = Trim(strBuffer)
    If Dir(strpath & "\PatchFtr.sys", vbHidden + vbSystem) <> "" Then
        checkfooter = strpath & "\PatchFtr.sys"
    Else
        checkfooter = ""
    End If
End Function

Public Function ftreditvalue() As String
Dim iHandle As Integer
Dim temp As Footer
Dim strpath As String
    strpath = checkfooter
    If strpath <> "" Then
            iHandle = FreeFile()
            Open strpath For Binary Access Read As #iHandle
            Get #iHandle, , temp
            Close #iHandle
            ftreditvalue = TrimChr(temp.FooterString)
            HStr.ReportFlag = (HStr.ReportFlag And 191)
    Else
        HStr.ReportFlag = (HStr.ReportFlag Or 64)
    End If
End Function
Public Sub CheckLoginStatus()
Dim sql As String
Dim rs As DAO.Recordset
On Error GoTo err
    Set DB = DAO.OpenDatabase(App.Path & "\PVT.MDB", dbDriverComplete, False, ";UID=;PWD=silbus")
    sql = "SELECT * FROM LOGINTABLE "
    Set rs = DB.OpenRecordset(sql, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs!SUPERUSER = "1"
        rs!SuperPassword = "1"
        rs.Update
        rs.Close
        Exit Sub
    End If
    rs.Close
err:
End Sub
Sub Main()
 
    
'     MsgBox Format(Time, " HH:NN:SS AM/PM")
'     MsgBox Format(Time, " HH:MM:SS AM/PM")
    WriteVersionInfo
    Call CheckLoginStatus

    loginform.Show
End Sub

Public Function dbexpdet() As Boolean
Dim hdl As Integer
Dim exp As EXPENSEDET

filename = App.Path & "\EXPENSEDET.DAT"
hdl = FreeFile()
If Dir(filename) <> "" Then Kill filename
Open filename For Binary Access Write As #hdl

RSql = "select * from EXPMASTER"
Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
If RES.RecordCount > 0 Then
    RES.MoveFirst
    Do While Not RES.EOF
    If RES.EOF = True Then Exit Do
            exp.ucType = TrimChr(RES!EXP_CODE) & Chr(0)
            exp.expname = TrimChr(RES!exp_name) & Chr(0)
    Put #hdl, , exp
        RES.MoveNext
Loop
  
End If
  RES.Close
Close #hdl
End Function

Public Function dbcrew() As Boolean
Dim hdl As Integer
Dim Crew As CREWDET

filename = App.Path & "\CREW.DAT"
hdl = FreeFile()
If Dir(filename) <> "" Then Kill filename


RSql = "select * from CREW"
Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
If RES.RecordCount > 0 Then
    Open filename For Binary Access Write As #hdl
    RES.MoveFirst
    Do While Not RES.EOF
        If RES.EOF = True Then Exit Do
                    Crew.EmpId = TrimChr(RES!EMPLOYEEID) & Chr(0)
                    Crew.EmpName = TrimChr(RES!EMPLOYEENAME) & Chr(0)
                    Crew.EmpType = TrimChr(RES!EmployeeTypeId) & Chr(0)
                    Crew.PassWord = IIf(IsNull(RES!PSWD), "111111", RES!PSWD) & Chr(0)
'                    Crew.DrvrID = TrimChr(RES!DR_ID) & Chr(0)
'                    Crew.Driver = TrimChr(RES!Dr_Name) & Chr(0)
'                    Crew.CndrID = TrimChr(RES!CDTR_ID) & Chr(0)
'                    Crew.Conductor = TrimChr(RES!cdtr_name) & Chr(0)
'                    Crew.ClnrID = TrimChr(RES!CLNR_ID) & Chr(0)
'                    Crew.Cleaner = TrimChr(RES!CLNR_NAME) & Chr(0)
'                    Crew.BusNo = TrimChr(RES!bus_no) & Chr(0)
'                    Crew.BusTypeName = TrimChr(RES!BusTypeName) & Chr(0)
'                    Crew.BusId = TrimChr(RES!BUSTYPEID) & Chr(0)
                    
                    Put #hdl, , Crew
        RES.MoveNext
    Loop
End If
RES.Close
Close #hdl
End Function
Public Function dbvehicle() As Boolean
Dim hdl As Integer
Dim vhcl As VEHICLE

filename = App.Path & "\VEHICLE.DAT"
hdl = FreeFile()
If Dir(filename) <> "" Then Kill filename

RSql = "select * from VEHICLETYPE"
Set RES = DB.OpenRecordset(RSql, dbOpenDynaset)
If RES.RecordCount > 0 Then
    Open filename For Binary Access Write As #hdl
    RES.MoveFirst
    Do While Not RES.EOF
        If RES.EOF = True Then Exit Do
                    vhcl.BUSID = TrimChr(RES!BUSID) & Chr(0)
                    vhcl.Busno = TrimChr(RES!Busno) & Chr(0)
                    Put #hdl, , vhcl
        RES.MoveNext
    Loop
End If
RES.Close
Close #hdl
End Function

