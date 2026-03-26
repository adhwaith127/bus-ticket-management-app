VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   720
      Top             =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str As String
Public condb As New ADODB.Connection
Public adoDBCon As New ADODB.Connection

Public Function getvalueQuery1(strSql As String) As String
'Author Mubeena
On Error GoTo lblErr
Dim oRS As New ADODB.Recordset
Dim QryValue As String
  Set oRS = condb.Execute(strSql)
  If oRS.EOF = False Then
        QryValue = oRS.Fields(0)
    Else
        QryValue = ""
    End If
    Set oRS = Nothing
    getvalueQuery1 = QryValue
    Exit Function
lblErr:
'MsgBox err.Description, vbOKOnly, prjTitle
QryValue = ""
End Function

Public Function ConnectMysqlDatabase() As Boolean
On Error GoTo CatchError
Dim intHandle As Integer
Dim strDatabase As String
Dim strUname As String
Dim strpswd, strHost, strString As String
Dim Fso As FileSystemObject
Dim FStream As TextStream
Set Fso = New FileSystemObject
''DoEvents
     If Dir(App.Path & "\HostName.ini", vbNormal) <> "" Then
           
         Set FStream = Fso.OpenTextFile(App.Path & "\HostName.ini", ForReading, False)
  
   
        If FStream.AtEndOfStream = False Then
          
            Dim strData() As String
            strString = FStream.ReadLine
            strData = Split(strString, ",")
            strHost = strData(0)
            strDatabase = strData(1)
            strUname = strData(2)
            strpswd = strData(3)
            Close #intHandle
        End If
    Else
        strString = "localhost"
        strHost = strString
    End If
    
    If adoDBCon.State <> adStateClosed Then adoDBCon.Close
    adoDBCon.CommandTimeout = 30
    adoDBCon.ConnectionTimeout = 30
    adoDBCon.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DB=" & strDatabase & ";SERVER=" & strHost & ";UID=" & strUname & ";PASSWORD=" & strpswd & ";PORT=3306;SOCKET=;OPTION=;STMT=;"
    'adoDBCon.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};DESC=;DB=" & strDatabase & ";SERVER=" & strHost & ";UID=" & strUname & ";PASSWORD=" & strPswd & ";PORT=3306;SOCKET=;OPTION=;STMT=;"
    'adoDBCon.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DB=frcs;SERVER=localhost;UID=root;PASSWORD=root;PORT=3306;SOCKET=;OPTION=;STMT=;"
    
    'DoEvents
    adoDBCon.Open
    'DoEvents
    If adoDBCon.State = adStateOpen Then
        ConnectMysqlDatabase = True
    End If
   ' DoEvents
    Exit Function
CatchError:
  ' Call WriteErrorLog("App.path/log.txt", "clsDBInfo", "ConnectDatabase", err.Number, err.Description)
    On Error Resume Next
    If adoDBCon.State = adStateOpen Then
        ConnectMysqlDatabase = True
    Else
       ConnectMysqlDatabase = False
    End If
       
        
End Function

Public Function CONNECTDB1()
    If condb.State <> 1 Then
        condb.Provider = "Microsoft.Jet.OLEDB.4.0"
        condb.ConnectionString = "Data Source=" & App.Path & "\Pvt.mdb"
        condb.Properties("Jet OLEDB:Database Password") = "silbus"
        condb.Open
    Else
'        condb.Close
'        condb.Provider = "Microsoft.Jet.OLEDB.4.0"
'        condb.ConnectionString = "Data Source=" & App.Path & "\Pvt.mdb"
'        condb.Properties("Jet OLEDB:Database Password") = "silbus"
'        condb.Open
    End If
End Function



Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
Dim sql As String
Dim exdesc, expname As String, sqlExp As String
Dim rs2 As New ADODB.Recordset
Dim rslogin1 As New ADODB.Recordset, rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
 If ConnectMysqlDatabase = True Then
    CONNECTDB1
 sql = "select * from ticketdetails"
 If rslogin1.State = 1 Then rslogin1.Close
   rslogin1.Open sql, adoDBCon, adOpenDynamic, adLockOptimistic
sql = ""
sql = "select * from TKTS"
   rs.Open sql, condb, adOpenDynamic, adLockOptimistic
   Do While Not rslogin1.EOF
   If rs1.State = 1 Then rs1.Close
     str = "select * from TKTS where Date='" & Format(rslogin1!TDate, "dd/mm/yyyy") & "' and PalmId='" & rslogin1!PalmID & "' and Schdule=" & rslogin1!ScheduleNo & " and  TripNo=" & rslogin1!TripNo & " and TicketNo=" & rslogin1!TicketNo & " and Amount=" & rslogin1!Amount & ""
      rs1.Open str, condb, adOpenDynamic, adLockOptimistic
      If rs1.EOF = True Then
        rs.AddNew
        rs!Date = Format(rslogin1!TDate, "dd/mm/yyyy")
        rs!Time = rslogin1!TTime
        rs!PalmID = rslogin1!PalmID
        rs!Schdule = rslogin1!ScheduleNo
        rs!TripNo = rslogin1!TripNo
        rs!TicketNo = rslogin1!TicketNo
        rs!Amount = rslogin1!Amount
        rs!Luggage = rslogin1!Luggcnt
        rs!FromStage = rslogin1!FromStage
        rs!ToStage = rslogin1!ToStage
        rs!Full = rslogin1!Fullcnt
        rs!Half = rslogin1!Halfcnt
        rs!st = rslogin1!Stcnt
        rs!Phy = rslogin1!phycnt
        rs!PassNo = rslogin1!PassNo
       rs.Update
       
    End If
    rs1.Close
       rslogin1.MoveNext
 '  If rs.State = 1 Then rs.Close
   Loop
   rslogin1.Close
    If rs.State = 1 Then rs.Close
    If rs1.State = 1 Then rs1.Close
     
    If rslogin1.State = 1 Then rslogin1.Close
    sql = "select * from rpt"
   rslogin1.Open sql, adoDBCon, adOpenDynamic, adLockOptimistic
sql = ""
sql = "select * from rpt"
   rs.Open sql, condb, adOpenDynamic, adLockOptimistic
   Do While Not rslogin1.EOF
   If rs1.State = 1 Then rs1.Close
     str = "select * from rpt where StartDate='" & Format(rslogin1!SDate, "dd/mm/yyyy") & "' and StartTime='" & rslogin1!STime & "' and PalmId='" & rslogin1!PalmID & "' and Schedule=" & rslogin1!SCHEDULE & " and  TripNo=" & rslogin1!TripNo & " and STicket='" & rslogin1!Starttkt & "' and ETicketNo='" & rslogin1!Endtkt & "' and TotalColl=" & rslogin1!NetCol & ""
      rs1.Open str, condb, adOpenDynamic, adLockOptimistic
      If rs1.EOF = True Then
        rs.AddNew
        rs!Date = Format(rslogin1!SDate, "dd/mm/yyyy")
        rs!STicket = rslogin1!Starttkt
        rs!ETicketNo = rslogin1!Endtkt
        rs!TotalColl = rslogin1!NetCol
        rs!FullColl = rslogin1!FullColl
        rs!HalfColl = rslogin1!HalfColl
        rs!PhyColl = rslogin1!PhyColl
        rs!LuggageColl = rslogin1!LugColl
        rs!STColl = rslogin1!StuColl
        rs!AdjustColl = rslogin1!AdjColl
        rs!Fulls = rslogin1!Full
        rs!Half = rslogin1!Half
        rs!Phy = rslogin1!Phy
        rs!Luggage = rslogin1!Lug
        rs!st = rslogin1!stu
        rs!Adjust = rslogin1!Adj
        rs!pass = rslogin1!pass
        rs!SCHEDULE = rslogin1!SCHEDULE
        rs!TripNo = rslogin1!TripNo
        rs!RouteCode = rslogin1!RoutNo
        rs!StartDate = Format(rslogin1!SDate, "dd/mm/yyyy")
        rs!StartTime = rslogin1!STime
        rs!EndDate = Format(rslogin1!EDate, "dd/mm/yyyy")
        rs!EndTime = rslogin1!ETime
        rs!expense = rslogin1!Expns
        rs!PalmID = rslogin1!PalmID
        rs!Busno = rslogin1!Busno
        rs!Driver = rslogin1!Driver
        rs!Conductor = rslogin1!Conductor
        rs!Cleaner = rslogin1!Cleaner
        rs!NoOfMisBill = rslogin1!NoOfMisBill
        rs!InHandAmount = rslogin1!InHandAmt
        rs!Free = rslogin1!Free
        rs!Conc = rslogin1!Conc
        rs!UpDownTrip = rslogin1!UpDownTrip
       rs.Update
      
    End If
'     sqlExp = ""
'    sqlExp = "Select sum(ExpAmt)as EXP from EXPENSE where TripMasterReferenceId='" & getvalueQuery1("Select Trip_Master_ID FROM RPT WHERE PALMID='" & rslogin1!PalmID & "' AND SCHEDULE=  " & rslogin1!ScheduleNo & " AND TRIPNO= " & rslogin1!TripNo & "  AND DATE = DateValue('" & Format(rslogin1!ExpDate, "dd/mm/yyyy") & "') ") & "' "
'
'            sqlExp = "Select sum(ExpAmt)as EXP from EXPENSE where TripMasterReferenceId='" & getvalueQuery1("Select Trip_Master_ID FROM RPT WHERE PALMID='" & rslogin1!PalmID & "'  AND SCHEDULE= " & rslogin1!ScheduleNo & " AND TRIPNO= " & rslogin1!TripNo & " AND DATE = DateValue('" & Format(rslogin1!ExpDate, "dd/mm/yyyy") & "') ") & "' "
'            Set RES8 = CNN.OpenRecordset(sqlExp, dbOpenDynaset)
'            'RES4.Edit
'            If (RES8("EXP")) > 0 Then
'           sql = "UPDATE RPT SET Expense=" & RES8("EXP") & " where Trip_Master_ID=" & getvalueQuery1("Select Trip_Master_ID FROM RPT WHERE PALMID='" & rslogin1!PalmID & "'  AND SCHEDULE= " & rslogin1!ScheduleNo & " AND TRIPNO= " & rslogin1!TripNo & "  AND DATE = DateValue('" & Format(rslogin1!ExpDate, "dd/mm/yyyy") & "') ") & ""
'           'Set RES7 = CNN.OpenRecordset(sql, dbOpenDynaset)
'           CNN.Execute (sql)
'           End If
     rslogin1.MoveNext
     ' If rs.State = 1 Then rs.Close
   Loop
   If rs.State = 1 Then rs.Close
   rslogin1.Close
   
     
    If rslogin1.State = 1 Then rslogin1.Close
    sql = "select * from expense"
   rslogin1.Open sql, adoDBCon, adOpenDynamic, adLockOptimistic
        Do While Not rslogin1.EOF

    If (getvalueQuery1("select Count(*) from EXPENSE where PALMID = '" & rslogin1!PalmID & "' AND rcpt_No= " & rslogin1!rcpt_No & " AND SCHEDULENO = " & rslogin1!ScheduleNo & " AND EXPCODE = '" & rslogin1!ExpCode & "' AND DATE = '" & Format(rslogin1!ExpDate, "dd/mm/yyyy") & "'")) = 0 Then
        
        
        If Trim(rslogin1!ExpCode) <> "1" Then
            
             rs2.Open "select * from EXPMASTER where EXP_CODE=" & rslogin1!ExpCode & "", condb, adOpenDynamic, adLockOptimistic
            expname = rs2!exp_name
            rs2.Close
        Else
            expname = "Diesel Entry"
        End If
        'sql = "insert into EXPENSE (TripMasterReferenceId,ExpCode ,ExpAmt, ExpName, Date, Time, PalmID,ScheduleNo,BusNo,DriverName,rcpt_No) values('"
        sql = "insert into EXPENSE values('" _
        & getvalueQuery1("Select Trip_Master_ID FROM RPT WHERE PALMID='" & rslogin1!PalmID & "' AND SCHEDULE= " & rslogin1!ScheduleNo & " AND TRIPNO= " & rslogin1!TripNo & " AND DATE = DateValue('" & Format(rslogin1!ExpDate, "dd/mm/yyyy") & "')") & "','" _
        & rslogin1!ExpCode & "'," _
        & rslogin1!ExpAmt & ",'" _
        & expname & "','" _
        & Format(rslogin1!ExpDate, "dd/mm/yyyy") & "','" _
        & rslogin1!ExpTime & "','" _
        & rslogin1!PalmID & "'," _
        & rslogin1!ScheduleNo & ",'" _
        & rslogin1!Busno & "','" _
        & rslogin1!DriverName & "'," _
        & rslogin1!rcpt_No & ")"
        condb.Execute sql
  End If
       
  rslogin1.MoveNext
   Loop
   
   If rslogin1.State = 1 Then rslogin1.Close
   
'    MsgBox "Mysql Connection Failed. Going To Abort"
'        End
'        Exit Sub
    End If
End Sub
