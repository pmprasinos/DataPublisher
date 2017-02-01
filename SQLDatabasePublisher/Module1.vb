
Imports System.Data
Imports System.Data.SqlClient
Imports WebfocusDLL
Imports System.Threading.Thread
Imports System.Runtime.CompilerServices
Imports Microsoft.Office.Interop
Imports System.Net
Imports System.Net.NetworkInformation
'Imports ClassLibrary1

Module module1
    Private Downloading As Boolean = False
    Dim LogInInfo As String()
    Dim Fuckups As Integer = 0
    Dim ConnectionString As String = "Server=SLREPORT01; Database=WFLocal; User Id=PrasinosApps; Password=Wyman123-;"
    Private tmp = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\test.temp" 'O.Path.Combine(My.Computer.FileSystem.SpecialDirectories.Temp, "snafu.fubar")
    Public TimeToDownload As String
    Public StartTime As String
    Dim UpdateTimes As Object()()



    Sub Main()
        Try

            NotificationEmails()

            'If My.Settings.NeedsUpdate Then
            '    My.Settings.Upgrade()
            '    My.Settings.NeedsUpdate = False
            '    My.Settings.Save()
            'End If

            Dim t As Date = Now
            Dim afterdate As String
            Dim beforedate As String = MakeWebfocusDate(Today.AddDays(1))
            Dim Dayrange As String = Today

            Dim wf As New WebfocusModule
            wf = wfLogin(wf)

            afterdate = MakeWebfocusDate(Today.AddDays(-50))


            Dim ScrapRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23scrapdatatqg&IBIMR_fex=pprasino/scrap_report.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/scrap_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&DISP_D=" & afterdate & "&LEDISP_D=" & beforedate & "&IBIMR_random=96021"
            ScrapRef = Replace(ScrapRef, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))

            beforedate = MakeWebfocusDate(Today.AddDays(1))
            If Hour(Now) < 3 Or Hour(Now) = 10 Then
                afterdate = MakeWebfocusDate(Today.AddDays(-10))
            Else
                afterdate = MakeWebfocusDate(Today.AddDays(-4))
            End If

            Dim ShipRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23salesshipmen&IBIMR_fex=pprasino/full_shipreport_by_lothtml.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/shipping_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&SHIPPED_D=" & afterdate & "&IBIMR_random=58708"
            ShipRef = Replace(ShipRef, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
            Dim TputRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23thruputrepor&IBIMR_fex=pprasino/esh_and_tput_for_flex_for_sql.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/thruput_detail_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&TP_DATE_COMPELTED=" & afterdate & "&LE_TP_DATE_COMPELTED=" & beforedate & "&IBIMR_random=31846"
            TputRef = Replace(TputRef, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
            Dim LaborRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23laborreporti&IBIMR_fex=pprasino/labor_part_detail_workorders_with_esh_for_sql_for_testing.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/labor_part_detail_workorders_with_esh.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&GECHARGE_DATE=" & afterdate & "&LECHARGE_DATE=" & beforedate & "&IBIMR_random=24311&"
            LaborRef = Replace(LaborRef, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))

            If Hour(Now) = 6 And Minute(Now) = 40 Then
                FullUpdate(wf)
            End If

            Debug.Print(TputRef)
            Console.WriteLine("Started at " & Now)

            'wf.GetReporthAsync(ShipRef, "ships")

            Dim RespNames() As String

            If Hour(Now) = 21 And Minute(Now) < 10 Then
                wf.GetReporthAsync("qavistes/qavistes.htm#routingandpa", "pprasinos:pprasino/ltsshtml.fex", "xtl")
                wf.GetReporthAsync("qavistes/qavistes.htm#routingandpa", "pprasinos:pprasino/allloy_part_data.fex", "partdata")
                Console.WriteLine("Pulling  WIP, Fingoods, Certs, Shipments, and  Timelines")
                My.Settings.LASTCERTPULL = Now

            ElseIf Minute(Now) < 20 Then
                Console.WriteLine("Pulling Opens, LabData, Fingoods, WIP and Shipments")
                'wf.GetReporthAsync("qavistes/qavistes.htm#certificateo", "pprasinos:pprasino/sl_wipfg_quality_check_inspbeyondhtml.fex", "certs")
                'wf.GetReporthAsync("qavistes/qavistes.htm#salesshipmen", "pprasinos:pprasino/custom_open_order_reportshtml.fex", "opens")
                wf.GetReporthAsync(ScrapRef, "scrap")

                'OpensUpdater(wf)
            ElseIf Minute(Now) < 30 Then
                Console.WriteLine("Pulling  WIP, Fingoods, Shipments, Tput, labor, and Opens")
                wf.GetReporthAsync(TputRef, "tput")
                My.Settings.LASTCERTPULL = Now
                My.Settings.LASTOPENPULL = Now
                '  If Hour(Now) = 1 Or Hour(Now) = 20 Then
                wf.GetReporthAsync(LaborRef, "labor")
            End If
            'wf.GetReporthAsync("qavistes/qavistes.htm#wipandshopco", "pprasinos:pprasino/customlotshtml.fex", "lots")
            'wf.GetReporthAsync("qavistes/qavistes.htm#salesshipmen", "pprasinos:pprasino/fingoodshtml.fex", "fingoods")

            My.Settings.LASTSHIPMENTPULL = Now

            UpdateAppend(wf, GetWFIds(wf.GetRequests))

            My.Settings.Save()

            Console.WriteLine("Lots Done in " & (Now - t).ToString)
            Console.WriteLine(Now - t)
            Threading.Thread.Sleep(20000)
            Exit Sub
emailerror:
            Fuckups = Fuckups + 1
            If Fuckups < 5 Then
                Threading.Thread.Sleep(1000)

            Else
                'LocalLib.EmailTools.EmailFile({"pprasinos@pccstructurals.com"}, "An error happened in SQLPublisher on line " & Err.Erl, "ErrorReport")
            End If
        Catch ex As Exception
            MsgBox("EX.MESSAGE")
            Threading.Thread.Sleep(10000000)
        End Try

    End Sub

    Public Function ShouldUpdate(MaxAge As Integer, TableName As String)
        ExecStoredProcedure("wflocal..getlastupdate", {"@tablename", "scrap"})

    End Function

    Public Function ExecStoredProcedure(Procedurename As String, Params As Object()) As String()
        Using cn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand
                cmd.CommandText = Procedurename
                cmd.Connection = cn
                For x = 0 To Params.Count / 2
                    cmd.Parameters.AddWithValue(Params(x), Params(x + 1))
                Next
                Try
                    cn.Open()
                    Using DR As SqlClient.SqlDataReader = cmd.ExecuteReader
                        'DR.VisibleFieldCount

                        Do While DR.Read

                        Loop
                    End Using
                Catch ex As Exception

                Finally
                    cn.Close()
                End Try
            End Using
        End Using

    End Function

    Private Function wfLogin(wf As WebfocusModule) As WebfocusModule
        If IsNothing(wf) Then wf = New WebfocusModule

        If Not wf.IsLoggedIn Then

            LogInInfo = GetUserPasswordandFex()
            wf.LogIn("pprasinos", "Wyman123-")
            Do Until wf.IsLoggedIn
                LogInInfo = GetUserPasswordandFex()
                wf.LogIn(LogInInfo(0), LogInInfo(1))
            Loop
        End If

        Return wf
    End Function

    Private Function FullUpdate(wf As WebfocusModule)
        Dim afterDate As String
        Dim beforeDate As String

        wfLogin(wf)
        Dim PARTLIST As New List(Of String)
        Using cn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand
                cmd.CommandText = "SELECT DISTINCT PARTNO FROM WFLOCAL..CERT_ERRORS WHERE ISNULL(DAYS_IN_WC,49) < 50 AND PARTNO NOT LIKE '%S'"
                cmd.Connection = cn
                cn.Open()
                Using DR As SqlClient.SqlDataReader = cmd.ExecuteReader
                    Do While DR.Read
                        PARTLIST.Add(DR("PARTNO"))
                    Loop
                End Using
                cn.Close()
                cmd.Parameters.Clear()
            End Using
        End Using
        PARTLIST.Sort()


        For I = 0 To PARTLIST.Count - 1 Step 1
            Try
                Dim PART As String = PARTLIST(I)
                PART = Trim(PART)
                Dim WipHistoryRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23wipandshopco&IBIMR_fex=pprasino/wo_move_history_8ball_for_sql.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/workorder_moves.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&PARTNO=" & PART & "&WORP_MPV=ab_gbv&&IBIMR_random=13866&"
                Console.WriteLine(PARTLIST.IndexOf(PART) & "   " & PART)
                wf.GetReporthAsync(WipHistoryRef, "wiphist")
                'Threading.Thread.Sleep(5000)
                UpdateAppend(wf, GetWFIds(wf.GetRequests))
                ' Threading.Thread.Sleep(1000)
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn("PPRASINOS", "Wyman123-")
            Catch EX1 As Exception
                Stop
            End Try

        Next


        For q = 0 To 20
            Console.Write(q & " ")
            Dim Span As Integer = 5
            beforeDate = MakeWebfocusDate(Today.AddDays(-q * Span))
            afterDate = MakeWebfocusDate(Today.AddDays(-1 - ((q + 1) * Span)))
            Console.WriteLine(beforeDate & "-" & afterDate)
            Dim InvRef1 As String = "qavistes/qavistes.htm#wipandshopco    pprasinos:pprasino/inventorybyms.fex "

            Dim LaborRef1 As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23laborreporti&IBIMR_fex=pprasino/labor_part_detail_workorders_with_esh_for_sql_for_testing.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/labor_part_detail_workorders_with_esh.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&GECHARGE_DATE=" & afterDate & "&LECHARGE_DATE=" & beforeDate & "&IBIMR_random=24311&"
            Dim TputRef1 As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23thruputrepor&IBIMR_fex=pprasino/esh_and_tput_for_flex_for_sql.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/thruput_detail_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&TP_DATE_COMPELTED=" & afterDate & "&LE_TP_DATE_COMPELTED=" & beforeDate & "&IBIMR_random=31846"
            Dim ScrapRef1 As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23scrapdatatqg&IBIMR_fex=pprasino/scrap_report_including_nodefect.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/scrap_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&DISP_D=" & afterDate & "&LEDISP_D=" & beforeDate & "&IBIMR_random=96021"
            TputRef1 = Replace(TputRef1, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
            ' ScrapRef1 = Replace(ScrapRef1, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
            ' LaborRef1 = Replace(LaborRef1, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))



            If Today.DayOfWeek = DayOfWeek.Monday Then wf.GetReporthAsync(TputRef1, "tput")
            If Today.DayOfWeek = DayOfWeek.Tuesday Then wf.GetReporthAsync(ScrapRef1, "scrap")
            If Today.DayOfWeek = DayOfWeek.Wednesday Then wf.GetReporthAsync(LaborRef1, "labor")

            '  Threading.Thread.Sleep(2000)
            UpdateAppend(wf, GetWFIds(wf.GetRequests))
            ' Threading.Thread.Sleep(1000)

            wf = Nothing
            wf = wfLogin(wf)

        Next q

    End Function


    Private Function GetPingMs(ByRef hostNameOrAddress As String)
        Dim ping As New System.Net.NetworkInformation.Ping
        Return ping.Send(hostNameOrAddress).RoundtripTime
        Threading.Thread.Sleep(1000)
    End Function

    Private Function GetUserPasswordandFex() As String()
        Dim h As New Random

        Dim Usernames() As String = {"hfaizi", "mreyes", "MALMARAZ", "MARJMAND", "HYANG", "GWONG", "VDELACRUZ", "JTIBAYAN", "JSOLIS", "ASINGH", "GREYES", "JPIMENTEL", "TOSULLIVAN", "MMARTIN", "VLOPEZ", "SLI", "JIMPERIAL", "JHERNANDEZ", "FHARO", "CGOUTAMA", "HGOMEZ", "EGONZALEZ", "CDAROSA"}

        Dim y As Integer = h.Next(0, Usernames.Length)
        Dim ps As String

        Dim FexAdd As String = "&IBIMR_sub_action=MR_MY_REPORT"
        If Usernames(y) <> "pprasinos" Then
            FexAdd = "&IBIMR_sub_action=MR_MY_REPORT&IBIMR_proxy_id=pprasino.htm&"
            ps = ChrW(112) & ChrW(97) & ChrW(115) & ChrW(115) & ChrW(50) & ChrW(48) & ChrW(49) & ChrW(53)
        Else
            ps = ChrW(87) & ChrW(121) & ChrW(109) & ChrW(97) & ChrW(110) & ChrW(49) & ChrW(50) & ChrW(51) & ChrW(45)
        End If
        Debug.Print(Usernames(y))
        Return {Usernames(y), ps, FexAdd}

    End Function


    Private Function GetWFIds(Requests As String) As String()
        Dim k() As String = Split(Requests, vbLf)
        For w = 0 To k.Length - 1
            k(w) = Mid(k(w), 3, 10)
            If Left(k(w), 1) <> "" Then k(w) = Right(k(w), Len(k(w)) - 1)
            If Left(k(w), 1) <> "" Then k(w) = Right(k(w), Len(k(w)) - 1)
            k(w) = Replace(k(w), " ", "")
        Next
        Return k
    End Function

    Private Function WithinString(String1 As String, String2 As String) As Boolean
        If InStr(String1, String2, CompareMethod.Text) <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function MakeWebfocusDate(Indate As Date) As String
        Dim vDay As String = Day(Indate)
        Dim Vmonth As String = Month(Indate)
        Dim vYear As String = Year(Indate)
        If Len(vDay) = 1 Then vDay = "0" & vDay
        If Len(Vmonth) = 1 Then Vmonth = "0" & Vmonth
        MakeWebfocusDate = Vmonth & vDay & vYear
    End Function

    Private Function NotificationEmails() As Int16

        Dim RawPull() As String = Split(FileIO.FileSystem.ReadAllText("\\slfs01\shared\prasinos\8ball\Notifications\Notifications.txt"), vbCrLf)

        Dim WOList As New List(Of String)


        Using cn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand
                cmd.CommandText = "select * from wflocal..NOTIFICATIONS a 
                                    left join wflocal..CERT_ERRORS b
                                    on a.WORKORDERNO=b.WORKORDERNO
                                    where a.OPERATIONNO<b.OPERATION
                                   "
                cmd.Connection = cn
                cn.Open()
                Using DR As SqlClient.SqlDataReader = cmd.ExecuteReader
                    Do While DR.Read
                        Dim MsgString As String = "This email Is to notify that lot " & DR("WORKORDERNO").ToString & " has reached Or passed operation " & DR("OPERATIONNO").ToString & Chr(10) & Chr(13) & vbCrLf & vbCrLf & "This Is an automated message"
                        EmailFile(DR("EMAIL").ToString, MsgString, "Movement notification:  " & DR("WORKORDERNO").ToString, True)
                        WOList.Add(DR("WORKORDERNO").ToString & "|" & DR("OPERATIONNO").ToString & "|" & DR("EMAIL").ToString)
                    Loop
                End Using

                For Each WO In WOList
                    cmd.CommandText = "DELETE FROM WFLOCAL..NOTIFICATIONS WHERE WORKORDERNO =@WORKORDERNO AND EMAIL=@EMAIL AND OPERATIONNO=@OPERATIONNO"
                    cmd.Parameters.AddWithValue("@WORKORDERNO", Split(WO, "|")(0))
                    cmd.Parameters.AddWithValue("@EMAIL", Split(WO, "|")(2))
                    cmd.Parameters.AddWithValue("@OPERATIONNO", Split(WO, "|")(1))
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                Next
                cn.Close()
                cmd.Parameters.Clear()
            End Using
        End Using

        Return 0
    End Function

    Sub EmailFile(Recipient As String, MessageBody As String, Subject As String, Optional Send As Boolean = False)

        Dim OutLookApp As New Outlook.Application
        Dim Mail As Outlook.MailItem = OutLookApp.CreateItem(Outlook.OlItemType.olMailItem)

        Dim mailRecipient As Outlook.Recipient

        mailRecipient = Mail.Recipients.Add(Recipient)
        mailRecipient.Resolve()

        Mail.Recipients.ResolveAll()

        Mail.HTMLBody = MessageBody
        Mail.Subject = Subject
        Mail.Save()
        If Send Then
            Mail.Send()
        Else
            Mail.Display()
        End If

    End Sub

    Private Sub RemoveTicket(TicketID As String)
        FileIO.FileSystem.CopyFile("\\slfs01\shared\prasinos\8ball\Notifications\Notifications.txt", My.Computer.FileSystem.SpecialDirectories.Desktop & "\Notifications.txt", True)

        Dim OutString As String = ""
        Dim instring() As String = Split(FileIO.FileSystem.ReadAllText(My.Computer.FileSystem.SpecialDirectories.Desktop & "\Notifications.txt"), vbCrLf)
        For Each textline In instring
            If InStr(textline, TicketID) = 0 Then OutString = OutString & textline & vbCrLf
        Next textline



        FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\8ball\Notifications\Notifications.txt", OutString, False)
    End Sub

    Private Sub UpdateAppend(WF As WebfocusDLL.WebfocusModule, RespNames() As String)
        'Threading.Thread.Sleep(60000)
        Dim RefFind() As String = {"ships", "fingoods", "lots", "certs", "scrap", "partdata", "xtl", "tput", "labor", "labor1", "wiphist"}
        Dim TableNames() As String = {"SHIPMENTS", "CERT_ERRORS", "CERT_ERRORS", "CERT_ERRORS", "SCRAP", "ALLOYS", "TIMELINE", "TPUT", "LABOR", "LABOR", "WIP_MOVE_HIST"}
        Dim UpdatedRows As Integer = 0
        Using cn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand
                cmd.Connection = cn

                Dim Query As String
                Dim TABLES() As String
                TABLES = TableNames

                If InStr(WF.GetRequests, "lots") <> 0 Then
                    cn.Open()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "UPDATE WFLOCAL.DBO.CERT_ERRORS SET ACTIVE = 2 WHERE ACTIVE <> 0"
                    cmd.ExecuteNonQuery()
                    cn.Close()
                End If

                For P = 0 To RespNames.Length - 1
                    If RespNames(P) = Nothing Or RespNames(P) = "opens" Then GoTo NEXTP
                    Dim j As New Object
                    j = WF.GetResponse(RespNames(P)).Response
                    Dim TableName As String
                    cn.Open()
                    Dim e As Integer
                    Dim t As Boolean
                    ' Try
                    For ind = 0 To RefFind.Length - 1
                        If RefFind(ind) = RespNames(P) Then TableName = TableNames(ind)
                    Next

                    Console.Write(TableName)
                    Console.CursorLeft = 0
                    cmd.CommandType = CommandType.Text
                    Query = "SELECT column_name, data_type FROM WFLOCAL.INFORMATION_SCHEMA.COLUMNS" & vbCrLf &
                "WHERE WFLOCAL.INFORMATION_SCHEMA.COLUMNS.TABLE_NAME='" & TableName & "'"
                    cmd.CommandText = Query

                    Dim ColumnInfo As New List(Of String())
                    Dim CSVColumns As String = ""
                    Dim CSVUPDATE As String = ""

                    Using dr As SqlDataReader = cmd.ExecuteReader
                        While dr.Read()
                            Dim Y As Integer = GetColumnNumber(j, dr("column_name").ToString)
                            If Y <> -1 Then
                                ColumnInfo.Add({dr("column_name").ToString, dr("data_type").ToString, Y})
                                CSVColumns = CSVColumns & "ROW." & dr("column_name").ToString & ", "
                                CSVUPDATE = CSVUPDATE & dr("column_name").ToString & " = @" & dr("column_name").ToString & ","
                            End If
                        End While
                    End Using
                    ColumnInfo.Add({"ACTIVE", "int", 0})

                    CSVColumns = CSVColumns & "ROW.ACTIVE, "

                    CSVUPDATE = CSVUPDATE & "ROW.ACTIVE = @ACTIVE,"

                    CSVColumns = Left(CSVColumns, Len(CSVColumns) - 2)
                    CSVUPDATE = Left(CSVUPDATE, Len(CSVUPDATE) - 1)

                    cmd.CommandType = CommandType.StoredProcedure
                    If TableName = "SCRAP" Then
                        cmd.CommandText = "WFLOCAL.DBO.UpdateScrap"
                    ElseIf TableName = "CERT_ERRORS" Then
                        cmd.CommandText = "WFLOCAL.DBO.UPDATEAPPENDWIP"
                    ElseIf TableName = "TIMELINE" Then
                        cmd.CommandText = "WFLOCAL.DBO.XTLupdateAppend"
                    ElseIf TableName = "ALLOYS" Then
                        cmd.CommandText = "	MERGE WFLOCAL..ALLOYS AS TARGET
                                            USING (SELECT @PARTNO, @ALLOY_DESCR, @MATERIAL_SPEC, @PART_DESCR, @PIECES_PER_MOLD, @SELLING_PRICE, @POUR_WEIGHT, @STOP_RELEASE, @PART_STATUS, @ROUT_REV) AS SOURCE 
                                                          (PARTNO,  ALLOY_DESCR,  MATERIAL_SPEC,  PART_DESCR,  PIECES_PER_MOLD,  SELLING_PRICE,  POUR_WEIGHT,  STOP_RELEASE,  PART_STATUS, ROUT_REV)
                                            ON TARGET.PARTNO=SOURCE.PARTNO
                                            WHEN MATCHED THEN
                                                    UPDATE SET	ALLOY_DESCR=SOURCE.ALLOY_DESCR,  
                                                                MATERIAL_SPEC=SOURCE.MATERIAL_SPEC,
                                                                PART_DESCR=SOURCE.PART_DESCR,
                                                                PIECES_PER_MOLD=SOURCE.PIECES_PER_MOLD,
                                                                SELLING_PRICE=SOURCE.SELLING_PRICE,
                                                                POUR_WEIGHT=SOURCE.POUR_WEIGHT,
                                                                STOP_RELEASE=SOURCE.STOP_RELEASE,
                                                                PART_STATUS=SOURCE.PART_STATUS,
                                                                ROUT_REV=SOURCE.ROUT_REV 
                                            WHEN NOT MATCHED THEN
                                                    INSERT (PARTNO,  ALLOY_DESCR,  MATERIAL_SPEC,  PART_DESCR,  PIECES_PER_MOLD,  SELLING_PRICE,  POUR_WEIGHT,  STOP_RELEASE,  PART_STATUS, ROUT_REV)
                                                    values (@PARTNO, @ALLOY_DESCR, @MATERIAL_SPEC, @PART_DESCR, @PIECES_PER_MOLD, @SELLING_PRICE, @POUR_WEIGHT, @STOP_RELEASE, @PART_STATUS, @ROUT_REV);
                                                "
                        cmd.CommandType = CommandType.Text
                    ElseIf TableName = "SHIPMENTS" Then
                        cmd.CommandText = "WFLOCAL.DBO.AddShipments"
                    ElseIf TableName = "TPUT" Then
                        cmd.CommandText = "WFLOCAL.DBO.UpdateThruput"
                    ElseIf TableName = "LABOR" Then
                        cmd.CommandText = "WFLOCAL.DBO.UpdateLabor"
                    ElseIf TableName = "WIP_MOVE_HIST" Then
                        cmd.CommandText = "WFLOCAL.DBO.UPDATE_WIP_HIST"
                    End If

                    Dim CT As Long = 1
                    For RowNum = 1 To j.length - 1

                        With cmd.Parameters
                            .Clear()
                            For Each Col In ColumnInfo

                                If Col(1) = "nvarchar" Or Col(1) = "nchar" Then

                                    .Add("@" & Col(0), SqlDbType.NVarChar).Value = j(RowNum)(Col(2))

                                ElseIf Col(1) = "float" And Col(0) <> "ACTIVE" Then
                                    Debug.Print(j(RowNum)(Col(2)))
                                    j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), "R", "1")
                                    j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), "N", "0")
                                    j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), "Y", "2")

                                    If Replace(j(RowNum)(Col(2)), ",", "") = "." Then
                                        .Add("@" & Col(0), SqlDbType.Float).Value = 0
                                    Else
                                        Dim s As Double = j(RowNum)(Col(2))
                                        .Add("@" & Col(0), SqlDbType.Float).Value = s
                                    End If
                                ElseIf InStr(Col(1), "smallint", CompareMethod.Text) <> 0 Then
                                    If Replace(j(RowNum)(Col(2)), ",", "") = "." Then
                                        .Add("@" & Col(0), SqlDbType.SmallInt).Value = 0
                                    Else
                                        Dim S As Int16 = Replace(j(RowNum)(Col(2)), ",", "") * 1
                                        .Add("@" & Col(0), SqlDbType.SmallInt).Value = S
                                    End If
                                ElseIf InStr(Col(1), "int", CompareMethod.Text) <> 0 And Col(0) <> "ACTIVE" Then
                                    If Replace(j(RowNum)(Col(2)), ",", "") = "." Then
                                        .Add("@" & Col(0), SqlDbType.Int).Value = 0
                                    Else
                                        Dim S As Integer = (0 & Replace(Replace(j(RowNum)(Col(2)), ",", ""), ".", "")) * 1
                                        .AddWithValue("@" & Col(0), S)
                                    End If
                                ElseIf InStr(Col(1), "Date", CompareMethod.Text) <> 0 Then

                                    Dim dt As DateTime = #1/1/1900#

                                    If Replace(j(RowNum)(Col(2)), ",", "") = "." Then

                                    Else
                                        dt = DateTime.Parse(j(RowNum)(Col(2)))
                                        'Debug.Print(dt.ToString)
                                        If dt.Year > 1900 Then
                                            .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
                                        Else
                                            dt = Now.AddYears(-100)
                                            .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
                                        End If
                                    End If
                                End If
                            Next
                            .Add("@ACTIVE", SqlDbType.Int).Value = 1
                        End With


                        t = True
                        Dim sqlTransaction As SqlClient.SqlTransaction = cn.BeginTransaction
                        cmd.Transaction = sqlTransaction
                        ' Try
                        cmd.ExecuteNonQuery()
                        sqlTransaction.Commit()

                        'Catch ex As Exception

                        '    Console.WriteLine("Trying To fix")
                        '    sqlTransaction.Rollback()
                        '    sqlTransaction.Dispose()
                        '    cn.Close()
                        '    Threading.Thread.Sleep(30000)
                        '    Try
                        '        cn.Open()

                        '        sqlTransaction = cn.BeginTransaction
                        '        cmd.Transaction = sqlTransaction
                        '        cmd.ExecuteNonQuery()
                        '        sqlTransaction.Commit()
                        '    Catch EX2 As Exception
                        '        sqlTransaction.Rollback()
                        '        MsgBox("Try To FIX?")
                        '        sqlTransaction.Rollback()
                        '        sqlTransaction.Dispose()
                        '        cn.Close()
                        '        Threading.Thread.Sleep(30000)
                        '        cn.Open()
                        '        sqlTransaction = cn.BeginTransaction
                        '        cmd.Transaction = sqlTransaction
                        '        cmd.ExecuteNonQuery()
                        '        sqlTransaction.Commit()
                        '    End Try

                        'End Try


                        'If e <> -1 Then Stop
                        t = False
                        CT = CT + 1

                        Console.CursorLeft = 0
                        Console.Write(CT & "/" & j.length & "       ")
                        'If RespNames(P) = "labor" Then Threading.Thread.Sleep(10)
                        '    If RespNames(P) = "tput" Then Threading.Thread.Sleep(10)

                    Next
                    Console.CursorLeft = 20
                    Console.WriteLine(TableName & " UPDATED Using " & RespNames(P))

                    ' Catch ex As Exception
                    ' MsgBox(j & "    " & t & "     " & ex.Message)
                    ' Finally
                    cn.Close()
                    ' End Try
                    '  Console.WriteLine("  (" & UpdatedRows & " Rows Updated)")
NEXTP:
                Next P

                If InStr(WF.GetRequests, "lots") <> 0 Then
                    cn.Open()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "UPDATE WFLOCAL.DBO.CERT_ERRORS Set ACTIVE = 0 WHERE ACTIVE = 2"
                    Dim RWS As Integer = cmd.ExecuteNonQuery()
                    '    If RWS <> -1 Then Stop
                    UpdatedRows = UpdatedRows + RWS
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "wflocal..cleanup"
                    cmd.Parameters.Clear()
                    cmd.ExecuteNonQuery()
                    cn.Close()
                End If

            End Using

        End Using

    End Sub


    Sub OpensUpdater(wf As WebfocusModule)

        Dim cdataset As New DataSet

        Dim j As Object = wf.GetResponse("opens").Response

        Dim ConnectionString As String = "Server=SLREPORT01; Database=WFLocal; persist security info=False; trusted_connection=Yes;"
        Using cn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand
                cmd.Connection = cn
                cn.Open()
                Dim Query As String = "UPDATE wflocal.dbo.OPEN_ORDERS Set ACTIVE = 2 WHERE ACTIVE <> 0"
                cmd.CommandText = Query
                cmd.ExecuteNonQuery()
                If Hour(Now) = 20 Then
                    cmd.CommandText = "DELETE FROM WFLOCAL.DBO.OPEN_ORDERS WHERE ACTIVE =2"
                    cmd.ExecuteNonQuery()
                End If

                Query = "Select column_name, data_type FROM WFLOCAL.INFORMATION_SCHEMA.COLUMNS" & vbCrLf &
                    "WHERE WFLOCAL.INFORMATION_SCHEMA.COLUMNS.TABLE_NAME='OPEN_ORDERS'"
                cmd.CommandText = Query

                Dim ColumnInfo As New List(Of String())
                Dim ColNumbers As New List(Of Integer)

                Using dr As SqlDataReader = cmd.ExecuteReader
                    While dr.Read()
                        Dim Y As Integer = GetColumnNumber(j, dr("column_name").ToString)
                        If Y <> -1 Then
                            ColumnInfo.Add({dr("column_name").ToString, dr("data_type").ToString, Y})
                        End If
                    End While
                End Using

                Dim QueryRoot As String = "INSERT INTO WFLOCAL.DBO.OPEN_ORDERS ("

                For Each col In ColumnInfo
                    QueryRoot = QueryRoot & col(0) & ", "
                Next
                QueryRoot = Left(QueryRoot, Len(QueryRoot) - 2) & ") values ("
                For Each col In ColumnInfo
                    QueryRoot = QueryRoot & "@" & col(0) & ", "
                Next
                QueryRoot = Left(QueryRoot, Len(QueryRoot) - 2) & ")"

                For RowNum = 1 To j.length - 1
                    With cmd.Parameters
                        Query = QueryRoot
                        .Clear()

                        For Each Col In ColumnInfo

                            If Col(1) = "nvarchar" Then
                                .Add("@" & Col(0), SqlDbType.NVarChar).Value = j(RowNum)(Col(2))
                                'Query = Replace(Query, "@" & Col(0), "'" & j(RowNum)(Col(2)) & "'")
                            ElseIf Col(1) = "float" Then
                                ' Debug.Print(j(RowNum)(Col(2)))
                                .Add("@" & Col(0), SqlDbType.Float).Value = Replace(j(RowNum)(Col(2)), ",", "")
                                'Query = Replace(Query, "@" & Col(0), CLng(j(RowNum)(Col(2))))
                            ElseIf Col(1) = "datetime" Then
                                Dim dt As DateTime = DateTime.Parse(j(RowNum)(Col(2)))
                                'Debug.Print(dt.ToString)
                                If dt.Year > 1900 Then
                                    .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
                                Else
                                    dt = Now.AddYears(-100)
                                    .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
                                End If
                                'Query = Replace(Query, "@" & Col(0), CLng(j(RowNum)(Col(2))))
                            End If
                        Next Col
                        .AddWithValue("@ACTIVE", 1)
                    End With

                    Console.Write(RowNum + 1 & "/" & j.length)

                    Console.CursorLeft = 0
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "WFLOCAL.DBO.OPENUPDATER"
                    Dim y As Integer = cmd.ExecuteNonQuery()
                    'If y <> -1 Then Stop
                Next RowNum
                cmd.CommandType = CommandType.Text
                cmd.CommandText = "UPDATE wflocal.dbo.OPEN_ORDERS SET ACTIVE = 0 WHERE ACTIVE = 2"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "INSERT INTO WFLOCAL.DBO.PO_REVIEW  (SALES_ORDER_NO, CUST_NO, SALES, USERNAME, ttimestamp, prel, pship, erel, eship)" & vbCrLf &
                                    "Select DISTINCT B.SALES_ORDER_NO, B.CUSTOMER_NO, B.ADDED_BY, B.ADDED_BY, getdate(), 1, 1, 1, 1" & vbCrLf &
                                    "From DBO.OPEN_ORDERS B" & vbCrLf &
                                    "Where Not EXISTS(SELECT distinct  B.SALES_ORDER_NO" & vbCrLf &
                                    "From DBO.PO_REVIEW" & vbCrLf &
                                    "Where PO_REVIEW.SALES_ORDER_NO = B.SALES_ORDER_NO" & vbCrLf &
                                    ")" & vbCrLf &
                                    "" & vbCrLf
                cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandText = "wflocal.dbo.CleanTickets"
                cmd.ExecuteNonQuery()
                Threading.Thread.Sleep(1)
            End Using
        End Using

    End Sub

    Private Function GetColumnNumber(InputTable()() As String, ColumLabel As String) As Integer

        Dim x As Integer = 0
        Do While ColumLabel <> InputTable(0)(x) And x < UBound(InputTable(0))
            x = x + 1
        Loop
        If x = UBound(InputTable(0)) And ColumLabel <> InputTable(0)(x) Then
            Return -1
            Debug.Print(ColumLabel)
        Else
            Return x
        End If
    End Function

End Module
