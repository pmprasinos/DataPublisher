
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
    Dim ConnectionString As String = "Server=SLREPORT01; Database=WFLocal; User Id=PrasinosApps; Password=Wyman123-; Connection Timeout = 10;"
    Private tmp = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\test.temp" 'O.Path.Combine(My.Computer.FileSystem.SpecialDirectories.Temp, "snafu.fubar")
    Public TimeToDownload As String
    Public StartTime As String
    Dim UpdateTimes As Object()()

    Public Function GetLastUpdate(TableName As String) As Date
        If IsNothing(UpdateTimes) Then Exit Function
        For x = 0 To UpdateTimes.Length - 1
            If UpdateTimes(x)(0) = TableName Then Return UpdateTimes(x)(1)
        Next
    End Function


    Private Function CheckIfRunning(ProcessName As String) As Integer
        Dim p() As Process = Process.GetProcessesByName(ProcessName)
        Return p.Count
    End Function

    Public Function ExecStoredProcedure(Procedurename As String, IsProcedure As Boolean, Optional Params As Object() = Nothing) As Object()()
        Dim StList As New List(Of Object())
        Using cn As New SqlConnection(ConnectionString)

            Using cmd As New SqlCommand
                cmd.CommandText = Procedurename
                cmd.Connection = cn
                cmd.CommandType = CommandType.Text
                If IsProcedure Then cmd.CommandType = CommandType.StoredProcedure
                If Not IsNothing(Params) Then
                    For x = 0 To (Params.Count / 2) - 1
                        cmd.Parameters.AddWithValue(Params(x), Params(x + 1))
                    Next
                End If
                Try
                    cn.Open()
                    Using DR As SqlClient.SqlDataReader = cmd.ExecuteReader
                        'DR.VisibleFieldCount

                        Do While DR.Read
                            Dim h(DR.VisibleFieldCount) As Object
                            DR.GetValues(h)
                            StList.Add(h)
                        Loop
                    End Using
                Catch ex As Exception

                Finally
                    cn.Close()
                End Try
            End Using
        End Using
        ExecStoredProcedure = StList.ToArray
    End Function

    Sub Main()
        Threading.Thread.Sleep(1000)
        Console.WriteLine("=====Do not close/disconnect from network until run complete=====")
        Console.WriteLine("Started at " & Now)
        Console.WriteLine()

        UpdateTimes = ExecStoredProcedure("wflocal..getlastupdate", True)

        If Environment.MachineName = "SLREPORT01" Or UCase(Environment.UserName) = "DATACOLLSL" Then
            If CheckIfRunning("SQLDatabasePublisher") <> 1 And Minute(Now) <> 20 Then
                If UCase(Environment.MachineName) <> "SLPPRASINOSLT01" Then Exit Sub
            ElseIf CheckIfRunning("SQLDatabasePublisher") > 2 Then
                System.Diagnostics.Process.Start("shutdown", "-r -f -t 00")
            ElseIf CheckIfRunning("EXCEL") > 0 Then

                If UCase(Environment.MachineName) = "SLREPORT01" Then Threading.Thread.Sleep(120000)
                FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\8ball\Logs.txt", Now() & "    EXCEL caused shutdown", True)
                If CheckIfRunning("EXCEL") > 0 And UCase(Environment.MachineName) = "SLREPORT01" Then System.Diagnostics.Process.Start("shutdown", "-r -f -t 00")
            End If
        End If
        Try

            If UCase(Environment.MachineName) <> "DATACOLLSL" Then NotificationEmails()

            Dim t As Date = Now

            Dim beforedate As String = MakeWebfocusDate(Today.AddDays(1))
            Dim afterdate As String = MakeWebfocusDate(Today.AddDays(-1))
            If Hour(Now) < 3 Or Hour(Now) = 14 Then afterdate = MakeWebfocusDate(Today.AddDays(-10))

            Dim adj As Integer = 0
            If (Hour(Now) >= 18 Or Hour(Now) <= 5) Then adj = 60
            If Environment.UserName <> "DATACOLLSL" Then adj = adj + 15


            Dim wf As New WebfocusModule
            wf = wfLogin(wf, True)

            Dim ScrapRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23scrapdatatqg&IBIMR_fex=pprasino/scrap_report.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/scrap_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&DISP_D=" & afterdate & "&LEDISP_D=" & beforedate & "&IBIMR_random=96021"
            ScrapRef = Replace(ScrapRef, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
            Dim ShipRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23salesshipmen&IBIMR_fex=pprasino/full_shipreport_by_lothtml.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/shipping_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&SHIPPED_D=" & afterdate & "&IBIMR_random=58708"
            ShipRef = Replace(ShipRef, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
            Dim TputRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23thruputrepor&IBIMR_fex=pprasino/esh_and_tput_for_flex_for_sql.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/thruput_detail_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&TP_DATE_COMPELTED=" & afterdate & "&LE_TP_DATE_COMPELTED=" & beforedate & "&IBIMR_random=31846"
            TputRef = Replace(TputRef, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
            Dim LaborRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23laborreporti&IBIMR_fex=pprasino/labor_part_detail_workorders_with_esh_for_sql_for_testing.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/labor_part_detail_workorders_with_esh.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&GECHARGE_DATE=" & afterdate & "&LECHARGE_DATE=" & beforedate & "&IBIMR_random=24311&"
            LaborRef = Replace(LaborRef, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))



            If Hour(Now) = 19 And DateDiff(DateInterval.Minute, GetLastUpdate("TIMELINE"), Now) > (60 * 24 * 15 + adj) Then
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn(LogInInfo(0), LogInInfo(1))
                UpdateStatus(1, "SUBMITTED", "TIMELINE", False)
                wf.GetReporthAsync("qavistes/qavistes.htm#routingandpa", "pprasinos:pprasino/ltsshtml.fex", "xtl")
                UpdateAppend(wf, GetWFIds(wf.GetRequests))
            End If

            If Hour(Now) = 19 And DateDiff(DateInterval.Minute, GetLastUpdate("ALLOYS"), Now) > (60 * 24 * 15 * adj) Then
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn(LogInInfo(0), LogInInfo(1))
                wf.GetReporthAsync("qavistes/qavistes.htm#routingandpa", "pprasinos:pprasino/allloy_part_data.fex", "partdata")
                UpdateStatus(1, "SUBMITTED", "ALLOY", False)
                UpdateAppend(wf, GetWFIds(wf.GetRequests))
            End If

            If Hour(Now) Mod 2 = 1 And Minute(Now) < 10 Then
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn(LogInInfo(0), LogInInfo(1))
                Console.WriteLine("UPDATING CDCS DATA")
                wf.GetReporthAsync("qavistes/qavistes.htm#certificateo", "pprasinos:pprasino/sl_wipfg_quality_check_inspbeyondhtml.fex", "certs")
                UpdateStatus(1, "SUBMITTED - CERTS", "CERT_ERRORS", False)
                UpdateAppend(wf, GetWFIds(wf.GetRequests))
            End If

            Threading.Thread.Sleep(1000)
            Console.WriteLine("SHIPMENTS IS " & DateDiff(DateInterval.Minute, GetLastUpdate("SHIPMENTS"), Now) & " MINUTES OLD (MAX: " & 9 + adj & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("SHIPMENTS"), Now) > 10 + adj Then
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn(LogInInfo(0), LogInInfo(1))
                wf.GetReporthAsync(ShipRef, "ships")
                UpdateStatus(1, "SUBMITTED", "SHIPMENTS", False)
                UpdateAppend(wf, GetWFIds(wf.GetRequests))
            End If

            Threading.Thread.Sleep(1000)
            Console.WriteLine("WIP IS " & DateDiff(DateInterval.Minute, GetLastUpdate("CERT_ERRORS"), Now) & " MINUTES OLD (MAX: " & 15 + adj & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("CERT_ERRORS"), Now) > 15 + adj Then
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn(LogInInfo(0), LogInInfo(1))
                wf.GetReporthAsync("qavistes/qavistes.htm#wipandshopco", "pprasinos:pprasino/customlotshtml.fex", "lots")
                wf.GetReporthAsync("qavistes/qavistes.htm#salesshipmen", "pprasinos:pprasino/fingoodshtml.fex", "fingoods")
                UpdateStatus(1, "SUBMITTED - LOTSANDFINGOODS", "CERT_ERRORS", False)
                UpdateAppend(wf, GetWFIds(wf.GetRequests))
            End If

            Threading.Thread.Sleep(1000)
            Console.WriteLine("TPUT IS " & DateDiff(DateInterval.Minute, GetLastUpdate("TPUT"), Now) & " MINUTES OLD (MAX: " & 60 + adj & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("TPUT"), Now) > 60 + adj Then
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn(LogInInfo(0), LogInInfo(1))
                wf.GetReporthAsync(TputRef, "tput")
                UpdateStatus(1, "SUBMITTED", "TPUT", False)
                UpdateAppend(wf, GetWFIds(wf.GetRequests))
            End If

            Threading.Thread.Sleep(1000)
            Console.WriteLine("LABOR IS " & DateDiff(DateInterval.Minute, GetLastUpdate("LABOR"), Now) & " MINUTES OLD (MAX: " & 340 & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("LABOR"), Now) > 340 Then
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn(LogInInfo(0), LogInInfo(1))
                wf.GetReporthAsync(LaborRef, "labor")
                UpdateStatus(1, "SUBMITTED", "LABOR", False)
                UpdateAppend(wf, GetWFIds(wf.GetRequests))
            End If

            Threading.Thread.Sleep(1000)
            Console.WriteLine("SCRAP IS " & DateDiff(DateInterval.Minute, GetLastUpdate("SCRAP"), Now) & " MINUTES OLD (MAX: " & 300 & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("SCRAP"), Now) > 300 Then
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn(LogInInfo(0), LogInInfo(1))
                wf.GetReporthAsync(ScrapRef, "scrap")
                UpdateStatus(1, "SUBMITTED", "scrap", False)
                UpdateAppend(wf, GetWFIds(wf.GetRequests))
            End If

            Threading.Thread.Sleep(1000)
            Console.WriteLine("OPEN ORDERS IS " & DateDiff(DateInterval.Minute, GetLastUpdate("OPEN_ORDERS"), Now) & " MINUTES OLD (MAX: " & 45 + adj & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("OPEN_ORDERS"), Now()) > 45 + adj Then
                wf = Nothing
                wf = New WebfocusModule
                wf.LogIn(LogInInfo(0), LogInInfo(1))
                wf.GetReporthAsync("qavistes/qavistes.htm#salesshipmen", "pprasinos:pprasino/custom_open_order_reportshtml.fex", "opens")
                UpdateStatus(1, "SUBMITTED", "OPEN_ORDERS", False)
                OpensUpdater(wf)
            End If

            Console.WriteLine()
            Console.WriteLine("Run Complete in " & (Now - t).ToString)
            Console.WriteLine()
            For x = 20 To 0 Step -1
                Threading.Thread.Sleep(1000)
                Console.Write("Form will close in " & x & "press any key")
                If Console.KeyAvailable Then Exit Sub
                Console.CursorLeft = 0
            Next

        Catch ex As Exception

            'MsgBox(Now & vbCrLf & vbCrLf & ex.Message)
            ' MsgBox(ex.InnerException.ToString)
            FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\8ball\log.txt", Now() & "   " & ex.Message.ToString & " || " & ex.InnerException.ToString, True)
            '            Threading.Thread.Sleep(10000000)
            MsgBox(ex.Message.ToString)
            MsgBox(ex.InnerException.ToString)
        End Try

    End Sub

    Public Function UpdateStatus(NewStatus As Integer, NewNotes As String, TableName As String, byuid As Boolean) As Guid
        ExecStoredProcedure("INSERT INTO WFLOCAL..PullStatus VALUES (GETDATE(), '" & TableName & "', '" & Environment.UserName & "', '" & Environment.MachineName & "', '" & NewNotes & "', " & NewStatus & ", NEWID(), GETDATE())", False)
        Return ExecStoredProcedure("Select UID from wflocal..PullStatus WHERE TABLENAME = '" & TableName & "' AND PULLNOTES = '" & NewNotes & "' AND MACHINENAME = '" & Environment.MachineName & "' AND PULLSTATUS = " & NewStatus, False)(0)(0)
    End Function

    Public Function UpdateStatus(NewStatus As Integer, NewNotes As String, uid As String) As Guid
        If uid <> "" Then ExecStoredProcedure("UPDATE WFLOCAL..PullStatus SET (TIMESTAMP = GETDATE(), PULLNOTES = '" & NewNotes & "', PULLSTATUS =" & NewStatus & "WHERE UID = '" & uid & "'", False)
        Return ExecStoredProcedure("Select UID from wflocal..PullStatus WHERE uid = '" & uid & "'", False)(0)(0)
    End Function
    Public Function ShouldUpdate(MaxAge As Integer, TableName As String)
        ExecStoredProcedure("wflocal..getlastupdate", {"@tablename", "scrap"})
    End Function

    Public Function ExecStoredProcedure(Procedurename As String, Params As Object(), Optional isProcedure As Boolean = False) As String()
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

    Private Function wfLogin(wf As WebfocusModule, Optional CredentialsOnly As Boolean = False) As WebfocusModule
        If IsNothing(wf) Then wf = New WebfocusModule

        If Not wf.IsLoggedIn Then

            LogInInfo = GetUserPasswordandFex()
            If Not CredentialsOnly Then
                wf.LogIn("pprasinos", "Wyman123-")
                Do Until wf.IsLoggedIn
                    LogInInfo = GetUserPasswordandFex()
                    wf.LogIn(LogInInfo(0), LogInInfo(1))
                Loop
            End If
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
                cmd.CommandText = "Select DISTINCT PARTNO FROM WFLOCAL..CERT_ERRORS WHERE ISNULL(DAYS_IN_WC, 49) < 50 And PARTNO Not Like '%S'"
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


    Private Function GetWFIds(Requests As String, Optional notarray As Boolean = False) As String()
        Dim k() As String = Split(Requests, vbLf)
        For w = 0 To k.Length - 1
            k(w) = Mid(k(w), 3, 10)
            If Left(k(w), 1) <> "" Then k(w) = Right(k(w), Len(k(w)) - 1)
            If Left(k(w), 1) <> "" Then k(w) = Right(k(w), Len(k(w)) - 1)
            If notarray Then
                k(0) = Replace(k(w), " ", "") & "  "
            Else
                k(w) = Replace(k(w), " ", "")
            End If
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
        Dim trans As SqlTransaction
        Dim RefFind() As String = {"ships", "fingoods", "lots", "certs", "scrap", "partdata", "xtl", "tput", "labor", "labor1", "wiphist"}
        Dim TableNames() As String = {"SHIPMENTS", "CERT_ERRORS", "CERT_ERRORS", "CERT_ERRORS", "SCRAP", "ALLOYS", "TIMELINE", "TPUT", "LABOR", "LABOR", "WIP_MOVE_HIST"}
        Dim UpdatedRows As Integer = 0
        Using cn As New SqlConnection(ConnectionString)
            cn.Open()
            Try
                Using cmd As New SqlCommand
                    trans = cn.BeginTransaction(Environment.UserName)
                    cmd.Connection = cn
                    cmd.Transaction = trans
                    Dim Query As String
                    Dim TABLES() As String
                    TABLES = TableNames

                    If InStr(WF.GetRequests, "lots") <> 0 Then
                        cmd.CommandType = CommandType.Text
                        cmd.CommandText = "UPDATE WFLOCAL.DBO.CERT_ERRORS SET ACTIVE = 2 WHERE ACTIVE <> 0"
                        cmd.ExecuteNonQuery()

                    End If

                    For P = 0 To RespNames.Length - 1
                        If RespNames(P) = Nothing Or RespNames(P) = "opens" Then GoTo NEXTP
                        Dim j As New Object
                        j = WF.GetResponse(RespNames(P)).Response

                        Dim TableName As String = ""

                        Dim e As Integer
                        Dim t As Boolean
                        ' Try
                        For ind = 0 To RefFind.Length - 1
                            If RefFind(ind) = RespNames(P) Then TableName = TableNames(ind)
                        Next

                        Dim UID As Guid
                        Try
                            UID = UpdateStatus(2, "RECIEVED", TableName, False)
                        Catch : End Try
                        ' Console.Write(TableName)
                        'Console.CursorLeft = 0
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
                                            USING (SELECT @PARTNO, @ALLOY_DESCR, @MATERIAL_SPEC, @PART_DESCR, @PIECES_PER_MOLD, @SELLING_PRICE, @POUR_WEIGHT, @STOP_RELEASE, @PART_STATUS, @ROUT_REV, @SHIP_WEIGHT) AS SOURCE 
                                                          (PARTNO,  ALLOY_DESCR,  MATERIAL_SPEC,  PART_DESCR,  PIECES_PER_MOLD,  SELLING_PRICE,  POUR_WEIGHT,  STOP_RELEASE,  PART_STATUS, ROUT_REV, SHIP_WEIGHT)
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
                                                                ROUT_REV=SOURCE.ROUT_REV,
                                                                SHIP_WEIGHT = SOURCE.SHIP_WEIGHT
                                            WHEN NOT MATCHED THEN
                                                    INSERT (PARTNO,  ALLOY_DESCR,  MATERIAL_SPEC,  PART_DESCR,  PIECES_PER_MOLD,  SELLING_PRICE,  POUR_WEIGHT,  STOP_RELEASE,  PART_STATUS, ROUT_REV, SHIP_WEIGHT)
                                                    values (@PARTNO, @ALLOY_DESCR, @MATERIAL_SPEC, @PART_DESCR, @PIECES_PER_MOLD, @SELLING_PRICE, @POUR_WEIGHT, @STOP_RELEASE, @PART_STATUS, @ROUT_REV, @SHIP_WEIGHT);
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
                        Console.Write("0")
                        Console.CursorLeft = 0

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

                            ' Try
                            cmd.ExecuteNonQuery()

                            t = False
                            CT = CT + 1

                            Console.CursorLeft = 0
                            Console.Write(CT & "/" & j.length & "        ")

                        Next
                        Console.CursorLeft = 20
                        Console.WriteLine(TableName & " UPDATED Using " & RespNames(P))


NEXTP:
                        Try
                            UpdateStatus(3, "UPDATED", TableName, False)
                        Catch : End Try
                    Next P

                    If InStr(WF.GetRequests, "lots") <> 0 Then
                        cmd.CommandType = CommandType.Text
                        cmd.CommandText = "UPDATE WFLOCAL.DBO.CERT_ERRORS Set ACTIVE = 0 WHERE ACTIVE = 2"
                        Dim RWS As Integer = cmd.ExecuteNonQuery()
                        '    If RWS <> -1 Then Stop
                        UpdatedRows = UpdatedRows + RWS
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "wflocal..cleanup"
                        cmd.Parameters.Clear()
                        cmd.ExecuteNonQuery()
                    End If
                End Using
                trans.Commit()
            Catch ex As Exception
                MsgBox(ex.ToString)
                MsgBox(ex.InnerException.ToString)
                Try
                    trans.Rollback()
                Catch ex2 As Exception
                    Console.WriteLine("Rollback Exception Type: {0}", ex2.GetType())
                    Console.WriteLine("  Message: {0}", ex2.Message)
                End Try

            End Try
        End Using

    End Sub


    Sub OpensUpdater(wf As WebfocusModule)
        Dim trans As SqlTransaction
        Dim cdataset As New DataSet

        Dim j As Object = wf.GetResponse("opens").Response

        Using cn As New SqlConnection(ConnectionString)
            Try
                cn.Open()
                Using cmd As New SqlCommand("", cn)
                    trans = cn.BeginTransaction(Environment.UserName)
                    cmd.Connection = cn
                    cmd.Transaction = trans
                    Dim Query As String = "UPDATE wflocal.dbo.OPEN_ORDERS Set ACTIVE = 2 WHERE ACTIVE <> 0"
                    cmd.CommandText = Query
                    cmd.ExecuteNonQuery()

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

                    Dim QueryRoot As String = ""

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
                        Console.CursorLeft = 0
                        Console.Write(RowNum + 1 & "/" & j.length & "       ")

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
                End Using
                trans.Commit()
                Console.CursorLeft = 20
                Console.WriteLine("OPEN_ORDERS" & " UPDATED Using " & "opens")
            Catch ex As Exception
                MsgBox("Commit Exception Type: {0}", ex.GetType().ToString)
                MsgBox("  Message: {0}", ex.Message.ToString)
                MsgBox(ex.InnerException.ToString)
                ' Attempt to roll back the transaction. 
                Try
                    trans.Rollback()
                Catch ex2 As Exception
                    ' This catch block will handle any errors that may have occurred 
                    ' on the server that would cause the rollback to fail, such as 
                    ' a closed connection.
                    MsgBox("Rollback Exception Type: {0}", ex2.GetType().ToString)
                    MsgBox("  Message: {0}", ex2.Message.ToString)

                End Try

            End Try
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
