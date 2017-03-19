
Imports System.Data
Imports System.Data.SqlClient
Imports WebfocusDLL
Imports System.Threading.Thread
Imports Microsoft.Office.Interop
Imports System.Net.Http
Imports mshtml
Imports ClassLibrary1
Imports SHDocVw
Imports Microsoft.VisualBasic
Imports System.Net

Module module1
    Dim LogInInfo As String()
    'Dim ConnectionString As String = "Server=SLPPRASINOSLT01; Database=WFLocal; User Id=PrasinosApps; Password=Wyman123-; Connection Timeout = 5;"
    Dim ConnectionString As String = "Server=SLREPORT01; Database=WFLocal; User Id=PrasinosApps; Password=Wyman123-; Connection Timeout = 5;"
    Private tmp = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\test.temp"
    Dim UpdateTimes As Object()()
    '######Define debugtext for testing on suffixed tables#####
    Dim DebugText As String = ""
    Dim IE As SHDocVw.InternetExplorer = Nothing
    Dim PullNumber As Integer = 0

    Sub Main()
        'If UCase(Environment.UserName) <> "PPRASINOS" Then Exit Sub
        ' If Not InStr(Environment.MachineName, "3DP") > 0 Then Exit Sub
        Console.WriteLine("=====Do not close or disconnect from network until run complete=====")
        Console.WriteLine()
        Console.WriteLine("Started at " & Now)

        '  If UCase(Environment.MachineName) <> "DATACOLLSL" Then NotificationEmails()

        Dim t As Date = Now
        Dim IsUser As Boolean = False

        '######BeforeDate and AfterDate are the range used for shipments, labor, tput, and wip_move_hist#####
        Dim BeforeDate As String = MakeWebfocusDate(Today.AddDays(1))
        Dim AfterDate As String = MakeWebfocusDate(Today)

        Dim adj As Integer = 0

        If Hour(Now) < 3 Then AfterDate = MakeWebfocusDate(Today.AddDays(-5))
        UpdateTimes = ExecStoredProcedure("wflocal..getlastupdate", True)

        '#####adj is added to the database age during hours where refresh rates can be extended#####
        If (Hour(Now) >= 18 Or Hour(Now) <= 5) Then adj = 60

        If UCase(Environment.UserName) <> "DATACOLLSL" Then
            adj = adj + 8
            AfterDate = MakeWebfocusDate(Today.AddDays(-4))
        Else
            adj = adj - 5
        End If

        '#####if there is an error in this program more than 2 times, the DataColl computer will be restarted#####
        If Environment.MachineName = "SLREPORT01" Or UCase(Environment.UserName) = "DATACOLLSL" Or UCase(Environment.UserName) = "PPRASINOS" Then
            If CheckIfRunning("SQLDatabasePublisher") > 1 Then
                System.Diagnostics.Process.Start("shutdown", "-r -f -t 00")
            ElseIf CheckIfRunning("EXCEL") > 0 And Environment.MachineName = "SLREPORT01" Then
                If UCase(Environment.MachineName) = "SLREPORT01" Then Threading.Thread.Sleep(120000)
                FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\8ball\Logs.txt", Now() & "    EXCEL caused shutdown", True)
                If CheckIfRunning("EXCEL") > 0 And UCase(Environment.MachineName) = "SLREPORT01" Then System.Diagnostics.Process.Start("shutdown", "-r -f -t 00")
            End If
        Else
            '#####Force update if pull is not automated###
            IsUser = True
            Console.Write("Enter FromDate using format 'MMDDYYYY': ")
            AfterDate = Console.ReadLine
            Console.Write("   Enter ToDate using format 'MMDDYYYY': " & MakeWebfocusDate(Today))
            Console.CursorLeft = Console.CursorLeft - 8
            BeforeDate = Console.ReadLine()
            If BeforeDate = "" Then BeforeDate = MakeWebfocusDate(Today)
            Console.WriteLine("Type 'Y' to delete and replace (else refresh)")
            If Console.ReadKey.KeyChar = "Y" Then

            End If

        End If
        'UpdateTimes(0)(0) = "OPEN_ORDERS" : UpdateTimes(0)(1) = Today.AddDays(-2)
        'UpdateTimes(1)(0) = "TPUT" : UpdateTimes(1)(1) = Today.AddDays(-2)
        'UpdateTimes(2)(0) = "SHIPMENTS" : UpdateTimes(2)(1) = Today.AddDays(-2)
        'UpdateTimes(3)(0) = "CERT_ERRORS" : UpdateTimes(3)(1) = Today.AddDays(-2)
        'Try
        Dim OpensRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=custom_open_order_reportshtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=209&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=13377"
        Dim TputRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=ESH_and_TPUT_FOR_FLEX_for_sql.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=340&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&LE_TP_DATE_COMPELTED=" + BeforeDate + "&TP_DATE_COMPELTED=" + AfterDate + "&BIP_rand=21066"
        Dim ScrapRef As String = "http://webfocus/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~pprasinos%252Fscrapdatatqg&BIP_item=scrap_report.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=560&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=10000&DISP_D=" + AfterDate + "&LEDISP_D=" + BeforeDate + "0&BIP_rand=74775"
        Dim ShipRef As String = "Http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=full_shipreport_by_lothtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=429&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&&SHIPPED_D=" + AfterDate + "&BIP_rand=7574"
        Dim WIPRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=customlotshtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=429&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=70094"
        Dim FGRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=fingoodshtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=429&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=36829"
        Dim CDCSRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=sl_wipfg_quality_check_inspbeyondhtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=209&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=20390"
        Dim TimeLineRef As String = "http://webfocus/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~pprasinos%252Froutingandpa&BIP_item=ltsshtml.fex&WF_STYLE_HEIGHT=1160&WF_STYLE_WIDTH=440&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=34035"
        Dim LaborRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=Labor_Part_Detail_ESH_FOR_SQL_for_testing.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=209&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&GECHARGE_DATE=" & AfterDate & "&LECHARGE_DATE=" & BeforeDate
        '#####define refrences for the reports#####
        ''''''Dim wf As New WebfocusModule
        '''''wfLogin(wf)

        '''''Dim LaborRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23laborreporti&IBIMR_fex=pprasino/labor_part_detail_workorders_with_esh_for_sql_for_testing.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/labor_part_detail_workorders_with_esh.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&GECHARGE_DATE=" & AfterDate & "&LECHARGE_DATE=" & BeforeDate & "&IBIMR_random=24311&"
        '''''LaborRef = Replace(LaborRef, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
        Dim Maxage As Integer = 0


            '#####TIMELINE and ALLOYS tables are updated once monthly, these are big reports
            If Day(Now) = 11 And DateDiff(DateInterval.Minute, GetLastUpdate("TIMELINE" & DebugText), Now) > ((60 * 24 * 15) + (12 * adj)) Then
                ExecStoredProcedure("update wflocal..TIMELINE set DWELL =31.6 WHERE OPERATION_NO = 20 AND PARTNO = '01296'", False)
                'If Minute(Now) Mod 10 = 0 Then Exit Sub
                ''''''wf = Nothing : wf = New WebfocusModule : wf.LogIn(LogInInfo(0), LogInInfo(1))
                UpdateStatus(1, "SUBMITTED", "TIMELINE", False)
                ''''''wf.GetReporthAsync("qavistes/qavistes.htm#routingandpa", "pprasinos:pprasino/ltsshtml.fex", "xtl")
                ''''''UpdateAppend(wf, GetWFIds(wf.GetRequests))
                UpdateAppend(TimeLineRef, "xtl")
                If Environment.UserName = "DATACOLLSL" Then Exit Sub
            End If

            If Day(Now) = 15 And DateDiff(DateInterval.Minute, GetLastUpdate("ALLOYS" & DebugText), Now) > ((60 * 24 * 15) - (12 * adj)) Then
                'ExecStoredProcedure("update wflocal..ALLOYS set ALLOY_DESCR = '347' WHERE PARTNO = '01296'", False)
                ''''''wf = Nothing : wf = New WebfocusModule : wf.LogIn(LogInInfo(0), LogInInfo(1))
                ''''''wf.GetReporthAsync("qavistes/qavistes.htm#routingandpa", "pprasinos:pprasino/allloy_part_data.fex", "partdata")
                'UpdateStatus(1, "SUBMITTED", "ALLOY", False)
                ''''''UpdateAppend(wf, GetWFIds(wf.GetRequests))
                'If Environment.UserName = "DATACOLLSL" Then Exit Sub
            End If

            If UCase(Environment.UserName) <> "DATACOLLSL" Then Threading.Thread.Sleep(50)
            Maxage = 14 + adj
            If Hour(Now) < 13 Then Maxage = Maxage + 20 'SHIPMENTS table does not need high refresh rate before 1:00PM
            Console.WriteLine("SHIPMENTS IS " & DateDiff(DateInterval.Minute, GetLastUpdate("SHIPMENTS" & DebugText), Now) & " MINUTES OLD (MAX: " & Maxage.ToString & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("SHIPMENTS" & DebugText), Now) > Maxage Then
                ''''''wf = Nothing : wf = New WebfocusModule : wf = wfLogin(wf)
                ''''''wf.GetReporthAsync(ShipRef, "ships")
                UpdateStatus(1, "SUBMITTED", "SHIPMENTS", False)
                ''''''UpdateAppend(wf, GetWFIds(wf.GetRequests))
                UpdateAppend(ShipRef, "ships")
            End If

            If UCase(Environment.UserName) <> "DATACOLLSL" Then Threading.Thread.Sleep(50)
            Maxage = 18 + adj
            Console.WriteLine("WIP IS " & DateDiff(DateInterval.Minute, GetLastUpdate("CERT_ERRORS" & DebugText), Now) & " MINUTES OLD (MAX: " & Maxage.ToString & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("CERT_ERRORS" & DebugText), Now) > Maxage Then
                ''''''wf = Nothing : wf = New WebfocusModule : wf.LogIn(LogInInfo(0), LogInInfo(1))
                ''''''wf.GetReporthAsync("qavistes/qavistes.htm#salesshipmen", "pprasinos:pprasino/fingoodshtml.fex", "fingoods")
                ''''''wf.GetReporthAsync("qavistes/qavistes.htm#wipandshopco", "pprasinos:pprasino/customlotshtml.fex", "lots")
                UpdateStatus(1, " SUBMITTED - LOTSANDFINGOODS", "CERT_ERRORS", False)
                ''''''UpdateAppend(wf, GetWFIds(wf.GetRequests))
                UpdateAppend(FGRef, "fingoods")
                UpdateAppend(WIPRef, "lots")


            End If

            If UCase(Environment.UserName) <> "DATACOLLSL" Then Threading.Thread.Sleep(50)
            Maxage = 90 + adj
            Console.WriteLine("TPUT IS " & DateDiff(DateInterval.Minute, GetLastUpdate("TPUT" & DebugText), Now) & " MINUTES OLD (MAX: " & Maxage.ToString() & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("TPUT" & DebugText), Now) > Maxage Then
                ''''''wf = Nothing : wf = New WebfocusModule : wf.LogIn(LogInInfo(0), LogInInfo(1))
                ''''''wf.GetReporthAsync(TputRef, "tput")
                UpdateStatus(1, "SUBMITTED", "TPUT", False)
                ''''''UpdateAppend(wf, GetWFIds(wf.GetRequests))
                UpdateAppend(TputRef, "tput")
            End If

            If UCase(Environment.UserName) <> "DATACOLLSL" Then Threading.Thread.Sleep(50)
            Maxage = 720 - adj
            Console.WriteLine("LABOR IS " & DateDiff(DateInterval.Minute, GetLastUpdate("LABOR" & DebugText), Now) & " MINUTES OLD (MAX: " & Maxage.ToString() & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("LABOR" & DebugText), Now) > Maxage Then
                UpdateStatus(1, "SUBMITTED", "LABOR", False)
                UpdateAppend(LaborRef, "labor")
            End If

            If UCase(Environment.UserName) <> "DATACOLLSL" Then Threading.Thread.Sleep(50)
            Maxage = 800 - adj
            Console.WriteLine("SCRAP IS " & DateDiff(DateInterval.Minute, GetLastUpdate("SCRAP" & DebugText), Now) & " MINUTES OLD (MAX: " & Maxage.ToString() & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("SCRAP" & DebugText), Now) > Maxage Then
                ''''''wf = Nothing : wf = New WebfocusModule : wf.LogIn(LogInInfo(0), LogInInfo(1))
                ''''''wf.GetReporthAsync(ScrapRef, "scrap")
                UpdateStatus(1, "SUBMITTED", "scrap", False)
                ''''''UpdateAppend(wf, GetWFIds(wf.GetRequests))
                UpdateAppend(ScrapRef, "scrap")
            End If

            If UCase(Environment.UserName) <> "DATACOLLSL" Then Threading.Thread.Sleep(50)
            Maxage = 55 + adj
            Console.WriteLine("OPEN ORDERS IS " & DateDiff(DateInterval.Minute, GetLastUpdate("OPEN_ORDERS" & DebugText), Now) & " MINUTES OLD (MAX: " & Maxage.ToString() & ")")
            If DateDiff(DateInterval.Minute, GetLastUpdate("OPEN_ORDERS" & DebugText), Now()) > Maxage Then
                ''''''wf = Nothing : wf = New WebfocusModule : wf.LogIn(LogInInfo(0), LogInInfo(1))
                ''''''wf.GetReporthAsync("qavistes/qavistes.htm#salesshipmen", "pprasinos:pprasino/custom_open_order_reportshtml.fex", "opens")
                UpdateStatus(1, "SUBMITTED", "OPEN_ORDERS", False)
                ''''''OpensUpdater(wf)
                OpensUpdater(OpensRef)

            End If

            If UCase(Environment.UserName) <> "DATACOLLSL" Then Threading.Thread.Sleep(50)
            If Environment.UserName = "DATACOLLSL" And Hour(Now) Mod 2 = 1 And Minute(Now) < 5 Then
                ''''''wf = Nothing : wf = New WebfocusModule : wf.LogIn(LogInInfo(0), LogInInfo(1))
                ''''''Console.WriteLine("UPDATING CDCS DATA")
                ''''''wf.GetReporthAsync("qavistes/qavistes.htm#certificateo", "pprasinos:pprasino/sl_wipfg_quality_check_inspbeyondhtml.fex", "certs")
                UpdateStatus(1, " SUBMITTED - CERTS", "CERT_ERRORS", False)
                UpdateAppend(CDCSRef, "certs")
                ''''''UpdateAppend(wf, GetWFIds(wf.GetRequests))

            End If

            Console.WriteLine()
            Console.WriteLine("Run Complete in " & (Now - t).ToString)
            Console.WriteLine()

            For x = 1000 To 0 Step -1
                Threading.Thread.Sleep(10)
                Console.Write("Form will close in " & CInt(x / 100) & " press any key to skip")
                If Console.KeyAvailable Or UCase(Environment.UserName) = "DATACOLLSL" Then Exit Sub
                Console.CursorLeft = 0
            Next

            ' Catch ex As Exception
            '  FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\8ball\log.txt", Now() & "   " & ex.Message.ToString & " || " & ex.InnerException.ToString, True)
            '   MsgBox(ex.Message.ToString)
            '  MsgBox(ex.InnerException.ToString)
            ' End Try
            If Not IsNothing(IE) Then IE.Quit()
    End Sub


    Public Function GetLastUpdate(TableName As String) As Object
        GetLastUpdate = #1/1/2000#
        If IsNothing(UpdateTimes) Then Exit Function
        For x = 0 To UpdateTimes.Length - 1
            If UCase(UpdateTimes(x)(0)) = UCase(TableName) Then Return UpdateTimes(x)(1)
        Next
    End Function


    Private Function CheckIfRunning(ProcessName As String) As Integer
        Dim p() As Process = Process.GetProcessesByName(ProcessName)
        Return p.Count
    End Function

    Public Function ExecStoredProcedure(Procedurename As String, IsProcedure As Boolean, Optional Params As Object() = Nothing) As Object()()
        Dim StList As New List(Of Object())
        Using cn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(Procedurename, cn)
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

                        Do While DR.Read
                            Dim h(DR.VisibleFieldCount) As Object
                            DR.GetValues(h)
                            StList.Add(h)
                        Loop
                    End Using
                Catch EX As Exception : Finally
                    cn.Close()
                End Try
            End Using
        End Using
        ExecStoredProcedure = StList.ToArray
    End Function

    Public Function UpdateStatus(NewStatus As Integer, NewNotes As String, TableName As String, byuid As Boolean) As Guid
        'Debug.Print("INSERT INTO WFLOCAL..PullStatus VALUES (GETDATE(), '" & TableName & "', '" & Environment.UserName & "', '" & Environment.MachineName & "', '" & NewNotes & "', " & NewStatus & ", NEWID(), GETDATE())")
        If PullNumber = 0 Then PullNumber = ExecStoredProcedure("SELECT MAX(PULLNUMBER) FROM WFLOCAL..PULLSTATUS", False)(0)(0) + 1
        ExecStoredProcedure("INSERT INTO WFLOCAL..PullStatus VALUES (GETDATE(), '" & TableName & "', '" & Environment.UserName & "', '" & Environment.MachineName & "', '" & NewNotes & "', " & NewStatus & ", NEWID(), GETDATE(), " & PullNumber & ")", False)
        'Debug.Print("Select UID from wflocal..PullStatus WHERE TABLENAME = '" & TableName & "' AND PULLNOTES = '" & NewNotes & "' AND MACHINENAME = '" & Environment.MachineName & "' AND PULLSTATUS = " & NewStatus)
        'Return ExecStoredProcedure("Select UID from wflocal..PullStatus WHERE TABLENAME = '" & TableName & "' AND PULLNOTES = '" & NewNotes & "' AND MACHINENAME = '" & Environment.MachineName & "' AND PULLSTATUS = " & NewStatus, False)(0)(0)
        Return Guid.NewGuid
    End Function

    'Public Function UpdateStatus(NewStatus As Integer, NewNotes As String, uid As String) As Guid
    '    If uid <> "" Then ExecStoredProcedure("UPDATE WFLOCAL..PullStatus SET (TIMESTAMP = GETDATE(), PULLNOTES = '" & NewNotes & "', PULLSTATUS =" & NewStatus & "WHERE UID = '" & uid & "'", False)
    '    Return ExecStoredProcedure("Select UID from wflocal..PullStatus WHERE uid = '" & uid & "'", False)(0)(0)
    'End Function


    'Private Function wfLogin(wf As WebfocusModule, Optional CredentialsOnly As Boolean = False) As WebfocusModule
    '    If IsNothing(wf) Then wf = New WebfocusModule
    '    If Not wf.IsLoggedIn Then
    '        LogInInfo = GetUserPasswordandFex()
    '        If Not CredentialsOnly Then
    '            wf.LogIn("pprasinos", "Wyman123-")
    '            Do Until wf.IsLoggedIn
    '                LogInInfo = GetUserPasswordandFex()
    '                wf.LogIn(LogInInfo(0), LogInInfo(1))
    '            Loop
    '        End If
    '    End If
    '    Return wf
    'End Function

    'Private Sub FullUpdate(wf As WebfocusModule)
    '    wfLogin(wf)
    '    Dim PARTLIST As New List(Of String)
    '    Using cn As New SqlConnection(ConnectionString)
    '        Using cmd As New SqlCommand("", cn)
    '            cmd.CommandText = "Select DISTINCT PARTNO FROM WFLOCAL..CERT_ERRORS WHERE ISNULL(DAYS_IN_WC, 49) < 50 And PARTNO Not Like '%S'"
    '            cn.Open()
    '            Using DR As SqlClient.SqlDataReader = cmd.ExecuteReader
    '                Do While DR.Read
    '                    PARTLIST.Add(DR("PARTNO"))
    '                Loop
    '            End Using
    '            cn.Close()
    '            cmd.Parameters.Clear()
    '        End Using
    '    End Using
    '    PARTLIST.Sort()

    'For I = 0 To PARTLIST.Count - 1 Step 1
    '        Try
    '            Dim PART As String = PARTLIST(I)
    '            PART = Trim(PART)
    '            Dim WipHistoryRef As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23wipandshopco&IBIMR_fex=pprasino/wo_move_history_8ball_for_sql.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/workorder_moves.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&PARTNO=" & PART & "&WORP_MPV=ab_gbv&&IBIMR_random=13866&"
    '            wf.GetReporthAsync(WipHistoryRef, "wiphist")
    '            UpdateAppend(wf, GetWFIds(wf.GetRequests))
    '            wf = Nothing
    '            wf = New WebfocusModule
    '            wf.LogIn("PPRASINOS", "Wyman123-")
    '        Catch EX1 As Exception
    '            Stop
    '        End Try
    '    Next


    'For q = 0 To 20
    '    Console.Write(q & " ")
    '    Dim Span As Integer = 5
    '    beforeDate = MakeWebfocusDate(Today.AddDays(-q * Span))
    '    afterDate = MakeWebfocusDate(Today.AddDays(-1 - ((q + 1) * Span)))
    '    Console.WriteLine(beforeDate & "-" & afterDate)
    '    Dim InvRef1 As String = "qavistes/qavistes.htm#wipandshopco    pprasinos:pprasino/inventorybyms.fex "

    '    Dim LaborRef1 As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23laborreporti&IBIMR_fex=pprasino/labor_part_detail_workorders_with_esh_for_sql_for_testing.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/labor_part_detail_workorders_with_esh.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&GECHARGE_DATE=" & afterDate & "&LECHARGE_DATE=" & beforeDate & "&IBIMR_random=24311&"
    '    Dim TputRef1 As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23thruputrepor&IBIMR_fex=pprasino/esh_and_tput_for_flex_for_sql.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/thruput_detail_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&TP_DATE_COMPELTED=" & afterDate & "&LE_TP_DATE_COMPELTED=" & beforeDate & "&IBIMR_random=31846"
    '    Dim ScrapRef1 As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23scrapdatatqg&IBIMR_fex=pprasino/scrap_report_including_nodefect.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/scrap_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&DISP_D=" & afterDate & "&LEDISP_D=" & beforeDate & "&IBIMR_random=96021"
    '    TputRef1 = Replace(TputRef1, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))

    '    If Today.DayOfWeek = DayOfWeek.Monday Then wf.GetReporthAsync(TputRef1, "tput")
    '    If Today.DayOfWeek = DayOfWeek.Tuesday Then wf.GetReporthAsync(ScrapRef1, "scrap")
    '    If Today.DayOfWeek = DayOfWeek.Wednesday Then wf.GetReporthAsync(LaborRef1, "labor")

    '    '  Threading.Thread.Sleep(2000)
    '    UpdateAppend(wf, GetWFIds(wf.GetRequests))
    '    ' Threading.Thread.Sleep(1000)

    '    wf = Nothing
    '    wf = wfLogin(wf)

    'Next q

    'End Sub


    Private Function GetPingMs(ByRef hostNameOrAddress As String)
        Dim ping As New System.Net.NetworkInformation.Ping
        Return ping.Send(hostNameOrAddress).RoundtripTime
        Threading.Thread.Sleep(1000)
    End Function

    'Private Function GetUserPasswordandFex() As String()
    '    Dim h As New Random
    '    Dim Usernames() As String = {"hfaizi", "mreyes", "MALMARAZ", "MARJMAND", "HYANG", "GWONG", "VDELACRUZ", "JTIBAYAN", "JSOLIS", "ASINGH", "GREYES", "JPIMENTEL", "TOSULLIVAN", "MMARTIN", "VLOPEZ", "SLI", "JIMPERIAL", "JHERNANDEZ", "FHARO", "CGOUTAMA", "HGOMEZ", "EGONZALEZ", "CDAROSA"}

    '    Dim y As Integer = h.Next(0, Usernames.Length)
    '    Dim ps As String
    '    Usernames(y) = "pprasinos"
    '    Dim FexAdd As String = "&IBIMR_sub_action=MR_MY_REPORT"
    '    If Usernames(y) <> "pprasinos" Then
    '        FexAdd = "&IBIMR_sub_action=MR_MY_REPORT&IBIMR_proxy_id=pprasino.htm&"
    '        ps = ChrW(112) & ChrW(97) & ChrW(115) & ChrW(115) & ChrW(50) & ChrW(48) & ChrW(49) & ChrW(53)
    '    Else
    '        ps = ChrW(87) & ChrW(121) & ChrW(109) & ChrW(97) & ChrW(110) & ChrW(49) & ChrW(50) & ChrW(51) & ChrW(45)
    '    End If
    '    Return {Usernames(y), ps, FexAdd}
    'End Function


    'Private Function GetWFIds(Requests As String, Optional notarray As Boolean = False) As String()
    '    Dim k() As String = Split(Requests, vbLf)
    '    For w = 0 To k.Length - 1
    '        k(w) = Mid(k(w), 3, 10)
    '        If Left(k(w), 1) <> "" Then k(w) = Right(k(w), Len(k(w)) - 1)
    '        If Left(k(w), 1) <> "" Then k(w) = Right(k(w), Len(k(w)) - 1)
    '        If notarray Then
    '            k(0) = Replace(k(w), " ", "") & "  "
    '        Else
    '            k(w) = Replace(k(w), " ", "")
    '        End If
    '    Next
    '    Return k

    'End Function

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
                                    where a.OPERATIONNO<b.OPERATION"

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

    Private Function GetWFReport(ref As String) As String()()
        'Debug.Print(ref)
        ' Debug.Print("")
        Dim doc As mshtml.HTMLDocument
        Try
            If IsNothing(IE) Then
                IE = New SHDocVw.InternetExplorerMedium
                IE.Visible = True
                '      Sleep(1000)
                '      'http://webfocus.pccstructurals.com/ibi_apps/
                IE.Navigate("http://webfocus.pccstructurals.com/ibi_apps/signin")
                Sleep(3000)
                For x = 0 To 10
                    Do Until IE.Busy = False And IE.ReadyState = 4 : Debug.Print(IE.ReadyState) : Debug.Print(IE.Busy) : Sleep(40) : Loop : Sleep(500)
                Next x

                doc = IE.Document
                doc.getElementById("SignonUserName").innerText = "gen_slan-8ball"
                doc.getElementById("SignonPassName").innerText = "Password17"
                doc.getElementById("SignonbtnLoginID").click()
                For x = 0 To 10
                    Do Until IE.Busy = False And IE.ReadyState = 4 : Debug.Print(IE.ReadyState) : Debug.Print(IE.Busy) : Sleep(40) : Loop : Sleep(500)
                Next x
            End If

            IE.Navigate(ref)

            For X = 0 To 10
                Do Until IE.Busy = False And IE.ReadyState = 4 : Sleep(10) : Loop : Sleep(100)
            Next X
            doc = IE.Document
            Dim i As Integer = 0
            Debug.Print(doc.all.length)
            If doc.all.length < 100 Then

                Do While i < doc.all.length
                    Dim element As Object = doc.all(i)
                    Try
                        If Not IsNothing(element.innerhtml) Then
                            'If Not IsNothing(element.id) Then Debug.Print("ID:  " & element.id)
                            ' If Not IsNothing(element.title) Then Debug.Print("TITLE:  " & element.title)
                            If InStr(element.innerhtml, "win.document.form1.action = ") > 0 Then
                                Dim RepURL As String = "http://webfocus" & Split(Split(element.innerhtml, "win.document.form1.action = " & Chr(34))(1), Chr(34) & ";")(0)
                                'Debug.Print(RepURL)
                                IE.Navigate(RepURL)
                                For X = 0 To 10
                                    Do Until IE.Busy = False And IE.ReadyState = 4 : Sleep(40) : Loop : Sleep(500)
                                Next X
                                Threading.Thread.Sleep(1000)
                                i = 1000000
                            End If
                        End If
                    Catch
                    End Try
                    i = i + 1
                Loop
            End If
            For X = 0 To 10
                Do Until IE.Busy = False And IE.ReadyState = 4 : Sleep(40) : Loop : Sleep(500)
            Next X
            doc = IE.Document
            Dim doc1 As String = doc.body.outerHTML
            IE.Navigate("http://webfocus.pccstructurals.com/ibi_apps/bip/portal/PCCStructuralsInc")

            Return ClassLibrary1.HTMLProcessor.ParseHtml(doc1)
        Catch ex As Exception
            IE.Visible = True

            FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\8ball\updater\error" & Day(Now) & Hour(Now) & Minute(Now) & ".txt", "ERROR ON LINE " & Erl() & vbCrLf & vbCrLf & ex.Message.ToString & vbCrLf & vbCrLf & vbCrLf & ex.InnerException.ToString, True)
        End Try
    End Function


    Private Sub UpdateAppend(ref As String, RespNames As String)
        Dim tab As String = ""
        Dim RefFind() As String = {"ships", "fingoods", "lots", "certs", "scrap", "partdata", "xtl", "tput", "labor", "labor1", "wiphist"}
        Dim TableNames() As String = {"SHIPMENTS", "CERT_ERRORS", "CERT_ERRORS", "CERT_ERRORS", "SCRAP", "ALLOYS", "TIMELINE", "TPUT", "LABOR", "LABOR", "WIP_MOVE_HIST"}
        Dim UpdatedRows As Integer = 0
        Using cn As New SqlConnection(ConnectionString)
            cn.Open()
            Try
                Using cmd As New SqlCommand("", cn)
                    cmd.CommandTimeout = 5
                    cmd.CommandType = CommandType.Text
                    '#updates one record so that other machines do not start a pull while one is waiting for a report
                    If InStr(RespNames, "lots") <> 0 Then
                        cmd.CommandText = "UPDATE WFLOCAL.DBO.CERT_ERRORS SET ACTIVE = 2 WHERE ACTIVE <> 0"
                        cmd.ExecuteNonQuery()
                    ElseIf InStr(RespNames, "ships") <> 0 Then
                        cmd.CommandText = "UPDATE WFLOCAL.DBO.SHIPMENTS SET INVOICE_NO = 'PACK(1).pdf' WHERE INVOICE_NO = 'PACK(1).pdf'"
                        cmd.ExecuteNonQuery()
                    ElseIf InStr(RespNames, "tput") Then
                        cmd.CommandText = "UPDATE WFLOCAL.DBO.TPUT SET TPUT_VALUE = 0 WHERE ESH = 7.9144 AND WORKORDERNO = '1012548-00169' "
                        cmd.ExecuteNonQuery()
                    End If


                    If RespNames = Nothing Or RespNames = "opens" Then GoTo NEXTP
                    Dim j As New Object
                    j = GetWFReport(ref)

                    Dim TableName As String = ""

                    For ind = 0 To RefFind.Length - 1
                        If RefFind(ind) = RespNames Then TableName = TableNames(ind) & DebugText
                    Next

                    Dim UID As Guid
                    Try : UID = UpdateStatus(2, "RECIEVED", TableName, False) : Catch : End Try

                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "SELECT column_name, data_type FROM wflocal.INFORMATION_SCHEMA.COLUMNS" & vbCrLf &
                            "WHERE wflocal.INFORMATION_SCHEMA.COLUMNS.TABLE_NAME='" & TableName & "'"

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
                    If TableName = "SCRAP" & DebugText Then
                        cmd.CommandText = "WFLOCAL.DBO.UpdateScrap"
                    ElseIf TableName = "CERT_ERRORS" & DebugText Then
                        cmd.CommandText = "WFLOCAL.DBO.UPDATEAPPENDWIP"
                    ElseIf TableName = "TIMELINE" & DebugText Then
                        cmd.CommandText = "WFLOCAL.DBO.XTLupdateAppend"
                    ElseIf TableName = "ALLOYS" & DebugText Then
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
                                                    values (@PARTNO, @ALLOY_DESCR, @MATERIAL_SPEC, @PART_DESCR, @PIECES_PER_MOLD, @SELLING_PRICE, @POUR_WEIGHT, @STOP_RELEASE, @PART_STATUS, @ROUT_REV, @SHIP_WEIGHT);"

                        cmd.CommandType = CommandType.Text
                    ElseIf TableName = "SHIPMENTS" & DebugText Then
                        cmd.CommandText = "WFLOCAL.DBO.AddShipments"
                    ElseIf TableName = "TPUT" & DebugText Then
                        cmd.CommandText = "WFLOCAL.DBO.UpdateThruput"
                    ElseIf TableName = "LABOR" & DebugText Then
                        cmd.CommandText = "WFLOCAL.DBO.UpdateLabor"
                    ElseIf TableName = "WIP_MOVE_HIST" & DebugText Then
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
                                    'Debug.Print(j(RowNum)(Col(2)))
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
                                    If Not Replace(j(RowNum)(Col(2)), ",", "") = "." Then
                                        dt = DateTime.Parse(j(RowNum)(Col(2)))
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
                        cmd.ExecuteNonQuery()
                        CT = CT + 1
                        Console.CursorLeft = 0
                        Console.Write(CT & "/" & j.length & "        ")
                    Next
                    Console.CursorLeft = 20
                    Console.WriteLine(TableName & " UPDATED Using " & RespNames)
                    tab = TableName
NEXTP:

                    Try : UpdateStatus(3, "UPDATED", tab, False) : Catch : End Try

                    If InStr(RespNames, "lots") <> 0 Then
                        cmd.CommandType = CommandType.Text
                        cmd.CommandText = "UPDATE WFLOCAL.DBO.CERT_ERRORS Set ACTIVE = 0 WHERE ACTIVE = 2"
                        Dim RWS As Integer = cmd.ExecuteNonQuery()
                        UpdatedRows = UpdatedRows + RWS
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "wflocal..cleanup"
                        cmd.Parameters.Clear()
                        cmd.ExecuteNonQuery()
                    End If
                End Using

            Catch ex As Exception
                MsgBox(ex.ToString)
                MsgBox(ex.InnerException.ToString)
            End Try
        End Using
    End Sub


    '    Private Sub UpdateAppend(WF As WebfocusDLL.WebfocusModule, RespNames() As String)
    '        Dim tab As String = ""
    '        Dim RefFind() As String = {"ships", "fingoods", "lots", "certs", "scrap", "partdata", "xtl", "tput", "labor", "labor1", "wiphist"}
    '        Dim TableNames() As String = {"SHIPMENTS", "CERT_ERRORS", "CERT_ERRORS", "CERT_ERRORS", "SCRAP", "ALLOYS", "TIMELINE", "TPUT", "LABOR", "LABOR", "WIP_MOVE_HIST"}
    '        Dim UpdatedRows As Integer = 0
    '        Using cn As New SqlConnection(ConnectionString)
    '            cn.Open()
    '            Try
    '                Using cmd As New SqlCommand("", cn)
    '                    cmd.CommandTimeout = 5
    '                    cmd.CommandType = CommandType.Text
    '                    '#updates one record so that other machines do not start a pull while one is waiting for a report
    '                    If InStr(WF.GetRequests, "lots") <> 0 Then
    '                        cmd.CommandText = "UPDATE WFLOCAL.DBO.CERT_ERRORS SET ACTIVE = 2 WHERE ACTIVE <> 0"
    '                        cmd.ExecuteNonQuery()
    '                    ElseIf InStr(WF.GetRequests, "ships") <> 0 Then
    '                        cmd.CommandText = "UPDATE WFLOCAL.DBO.SHIPMENTS SET INVOICE_NO = 'PACK(1).pdf' WHERE INVOICE_NO = 'PACK(1).pdf'"
    '                        cmd.ExecuteNonQuery()
    '                    ElseIf InStr(WF.GetRequests, "tput") Then
    '                        cmd.CommandText = "UPDATE WFLOCAL.DBO.TPUT SET TPUT_VALUE = 0 WHERE ESH = 7.9144 AND WORKORDERNO = '1012548-00169' "
    '                        cmd.ExecuteNonQuery()
    '                    End If

    '                    For P = 0 To RespNames.Length - 1
    '                        If RespNames(P) = Nothing Or RespNames(P) = "opens" Then GoTo NEXTP
    '                        Dim j As New Object
    '                        j = WF.GetResponse(RespNames(P)).Response

    '                        Dim TableName As String = ""

    '                        For ind = 0 To RefFind.Length - 1
    '                            If RefFind(ind) = RespNames(P) Then TableName = TableNames(ind) & DebugText
    '                        Next

    '                        Dim UID As Guid
    '                        Try : UID = UpdateStatus(2, "RECIEVED", TableName, False) : Catch : End Try

    '                        cmd.CommandType = CommandType.Text
    '                        cmd.CommandText = "SELECT column_name, data_type FROM wflocal.INFORMATION_SCHEMA.COLUMNS" & vbCrLf &
    '                            "WHERE wflocal.INFORMATION_SCHEMA.COLUMNS.TABLE_NAME='" & TableName & "'"

    '                        Dim ColumnInfo As New List(Of String())
    '                        Dim CSVColumns As String = ""
    '                        Dim CSVUPDATE As String = ""

    '                        Using dr As SqlDataReader = cmd.ExecuteReader
    '                            While dr.Read()
    '                                Dim Y As Integer = GetColumnNumber(j, dr("column_name").ToString)
    '                                If Y <> -1 Then
    '                                    ColumnInfo.Add({dr("column_name").ToString, dr("data_type").ToString, Y})
    '                                    CSVColumns = CSVColumns & "ROW." & dr("column_name").ToString & ", "
    '                                    CSVUPDATE = CSVUPDATE & dr("column_name").ToString & " = @" & dr("column_name").ToString & ","
    '                                End If
    '                            End While
    '                        End Using
    '                        ColumnInfo.Add({"ACTIVE", "int", 0})
    '                        CSVColumns = CSVColumns & "ROW.ACTIVE, "
    '                        CSVUPDATE = CSVUPDATE & "ROW.ACTIVE = @ACTIVE,"
    '                        CSVColumns = Left(CSVColumns, Len(CSVColumns) - 2)
    '                        CSVUPDATE = Left(CSVUPDATE, Len(CSVUPDATE) - 1)

    '                        cmd.CommandType = CommandType.StoredProcedure
    '                        If TableName = "SCRAP" & DebugText Then
    '                            cmd.CommandText = "WFLOCAL.DBO.UpdateScrap"
    '                        ElseIf TableName = "CERT_ERRORS" & DebugText Then
    '                            cmd.CommandText = "WFLOCAL.DBO.UPDATEAPPENDWIP"
    '                        ElseIf TableName = "TIMELINE" & DebugText Then
    '                            cmd.CommandText = "WFLOCAL.DBO.XTLupdateAppend"
    '                        ElseIf TableName = "ALLOYS" & DebugText Then
    '                            cmd.CommandText = "	MERGE WFLOCAL..ALLOYS AS TARGET
    '                                            USING (SELECT @PARTNO, @ALLOY_DESCR, @MATERIAL_SPEC, @PART_DESCR, @PIECES_PER_MOLD, @SELLING_PRICE, @POUR_WEIGHT, @STOP_RELEASE, @PART_STATUS, @ROUT_REV, @SHIP_WEIGHT) AS SOURCE 
    '                                                          (PARTNO,  ALLOY_DESCR,  MATERIAL_SPEC,  PART_DESCR,  PIECES_PER_MOLD,  SELLING_PRICE,  POUR_WEIGHT,  STOP_RELEASE,  PART_STATUS, ROUT_REV, SHIP_WEIGHT)
    '                                            ON TARGET.PARTNO=SOURCE.PARTNO
    '                                            WHEN MATCHED THEN
    '                                                    UPDATE SET	ALLOY_DESCR=SOURCE.ALLOY_DESCR,  
    '                                                                MATERIAL_SPEC=SOURCE.MATERIAL_SPEC,
    '                                                                PART_DESCR=SOURCE.PART_DESCR,
    '                                                                PIECES_PER_MOLD=SOURCE.PIECES_PER_MOLD,
    '                                                                SELLING_PRICE=SOURCE.SELLING_PRICE,
    '                                                                POUR_WEIGHT=SOURCE.POUR_WEIGHT,
    '                                                                STOP_RELEASE=SOURCE.STOP_RELEASE,
    '                                                                PART_STATUS=SOURCE.PART_STATUS,
    '                                                                ROUT_REV=SOURCE.ROUT_REV,
    '                                                                SHIP_WEIGHT = SOURCE.SHIP_WEIGHT
    '                                            WHEN NOT MATCHED THEN
    '                                                    INSERT (PARTNO,  ALLOY_DESCR,  MATERIAL_SPEC,  PART_DESCR,  PIECES_PER_MOLD,  SELLING_PRICE,  POUR_WEIGHT,  STOP_RELEASE,  PART_STATUS, ROUT_REV, SHIP_WEIGHT)
    '                                                    values (@PARTNO, @ALLOY_DESCR, @MATERIAL_SPEC, @PART_DESCR, @PIECES_PER_MOLD, @SELLING_PRICE, @POUR_WEIGHT, @STOP_RELEASE, @PART_STATUS, @ROUT_REV, @SHIP_WEIGHT);"

    '                            cmd.CommandType = CommandType.Text
    '                        ElseIf TableName = "SHIPMENTS" & DebugText Then
    '                            cmd.CommandText = "WFLOCAL.DBO.AddShipments"
    '                        ElseIf TableName = "TPUT" & DebugText Then
    '                            cmd.CommandText = "WFLOCAL.DBO.UpdateThruput"
    '                        ElseIf TableName = "LABOR" & DebugText Then
    '                            cmd.CommandText = "WFLOCAL.DBO.UpdateLabor"
    '                        ElseIf TableName = "WIP_MOVE_HIST" & DebugText Then
    '                            cmd.CommandText = "WFLOCAL.DBO.UPDATE_WIP_HIST"
    '                        End If

    '                        Dim CT As Long = 1
    '                        Console.Write("0")
    '                        Console.CursorLeft = 0

    '                        For RowNum = 1 To j.length - 1
    '                            With cmd.Parameters
    '                                .Clear()
    '                                For Each Col In ColumnInfo
    '                                    If Col(1) = "nvarchar" Or Col(1) = "nchar" Then
    '                                        .Add("@" & Col(0), SqlDbType.NVarChar).Value = j(RowNum)(Col(2))
    '                                    ElseIf Col(1) = "float" And Col(0) <> "ACTIVE" Then
    '                                        Debug.Print(j(RowNum)(Col(2)))
    '                                        j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), "R", "1")
    '                                        j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), "N", "0")
    '                                        j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), "Y", "2")

    '                                        If Replace(j(RowNum)(Col(2)), ",", "") = "." Then
    '                                            .Add("@" & Col(0), SqlDbType.Float).Value = 0
    '                                        Else
    '                                            Dim s As Double = j(RowNum)(Col(2))
    '                                            .Add("@" & Col(0), SqlDbType.Float).Value = s
    '                                        End If
    '                                    ElseIf InStr(Col(1), "smallint", CompareMethod.Text) <> 0 Then
    '                                        If Replace(j(RowNum)(Col(2)), ",", "") = "." Then
    '                                            .Add("@" & Col(0), SqlDbType.SmallInt).Value = 0
    '                                        Else
    '                                            Dim S As Int16 = Replace(j(RowNum)(Col(2)), ",", "") * 1
    '                                            .Add("@" & Col(0), SqlDbType.SmallInt).Value = S
    '                                        End If
    '                                    ElseIf InStr(Col(1), "int", CompareMethod.Text) <> 0 And Col(0) <> "ACTIVE" Then
    '                                        If Replace(j(RowNum)(Col(2)), ",", "") = "." Then
    '                                            .Add("@" & Col(0), SqlDbType.Int).Value = 0
    '                                        Else
    '                                            Dim S As Integer = (0 & Replace(Replace(j(RowNum)(Col(2)), ",", ""), ".", "")) * 1
    '                                            .AddWithValue("@" & Col(0), S)
    '                                        End If
    '                                    ElseIf InStr(Col(1), "Date", CompareMethod.Text) <> 0 Then
    '                                        Dim dt As DateTime = #1/1/1900#
    '                                        If Not Replace(j(RowNum)(Col(2)), ",", "") = "." Then
    '                                            dt = DateTime.Parse(j(RowNum)(Col(2)))
    '                                            If dt.Year > 1900 Then
    '                                                .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
    '                                            Else
    '                                                dt = Now.AddYears(-100)
    '                                                .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
    '                                            End If
    '                                        End If
    '                                    End If
    '                                Next
    '                                .Add("@ACTIVE", SqlDbType.Int).Value = 1
    '                            End With
    '                            cmd.ExecuteNonQuery()
    '                            CT = CT + 1
    '                            Console.CursorLeft = 0
    '                            Console.Write(CT & "/" & j.length & "        ")
    '                        Next
    '                        Console.CursorLeft = 20
    '                        Console.WriteLine(TableName & " UPDATED Using " & RespNames(P))
    '                        tab = TableName
    'NEXTP:
    '                    Next P
    '                    Try : UpdateStatus(3, "UPDATED", tab, False) : Catch : End Try

    '                    If InStr(WF.GetRequests, "lots") <> 0 Then
    '                        cmd.CommandType = CommandType.Text
    '                        cmd.CommandText = "UPDATE WFLOCAL.DBO.CERT_ERRORS Set ACTIVE = 0 WHERE ACTIVE = 2"
    '                        Dim RWS As Integer = cmd.ExecuteNonQuery()
    '                        UpdatedRows = UpdatedRows + RWS
    '                        cmd.CommandType = CommandType.StoredProcedure
    '                        cmd.CommandText = "wflocal..cleanup"
    '                        cmd.Parameters.Clear()
    '                        cmd.ExecuteNonQuery()
    '                    End If
    '                End Using

    '            Catch ex As Exception
    '                MsgBox(ex.ToString)
    '                MsgBox(ex.InnerException.ToString)
    '            End Try
    '        End Using
    '    End Sub


    'Sub OpensUpdater(wf As WebfocusModule)

    '    Dim j As Object = wf.GetResponse("opens").Response
    '    UpdateStatus(2, "RECIEVED", "OPEN_ORDERS", False)
    '    Using cn As New SqlConnection(ConnectionString)
    '        Try
    '            cn.Open()
    '            Using cmd As New SqlCommand("", cn)
    '                cmd.CommandTimeout = 5

    '                cmd.CommandText = "UPDATE wflocal.dbo.OPEN_ORDERS Set ACTIVE = 2 WHERE ACTIVE <> 0"
    '                cmd.ExecuteNonQuery()
    '                cmd.CommandText = " Select column_name, data_type 
    '                                    FROM WFLOCAL.INFORMATION_SCHEMA.COLUMNS
    '                                    WHERE WFLOCAL.INFORMATION_SCHEMA.COLUMNS.TABLE_NAME ='OPEN_ORDERS'"

    '                Dim ColumnInfo As New List(Of String())
    '                Dim ColNumbers As New List(Of Integer)

    '                Using dr As SqlDataReader = cmd.ExecuteReader
    '                    While dr.Read()
    '                        Dim Y As Integer = GetColumnNumber(j, dr("column_name").ToString)
    '                        If Y <> -1 Then
    '                            ColumnInfo.Add({dr("column_name").ToString, dr("data_type").ToString, Y})
    '                        End If
    '                    End While
    '                End Using

    '                For RowNum = 1 To j.length - 1
    '                    With cmd.Parameters
    '                        .Clear()
    '                        For Each Col In ColumnInfo
    '                            If Col(1) = "nvarchar" Then
    '                                .Add("@" & Col(0), SqlDbType.NVarChar).Value = j(RowNum)(Col(2))
    '                            ElseIf Col(1) = "float" Then
    '                                .Add("@" & Col(0), SqlDbType.Float).Value = Replace(j(RowNum)(Col(2)), ",", "")
    '                            ElseIf Col(1) = "datetime" Then
    '                                Dim dt As DateTime = DateTime.Parse(j(RowNum)(Col(2)))
    '                                If dt.Year > 1900 Then
    '                                    .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
    '                                Else
    '                                    dt = Now.AddYears(-100)
    '                                    .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
    '                                End If
    '                            End If
    '                        Next Col
    '                        .AddWithValue("@ACTIVE", 1)
    '                    End With
    '                    Console.CursorLeft = 0
    '                    Console.Write(RowNum + 1 & "/" & j.length & "       ")

    '                    cmd.CommandType = CommandType.StoredProcedure
    '                    cmd.CommandText = "WFLOCAL.DBO.OPENUPDATER"
    '                    Dim y As Integer = cmd.ExecuteNonQuery()
    '                    'If y <> -1 Then Stop
    '                Next RowNum
    '                cmd.CommandType = CommandType.Text
    '                cmd.CommandText = "UPDATE wflocal.dbo.OPEN_ORDERS SET ACTIVE = 0 WHERE ACTIVE = 2"
    '                cmd.ExecuteNonQuery()
    '                cmd.CommandText = "INSERT INTO WFLOCAL.DBO.PO_REVIEW  (SALES_ORDER_NO, CUST_NO, SALES, USERNAME, ttimestamp, prel, pship, erel, eship)
    '                                Select DISTINCT B.SALES_ORDER_NO, B.CUSTOMER_NO, B.ADDED_BY, B.ADDED_BY, getdate(), 1, 1, 1, 1
    '                                From DBO.OPEN_ORDERS B
    '                                Where Not EXISTS(Select distinct  B.SALES_ORDER_NO
    '                                From DBO.PO_REVIEW
    '                                Where PO_REVIEW.SALES_ORDER_NO = B.SALES_ORDER_NO)"
    '                cmd.ExecuteNonQuery()
    '                cmd.Parameters.Clear()
    '                cmd.CommandType = CommandType.StoredProcedure
    '                cmd.CommandText = "wflocal.dbo.CleanTickets"
    '                cmd.ExecuteNonQuery()
    '            End Using
    '            Console.CursorLeft = 20
    '            Console.WriteLine("OPEN_ORDERS" & " UPDATED Using " & "opens")
    '        Catch ex As Exception
    '            MsgBox(ex.GetType().ToString)
    '            MsgBox(ex.Message.ToString)
    '            MsgBox(ex.InnerException.ToString)
    '        End Try
    '    End Using
    '    UpdateStatus(3, "UPDATED", "OPEN_ORDERS", False)
    'End Sub

    Sub OpensUpdater(ref As String)

        Dim j As Object = GetWFReport(ref)
        UpdateStatus(2, "RECIEVED", "OPEN_ORDERS", False)
        Using cn As New SqlConnection(ConnectionString)
            Try
                cn.Open()
                Using cmd As New SqlCommand("", cn)
                    cmd.CommandTimeout = 5

                    cmd.CommandText = "UPDATE wflocal.dbo.OPEN_ORDERS Set ACTIVE = 2 WHERE ACTIVE <> 0"
                    cmd.ExecuteNonQuery()
                    cmd.CommandText = " Select column_name, data_type 
                                        FROM WFLOCAL.INFORMATION_SCHEMA.COLUMNS
                                        WHERE WFLOCAL.INFORMATION_SCHEMA.COLUMNS.TABLE_NAME ='OPEN_ORDERS'"

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

                    For RowNum = 1 To j.length - 1
                        With cmd.Parameters
                            .Clear()
                            For Each Col In ColumnInfo
                                If Col(1) = "nvarchar" Then
                                    .Add("@" & Col(0), SqlDbType.NVarChar).Value = j(RowNum)(Col(2))
                                ElseIf Col(1) = "float" Then
                                    .Add("@" & Col(0), SqlDbType.Float).Value = Replace(j(RowNum)(Col(2)), ",", "")
                                ElseIf Col(1) = "datetime" Then
                                    Dim dt As DateTime = DateTime.Parse(j(RowNum)(Col(2)))
                                    If dt.Year > 1900 Then
                                        .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
                                    Else
                                        dt = Now.AddYears(-100)
                                        .Add("@" & Col(0), SqlDbType.DateTime).Value = dt
                                    End If
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
                    cmd.CommandText = "INSERT INTO WFLOCAL.DBO.PO_REVIEW  (SALES_ORDER_NO, CUST_NO, SALES, USERNAME, ttimestamp, prel, pship, erel, eship)
                                    Select DISTINCT B.SALES_ORDER_NO, B.CUSTOMER_NO, B.ADDED_BY, B.ADDED_BY, getdate(), 1, 1, 1, 1
                                    From DBO.OPEN_ORDERS B
                                    Where Not EXISTS(Select distinct  B.SALES_ORDER_NO
                                    From DBO.PO_REVIEW
                                    Where PO_REVIEW.SALES_ORDER_NO = B.SALES_ORDER_NO)"
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "wflocal.dbo.CleanTickets"
                    cmd.ExecuteNonQuery()
                End Using
                Console.CursorLeft = 20
                Console.WriteLine("OPEN_ORDERS" & " UPDATED Using " & "opens")
            Catch ex As Exception
                MsgBox(ex.GetType().ToString)
                MsgBox(ex.Message.ToString)
                MsgBox(ex.InnerException.ToString)
            End Try
        End Using
        UpdateStatus(3, "UPDATED", "OPEN_ORDERS", False)
    End Sub

    Private Function GetColumnNumber(InputTable()() As String, ColumLabel As String) As Integer
        Dim x As Integer = 0
        Do While ColumLabel <> InputTable(0)(x) And x < UBound(InputTable(0))
            x = x + 1
        Loop
        If x = UBound(InputTable(0)) And ColumLabel <> InputTable(0)(x) Then
            Return -1
        Else
            Return x
        End If
    End Function

End Module
