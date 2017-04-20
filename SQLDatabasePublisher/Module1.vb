
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
    '######WEBFOCUSPULL SET TO FALSE TO ALLOW REPORTS FROM \\slfs01\public\visdownloads  ##############
    Dim WebFocusPull As Boolean = True

    Sub Main()
        'If UCase(Environment.UserName) <> "PPRASINOS" Then Exit Sub
        ' If Not InStr(Environment.MachineName, "3DP") > 0 Then Exit Sub
        Console.WriteLine("=====Do not close or disconnect from network until run complete=====")
        Console.WriteLine()
        Console.WriteLine("Started at " & Now)
        Dim theProcesses() As Process = System.Diagnostics.Process.GetProcessesByName("iexplore")
        For Each currentProcess As Process In theProcesses

            'Get the currentProcess MainWindowTitle and see if that title matches the title of the action cancelled IE instance:
            currentProcess.Kill()
        Next
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
            AfterDate = MakeWebfocusDate(Today.AddDays(-10))
        Else
            adj = adj - 5
        End If

        '#####if there is an error in this program more than 2 times, the DataColl computer will be restarted#####
        If Environment.MachineName = "SLREPORT01" Or UCase(Environment.UserName) = "DATACOLLSL" Or UCase(Environment.UserName) = "PPRASINOS" Then
            If CheckIfRunning("SQLDatabasePublisher") > 1 Or Hour(Now) = 10 Then
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
        AfterDate = MakeWebfocusDate(Today.AddDays(-7))
        UpdateTimes(1)(0) = "TPUT" : UpdateTimes(1)(1) = Today.AddDays(-2)
        UpdateTimes(2)(0) = "SHIPMENTS" : UpdateTimes(2)(1) = Today.AddDays(-2)
        'UpdateTimes(3)(0) = "CERT_ERRORS" : UpdateTimes(3)(1) = Today.AddDays(-2)
        'Try
        Dim OpensRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=custom_open_order_reportshtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=209&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=13377"
        Dim TputRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=ESH_and_TPUT_FOR_FLEX_for_sql.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=340&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&LE_TP_DATE_COMPELTED=" + BeforeDate + "&TP_DATE_COMPELTED=" + AfterDate + "&BIP_rand=21066"
        Dim ScrapRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=scrap_report.fex&WF_STYLE_HEIGHT=140&WF_STYLE_WIDTH=330&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&DISP_D=" + AfterDate + "&LEDISP_D=" + BeforeDate + "&BIP_rand=74775"
        Dim ShipRef As String = "Http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=full_shipreport_by_lothtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=429&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&&SHIPPED_D=" + AfterDate + "&BIP_rand=7574"
        Dim WIPRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=customlotshtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=429&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=70094"
        Dim FGRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=fingoodshtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=429&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=36829"
        Dim CDCSRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=sl_wipfg_quality_check_inspbeyondhtml.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=209&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=20390"
        Dim TimeLineRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=ltsshtml.fex&WF_STYLE_HEIGHT=140&WF_STYLE_WIDTH=330&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&BIP_rand=83356"
        Dim LaborRef As String = "http://webfocus.pccstructurals.com/ibi_apps/run.bip?BIP_REQUEST_TYPE=BIP_RUN&BIP_folder=IBFS%253A%252FWFC%252FRepository%252Fqavistes%252F~gen_slan-8ball&BIP_item=Labor_Part_Detail_ESH_FOR_SQL_for_testing.fex&WF_STYLE_HEIGHT=353&WF_STYLE_WIDTH=209&WF_STYLE_UNITS=PIXELS&IBIWF_redirNewWindow=true&WF_STYLE=IBFS%3A%2FFILE%2FIBI_HTML_DIR%2Fjavaassist%2Fintl%2FEN%2Fcombine_templates%2FENInformationBuilders_Medium1.sty&WF_THEME=BIPFlat&BIP_CACHE=100000&GECHARGE_DATE=" & AfterDate & "&LECHARGE_DATE=" & BeforeDate
        
        If Not WebFocusPull Then
            OpensRef = "\\slfs01\public\VisDownloads\Sales\slan_open_orders.csv"
            TputRef = "\\slfs01\public\VisDownloads\TOC\slan_toc_esh.csv"
            ScrapRef = "\\slfs01\public\VisDownloads\WIP\slan_scrap.csv"
            ShipRef = " \\slfs01\public\VisDownloads\Sales\slan_shipments.csv"
            WIPRef = "\\slfs01\public\VisDownloads\WIP\slan_wip.csv"
            FGRef = "\\slfs01\public\VisDownloads\WIP\Slan_fg.csv"
            CDCSRef = "\\slfs01\public\VisDownloads\WIP\slan_quality_check.csv"
            TimeLineRef = "\\slfs01\public\VisDownloads\Part\slan_part_lt.csv"
            LaborRef = "\\slfs01\public\VisDownloads\Labor\slan_wo_labor.csv"
        End if

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
            'UpdateAppend(ScrapRef, "scrap")
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

   


    Private Function GetPingMs(ByRef hostNameOrAddress As String)
        Dim ping As New System.Net.NetworkInformation.Ping
        Return ping.Send(hostNameOrAddress).RoundtripTime
        Threading.Thread.Sleep(1000)
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
                                    where a.OPERATIONNO<b.OPERATION"

                cmd.Connection = cn
                cn.Open()
                Using DR As SqlClient.SqlDataReader = cmd.ExecuteReader
                    Do While DR.Read
                        Dim MsgString As String = "This email Is to notify that lot " & DR("WORKORDERNO").ToString & " has reached Or passed operation " & DR("OPERATIONNO").ToString & Chr(10) & Chr(13) & vbCrLf & vbCrLf & "This Is an automated message"
                        'EmailFile(DR("EMAIL").ToString, MsgString, "Movement notification:  " & DR("WORKORDERNO").ToString, True)
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

    'Sub EmailFile(Recipient As String, MessageBody As String, Subject As String, Optional Send As Boolean = False)

    '    Dim OutLookApp As New Outlook.Application
    '    Dim Mail As Outlook.MailItem = OutLookApp.CreateItem(Outlook.OlItemType.olMailItem)
    '    Dim mailRecipient As Outlook.Recipient
    '    mailRecipient = Mail.Recipients.Add(Recipient)
    '    mailRecipient.Resolve()
    '    Mail.Recipients.ResolveAll()
    '    Mail.HTMLBody = MessageBody
    '    Mail.Subject = Subject
    '    Mail.Save()
    '    If Send Then
    '        Mail.Send()
    '    Else
    '        Mail.Display()
    '    End If

    'End Sub

    Private Function GetWFReport(ref As String) As String()()
        'Debug.Print(ref)
        ' Debug.Print("")
        Dim doc As mshtml.HTMLDocument
        'Try
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
            Sleep(4000)
            IE.Navigate(ref)

            For X = 0 To 10
                Do Until IE.Busy = False And IE.ReadyState = 4 : Sleep(10) : Loop : Sleep(10)
            Next X

            doc = IE.Document
            Dim i As Integer = 0
            Debug.Print(doc.all.length)
            If doc.all.length < 100 Then

                Do While i < doc.all.length
                    Dim element As Object = doc.all(i)
                    Try
                        If Not IsNothing(element.innerhtml) Then
                            ' If Not IsNothing(element.id) Then Debug.Print("ID:  " & element.id)
                            ' If Not IsNothing(element.title) Then Debug.Print("TITLE:  " & element.title)
                            If InStr(element.innerhtml, "win.document.form1.action = ") > 0 Then
                                Dim RepURL As String = "http://webfocus" & Split(Split(element.innerhtml, "win.document.form1.action = " & Chr(34))(1), Chr(34) & ";")(0)
                                'Debug.Print(RepURL)
                                IE.Navigate(RepURL)
                                For X = 0 To 10
                                    Do Until IE.Busy = False And IE.ReadyState = 4 : Sleep(40) : Loop : Sleep(10)
                                Next X
                                Threading.Thread.Sleep(100)
                                i = 1000000
                            End If
                        End If
                    Catch
                    End Try
                    i = i + 1
                Loop
            End If
            For X = 0 To 10
                Do Until IE.Busy = False And IE.ReadyState = 4 : Sleep(40) : Loop : Sleep(10)
            Next X
            doc = IE.Document
            Dim doc1 As String = doc.body.outerHTML
            IE.Navigate("http://webfocus.pccstructurals.com/ibi_apps/bip/portal/PCCStructuralsInc")

            Return ClassLibrary1.HTMLProcessor.ParseHtml(doc1)
            'Catch ex As Exception
            IE.Visible = True

            'FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\8ball\updater\error" & Day(Now) & Hour(Now) & Minute(Now) & ".txt", "ERROR ON LINE " & Erl() & vbCrLf & vbCrLf & ex.Message.ToString & vbCrLf & vbCrLf & vbCrLf & ex.InnerException.ToString, True)
        'End Try
    End Function


    Private Sub UpdateAppend(ref As String, RespNames As String)
        Dim tab As String = ""
        Dim RefFind() As String = {"ships", "fingoods", "lots", "certs", "scrap", "partdata", "xtl", "tput", "labor", "labor1", "wiphist"}
        Dim TableNames() As String = {"SHIPMENTS", "CERT_ERRORS", "CERT_ERRORS", "CERT_ERRORS", "SCRAP", "ALLOYS", "TIMELINE", "TPUT", "LABOR", "LABOR", "WIP_MOVE_HIST"}
        Dim UpdatedRows As Integer = 0
        Using cn As New SqlConnection(ConnectionString)
            cn.Open()
            ' Try
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
                    If WebFocusPull Then
                        j = GetWFReport(ref)
                        Dim headerst As String = ""
                        For Each s As String In j(0)
                            headerst = headerst & s & ","
                        Next
                        headerst = Left(headerst, (Len(headerst) - 1)) & vbCrLf
                    'If Environment.MachineName = "SLPPRASINOSLT01" or Environment.MachineName = "SLAN-1ZNFXZ1" Then
                     FileIO.FileSystem.WriteAllText("\\slfs01\public\VisDownloads\" & RespNames & ".csv", headerst, False)
                Else

                    ' j = FileIO.FileSystem.ReadAllText(Replace(ref, ".csv", "_headers.csv")) & FileIO.FileSystem.ReadAllText(ref)
                    Dim t As String() = Split(FileIO.FileSystem.ReadAllText(Replace(ref, ".csv", "_headers.csv")) & FileIO.FileSystem.ReadAllText(ref), vbCrLf)
                    Dim TempList As New List(Of String())
                        For Each s In t
                        TempList.Add(Split(Replace(s, "CE, W", ""), ","))
                    Next
                        j = TempList.ToArray
                    End If

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

                For RowNum = 1 To j.length - 2
                    With cmd.Parameters
                        .Clear()
                        For Each Col In ColumnInfo
                            'j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), Chr(34), "")
                            j(RowNum)(Col(2)) = Trim(j(RowNum)(Col(2)))
                            'j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), "E, W", "")
                            j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), Chr(34), "")
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
                                    Dim S As Integer = (0 & Replace((Replace(Replace(j(RowNum)(Col(2)), " ", ""), ".", "")), Chr(34), "")) * 1
                                    .AddWithValue("@" & Col(0), S)
                                End If
                            ElseIf InStr(Col(1), "Date", CompareMethod.Text) <> 0 Then
                                Dim dt As DateTime = #1/1/1900#
                                If Not Replace(j(RowNum)(Col(2)), ",", "") = "." Then
                                    If j(RowNum)(Col(2)) <> "0" Then
                                        Try
                                            If Len(j(RowNum)(Col(2))) = 8 Then dt = DateTime.ParseExact(j(RowNum)(Col(2)), "MMddyyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                            If Len(j(RowNum)(Col(2))) >= 9 Then dt = DateTime.ParseExact(Left(j(RowNum)(Col(2)), 14), "yyyyMMddHHmmss", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                        Catch
                                            dt = DateTime.Parse(j(RowNum)(Col(2)), System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                        End Try
                                    Else
                                        dt = #1/1/1900#
                                    End If
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

            'Catch ex As Exception
            '    MsgBox(Erl() & ex.ToString)
            '    MsgBox(ex.InnerException.ToString)
            'End Try
        End Using
    End Sub



    Sub OpensUpdater(ref As String)
        Dim j As New Object 
        If WebFocusPull Then
            j = GetWFReport(ref)
            Dim headerst As String = ""
            For Each s As String In j(0)
                headerst = headerst & s & ","
            Next
            headerst = Left(headerst, (Len(headerst) - 1)) & vbCrLf
            'If Environment.MachineName = "SLPPRASINOSLT01" or Environment.MachineName = "SLAN-1ZNFXZ1" Then
            FileIO.FileSystem.WriteAllText("\\slfs01\public\VisDownloads\" & "OPENS" & ".csv", headerst, False)
        Else
            Dim t As String() = Split(FileIO.FileSystem.ReadAllText(Replace(ref, ".csv", "_headers.csv")) & FileIO.FileSystem.ReadAllText(ref), vbCrLf)
            Dim TempList As New List(Of String())
            For Each s In t
                TempList.Add(Split(Replace(s, "CE, W", ""), ","))
            Next
            j = TempList.ToArray

        End if
        UpdateStatus(2, "RECIEVED", "OPEN_ORDERS", False)
        Using cn As New SqlConnection(ConnectionString)
            ' Try
            cn.Open()
                Using cmd As New SqlCommand("", cn)
                    cmd.CommandTimeout = 5

                    cmd.CommandText = "UPDATE wflocal.dbo.OPEN_ORDERS Set ACTIVE = 2 WHERE ACTIVE <> 0"
                    cmd.ExecuteNonQuery()
                    cmd.CommandText = " Select column_name, data_type 
                                        From WFLOCAL.INFORMATION_SCHEMA.COLUMNS
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

                For RowNum = 1 To j.length - 2

                    With cmd.Parameters
                        .Clear()
                        For Each Col In ColumnInfo
                            j(RowNum)(Col(2)) = Replace(j(RowNum)(Col(2)), Chr(34), "")
                            j(RowNum)(Col(2)) = Trim(j(RowNum)(Col(2)))


                            If Col(1) = "nvarchar" Then
                                .Add("@" & Col(0), SqlDbType.NVarChar).Value = j(RowNum)(Col(2))
                            ElseIf Col(1) = "float" Then
                                .Add("@" & Col(0), SqlDbType.Float).Value = Replace(j(RowNum)(Col(2)), ",", "")
                            ElseIf Col(1) = "datetime" Then
                                Dim dt As DateTime
                                Try
                                    dt = DateTime.ParseExact(j(RowNum)(Col(2)), "MMddyyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo).Date
                                Catch
                                    dt = #1/1/1900#
                                    GoTo skip
                                End Try
                                If dt.Year > 1900 Then
                                    .Add("@" & Col(0), SqlDbType.Date).Value = dt
                                Else
                                    dt = Now.AddYears(-100)
                                    .Add("@" & Col(0), SqlDbType.Date).Value = dt
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
skip:
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
            'Catch ex As Exception
            '    MsgBox(ex.GetType().ToString)
            '    MsgBox(ex.Message.ToString)
            '    MsgBox(ex.InnerException.ToString)
            'End Try
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
