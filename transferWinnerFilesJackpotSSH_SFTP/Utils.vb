Imports System.IO
Imports System.Text
Imports System.Xml
Imports Microsoft.Win32

Module Utils

    Public appName As String = "SETTINGS LOCAL APPS"

    Public Function GetSetting(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Return Key.GetValue(Keyname, DefVal)


        Catch
            Return DefVal
        End Try
    End Function

    Public Sub SalvarSetting(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Sub SalvarSettingConfiServerDB(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string database")
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Function GetSettingConfigServerDB(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string database")
            Return Key.GetValue(Keyname, DefVal)
        Catch
            Return DefVal
        End Try
    End Function

    Public Sub SalvarSettingConfiHost(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string host")
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Function GetSettingConfigHost(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string host")
            Return Key.GetValue(Keyname, DefVal)
        Catch
            Return DefVal
        End Try
    End Function

    Sub getHostModeNew(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal hostNumberArray() As Integer, ByRef errCodeInt As Integer, ByRef errMessageStr As String)


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum, posInt, numInt, isIcludedInt, spareInt As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Sub
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Sub
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Sub
            'Else
            '    MsgBox("Funciono OpenSessionChannel", MsgBoxStyle.OkOnly, "Funciono OpenSessionChannel. ")

        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Sub
            'Else
            '    MsgBox("Funciono SendReqPty", MsgBoxStyle.OkOnly, "Funciono SendReqPty. ")

        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Sub
            'Else
            '    MsgBox("Funciono SendReqShell", MsgBoxStyle.OkOnly, "Funciono SendReqShell. ")
        End If


        errCodeInt = SSHUntilMatch(ssh, "Enter Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: ", MsgBoxStyle.OkOnly, "FUNCIONO.")

        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "10", "Enter Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: 10 ", MsgBoxStyle.OkOnly, "FUNCIONO.")
        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "01", "Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: 01 ", MsgBoxStyle.OkOnly, "FUNCIONO.")


        End If

        errCodeInt = sentStringToSSH(ssh, "PEERS", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        ssh.Disconnect()


        posInt = InStr(1, cmdOutputStr, "live_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 10, 2))
        hostNumberArray(0) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "backup_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 11, 2))
        hostNumberArray(1) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "wait_for_spare:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 16, 2))
        spareInt = numInt + 1


        If ((spareInt = hostNumberArray(0)) Or (spareInt = hostNumberArray(1))) Then

            For i = 1 To 5
                isIcludedInt = Array.IndexOf(hostNumberArray, i)
                If isIcludedInt < 0 Then
                    hostNumberArray(2) = i
                    Exit For
                End If
            Next
        Else
            hostNumberArray(2) = spareInt

        End If


        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(3) = i
                Exit For
            End If
        Next

        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(4) = i
                Exit For
            End If
        Next

    End Sub
    Sub getHostMode(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal hostNumberArray() As Integer, ByRef errCodeInt As Integer, ByRef errMessageStr As String)


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum, posInt, numInt, isIcludedInt, spareInt As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Sub
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Sub
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Sub
        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Sub
        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Sub
        End If


        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            errCodeInt = -7
            errMessageStr = "Error related to ChannelReadAndPoll: " & ssh.LastErrorText
            'MsgBox("NO funciono ChannelReadAndPoll " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono ChannelReadAndPoll. ")
            Exit Sub
        End If


        cmdOutputStr = ssh.GetReceivedText(channelNum, "ansi")
        If (ssh.LastMethodSuccess <> True) Then
            errCodeInt = -8
            errMessageStr = "Error related to GetReceivedText: " & ssh.LastErrorText
            'MsgBox("NO funciono GetReceivedText " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono GetReceivedText. ")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If




        errCodeInt = sentStringToSSH(ssh, "cd gtms/bin", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If


        errCodeInt = sentStringToSSH(ssh, "gxvision", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        errCodeInt = sentStringToSSH(ssh, "PEERS", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        ssh.Disconnect()


        posInt = InStr(1, cmdOutputStr, "live_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 10, 2))
        hostNumberArray(0) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "backup_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 11, 2))
        hostNumberArray(1) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "wait_for_spare:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 16, 2))
        spareInt = numInt + 1

        If ((spareInt = hostNumberArray(0)) Or (spareInt = hostNumberArray(1))) Then

            For i = 1 To 5
                isIcludedInt = Array.IndexOf(hostNumberArray, i)
                If isIcludedInt < 0 Then
                    hostNumberArray(2) = i
                    Exit For
                End If
            Next
        Else
            hostNumberArray(2) = spareInt

        End If


        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(3) = i
                Exit For
            End If
        Next

        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(4) = i
                Exit For
            End If
        Next



    End Sub




    Sub getWinnerFilesSFTP(ByVal strFilePath As String, ByVal currentDay As Date, ByRef errorStr As String)

        Dim nameDayWeekStr, drawNumberStr As String
        Dim drawNumberInt, cont As Integer
        Dim toKnowDate As Date
        Dim codErrorInt As Integer = 0
        Dim errMessageStr As String = ""
        Dim newJackPotDbl As Double
        Dim currentDayStr, newJackPotStr, conStringHostStr, IpStr, userNameHostStr, passwordHostStr As String
        Dim file As System.IO.StreamWriter
        Dim substrings(), monthStr, dayStr, yearStr As String




        'substrings(0) -> username
        'substrings(1) -> IP ESTE1
        'substrings(2) -> Password ESTE1
        'substrings(3) -> IP ESTE2
        'substrings(4) -> Password ESTE2
        'substrings(5) -> IP ESTE3
        'substrings(6) -> Password ESTE3
        'substrings(7) -> IP ESTE4
        'substrings(8) -> Password ESTE4
        'substrings(9) -> IP ESTE5
        'substrings(10) -> Password ESTE5



        conStringHostStr = GetSettingConfigHost(appName, "conStringHost", "").ToString()
        cont = conStringHostStr.Length
        If cont <> 0 Then
            substrings = conStringHostStr.Split("|")
            IpStr = substrings(1)
            userNameHostStr = substrings(0)
            passwordHostStr = substrings(2)
        Else
            IpStr = ""
            userNameHostStr = ""
            passwordHostStr = ""

        End If

        monthStr = CStr(currentDay.Month).PadLeft(2, "0")
        dayStr = CStr(currentDay.Day).PadLeft(2, "0")
        yearStr = CStr(currentDay.Year).PadLeft(4, "0")

        'currentDay = Now
        toKnowDate = currentDay.AddDays(-1) 'Normalmente debe estar con -1, es decir con la fecha anterior

        'toKnowDate = DateTimePicker1.Value
        nameDayWeekStr = WeekdayName(Weekday(toKnowDate))

        'Shell("NET USE Y: \\10.5.165.98\SharedDocuments Formula7 /USER:cvegabello")

        Select Case nameDayWeekStr

            Case "Sunday", "Monday", "Thursday"

                drawNumberInt = getDrawNumberXProductXday(10, toKnowDate, 0)
                drawNumberStr = Trim(Str(drawNumberInt))
                drawNumberStr = drawNumberStr.PadLeft(6, "0")
                codErrorInt = downloadFileSFTP(IpStr, 22, userNameHostStr, passwordHostStr, "/rptfiles1/progam/csh5/reports/winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep", strFilePath & "winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep", errorStr)
                If codErrorInt = 0 Then
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep" & " was successfully transferred")
                    file.Close()
                Else
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  " & errorStr)
                    file.Close()
                End If


                drawNumberInt = getDrawNumberXProductXday(13, toKnowDate, 0)
                drawNumberStr = Trim(Str(drawNumberInt))
                drawNumberStr = drawNumberStr.PadLeft(4, "0")
                codErrorInt = downloadFileSFTP(IpStr, 22, userNameHostStr, passwordHostStr, "/rptfiles1/progam/life/reports/winners_cl_" & drawNumberStr & ".xml", strFilePath & "winners_cl_" & drawNumberStr & ".xml", errorStr)

                If codErrorInt = 0 Then
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  winners_cl_" & drawNumberStr & ".xml" & " was successfully transferred")
                    file.Close()

                Else
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  " & errorStr)
                    file.Close()
                End If


            Case "Tuesday", "Friday"

                drawNumberInt = getDrawNumberXProductXday(10, toKnowDate, 0)
                drawNumberStr = Trim(Str(drawNumberInt))
                drawNumberStr = drawNumberStr.PadLeft(6, "0")
                codErrorInt = downloadFileSFTP(IpStr, 22, userNameHostStr, passwordHostStr, "/rptfiles1/progam/csh5/reports/winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep", strFilePath & "winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep", errorStr)
                If codErrorInt = 0 Then
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep" & " was successfully transferred")
                    file.Close()
                Else
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  " & errorStr)
                    file.Close()
                End If


                drawNumberInt = getDrawNumberXProductXday(13, toKnowDate, 0)
                drawNumberStr = Trim(Str(drawNumberInt))
                drawNumberStr = drawNumberStr.PadLeft(4, "0")
                codErrorInt = downloadFileSFTP(IpStr, 22, userNameHostStr, passwordHostStr, "/rptfiles1/progam/life/reports/winners_cl_" & drawNumberStr & ".xml", strFilePath & "winners_cl_" & drawNumberStr & ".xml", errorStr)

                If codErrorInt = 0 Then
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  winners_cl_" & drawNumberStr & ".xml" & " was successfully transferred")
                    file.Close()

                Else
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  " & errorStr)
                    file.Close()
                End If


                drawNumberInt = getDrawNumberXProductXday(12, toKnowDate, 0)
                drawNumberStr = Trim(Str(drawNumberInt))
                drawNumberStr = drawNumberStr.PadLeft(4, "0")
                codErrorInt = downloadFileSFTP(IpStr, 22, userNameHostStr, passwordHostStr, "/rptfiles1/progam/bigg/reports/winners_mm_" & drawNumberStr & ".xml", strFilePath & "winners_mm_" & drawNumberStr & ".xml", errorStr)

                If codErrorInt = 0 Then
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  winners_mm_" & drawNumberStr & ".xml" & " was successfully transferred")
                    file.Close()

                    newJackPotDbl = getJackpotFromSSH(IpStr, "xfer", "Welcome1", 12, currentDay, codErrorInt, errMessageStr) 'Bigg
                    If codErrorInt <> 0 Then
                        file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                        file.WriteLine(Now & " " & "ErrorSSH: " & errMessageStr)
                        file.Close()
                    Else
                        file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                        file.WriteLine(Now & " " & "Jackpot MegaMillions OK")
                        file.Close()
                    End If

                    drawNumberInt = getDrawNumberXProductXday(12, currentDay, 0)
                    drawNumberStr = Trim(Str(drawNumberInt))
                    drawNumberStr = drawNumberStr.PadLeft(6, "0")

                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "jackpot_" & drawNumberStr & "_bigg.txt", False)
                    currentDayStr = currentDay
                    currentDayStr = currentDayStr.PadRight(40, " ")
                    newJackPotStr = Trim(Str(newJackPotDbl))
                    newJackPotStr = newJackPotStr.PadLeft(20, "0")
                    file.WriteLine(currentDayStr & newJackPotStr)
                    file.Close()

                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & " " & "File Jackpot MegaMillions created OK")
                    file.Close()

                Else
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  " & errorStr)
                    file.Close()
                End If


            Case "Wednesday", "Saturday"

                drawNumberInt = getDrawNumberXProductXday(10, toKnowDate, 0)
                drawNumberStr = Trim(Str(drawNumberInt))
                drawNumberStr = drawNumberStr.PadLeft(6, "0")
                codErrorInt = downloadFileSFTP(IpStr, 22, userNameHostStr, passwordHostStr, "/rptfiles1/progam/csh5/reports/winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep", strFilePath & "winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep", errorStr)
                If codErrorInt = 0 Then
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep" & " was successfully transferred")
                    file.Close()
                Else
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  " & errorStr)
                    file.Close()
                End If
                '=============================================================================================================

                drawNumberInt = getDrawNumberXProductXday(13, toKnowDate, 0)
                drawNumberStr = Trim(Str(drawNumberInt))
                drawNumberStr = drawNumberStr.PadLeft(4, "0")
                codErrorInt = downloadFileSFTP(IpStr, 22, userNameHostStr, passwordHostStr, "/rptfiles1/progam/life/reports/winners_cl_" & drawNumberStr & ".xml", strFilePath & "winners_cl_" & drawNumberStr & ".xml", errorStr)

                If codErrorInt = 0 Then
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  winners_cl_" & drawNumberStr & ".xml" & " was successfully transferred")
                    file.Close()

                Else
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  " & errorStr)
                    file.Close()
                End If

                '===============================================================================================================================
                drawNumberInt = getDrawNumberXProductXday(8, toKnowDate, 0)
                drawNumberStr = Trim(Str(drawNumberInt))
                drawNumberStr = drawNumberStr.PadLeft(6, "0")
                codErrorInt = downloadFileSFTP(IpStr, 22, userNameHostStr, passwordHostStr, "/rptfiles1/progam/loto/reports/winner_summary_report_p008_d" & drawNumberStr & "_wincnt_english.rep", strFilePath & "winner_summary_report_p008_d" & drawNumberStr & "_wincnt_english.rep", errorStr)

                If codErrorInt = 0 Then
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  winner_summary_report_p008_d" & drawNumberStr & "_wincnt_english.rep" & " was successfully transferred")
                    file.Close()

                    'newJackPotDbl = getJackpotFromSSH(IpStr, "xfer", "Welcome1", 8, currentDay, codErrorInt, errMessageStr) 'Lotto
                    newJackPotDbl = getJackpotFromSSH(IpStr, "xfer", "Welcome1", 8, toKnowDate, codErrorInt, errMessageStr) 'Lotto
                    If codErrorInt <> 0 Then
                        file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                        file.WriteLine(Now & " " & "ErrorSSH: " & errMessageStr)
                        file.Close()
                    Else
                        file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                        file.WriteLine(Now & " " & "Jackpot Lotto OK")
                        file.Close()
                    End If

                    drawNumberInt = getDrawNumberXProductXday(8, currentDay, 0)
                    drawNumberStr = Trim(Str(drawNumberInt))
                    drawNumberStr = drawNumberStr.PadLeft(6, "0")

                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "jackpot_" & drawNumberStr & "_loto.txt", False)
                    currentDayStr = currentDay
                    currentDayStr = currentDayStr.PadRight(40, " ")
                    newJackPotStr = Trim(Str(newJackPotDbl))
                    newJackPotStr = newJackPotStr.PadLeft(20, "0")
                    file.WriteLine(currentDayStr & newJackPotStr)
                    file.Close()

                Else
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  " & errorStr)
                    file.Close()
                End If
                '===================================================================================================================

                drawNumberInt = getDrawNumberXProductXday(15, toKnowDate, 0)
                drawNumberStr = Trim(Str(drawNumberInt))
                drawNumberStr = drawNumberStr.PadLeft(4, "0")
                codErrorInt = downloadFileSFTP(IpStr, 22, userNameHostStr, passwordHostStr, "/rptfiles1/progam/pwrb/reports/winners_pb_" & drawNumberStr & ".xml", strFilePath & "winners_pb_" & drawNumberStr & ".xml", errorStr)

                If codErrorInt = 0 Then
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  winners_pb_" & drawNumberStr & ".xml" & " was successfully transferred")
                    file.Close()

                    newJackPotDbl = getJackpotFromSSH(IpStr, "xfer", "Welcome1", 15, currentDay, codErrorInt, errMessageStr) 'PWRB
                    If codErrorInt <> 0 Then
                        file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                        file.WriteLine(Now & " " & "ErrorSSH: " & errMessageStr)
                        file.Close()
                    Else
                        file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                        file.WriteLine(Now & " " & "Jackpot PowerBall OK")
                        file.Close()
                    End If

                    drawNumberInt = getDrawNumberXProductXday(15, currentDay, 0)
                    drawNumberStr = Trim(Str(drawNumberInt))
                    drawNumberStr = drawNumberStr.PadLeft(6, "0")

                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "jackpot_" & drawNumberStr & "_pwrb.txt", False)
                    currentDayStr = currentDay
                    currentDayStr = currentDayStr.PadRight(40, " ")
                    newJackPotStr = Trim(Str(newJackPotDbl))
                    newJackPotStr = newJackPotStr.PadLeft(20, "0")
                    file.WriteLine(currentDayStr & newJackPotStr)
                    file.Close()

                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & " " & "File Jackpot Powerball created OK")
                    file.Close()


                Else
                    file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    file.WriteLine(Now & "  " & errorStr)
                    file.Close()
                End If

        End Select



        

    End Sub


    Function getDrawNumberXProductXday(ByVal codSysProduct As Integer, ByVal toKnowDate As DateTime, ByVal dayDraw As Integer) As Integer

        Dim num1 As Integer
        Dim num3 As Integer
        Dim dayWeekInt As Integer
        Dim moduloRes As Integer
        Dim cocienteRes As Integer

        num1 = convertDateInt(toKnowDate)
        dayWeekInt = Weekday(toKnowDate)

        Select Case codSysProduct
            Case 8 'Loto

                Select Case dayWeekInt
                    Case 5
                        moduloRes = (num1 + 2) Mod 7
                        cocienteRes = ((num1 + 2) \ 7) * 2
                    Case 6
                        moduloRes = (num1 + 1) Mod 7
                        cocienteRes = ((num1 + 1) \ 7) * 2
                    Case Else
                        moduloRes = num1 Mod 7
                        cocienteRes = (num1 \ 7) * 2

                End Select
                If (moduloRes < 1) Then
                    num3 = cocienteRes - 11488
                Else
                    num3 = cocienteRes - 11487
                End If
                num3 += 2357



            Case 9 'PCK3 -NUMBERS-

                Select Case dayDraw
                    Case 1 'Midday
                        num3 = ((2 * num1) - 69216)
                    Case 2 'Evenning
                        num3 = (((2 * num1) - 69216) + 1)
                End Select

            Case 10
                num3 = (num1 - 35525)

            Case 11

            Case 12 'BIGG
                If ((num1 Mod 7) < 4) Then
                    num3 = (((num1 \ 7) * 2) - 10682)
                Else
                    num3 = (((num1 \ 7) * 2) - 10681)
                End If

            Case 13 'LIFE

                If num1 <= 43647 Then
                    Select Case dayWeekInt
                        Case 6
                            num3 = (((num1 \ 7) * 2) - 11941)

                        Case Else
                            If ((num1 Mod 7) < 3) Then
                                num3 = (((num1 \ 7) * 2) - 11943)
                            Else
                                num3 = (((num1 \ 7) * 2) - 11942)
                            End If

                    End Select
                Else
                    num3 = (num1 - 43120)

                End If


            Case 14

                Select Case dayDraw
                    Case 1 'Midday
                        num3 = ((2 * num1) - 69216)
                    Case 2 'Evenning
                        num3 = (((2 * num1) - 69216) + 1)
                End Select


            Case 15 'PowerBall

                Select Case dayWeekInt
                    Case 5
                        moduloRes = (num1 + 2) Mod 7
                        cocienteRes = ((num1 + 2) \ 7) * 2
                    Case 6
                        moduloRes = (num1 + 1) Mod 7
                        cocienteRes = ((num1 + 1) \ 7) * 2
                    Case Else
                        moduloRes = num1 Mod 7
                        cocienteRes = (num1 \ 7) * 2

                End Select
                If (moduloRes < 1) Then
                    num3 = cocienteRes - 11488
                Else
                    num3 = cocienteRes - 11487
                End If

            Case 27 'DKNO
                num3 = (num1 - 31979)

        End Select
        Return num3
    End Function


    Function convertDateInt(ByVal toKnowDate As Date) As Integer
        Dim initDate As Date
        Dim diff1 As TimeSpan
        initDate = "1/1/1900"

        diff1 = toKnowDate.Subtract(initDate)
        convertDateInt = diff1.Days + 2

    End Function
    Function returnCDC(ByVal toKnowDate As Date) As Integer
        Dim initDate As Date
        Dim diff1 As TimeSpan
        initDate = "1/1/1986"

        diff1 = toKnowDate.Subtract(initDate)
        returnCDC = diff1.Days + 1

    End Function

End Module
