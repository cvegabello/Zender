Public Class frmMain


    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cdcInt As Integer

        Shell("NET USE H: \\10.5.165.98\SharedDocuments Formula4 /USER:myigt\cvegabello")
        'MsgBox("NET USE Y:", MsgBoxStyle.Exclamation, "NET")
        DateTimePicker1.Value = Now
        cdcInt = returnCDC(DateTimePicker1.Value)
        Label1.Text = cdcInt
        GroupBox1.Enabled = False
        ProgressBar1.Maximum = 100
        okBtn.Visible = False
        cancelBtn.Visible = False
        Timer1.Enabled = True

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim errCode As Integer
        Dim errorStr As String = ""
        Dim fecha As Date
        Dim file As System.IO.StreamWriter
        Dim monthStr, dayStr, yearStr, strFilePath As String
        Dim hostNumberArray(5) As Integer
        Dim conStrinHostStr As String
        Dim substrings() As String
        Static count As Integer = 0
        count = count + 10
        If count <= 100 Then
            ProgressBar1.Value = count
        Else
            Timer1.Enabled = False

            strFilePath = Utils.GetSetting("InfoBoard", "UbiWinFileLocation", "").ToString()

            fecha = Format(Now, "MM/dd/yyyy")
            monthStr = CStr(fecha.Month).PadLeft(2, "0")
            dayStr = CStr(fecha.Day).PadLeft(2, "0")
            yearStr = CStr(fecha.Year).PadLeft(4, "0")

            Utils.getWinnerFilesSFTP(strFilePath, fecha, errorStr)

            conStrinHostStr = GetSettingConfigHost(appName, "conStringHost", "").ToString()
            substrings = conStrinHostStr.Split("|")
            getHostModeNew(substrings(3), substrings(0), substrings(4), hostNumberArray, errCode, errorStr)

            If errCode = 0 Then
                file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "hostMode.txt", False)
                file.WriteLine(hostNumberArray(0) & "|" & hostNumberArray(1) & "|" & hostNumberArray(2) & "|" & hostNumberArray(3) & "|" & hostNumberArray(4))
                file.Close()

                file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                file.WriteLine(Now & ".  File with host mode was generated and transferred successfully")
                file.Close()

            Else
                file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                file.WriteLine(Now & "  " & errorStr)
                file.Close()

            End If
            file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "flagFile.txt", True)
            file.WriteLine(Now & ".  Ready.")
            file.Close()

            Me.Dispose()

        End If



    End Sub

    Private Sub okBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles okBtn.Click


        Dim errCode As Integer
        Dim errorStr As String = ""
        Dim fecha As Date
        Dim file As System.IO.StreamWriter
        Dim monthStr, dayStr, yearStr, strFilePath As String
        Dim hostNumberArray(5) As Integer
        Dim conStrinHostStr As String
        Dim substrings() As String

        Me.Cursor = Cursors.WaitCursor
        strFilePath = Utils.GetSetting("InfoBoard", "UbiWinFileLocation", "").ToString()

        fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
        'fecha = Format(Now, "MM/dd/yyyy")
        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")


        Utils.getWinnerFilesSFTP(strFilePath, fecha, errorStr)

        conStrinHostStr = GetSettingConfigHost(appName, "conStringHost", "").ToString()
        substrings = conStrinHostStr.Split("|")
        getHostModeNew(substrings(3), substrings(0), substrings(4), hostNumberArray, errCode, errorStr)

        If errCode = 0 Then
            file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "hostMode.txt", False)
            file.WriteLine(hostNumberArray(0) & "|" & hostNumberArray(1) & "|" & hostNumberArray(2) & "|" & hostNumberArray(3) & "|" & hostNumberArray(4))
            file.Close()

            file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
            file.WriteLine(Now & ".  File with host mode was generated and transferred successfully")
            file.Close()

        Else
            file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
            file.WriteLine(Now & "  " & errorStr)
            file.Close()

        End If
        file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "flagFile.txt", True)
        file.WriteLine(Now & ".  Ready.")
        file.Close()

        Me.Dispose()


    End Sub

    Private Sub stopBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopBtn.Click
        Timer1.Enabled = False
        stopBtn.Visible = False
        ProgressBar1.Visible = False
        GroupBox1.Enabled = True

        okBtn.Visible = True
        cancelBtn.Visible = True
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim dateSelected As Date
        Dim cdcInt As Integer


        dateSelected = DateTimePicker1.Value
        cdcInt = returnCDC(dateSelected)
        Label1.Text = cdcInt

    End Sub

    Private Sub cancelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelBtn.Click
        Me.Dispose()
    End Sub
End Class
