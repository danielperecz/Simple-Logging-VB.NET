Public Class Log

    ' Log file settings.
    Private ReadOnly LINE_LIMIT As Integer = 10000          ' Max line limit for file. If file reaches this number, a new file will be created.
    Private ReadOnly LOG_TYPE_PADDING As Integer = 6
    Private ReadOnly METHOD_PADDING As Integer = 40
    Private ReadOnly RESULT_PADDING As Integer = 5

    Public Sub WriteToFile(Method As String, LogType As String, Optional Result As String = "None", Optional Payload As String = "None")
        ' Write new line to a text file without deleting any existing data the file may contain.

        Dim file As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(file:=SelectFile(), append:=True)
        file.WriteLine(
            Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " | " +
            LogType.ToUpper().PadRight(LOG_TYPE_PADDING, " ") + " | " +
            Method.PadRight(METHOD_PADDING, " ") + " | " +
            Result.PadRight(RESULT_PADDING, " ") + " | " +
            Payload
        )
        file.Close()

    End Sub

    Private Function SelectFile() As String
        ' Return full path to log file to be used.
        
        Dim newFilePath As String = GlobalVariables.Logs & "\" & Date.Now.ToString("yyyy.MM.dd_HH.mm.ss") & ".txt"
        Dim folder As IO.DirectoryInfo = New IO.DirectoryInfo(GlobalVariables.Logs)

        If folder.GetFiles.Length = 0 Then
            ' If Logs folder is empty.
            Return newFilePath

        Else
            ' If Logs folder is not empty.

            Dim mostRecentFile As IO.FileInfo = folder.EnumerateFiles("*.txt").OrderByDescending(Function(f) f.LastWriteTime).FirstOrDefault()
            Dim lines() As String = IO.File.ReadAllLines(mostRecentFile.FullName)

            If lines.Length >= LINE_LIMIT Then
                ' If most recent file is larger than, or equal to LINE_LIMIT.
                Return newFilePath

            Else
                ' If newest file has less lines than the LINE_LIMIT: select this file.
                Return mostRecentFile.FullName

            End If

        End If

    End Function

End Class
