' Simple but dynamic class for creating log files. Example line from a log file:
' 2020-04-29 13:13:24.495 | 1  | admin | btnSubmitDelete_Click | Deleted hours worked. | Original Range: 2020-04-24 09:00 to 2020-04-24 17:00

Public Class Log

    ' Log file settings.
    Private ReadOnly LINE_LIMIT As Integer = 10000          ' Max line limit for file. If file reaches this number, a new file will be created for next log.
    Private ReadOnly EMPLOYEE_ID_PADDING As Integer = 5
    Private ReadOnly ROLE_PADDING As Integer = 5
    Private ReadOnly METHOD_PADDING As Integer = 40
    Private ReadOnly DESCRIPTION_PADDING As Integer = 60

    ' Log file data.
    Private CurrentDateTime As String
    Private Employee_ID As String
    Private Role As String
    Private Method As String
    Private Description As String
    Private Payload As String

    Public Sub New(CurrentDateTime As Date, Employee_ID As String, Role As String, Method As String, Description As String, Optional Payload As String = "None")

        Me.CurrentDateTime = CurrentDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
        Me.Employee_ID = Employee_ID
        Me.Role = Role
        Me.Method = Method
        Me.Description = Description
        Me.Payload = Payload
        
        WriteToFile()

    End Sub

    Private Sub WriteToFile()
        ' Write new line to a text file without deleting any existing data the file may contain.

        Dim file As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(file:=SelectFile(), append:=True)
        file.WriteLine(CurrentDateTime & " | " &
                       Employee_ID.PadRight(EMPLOYEE_ID_PADDING, " ") & " | " &
                       Role.PadRight(ROLE_PADDING, " ") & " | " &
                       Method.PadRight(METHOD_PADDING, " ") & " | " &
                       Description.PadRight(DESCRIPTION_PADDING, " ") & " | " &
                       Payload)
        file.Close()

    End Sub

    Private Function SelectFile() As String
        ' Return full path to log file to be used.
        
        ' LogsPath is a path to your log folder. E.g.: "C:\Application\Logs".
        Dim newFilePath As String = LogsPath & "\" & DateTime.Now.ToString("yyyy.MM.dd_HH.mm.ss") & ".txt"
        Dim folder As IO.DirectoryInfo = New IO.DirectoryInfo(LogsPath)

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
