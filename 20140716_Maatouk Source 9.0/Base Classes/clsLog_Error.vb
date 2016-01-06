Public Class clsLog_Error
    Inherits Object

    Private Const log_PROCESS_ORDERS As String = "Log_ProcessOrders.txt"
    Private Const log_INVOICING As String = "Log_Invoicing.txt"

    Private oFSO As Scripting.FileSystemObject

    Public Enum Log As Integer
        lg_PROCESS_ORDER = 1
        lg_INVOICING
    End Enum

    Public Sub New()
        MyBase.New()
        oFSO = New Scripting.FileSystemObject
    End Sub

    Public Sub WriteToLog(ByVal sText As String, ByVal Type As Log)
        Dim sLogPath As String
        Dim sLogFilePath As String
        Dim oStream As Scripting.TextStream
        Try
            sLogPath = oApplication.Utilities.getApplicationPath() & "\Log"

            If Not oFSO.FolderExists(sLogPath) Then
                oFSO.CreateFolder(sLogPath)
            End If

            Select Case Type
                Case Log.lg_PROCESS_ORDER
                    sLogFilePath = sLogPath & "\" & log_PROCESS_ORDERS

                Case Log.lg_INVOICING
                    sLogFilePath = sLogPath & "\" & log_INVOICING

            End Select

            If Not oFSO.FileExists(sLogFilePath) Then
                oStream = oFSO.CreateTextFile(sLogFilePath, True)
            Else
                oStream = oFSO.OpenTextFile(sLogFilePath, Scripting.IOMode.ForAppending, True, Scripting.Tristate.TristateUseDefault)
            End If

            sText = sText & vbCrLf & vbCrLf
            oStream.Write(sText)

        Catch ex As Exception
            Throw (ex)
        Finally
            oStream.Close()
            oStream = Nothing
        End Try
    End Sub

    Public Sub DeleteFile(ByVal Type As Log)
        Dim sLogFilePath As String

        Select Case Type
            Case Log.lg_PROCESS_ORDER
                sLogFilePath = oApplication.Utilities.getApplicationPath() & "\Log\" & log_PROCESS_ORDERS

            Case Log.lg_INVOICING
                sLogFilePath = oApplication.Utilities.getApplicationPath() & "\Log\" & log_INVOICING

        End Select
        If oFSO.FileExists(sLogFilePath) Then
            oFSO.DeleteFile(sLogFilePath)
        End If
    End Sub

    Public Sub ShowLogFile(ByVal Type As Log)
        Dim sLogFilePath As String

        Select Case Type
            Case Log.lg_PROCESS_ORDER
                sLogFilePath = oApplication.Utilities.getApplicationPath() & "\Log\" & log_PROCESS_ORDERS

            Case Log.lg_INVOICING
                sLogFilePath = oApplication.Utilities.getApplicationPath() & "\Log\" & log_INVOICING

        End Select

        Shell("Notepad.exe " & sLogFilePath, AppWinStyle.NormalFocus)

    End Sub

    Protected Overrides Sub Finalize()
        oFSO = Nothing
    End Sub

End Class
