Imports System.Diagnostics
Imports System.IO


Module Startup_Module

    Dim WorkingFolder As String = ""

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Video Directory Screencap Preview Page Generator's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Sub main(ByVal args() As String)
        If args.Length > 0 Then
            Dim dinfo As DirectoryInfo = New DirectoryInfo(args(0))
            If dinfo.Exists = True Then
                WorkingFolder = dinfo.FullName
            End If
            dinfo = Nothing
        End If

        Dim ApplicationName As String
        ApplicationName = "Video Directory Screencap Preview Page Generator"
        Try
            Dim aModuleName As String = Diagnostics.Process.GetCurrentProcess.MainModule.ModuleName
            Dim aProcName As String = System.IO.Path.GetFileNameWithoutExtension(aModuleName)

            If Process.GetProcessesByName(aProcName).Length > 2 Then
                Dim Display_Message1 As New Display_Message("Another Instance of " & ApplicationName & " is already running. Only two instances of " & ApplicationName & " is allowed to run at any time")
                Display_Message1.ShowDialog()
                Application.Exit()
            Else
                Dim Splash As New Splash_Screen
                Splash.Show()
                While Splash.dataloaded = False
                End While

                Dim monitor As New Main_Screen(Splash)
                monitor.WorkingFolder = WorkingFolder
                monitor.ShowDialog()
                If Not monitor Is Nothing Then
                    monitor.Visible = False
                End If
                If Not monitor Is Nothing Then
                    While monitor.dataloaded = False
                    End While
                End If
                If Not monitor Is Nothing Then
                    monitor.Visible = True
                End If
                Application.Exit()
            End If
        Catch ex As Exception
            Error_Handler(ex, "Application Launch")
        End Try
    End Sub
End Module
