Imports System.IO
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports WaitForMe.My

Module Module1
    ' Constants
    Private Const ProcessName As String = "League of Legends"
    Private Const DefaultPauseTime As UShort = 1300
    Private Const DefaultResumeTime As UShort = 5
    Private Const ConfigFilePath As String = "Settings.config"

    ' Configuration
    Private PauseTime As UShort = DefaultPauseTime
    Private ResumeTime As UShort = DefaultResumeTime

    ' Timer and delay settings
    Dim Timer As Timer
    Dim WaitingTime As UShort
    Private DotCounter As UInteger = 0

    ' Region for startup-related methods
#Region "Startup"
    Sub Main()
        Try
            ' Check for admin privileges
            If Not IsRunningAsAdmin() Then
                Console.WriteLine("Error: No Admin Privileges!")
                Console.ReadKey()
                Environment.Exit(0)
            End If

            ' Set up Ctrl+C handler for cleanup
            SetConsoleCtrlHandler(New ConsoleCtrlDelegate(AddressOf MyConsoleCtrlHandler), True)
            AddHandler AppDomain.CurrentDomain.ProcessExit, AddressOf OnApplicationExit
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CurrentDomain_UnhandledException
        Catch ex As Exception
            ' Handle exceptions
        End Try

        ' Load settings from config file
        LoadSettings()

        Try
            ' Check for command-line arguments
            If Environment.GetCommandLineArgs().Length > 1 Then
                If Not ProcessCommandLineArguments() Then
                    UserInputDelay()
                End If
            Else
                UserInputDelay()
            End If

            Console.WriteLine($"Delay set to: {WaitingTime} Minutes")
            Console.WriteLine("Waiting for Process")
        Catch ex As Exception
            Environment.Exit(0)
        End Try

        ' Start the timer to check the process status
        Try
            If Timer Is Nothing Then
                Timer = New Timer(AddressOf CheckProcessStatus, Nothing, 0, 100)
            End If

            ' Keep the console application running
            Do
                Console.ReadLine()
            Loop
            NormalExit(True)
        Catch ex As Exception
            NormalExit(False)
        End Try
    End Sub

    ' Function to check if the application is running with admin privileges
    Function IsRunningAsAdmin() As Boolean
        Try
            Dim identity As WindowsIdentity = WindowsIdentity.GetCurrent()
            Dim principal As New WindowsPrincipal(identity)
            Return principal.IsInRole(WindowsBuiltInRole.Administrator)
        Catch ex As Exception
            Return True
        End Try
    End Function

    ' Function to get user input for delay
    Sub UserInputDelay()
        Dim key As Integer
        Console.WriteLine("How long do you want to suspend the loading process? (1-4 Min):")
        If Not Integer.TryParse(Console.ReadLine(), key) OrElse key < 1 OrElse key > 4 Then
            Console.WriteLine("Invalid Input!")
            Console.WriteLine("Press any Key to Exit.")
            Console.ReadKey()
            NormalExit(False)
        Else
            WaitingTime = CUShort(key)
        End If
    End Sub

    ' Function to process command-line arguments
    Function ProcessCommandLineArguments() As Boolean
        Try
            Dim arguments() As String = Environment.GetCommandLineArgs()
            If arguments.Length = 3 AndAlso (arguments(1).ToLower() = "/delay" OrElse arguments(1).ToLower() = "-delay") Then
                Dim delayValue As Integer
                If Integer.TryParse(arguments(2), delayValue) AndAlso delayValue >= 1 AndAlso delayValue <= 4 Then
                    WaitingTime = delayValue
                    Return True
                Else
                    Console.WriteLine("Invalid value for /delay. Please provide a value between 1 and 4.")
                    Return False
                End If
            Else
                Console.WriteLine("Invalid command-line arguments. Usage: /delay [1-4]")
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

    ' Region for process manipulation methods
#Region "Process Manipulation"
    ' Enum for thread access rights
    <Flags>
    Public Enum ThreadAccess As Integer
        TERMINATE = &H1
        SUSPEND_RESUME = &H2
        GET_CONTEXT = &H8
        SET_CONTEXT = &H10
        SET_INFORMATION = &H20
        QUERY_INFORMATION = &H40
        SET_THREAD_TOKEN = &H80
        IMPERSONATE = &H100
        DIRECT_IMPERSONATION = &H200
    End Enum

    ' External kernel32.dll functions for thread manipulation
    <System.Runtime.InteropServices.DllImport("kernel32.dll")>
    Public Function SuspendThread(hThread As IntPtr) As UInteger
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll")>
    Public Function ResumeThread(hThread As IntPtr) As UInteger
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll")>
    Public Function OpenThread(dwDesiredAccess As ThreadAccess, bInheritHandle As Boolean, dwThreadId As UInteger) As IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Public Function CloseHandle(hObject As IntPtr) As Boolean
    End Function

    ' Function to check the process status
    Private Sub CheckProcessStatus(state As Object)
        Try
            Dim process As Process = Process.GetProcessesByName(ProcessName).FirstOrDefault()
            If process IsNot Nothing Then
                ' Stop the timer if the process is found
                Timer.Dispose()
                Console.WriteLine($"Process {process.Id} ({process.ProcessName}) found. Slowing down.")

                ' Slow down the process for the specified time
                SlowDownProcessForDuration(process, TimeSpan.FromMinutes(WaitingTime))
            Else
                ' The process is not running
                DotCounter += 1
                If DotCounter Mod 100 = 0 Or DotCounter = 4294967290 Then
                    Thread.Sleep(100)
                    Console.Write(".")
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Error:" + ex.ToString)
            NormalExit(False)
        End Try
    End Sub

    ' Function to change process priority
    Sub ChangeProcessPriority(process As Process, Slow As Boolean)
        Try
            If Slow Then
                process.PriorityClass = ProcessPriorityClass.BelowNormal
                Console.WriteLine($"{ProcessName} Priority: Low")
            Else
                process.PriorityClass = ProcessPriorityClass.Normal
                Console.WriteLine($"{ProcessName} Priority: Normal")
            End If
        Catch ex As Exception
            Console.WriteLine($"Error: Cannot change Priority {ex.Message}")
            Console.WriteLine("This Application will still try to delay the Game start. This Function is not crucial")
        End Try
    End Sub

    ' Function to slow down the process for a specified duration
    Private Sub SlowDownProcessForDuration(process As Process, duration As TimeSpan)
        Try
            ChangeProcessPriority(process, True)
        Catch ex As Exception
        End Try

        Try
            Do While Not process.HasExited
                PauseProcess(process, PauseTime)
                ResumeProcess(process, ResumeTime)

                ' Check if the slowdown duration has been reached
                If DateTime.Now - process.StartTime > duration Then
                    NormalExit(True)
                End If
            Loop
        Catch ex As UnauthorizedAccessException
            Console.WriteLine("Error: Access Violation. No Rights. Try to run the Application as Administrator!")
            Console.ReadKey()
            NormalExit(False)
        Catch ex As Exception
            Console.WriteLine("Error:" + ex.ToString)
            NormalExit(False)
        End Try
    End Sub

    ' Function to pause the process for a specified time
    Private Sub PauseProcess(process As Process, milliseconds As Integer)
        Try
            If Not process.HasExited Then
                Dim threadIds As Integer() = New Integer(process.Threads.Count - 1) {}
                For i As Integer = 0 To process.Threads.Count - 1
                    threadIds(i) = process.Threads(i).Id
                Next

                ' Pause all threads of the process
                For Each threadId As Integer In threadIds
                    Dim threadToPause As ProcessThread = process.Threads.Cast(Of ProcessThread)().FirstOrDefault(Function(t) t.Id = threadId)
                    If threadToPause IsNot Nothing Then
                        Dim threadHandle As IntPtr = OpenThread(ThreadAccess.SUSPEND_RESUME, False, CType(threadToPause.Id, UInteger))
                        If threadHandle <> IntPtr.Zero Then
                            SuspendThread(threadHandle)
                            CloseHandle(threadHandle)
                        End If
                    End If
                Next

                ' Pause the main thread to achieve the delay
                Thread.Sleep(milliseconds)
            End If
        Catch ex As UnauthorizedAccessException
            Console.WriteLine("Error: Access Violation. No Rights. Try to run the Application as Administrator!")
            Console.ReadKey()
            NormalExit(False)
        Catch ex As Exception
            Console.WriteLine("Error:" + ex.Message)
            NormalExit(True)
        End Try
    End Sub

    ' Function to resume the process after a specified time
    Private Sub ResumeProcess(process As Process, milliseconds As Integer)
        Try
            If Not process.HasExited Then
                Dim threadIds As Integer() = New Integer(process.Threads.Count - 1) {}
                For i As Integer = 0 To process.Threads.Count - 1
                    threadIds(i) = process.Threads(i).Id
                Next

                ' Resume all threads
                For Each threadId As Integer In threadIds
                    Dim threadToResume As ProcessThread = process.Threads.Cast(Of ProcessThread)().FirstOrDefault(Function(t) t.Id = threadId)
                    If threadToResume IsNot Nothing Then
                        Dim threadHandle As IntPtr = OpenThread(ThreadAccess.SUSPEND_RESUME, False, CType(threadToResume.Id, UInteger))
                        If threadHandle <> IntPtr.Zero Then
                            Dim suspendCount As Integer = 0
                            Do
                                suspendCount = ResumeThread(threadHandle)
                            Loop While suspendCount > 0
                            CloseHandle(threadHandle)
                        End If
                    End If
                Next

                ' Pause the main thread to achieve the delay
                Thread.Sleep(milliseconds)
            End If
        Catch ex As UnauthorizedAccessException
            Console.WriteLine("Error: Access Violation. No Rights. Try to run the Application as Administrator!")
            Console.ReadKey()
            NormalExit(False)
        Catch ex As Exception
            Console.WriteLine("Error:" + ex.Message)
            NormalExit(True)
        End Try
    End Sub
#End Region

    ' Region for cleanup and exit-related methods
#Region "Retarded Exit"
    Dim secondsRemaining As Integer = 5
    Dim cancellationTokenSource As New Threading.CancellationTokenSource()
    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function SetConsoleCtrlHandler(ByVal handlerRoutine As ConsoleCtrlDelegate, ByVal add As Boolean) As Boolean  'Funktion wird benötigt, um den Prozess wiederherzustellen beim beenden
    End Function
    Private Delegate Function ConsoleCtrlDelegate(ByVal ctrlType As CtrlType) As Boolean
    Private Enum CtrlType
        CTRL_C_EVENT = 0
        CTRL_BREAK_EVENT = 1
        CTRL_CLOSE_EVENT = 2
        CTRL_LOGOFF_EVENT = 5
        CTRL_SHUTDOWN_EVENT = 6
    End Enum
    ' Function for normal exit handling
    Async Sub NormalExit(ExitMsg As Boolean)
        Try
            Dim process As Process = Nothing
            process = Process.GetProcessesByName(ProcessName).FirstOrDefault()
            If process IsNot Nothing Then
                Restore(process)
            End If
        Catch ex As Exception
        End Try

        Try
            If ExitMsg = False Then
                cancellationTokenSource.Cancel()
                KillOwnProcess()
                Exit Sub
            Else
                Dim updateTask = Task.Run(Async Function()
                                              While secondsRemaining > 0 AndAlso Not Console.KeyAvailable
                                                  UpdateLastCharacter(secondsRemaining.ToString)
                                                  Await Task.Delay(1000)
                                                  secondsRemaining -= 1
                                              End While
                                          End Function, cancellationTokenSource.Token)
                updateTask.Wait()
                cancellationTokenSource.Cancel()
            End If

            KillOwnProcess()
            Exit Sub
        Catch ex As Exception
            KillOwnProcess()
            Exit Sub
        End Try
    End Sub

    ' Function to update the last character in the console
    Sub UpdateLastCharacter(newLastCharacter As String)
        Console.SetCursorPosition(0, Console.CursorTop)
        Console.Write("Press any Key to Exit. Application will close in seconds: " + newLastCharacter.PadLeft(2) + "   ")
    End Sub

    ' Console Ctrl handler
    Private Function MyConsoleCtrlHandler(ByVal ctrlType As CtrlType) As Boolean
        NormalExit(False)
        Return False
    End Function

    ' Function to restore process settings
    Sub Restore(ByVal process As Process)
        Try
            ResumeProcess(process, 0)
            ChangeProcessPriority(process, False)
            If process IsNot Nothing Then
                ' Uncomment the following lines if needed for handling console window
                ' ShowWindow(process.MainWindowHandle, SW_RESTORE)
                ' SetForegroundWindow(process.MainWindowHandle)
            End If
            Console.WriteLine($"{ProcessName} is working normally now.")
        Finally
            Try
                If process IsNot Nothing Then
                    process.Dispose()
                End If
                If Timer IsNot Nothing Then
                    Timer.Dispose()
                End If
            Catch ex As Exception
            End Try
        End Try
    End Sub

    ' Function to handle application exit event
    Sub OnApplicationExit(sender As Object, e As EventArgs)
        NormalExit(False)
    End Sub

    ' Function to handle unhandled exceptions
    Private Sub CurrentDomain_UnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        Try
            Dim exception As Exception = DirectCast(e.ExceptionObject, Exception)
            Console.WriteLine(exception.Message)
        Catch ex As Exception
        End Try

        Try
            NormalExit(False)
        Catch ex As Exception
            KillOwnProcess()
        End Try
    End Sub

    ' Function to kill the application process
    Sub KillOwnProcess()
        Try
            Dim processName As String = Process.GetCurrentProcess().ProcessName
            Dim cmdProcess As New Process()
            cmdProcess.StartInfo.CreateNoWindow = True
            cmdProcess.StartInfo.UseShellExecute = False
            cmdProcess.StartInfo.FileName = "cmd.exe"
            cmdProcess.StartInfo.Arguments = $"/c taskkill /F /IM {processName}.exe"
            cmdProcess.Start()
            cmdProcess.WaitForExit()
        Catch ex As Exception
        End Try
    End Sub
#End Region

    ' Region for settings-related methods
#Region "Settings"
    ' Function to load settings from the config file
    Sub LoadSettings()
        Try
            If File.Exists(ConfigFilePath) Then
                Dim xmlDoc As XDocument = XDocument.Load(ConfigFilePath)
                Dim rootElement As XElement = xmlDoc.Element("Settings")

                If rootElement IsNot Nothing Then
                    Settings.PauseTime = Convert.ToUInt16(rootElement.Element("PauseTime").Value)
                    Settings.ResumeTime = Convert.ToUInt16(rootElement.Element("ResumeTime").Value)

                    If Settings.PauseTime >= 10 AndAlso Settings.ResumeTime >= 1 Then
                        Console.WriteLine("Settings loaded!")
                        PauseTime = Settings.PauseTime
                        Console.WriteLine($"PauseTime: {PauseTime}")
                        ResumeTime = Settings.ResumeTime
                        Console.WriteLine($"ResumeTime: {ResumeTime}")
                    End If
                End If
            Else
                Console.WriteLine("No Config found!")
                Console.WriteLine("Using Default Settings:")
                Console.WriteLine($"ResumeTime: {ResumeTime}{vbNewLine}PauseTime: {PauseTime}")
                CreateSettings()
            End If
        Catch ex As Exception
            Console.WriteLine($"Error: {ex.Message}")
            Console.WriteLine("Using Default Settings:")
            Console.WriteLine($"ResumeTime: {ResumeTime}{vbNewLine}PauseTime: {PauseTime}")
        End Try
    End Sub

    ' Function to create default settings if config file is not found
    Sub CreateSettings()
        Try
            Dim xmlDoc As New XDocument(
                New XElement("Settings",
                    New XElement("PauseTime", Settings.PauseTime),
                    New XElement("ResumeTime", Settings.ResumeTime)
                )
            )
            xmlDoc.Save(ConfigFilePath)
            Console.WriteLine($"Settings File created: {ConfigFilePath}")
        Catch ex As Exception
            Console.WriteLine($"Error creating settings: {ex.Message}")
        End Try
    End Sub
#End Region
End Module
