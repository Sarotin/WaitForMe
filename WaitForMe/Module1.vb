Imports WaitForMe.My
Imports System.IO
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Security.Principal

Module Module1
    Dim timer As Timer
    Dim WaitingTime As UInt16
    Private dotCounter As UInteger = 0
    Private Const PName As String = "League of Legends"
    Dim PauseTime As UShort = 1300
    Dim ResumeTime As UShort = 5
    ReadOnly Params As String


#Region "Startup"
    Sub Main()
        Try
            If IsRunningAsAdmin() = False Then
                Console.WriteLine("Error: No Admin Privileges!")
                Console.ReadKey()
                Environment.Exit(0)
            End If
            SetConsoleCtrlHandler(New ConsoleCtrlDelegate(AddressOf MyConsoleCtrlHandler), True)
            AddHandler AppDomain.CurrentDomain.ProcessExit, AddressOf OnApplicationExit

            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CurrentDomain_UnhandledException 'ExceptionMsgBox
        Catch ex As Exception
        End Try

        Try
            LoadSettings()
        Catch ex As Exception
        End Try

        Try
            If Environment.GetCommandLineArgs().Length > 1 Then
                If ProcessCommandLineArguments() = False Then
                    UserInputDelay()
                End If
            Else
                UserInputDelay()
            End If
            Console.WriteLine($"Delay set to: {WaitingTime}" + " Minutes")
            Console.WriteLine("Waiting for Process")
        Catch ex As Exception
            Environment.Exit(0)
        End Try

        Try 'Extra Try, da Timer nicht läuft
            If timer Is Nothing Then
                timer = New Timer(AddressOf CheckProcessStatus, Nothing, 0, 100)
            End If

            'Später ersetzen
            Do
                Console.ReadLine()
            Loop
            NormalExit(True)
        Catch ex As Exception
            NormalExit(False)
        End Try
    End Sub
    Function IsRunningAsAdmin() As Boolean
        Try
            Dim identity As WindowsIdentity = WindowsIdentity.GetCurrent()
            Dim principal As New WindowsPrincipal(identity)

            'Überprüfe, ob der aktuelle Benutzer Mitglied der Administratorengruppe ist
            Return principal.IsInRole(WindowsBuiltInRole.Administrator)
        Catch ex As Exception
            'Kein return false, es kann sein das User Adminrechte hat.
            Return True
        End Try
    End Function
    Sub UserInputDelay()
        'Parameter
        Dim key As Integer
        Console.WriteLine("How long do you want to suspend the loading process? (1-4 Min):")
        If Not Integer.TryParse(Console.ReadLine(), key) OrElse key < 1 OrElse key > 4 Then
            Console.WriteLine("Invalid Input!")
            Console.WriteLine("Press any Key to Exit.")
            Console.ReadKey()
            NormalExit(False)
        Else

            WaitingTime = CUShort(key)  'Übergebe eingegebenen Wert
        End If

    End Sub
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
#Region "Process Manipulation"
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

    Private Sub CheckProcessStatus(state As Object)
        Try
            Dim process As Process = Process.GetProcessesByName(PName).FirstOrDefault()
            If process IsNot Nothing Then
                'Stoppe, wenn der Prozess läuft
                timer.Dispose()
                Console.WriteLine($"Process {process.Id} ({process.ProcessName}) found. Slowing down.")


                'Verlangsamen Sie den Prozess für die angegebene Zeit
                SlowDownProcessForDuration(process, TimeSpan.FromMinutes(WaitingTime))
            Else
                'Der Prozess läuft nicht
                dotCounter += 1
                If dotCounter Mod 100 = 0 Or dotCounter = 4294967290 Then ' Überprüfe ob ein Vielfaches von 100 // 4294967295(U32BIT) Falls die Anwendung Jahre läuft :)
                    Thread.Sleep(100)
                    Console.Write(".")
                End If

            End If
        Catch ex As Exception
            Console.WriteLine("Error:" + ex.ToString)
            NormalExit(False)
        End Try
    End Sub
    Sub ChangeProcessPriority(process As Process, Slow As Boolean)
        Try
            ' Ändere die Priorität
            If Slow = True Then
                process.PriorityClass = ProcessPriorityClass.BelowNormal  'Idle könnte zu Problemen führen.
                Console.WriteLine(PName + " Priority: Low")
            Else
                process.PriorityClass = ProcessPriorityClass.Normal
                Console.WriteLine(PName + " Priority: Normal")
            End If

        Catch ex As Exception
            Console.WriteLine($"Error: Cannot change Priority {ex.Message}")
            Console.WriteLine("This Application will still try to delay the Game start. This Function is not crucial")
        End Try
    End Sub


    Private Sub SlowDownProcessForDuration(process As Process, duration As TimeSpan)
        Try
            ChangeProcessPriority(process, True)
        Catch ex As Exception

        End Try
        Try
            Do While Not process.HasExited
                PauseProcess(process, PauseTime)
                ResumeProcess(process, ResumeTime)

                'Überprüfen Sie, ob die Verlangsamungsdauer erreicht wurde
                If DateTime.Now - process.StartTime > duration Then
                    NormalExit(True)
                End If
            Loop


        Catch ex As UnauthorizedAccessException
            Console.WriteLine("Error: Access Violation. No Rights. Try to run the Application as Adminstrator!")
            Console.ReadKey()
            NormalExit(False)

        Catch ex As Exception
            Console.WriteLine("Error:" + ex.ToString)
            NormalExit(False)
        End Try
    End Sub
    Private Sub SlowDownProcessForDuration_OLD(process As Process, duration As TimeSpan)
        Try
            Do
                Try
                    If process.HasExited Then Exit Do ' Schleife beenden, wenn der Prozess beendet wurde
                    ChangeProcessPriority(process, True)
                    PauseProcess(process, PauseTime)
                    ResumeProcess(process, ResumeTime)
                    If DateTime.Now - process.StartTime > duration Then 'Durch process.StartTime wird Sichergestellt, dass der Prozess nicht zu lange angehalten wird.
                        NormalExit(True)
                    End If
                Catch ex As UnauthorizedAccessException
                    Console.WriteLine("Error: Access Violation. No Rights. Try to run the Application as Adminstrator!")
                    Console.ReadKey()
                    NormalExit(False)
                Catch ex As Exception
                    Console.WriteLine("Error:" + ex.Message)  'Schleife wird fortgesetzt
                End Try
            Loop
        Catch ex As Exception
            Console.WriteLine("Outer Error:" + ex.Message)
            ' Behandle den äußeren Fehler, wenn die Schleife aufgrund eines schwerwiegenden Fehlers verlassen wird
            NormalExit(True)
        End Try
    End Sub

    Private Sub PauseProcess(process As Process, milliseconds As Integer)
        ' Pausieren Sie den Prozess für die angegebene Zeit
        Try
            If Not process.HasExited Then
                Dim threadIds As Integer() = New Integer(process.Threads.Count - 1) {}
                For i As Integer = 0 To process.Threads.Count - 1
                    threadIds(i) = process.Threads(i).Id
                Next
                ' Pausieren Sie alle Threads des Prozesses
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
                ' Pausieren Sie den Hauptthread, um die Verzögerung zu erreichen
                Thread.Sleep(milliseconds) 'XX2
            End If
        Catch ex As UnauthorizedAccessException
            Console.WriteLine("Error: Access Violation. No Rights. Try to run the Application as Adminstrator!")
            Console.ReadKey()
            NormalExit(False)
        Catch ex As Exception
            Console.WriteLine("Error:" + ex.Message)
            NormalExit(True)
        End Try
    End Sub


    Private Sub ResumeProcess(process As Process, milliseconds As Integer)

        Try
            If Not process.HasExited Then
                Dim threadIds As Integer() = New Integer(process.Threads.Count - 1) {}
                For i As Integer = 0 To process.Threads.Count - 1
                    threadIds(i) = process.Threads(i).Id
                Next

                'Resume alle Threads
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

                'Pausieren Sie den Hauptthread, um die Verzögerung zu erreichen
                Thread.Sleep(milliseconds) 'XX2
            End If
        Catch ex As UnauthorizedAccessException
            Console.WriteLine("Error: Access Violation. No Rights. Try to run the Application as Adminstrator!")
            Console.ReadKey()
            NormalExit(False)

        Catch ex As Exception
            Console.WriteLine("Error:" + ex.Message)
            NormalExit(True)
        End Try
    End Sub
#End Region
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
    Async Sub NormalExit(ExitMsg As Boolean)
        Try
            Dim Prozess As Process = Nothing
            Prozess = Process.GetProcessesByName(PName).FirstOrDefault()
            If Prozess IsNot Nothing Then
                Restore(Prozess)
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
                updateTask.Wait() 'Warte auf eine Benutzereingabe
                cancellationTokenSource.Cancel() 'Breche die Aufgabe ab, falls nicht abgeschlossen
            End If
            KillOwnProcess()
            Exit Sub
        Catch ex As Exception
            KillOwnProcess()
            Exit Sub
        End Try
    End Sub
    Sub UpdateLastCharacter(newLastCharacter As String)
        Console.SetCursorPosition(0, Console.CursorTop)
        Console.Write("Press any Key to Exit. Application will close in seconds: " + newLastCharacter.PadLeft(2) + "   ")
    End Sub
    Private Function MyConsoleCtrlHandler(ByVal ctrlType As CtrlType) As Boolean  'Beim beenden ausgeführt
        NormalExit(False)
        Return False ' oder True, je nach Bedarf       
    End Function
    Sub Restore(ByVal process As Process)
        Try
            ResumeProcess(process, 0) 'Prozess zur Sicherheit nochmal resumen.
            ChangeProcessPriority(process, False) 'Setzt Priorität auf Normal
            If process IsNot Nothing Then
                '     ShowWindow(process.MainWindowHandle, SW_RESTORE)
                '     SetForegroundWindow(process.MainWindowHandle)
            End If
            Console.WriteLine(PName + " is working normal now.")
        Finally 'Der Finally-Block wird immer ausgeführt, unabhängig davon, ob ein Fehler auftritt oder nicht. 
            Try
                If process IsNot Nothing Then
                    process.Dispose() ' Hier kannst du den Prozess freigeben, falls erforderlich
                End If
                If timer IsNot Nothing Then
                    timer.Dispose() ' Versuche den Timer zu dispoieren, falls er nicht Nothing ist
                End If
            Catch ex As Exception
            End Try
        End Try
    End Sub

    Sub OnApplicationExit(sender As Object, e As EventArgs)
        NormalExit(False)
    End Sub
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
    Sub KillOwnProcess()
        Try
            Dim processName As String = Process.GetCurrentProcess().ProcessName
            Dim cmdProcess As New Process()
            cmdProcess.StartInfo.CreateNoWindow = True
            cmdProcess.StartInfo.UseShellExecute = False
            cmdProcess.StartInfo.FileName = "cmd.exe"
            cmdProcess.StartInfo.Arguments = "/c taskkill /F /IM " & processName & ".exe"
            cmdProcess.Start()
            cmdProcess.WaitForExit()
        Catch ex As Exception
        End Try
    End Sub
#End Region
#Region "Settings"
    ReadOnly configFilePath As String = Path.Combine(My.Application.Info.DirectoryPath & "\" & "Settings.config")
    Sub LoadSettings()

        Try
            If File.Exists(configFilePath) Then
                Dim xmlDoc As XDocument = XDocument.Load(configFilePath)
                Dim rootElement As XElement = xmlDoc.Element("Settings")

                If rootElement IsNot Nothing Then
                    Settings.PauseTime = Convert.ToUInt16(rootElement.Element("PauseTime").Value)
                    Settings.ResumeTime = Convert.ToUInt16(rootElement.Element("ResumeTime").Value)
                    If Settings.PauseTime >= 10 And Settings.ResumeTime >= 1 Then
                        Console.WriteLine("Settings loaded!")
                        PauseTime = Settings.PauseTime
                        Console.WriteLine("PauseTime: " + PauseTime.ToString)
                        ResumeTime = Settings.ResumeTime
                        Console.WriteLine("ResumeTime: " + ResumeTime.ToString)
                    End If
                End If
            Else
                Console.WriteLine("No Config found!")
                Console.WriteLine("Using Default Settings:")
                Console.WriteLine("ResumeTime: " + ResumeTime.ToString + vbNewLine + "PauseTime: " + PauseTime.ToString)
                CreateSettings()
            End If
        Catch ex As Exception
            Console.WriteLine("Error:" + ex.Message)
            Console.WriteLine("Using Default Settings:")
            Console.WriteLine("ResumeTime: " + ResumeTime.ToString + vbNewLine + "PauseTime: " + PauseTime.ToString)

        End Try
    End Sub

    Sub CreateSettings()
        Try
            Dim xmlDoc As New XDocument(
                New XElement("Settings",
                    New XElement("PauseTime", Settings.PauseTime),
                    New XElement("ResumeTime", Settings.ResumeTime)
                )
            )
            xmlDoc.Save(configFilePath)
            Console.WriteLine("Settings File created: " + configFilePath.ToString)
        Catch ex As Exception
            Console.WriteLine("Error creating settings: + " + ex.Message)
        End Try
    End Sub
#End Region
End Module



