Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.Threading

Module Program_PneumaticCircuitHelper

    <STAThread()>
    Public Sub Main()

        Dim logPath As String = ""
        Try
            Dim exeFolder As String = AppDomain.CurrentDomain.BaseDirectory
            logPath = Path.Combine(exeFolder, "PneumaticCircuitHelper_startup_log.txt")

            ' Reset log each run
            Try
                File.WriteAllText(logPath, "=== PneumaticCircuitHelper Startup Log ===" & Environment.NewLine)
                File.AppendAllText(logPath, "Time: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & Environment.NewLine)
                File.AppendAllText(logPath, "BaseDir: " & exeFolder & Environment.NewLine)
            Catch
            End Try

            ' Global exception hooks (helps catch hidden startup crashes)
            AddHandler Application.ThreadException,
                Sub(sender As Object, e As ThreadExceptionEventArgs)
                    Try
                        File.AppendAllText(logPath,
                            Environment.NewLine &
                            "[Application.ThreadException]" & Environment.NewLine &
                            e.Exception.ToString() & Environment.NewLine)
                    Catch
                    End Try

                    MessageBox.Show("Unhandled UI error:" & Environment.NewLine & e.Exception.Message &
                                    Environment.NewLine & Environment.NewLine &
                                    "See log: " & logPath,
                                    "Pneumatic Circuit Helper - Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Sub

            AddHandler AppDomain.CurrentDomain.UnhandledException,
                Sub(sender As Object, e As UnhandledExceptionEventArgs)
                    Try
                        Dim ex As Exception = TryCast(e.ExceptionObject, Exception)
                        If ex IsNot Nothing Then
                            File.AppendAllText(logPath,
                                Environment.NewLine &
                                "[AppDomain.UnhandledException]" & Environment.NewLine &
                                ex.ToString() & Environment.NewLine)
                        Else
                            File.AppendAllText(logPath,
                                Environment.NewLine &
                                "[AppDomain.UnhandledException] Non-Exception object" & Environment.NewLine)
                        End If
                    Catch
                    End Try
                End Sub

            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)

            Try
                File.AppendAllText(logPath, "Creating form..." & Environment.NewLine)
            Catch
            End Try

            Dim f As New PneumaticCircuitHelperForm()

            Try
                File.AppendAllText(logPath, "Running form..." & Environment.NewLine)
            Catch
            End Try

            Application.Run(f)

            Try
                File.AppendAllText(logPath, "Application closed normally." & Environment.NewLine)
            Catch
            End Try

        Catch ex As Exception
            Try
                If logPath <> "" Then
                    File.AppendAllText(logPath,
                        Environment.NewLine &
                        "[Main Catch]" & Environment.NewLine &
                        ex.ToString() & Environment.NewLine)
                End If
            Catch
            End Try

            MessageBox.Show("Startup failed:" & Environment.NewLine & ex.Message &
                            Environment.NewLine & Environment.NewLine &
                            "Check startup log in EXE folder.",
                            "Pneumatic Circuit Helper - Startup Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

End Module