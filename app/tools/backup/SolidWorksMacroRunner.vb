Option Strict On
Option Explicit On

Public Class SolidWorksMacroRunner

    Public Shared Function Run(
        swApp As Object,
        macroPath As String,
        moduleName As Object,
        procName As Object,
        ByRef errMsg As String) As Boolean

        errMsg = ""

        Try
            Dim longErr As Integer = 0
            Dim ok As Boolean =
                swApp.RunMacro2(macroPath, moduleName, procName, 2, longErr)

            If Not ok Then
                errMsg = "RunMacro2 failed. Error = " & longErr
            End If

            Return ok

        Catch ex As Exception
            errMsg = ex.Message
            Return False
        End Try

    End Function

End Class
