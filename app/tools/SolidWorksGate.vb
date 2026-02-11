Option Strict off
Option Explicit On

Imports System.Runtime.InteropServices

Public Module SolidWorksGate

    Public Function GetSolidWorksApp(ByRef swApp As Object, ByRef errMsg As String) As Boolean
        Try
            swApp = Marshal.GetActiveObject("SldWorks.Application")
            Return True
        Catch
            errMsg = "SolidWorks is not running."
            swApp = Nothing
            Return False
        End Try
    End Function

    Public Function HasActiveDocument(swApp As Object, ByRef model As Object, ByRef errMsg As String) As Boolean
        Try
            model = swApp.ActiveDoc
            If model Is Nothing Then
                errMsg = "No active document found in SolidWorks."
                Return False
            End If
            Return True
        Catch
            errMsg = "Unable to read active SolidWorks document."
            model = Nothing
            Return False
        End Try
    End Function

    Public Function ValidateDocumentType(model As Object, requiredType As Integer, ByRef errMsg As String) As Boolean
        Try
            Dim docType As Integer = CInt(CallByName(model, "GetType", CallType.Method))

            If docType <> requiredType Then
                Select Case requiredType
                    Case 2
                        errMsg = "Active document is not an ASSEMBLY (.SLDASM)."
                    Case 1
                        errMsg = "Active document is not a PART (.SLDPRT)."
                    Case 3
                        errMsg = "Active document is not a DRAWING (.SLDDRW)."
                    Case Else
                        errMsg = "Invalid document type requirement."
                End Select
                Return False
            End If

            Return True
        Catch
            errMsg = "Failed to validate document type."
            Return False
        End Try
    End Function

End Module
