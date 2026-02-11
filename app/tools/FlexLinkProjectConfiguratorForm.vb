Option Strict On
Option Explicit On

Imports System
Imports System.Windows.Forms

' MainForm.vb calls: New FlexLinkProjectConfiguratorForm()
' So this exact class name MUST exist.

Public Class FlexLinkProjectConfiguratorForm
    Inherits FlexLinkCalculatorForm

    Public Sub New()
        MyBase.New()
        Me.Text = "FlexLink Project Configurator"
    End Sub

End Class
