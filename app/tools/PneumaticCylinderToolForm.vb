Option Strict On
Option Explicit On

Imports System
Imports System.Windows.Forms

' Clean launcher name (tidy).
' This just opens the Pneumatic Cylinder Calculator UI.

Public Class PneumaticCylinderToolForm
    Inherits PneumaticCylinderCalculatorForm

    Public Sub New()
        MyBase.New()
        Me.Text = "Pneumatic Cylinder Calculator"
    End Sub

End Class
