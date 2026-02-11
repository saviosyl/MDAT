Option Strict On
Option Explicit On

Imports System
Imports System.Drawing

' ============================================================
' UITheme (DEFAULT NAMESPACE)
' - Must be Public Module named "UITheme"
' - Used by AboutForm, EngineeringNotepad, and theme reading
' ============================================================
Public Module UITheme

    ' ================= APP TEXT =================
    Public Const AppName As String = "MDAT – Mechanical Design Automation Tool"
    Public Const AppTagline As String = "MetaMech Mechanical Design Automation"

    ' ================= COLORS (DARK BASE) =================
    Public ReadOnly BG_MAIN As Color = Color.FromArgb(18, 22, 30)
    Public ReadOnly BG_PANEL As Color = Color.FromArgb(28, 34, 44)
    Public ReadOnly BG_CARD As Color = Color.FromArgb(24, 28, 38)

    Public ReadOnly TEXT_PRIMARY As Color = Color.White
    Public ReadOnly TEXT_MUTED As Color = Color.FromArgb(200, 200, 200)

    Public ReadOnly ACCENT_PRIMARY As Color = Color.FromArgb(150, 90, 190)

    Public ReadOnly BTN_BG As Color = Color.FromArgb(40, 45, 58)
    Public ReadOnly BTN_TEXT As Color = Color.White

End Module
