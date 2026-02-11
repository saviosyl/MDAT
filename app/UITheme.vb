Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

' ============================================================
' UITheme — Premium theme foundation
' Provides Light/Dark color palettes, font definitions,
' and helper methods for applying themes to forms and controls.
'
' BACKWARD COMPAT: Old module-level constants are preserved
' as Shared ReadOnly fields so AboutForm, EngineeringNotepad,
' and any other consumer keeps compiling unchanged.
' ============================================================
Public Class UITheme

    ' ================= THEME ENUM =================
    Public Enum ThemeKind
        Light
        Dark
    End Enum

    ' ================= CURRENT STATE =================
    Private Shared _current As ThemeKind = ThemeKind.Light

    Public Shared Property CurrentTheme As ThemeKind
        Get
            Return _current
        End Get
        Set(value As ThemeKind)
            _current = value
        End Set
    End Property

    Public Shared ReadOnly Property IsDark As Boolean
        Get
            Return _current = ThemeKind.Dark
        End Get
    End Property

    ' ================= APP TEXT (backward compat) =================
    Public Shared ReadOnly AppName As String = "MDAT – Mechanical Design Automation Tool"
    Public Shared ReadOnly AppTagline As String = "MetaMech Mechanical Design Automation"

    ' ================= LIGHT PALETTE =================
    Private Shared ReadOnly L_BG_MAIN As Color = Color.FromArgb(248, 249, 250)       ' #F8F9FA
    Private Shared ReadOnly L_BG_PANEL As Color = Color.FromArgb(255, 255, 255)       ' #FFFFFF
    Private Shared ReadOnly L_BG_CARD As Color = Color.FromArgb(255, 255, 255)        ' #FFFFFF
    Private Shared ReadOnly L_ACCENT As Color = Color.FromArgb(0, 180, 255)           ' #00B4FF
    Private Shared ReadOnly L_ACCENT_TEAL As Color = Color.FromArgb(0, 180, 216)      ' #00B4D8
    Private Shared ReadOnly L_TEXT_PRIMARY As Color = Color.FromArgb(26, 26, 46)       ' #1A1A2E
    Private Shared ReadOnly L_TEXT_SECONDARY As Color = Color.FromArgb(108, 117, 125)  ' #6C757D
    Private Shared ReadOnly L_BORDER As Color = Color.FromArgb(224, 224, 224)          ' #E0E0E0
    Private Shared ReadOnly L_SUCCESS As Color = Color.FromArgb(40, 167, 69)           ' #28A745
    Private Shared ReadOnly L_WARNING As Color = Color.FromArgb(255, 193, 7)           ' #FFC107
    Private Shared ReadOnly L_ERROR As Color = Color.FromArgb(220, 53, 69)             ' #DC3545
    Private Shared ReadOnly L_HEADER_BG As Color = Color.FromArgb(26, 26, 46)          ' #1A1A2E
    Private Shared ReadOnly L_HEADER_TEXT As Color = Color.FromArgb(255, 255, 255)      ' #FFFFFF
    Private Shared ReadOnly L_LOG_BG As Color = Color.FromArgb(30, 30, 46)             ' #1E1E2E
    Private Shared ReadOnly L_LOG_TEXT As Color = Color.FromArgb(212, 212, 212)         ' #D4D4D4
    Private Shared ReadOnly L_BTN_BG As Color = Color.FromArgb(240, 242, 245)          ' #F0F2F5
    Private Shared ReadOnly L_BTN_HOVER As Color = Color.FromArgb(227, 232, 239)       ' #E3E8EF
    Private Shared ReadOnly L_BTN_SELECTED As Color = Color.FromArgb(0, 180, 255)      ' #00B4FF
    Private Shared ReadOnly L_FOOTER_BG As Color = Color.FromArgb(240, 242, 245)       ' #F0F2F5

    ' ================= DARK PALETTE =================
    Private Shared ReadOnly D_BG_MAIN As Color = Color.FromArgb(11, 30, 52)            ' #0B1E34
    Private Shared ReadOnly D_BG_PANEL As Color = Color.FromArgb(15, 42, 68)           ' #0F2A44
    Private Shared ReadOnly D_BG_CARD As Color = Color.FromArgb(22, 44, 69)            ' #162C45
    Private Shared ReadOnly D_ACCENT As Color = Color.FromArgb(0, 180, 255)            ' #00B4FF
    Private Shared ReadOnly D_ACCENT_TEAL As Color = Color.FromArgb(0, 229, 208)       ' #00E5D0
    Private Shared ReadOnly D_TEXT_PRIMARY As Color = Color.FromArgb(234, 244, 255)     ' #EAF4FF
    Private Shared ReadOnly D_TEXT_SECONDARY As Color = Color.FromArgb(169, 199, 232)   ' #A9C7E8
    Private Shared ReadOnly D_BORDER As Color = Color.FromArgb(30, 58, 95)             ' #1E3A5F
    Private Shared ReadOnly D_SUCCESS As Color = Color.FromArgb(0, 229, 208)            ' #00E5D0
    Private Shared ReadOnly D_WARNING As Color = Color.FromArgb(255, 193, 7)            ' #FFC107
    Private Shared ReadOnly D_ERROR As Color = Color.FromArgb(244, 67, 54)              ' #F44336
    Private Shared ReadOnly D_HEADER_BG As Color = Color.FromArgb(7, 21, 37)            ' #071525
    Private Shared ReadOnly D_HEADER_TEXT As Color = Color.FromArgb(255, 255, 255)       ' #FFFFFF
    Private Shared ReadOnly D_LOG_BG As Color = Color.FromArgb(5, 14, 26)               ' #050E1A
    Private Shared ReadOnly D_LOG_TEXT As Color = Color.FromArgb(169, 199, 232)          ' #A9C7E8
    Private Shared ReadOnly D_BTN_BG As Color = Color.FromArgb(22, 44, 69)              ' #162C45
    Private Shared ReadOnly D_BTN_HOVER As Color = Color.FromArgb(30, 58, 95)           ' #1E3A5F
    Private Shared ReadOnly D_BTN_SELECTED As Color = Color.FromArgb(0, 180, 255)       ' #00B4FF
    Private Shared ReadOnly D_FOOTER_BG As Color = Color.FromArgb(10, 25, 41)           ' #0A1929

    ' ================= CURRENT-THEME ACCESSORS =================
    Public Shared ReadOnly Property BgMain As Color
        Get
            Return If(IsDark, D_BG_MAIN, L_BG_MAIN)
        End Get
    End Property
    Public Shared ReadOnly Property BgPanel As Color
        Get
            Return If(IsDark, D_BG_PANEL, L_BG_PANEL)
        End Get
    End Property
    Public Shared ReadOnly Property BgCard As Color
        Get
            Return If(IsDark, D_BG_CARD, L_BG_CARD)
        End Get
    End Property
    Public Shared ReadOnly Property Accent As Color
        Get
            Return If(IsDark, D_ACCENT, L_ACCENT)
        End Get
    End Property
    Public Shared ReadOnly Property AccentTeal As Color
        Get
            Return If(IsDark, D_ACCENT_TEAL, L_ACCENT_TEAL)
        End Get
    End Property
    Public Shared ReadOnly Property TextPrimary As Color
        Get
            Return If(IsDark, D_TEXT_PRIMARY, L_TEXT_PRIMARY)
        End Get
    End Property
    Public Shared ReadOnly Property TextSecondary As Color
        Get
            Return If(IsDark, D_TEXT_SECONDARY, L_TEXT_SECONDARY)
        End Get
    End Property
    Public Shared ReadOnly Property Border As Color
        Get
            Return If(IsDark, D_BORDER, L_BORDER)
        End Get
    End Property
    Public Shared ReadOnly Property Success As Color
        Get
            Return If(IsDark, D_SUCCESS, L_SUCCESS)
        End Get
    End Property
    Public Shared ReadOnly Property Warning As Color
        Get
            Return If(IsDark, D_WARNING, L_WARNING)
        End Get
    End Property
    Public Shared ReadOnly Property [Error] As Color
        Get
            Return If(IsDark, D_ERROR, L_ERROR)
        End Get
    End Property
    Public Shared ReadOnly Property HeaderBg As Color
        Get
            Return If(IsDark, D_HEADER_BG, L_HEADER_BG)
        End Get
    End Property
    Public Shared ReadOnly Property HeaderText As Color
        Get
            Return If(IsDark, D_HEADER_TEXT, L_HEADER_TEXT)
        End Get
    End Property
    Public Shared ReadOnly Property LogBg As Color
        Get
            Return If(IsDark, D_LOG_BG, L_LOG_BG)
        End Get
    End Property
    Public Shared ReadOnly Property LogText As Color
        Get
            Return If(IsDark, D_LOG_TEXT, L_LOG_TEXT)
        End Get
    End Property
    Public Shared ReadOnly Property BtnBg As Color
        Get
            Return If(IsDark, D_BTN_BG, L_BTN_BG)
        End Get
    End Property
    Public Shared ReadOnly Property BtnHover As Color
        Get
            Return If(IsDark, D_BTN_HOVER, L_BTN_HOVER)
        End Get
    End Property
    Public Shared ReadOnly Property BtnSelected As Color
        Get
            Return If(IsDark, D_BTN_SELECTED, L_BTN_SELECTED)
        End Get
    End Property
    Public Shared ReadOnly Property FooterBg As Color
        Get
            Return If(IsDark, D_FOOTER_BG, L_FOOTER_BG)
        End Get
    End Property

    ' ================= FONTS =================
    Public Shared ReadOnly Property FontTitle As Font
        Get
            Return New Font("Segoe UI", 14.0F, FontStyle.Bold)
        End Get
    End Property
    Public Shared ReadOnly Property FontSubtitle As Font
        Get
            Return New Font("Segoe UI", 10.0F, FontStyle.Regular)
        End Get
    End Property
    Public Shared ReadOnly Property FontBody As Font
        Get
            Return New Font("Segoe UI", 9.5F, FontStyle.Regular)
        End Get
    End Property
    Public Shared ReadOnly Property FontLog As Font
        Get
            Return New Font("Consolas", 9.0F, FontStyle.Regular)
        End Get
    End Property
    Public Shared ReadOnly Property FontButton As Font
        Get
            Return New Font("Segoe UI", 9.5F, FontStyle.Bold)
        End Get
    End Property
    Public Shared ReadOnly Property FontSmall As Font
        Get
            Return New Font("Segoe UI", 8.0F, FontStyle.Regular)
        End Get
    End Property

    ' ================= HELPER: SWITCH THEME =================
    Public Shared Sub SetTheme(kind As ThemeKind)
        _current = kind
    End Sub

    ' ================= HELPER: APPLY TO FORM =================
    Public Shared Sub ApplyToForm(frm As Form)
        If frm Is Nothing Then Return
        Try
            frm.BackColor = BgMain
            frm.ForeColor = TextPrimary
        Catch
        End Try
    End Sub

    ' ================= HELPER: APPLY TO BUTTON =================
    Public Shared Sub ApplyToButton(btn As Button, isSelected As Boolean)
        If btn Is Nothing Then Return
        Try
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 1
            btn.FlatAppearance.BorderColor = Accent

            If isSelected Then
                btn.BackColor = BtnSelected
                btn.ForeColor = Color.White
            Else
                btn.BackColor = BtnBg
                btn.ForeColor = TextPrimary
            End If
        Catch
        End Try
    End Sub

    ' ================= BACKWARD COMPATIBILITY =================
    ' Old code (AboutForm, EngineeringNotepad, etc.) references these
    ' as UITheme.BG_MAIN, UITheme.TEXT_PRIMARY, etc.
    ' Keep them as Shared ReadOnly fields with the original dark-base values.
    Public Shared ReadOnly BG_MAIN As Color = Color.FromArgb(18, 22, 30)
    Public Shared ReadOnly BG_PANEL As Color = Color.FromArgb(28, 34, 44)
    Public Shared ReadOnly BG_CARD As Color = Color.FromArgb(24, 28, 38)

    Public Shared ReadOnly TEXT_PRIMARY As Color = Color.White
    Public Shared ReadOnly TEXT_MUTED As Color = Color.FromArgb(200, 200, 200)

    Public Shared ReadOnly ACCENT_PRIMARY As Color = Color.FromArgb(150, 90, 190)

    Public Shared ReadOnly BTN_BG As Color = Color.FromArgb(40, 45, 58)
    Public Shared ReadOnly BTN_TEXT As Color = Color.White

End Class
