Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Module ThemeApplier

    ' ============================================================
    ' ApplyTheme
    ' - Safe: only changes colors/fonts, never changes layout/positions
    ' - Keeps "special" colored buttons (Green/Yellow/Red) as-is
    ' ============================================================
    Public Sub ApplyTheme(target As Control, bg As Color, panel As Color, accent As Color, isDark As Boolean)
        If target Is Nothing Then Exit Sub

        Try
            ApplyToOne(target, bg, panel, accent, isDark)
            ApplyChildren(target, bg, panel, accent, isDark)
        Catch
            ' never let theme crash the app
        End Try
    End Sub

    Private Sub ApplyChildren(parent As Control, bg As Color, panel As Color, accent As Color, isDark As Boolean)
        Try
            For Each c As Control In parent.Controls
                Try
                    ApplyToOne(c, bg, panel, accent, isDark)
                    If c.HasChildren Then
                        ApplyChildren(c, bg, panel, accent, isDark)
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub

    Private Sub ApplyToOne(c As Control, bg As Color, panel As Color, accent As Color, isDark As Boolean)
        If c Is Nothing Then Exit Sub

        Dim textCol As Color = If(isDark, Color.White, Color.Black)
        Dim faintText As Color = If(isDark, Color.FromArgb(210, 210, 210), Color.FromArgb(60, 60, 60))

        Try
            ' ---- FORM ----
            If TypeOf c Is Form Then
                c.BackColor = bg
                c.ForeColor = textCol
                Exit Sub
            End If

            ' ---- PANELS ----
            If TypeOf c Is Panel Then
                c.BackColor = panel
                c.ForeColor = textCol
                Exit Sub
            End If

            ' ---- GROUPBOX ----
            If TypeOf c Is GroupBox Then
                c.BackColor = panel
                c.ForeColor = textCol
                Exit Sub
            End If

            ' ---- LABEL ----
            If TypeOf c Is Label Then
                c.BackColor = Color.Transparent
                c.ForeColor = textCol
                Exit Sub
            End If

            ' ---- LINKLABEL ----
            If TypeOf c Is LinkLabel Then
                Dim ll As LinkLabel = DirectCast(c, LinkLabel)
                ll.LinkColor = accent
                ll.ActiveLinkColor = accent
                ll.VisitedLinkColor = accent
                ll.ForeColor = accent
                Exit Sub
            End If

            ' ---- TEXTBOX ----
            If TypeOf c Is TextBox Then
                Dim tb As TextBox = DirectCast(c, TextBox)

                If tb.ReadOnly Then
                    tb.BackColor = panel
                    tb.ForeColor = faintText
                Else
                    tb.BackColor = bg
                    tb.ForeColor = textCol
                End If

                Exit Sub
            End If

            ' ---- RICHTEXTBOX ----
            If TypeOf c Is RichTextBox Then
                Dim rtb As RichTextBox = DirectCast(c, RichTextBox)
                rtb.BackColor = bg
                rtb.ForeColor = textCol
                Exit Sub
            End If

            ' ---- COMBOBOX ----
            If TypeOf c Is ComboBox Then
                Dim cb As ComboBox = DirectCast(c, ComboBox)
                cb.BackColor = bg
                cb.ForeColor = textCol
                Exit Sub
            End If

            ' ---- LISTBOX ----
            If TypeOf c Is ListBox Then
                Dim lb As ListBox = DirectCast(c, ListBox)
                lb.BackColor = bg
                lb.ForeColor = textCol
                Exit Sub
            End If

            ' ---- CHECKBOX / RADIO ----
            If TypeOf c Is CheckBox OrElse TypeOf c Is RadioButton Then
                c.BackColor = Color.Transparent
                c.ForeColor = textCol
                Exit Sub
            End If

            ' ---- BUTTON ----
            If TypeOf c Is Button Then
                Dim b As Button = DirectCast(c, Button)

                ' Keep special action colors (common in your calculators)
                If IsSpecialActionButtonColor(b.BackColor) Then
                    b.ForeColor = Color.White
                    Exit Sub
                End If

                b.FlatStyle = FlatStyle.Flat
                b.FlatAppearance.BorderSize = 1
                b.FlatAppearance.BorderColor = accent

                b.BackColor = If(isDark, Color.FromArgb(40, 45, 58), Color.White)
                b.ForeColor = textCol
                Exit Sub
            End If

            ' ---- TABCONTROL ----
            If TypeOf c Is TabControl Then
                c.BackColor = panel
                c.ForeColor = textCol
                Exit Sub
            End If

            ' ---- TABPAGE ----
            If TypeOf c Is TabPage Then
                c.BackColor = bg
                c.ForeColor = textCol
                Exit Sub
            End If

            ' ---- DATAGRIDVIEW ----
            If TypeOf c Is DataGridView Then
                Dim gv As DataGridView = DirectCast(c, DataGridView)
                gv.BackgroundColor = bg
                gv.GridColor = panel
                gv.DefaultCellStyle.BackColor = bg
                gv.DefaultCellStyle.ForeColor = textCol
                gv.ColumnHeadersDefaultCellStyle.BackColor = panel
                gv.ColumnHeadersDefaultCellStyle.ForeColor = textCol
                gv.EnableHeadersVisualStyles = False
                Exit Sub
            End If

            ' ---- DEFAULT FALLBACK ----
            ' Do not over-touch unknown controls; keep stable.
            c.ForeColor = textCol

        Catch
            ' never crash
        End Try
    End Sub

    Private Function IsSpecialActionButtonColor(c As Color) As Boolean
        ' common "action" colors used in your engineering tools
        If c = Color.LightGreen Then Return True
        If c = Color.Gold Then Return True
        If c = Color.OrangeRed Then Return True
        If c = Color.Red Then Return True
        If c = Color.Green Then Return True
        Return False
    End Function

End Module
