Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Globalization
Imports System.Net
Imports System.Threading
Imports System.Diagnostics

Public Class MainForm
    Inherits Form

    '========================================================
    ' APP
    '========================================================
    Private Const APP_NAME As String = "MDAT – Mechanical Design Automation Tool"
    Private Const CONFIG_FILE As String = "config.txt"

    Private Const EXPIRY_WARNING_DAYS As Integer = 30
    Private Const EXPIRY_CRITICAL_DAYS As Integer = 7

    ' Persisted timing stats (per slot)
    Private Const ACTION_TIMES_FILE As String = "action_times.txt"

    '========================================================
    ' TELEMETRY CONFIG (ADDITIVE)
    '========================================================
    Private syncUrl As String = ""
    Private telemetryToken As String = ""
    Private telemetryEnabled As Boolean = True ' default ON unless TELEMETRY=OFF

    '========================================================
    ' MACRO DELIVERY (R2 via Worker) - ADDITIVE
    '========================================================
    Private macroDeliveryUrl As String = ""
    Private macroToken As String = ""
    Private macroCacheDir As String = "" ' default computed
    Private hideMacroCache As Boolean = True

    ' Option C: auto-clean cache after each successful run
    Private autoCleanMacroCacheAfterSuccess As Boolean = True

    ' Cleanup retry (handles "file in use" from SolidWorks)
    Private cleanupTimer As System.Windows.Forms.Timer = Nothing
    Private cleanupRetryCount As Integer = 0
    Private Const CLEANUP_MAX_RETRIES As Integer = 30 ' ~60 sec if interval=2000ms

    '========================================================
    ' SELECTED ASSEMBLY
    '========================================================
    Private selectedAssemblyPath As String = String.Empty

    '========================================================
    ' SOLIDWORKS OBJECTS (late-binding)
    '========================================================
    Private swApp As Object = Nothing
    Private swModel As Object = Nothing

    '========================================================
    ' UI STATE
    '========================================================
    Private selectedButton As Button = Nothing

    '========================================================
    ' HEADER CONTROLS
    '========================================================
    Private pnlHeader As Panel
    Private pnlHeaderLine As Panel
    Private picLogo As PictureBox
    Private lblTitle As Label
    Private lblSubtitle As Label
    Private btnTheme As Button
    Private btnLicense As Button
    Private lblVersion As Label
    Private lblLicence As Label
    Private lblValidity As Label

    '========================================================
    ' FOOTER PANEL (premium status bar)
    '========================================================
    Private pnlFooter As Panel
    Private pnlFooterBorder As Panel
    Private lblFooterLeft As Label
    Private lblFooterRight As Label

    '========================================================
    ' LOG HEADER LABEL
    '========================================================
    Private lblLogHeader As Label

    '========================================================
    ' CENTRE LABELS (section headers)
    '========================================================
    Private lblSwVersion As Label
    Private lblAssembly As Label

    '========================================================
    ' CENTRE
    '========================================================
    Private cmbSW As ComboBox
    Private btnSelectFile As Button
    Private txtLog As TextBox

    '========================================================
    ' FOOTER
    '========================================================
    Private lblFooter As Label

    '========================================================
    ' SIDE PANELS
    '========================================================
    Private pnlDesign As Panel
    Private pnlEngineering As Panel
    Private pnlDesignContent As Panel
    Private pnlEngineeringContent As Panel

    '========================================================
    ' MACROS
    '========================================================
    Private macroDisplayMap As New Dictionary(Of Integer, String)()
    Private macroExecMap As New Dictionary(Of Integer, String)()

    '========================================================
    ' LICENSE STATE
    '========================================================
    Private activeLicense As LicenseInfo = Nothing
    Private currentTier As Integer = 0
    Private licenseValid As Boolean = False

    '========================================================
    ' THEME
    '========================================================
    Private Enum ThemeMode
        PANEL_LIGHT
        PANEL_DARK
        PANEL_MM
        PANEL_CUSTOM
    End Enum

    Private currentTheme As ThemeMode = ThemeMode.PANEL_LIGHT
    Private customAccent As Color = Color.FromArgb(0, 180, 255)

    Private BG_LIGHT As Color = Color.FromArgb(248, 249, 250)
    Private PANEL_LIGHT As Color = Color.FromArgb(255, 255, 255)

    Private BG_DARK As Color = Color.FromArgb(11, 30, 52)
    Private PANEL_DARK As Color = Color.FromArgb(15, 42, 68)

    Private BG_MM As Color = Color.FromArgb(242, 240, 248)
    Private PANEL_MM As Color = Color.FromArgb(225, 220, 240)

    Private currentThemeIsDark As Boolean = False

    '========================================================
    ' ACTION TIMING
    '========================================================
    Private Class ActionStat
        Public Count As Integer
        Public AvgSeconds As Double
    End Class

    Private actionStart As New Dictionary(Of Integer, DateTime)()
    Private actionStats As New Dictionary(Of Integer, ActionStat)()

    '========================================================
    ' INIT
    '========================================================
    Public Sub New()
        Try
            Me.Text = APP_NAME
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.Font = New Font("Segoe UI", 10)
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.MinimumSize = New Size(1100, 700)

            ' Set application icon (taskbar + title bar)
            Try
                Dim icoPath As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "mdat.ico")
                If IO.File.Exists(icoPath) Then
                    Me.Icon = New Icon(icoPath)
                End If
            Catch
            End Try

            Dim wa As Rectangle = Screen.PrimaryScreen.WorkingArea
            Me.Size = New Size(CInt(wa.Width * 0.92), CInt(wa.Height * 0.88))

            BuildHeader()
            BuildCentre()
            BuildSidePanels()
            BuildFooter()

            LoadMacrosFromConfig()
            LoadActionStats()

            ' ---- LICENSE LOAD ----
            Try
                activeLicense = Licensing.GetLicenseInfo()
                ResolveTierFromLicense(activeLicense)
            Catch
                licenseValid = False
            End Try

            ' ---- SEAT + TELEMETRY + MACRO DELIVERY CONFIG ----
            Try
                LoadSeatConfigFromConfig()
            Catch
            End Try

            ' Ensure cache folder exists early (best-effort)
            Try
                EnsureMacroCacheFolder()
            Catch
            End Try

            ' ---- TELEMETRY FLUSH ----
            Try
                If telemetryEnabled AndAlso syncUrl <> "" AndAlso telemetryToken <> "" Then
                    TelemetryQueue.Flush(syncUrl, telemetryToken, 8000)
                End If
            Catch
            End Try

            ' ---- LOAD SAVED THEME ----
            Try
                Dim savedTheme As String = UISettings.LoadThemeName()
                If savedTheme = "Dark" Then
                    currentTheme = ThemeMode.PANEL_DARK
                    UITheme.SetTheme(UITheme.ThemeKind.Dark)
                Else
                    currentTheme = ThemeMode.PANEL_LIGHT
                    UITheme.SetTheme(UITheme.ThemeKind.Light)
                End If
            Catch
            End Try

            BuildButtons()
            ApplyTierLocks()
            ApplyTheme()

            LogA("✅", "Application ready.")
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "MDAT Startup Failure", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Application.[Exit]()
        End Try
    End Sub

    '========================================================
    ' FORCE FIRST LAYOUT (OLD UI BEHAVIOUR)
    '========================================================
    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        HeaderResize(Nothing, EventArgs.Empty)
    End Sub

    '========================================================
    ' HEADER (OLD UI FROM MainForm - old.txt)
    '========================================================
    Private Sub BuildHeader()
        pnlHeader = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 110,
            .BackColor = Color.FromArgb(15, 23, 42)
        }
        Me.Controls.Add(pnlHeader)

        picLogo = New PictureBox() With {
            .Size = New Size(280, 75),
            .Location = New Point(8, 10),
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BackColor = Color.Transparent
        }
        pnlHeader.Controls.Add(picLogo)

        Dim logoPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets\logo\logo.png")
        Try
            If File.Exists(logoPath) Then
                Using fs As New FileStream(logoPath, FileMode.Open, FileAccess.Read)
                    Using bmp As New Bitmap(fs)
                        picLogo.Image = New Bitmap(bmp)
                    End Using
                End Using
            End If
        Catch
        End Try

        lblTitle = New Label() With {
            .Text = "MetaMech Mechanical Design Automation",
            .Font = New Font("Segoe UI Semibold", 20, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.Transparent,
            .AutoSize = True,
            .Location = New Point(picLogo.Right + 8, 18)
        }
        pnlHeader.Controls.Add(lblTitle)

        lblSubtitle = New Label() With {
            .Text = "Designed by MetaMech Solutions",
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .ForeColor = Color.FromArgb(169, 199, 232),
            .BackColor = Color.Transparent,
            .AutoSize = True,
            .Location = New Point(picLogo.Right + 8, 52)
        }
        pnlHeader.Controls.Add(lblSubtitle)

        lblVersion = New Label() With {
            .Text = "v1.0",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .ForeColor = Color.FromArgb(120, 140, 170),
            .BackColor = Color.Transparent,
            .AutoSize = True
        }
        pnlHeader.Controls.Add(lblVersion)

        btnTheme = New Button() With {
            .Text = "🌙 DARK",
            .Size = New Size(85, 26),
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .FlatStyle = FlatStyle.Flat,
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(15, 23, 42),
            .Cursor = Cursors.Hand
        }
        btnTheme.FlatAppearance.BorderSize = 1
        btnTheme.FlatAppearance.BorderColor = Color.FromArgb(60, 80, 110)
        btnTheme.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, 40, 60)
        AddHandler btnTheme.Click, AddressOf ShowThemeMenu
        pnlHeader.Controls.Add(btnTheme)

        btnLicense = New Button() With {
            .Text = "LICENSE",
            .Size = New Size(85, 26),
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .FlatStyle = FlatStyle.Flat,
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(15, 23, 42),
            .Cursor = Cursors.Hand
        }
        btnLicense.FlatAppearance.BorderSize = 1
        btnLicense.FlatAppearance.BorderColor = Color.FromArgb(60, 80, 110)
        btnLicense.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, 40, 60)
        AddHandler btnLicense.Click, AddressOf ShowLicensePopup
        pnlHeader.Controls.Add(btnLicense)

        lblLicence = New Label() With {
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.Transparent,
            .AutoSize = True
        }
        pnlHeader.Controls.Add(lblLicence)

        lblValidity = New Label() With {
            .Font = New Font("Segoe UI", 8),
            .ForeColor = Color.FromArgb(169, 199, 232),
            .BackColor = Color.Transparent,
            .AutoSize = True
        }
        pnlHeader.Controls.Add(lblValidity)

        AddHandler pnlHeader.Resize, AddressOf HeaderResize

        pnlHeaderLine = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 3,
            .BackColor = Color.FromArgb(0, 180, 255)
        }
        Me.Controls.Add(pnlHeaderLine)
    End Sub

    Private Sub HeaderResize(sender As Object, e As EventArgs)
        If pnlHeader Is Nothing Then Return

        Dim rightMargin As Integer = 20

        ' Row 1: lblLicence
        If lblLicence IsNot Nothing Then
            lblLicence.Location = New Point(pnlHeader.Width - lblLicence.Width - rightMargin, 10)
            lblLicence.ForeColor = Color.White
            lblLicence.BackColor = Color.Transparent
        End If

        ' Row 2: lblValidity
        If lblValidity IsNot Nothing Then
            lblValidity.Location = New Point(pnlHeader.Width - lblValidity.Width - rightMargin, 30)
            lblValidity.BackColor = Color.Transparent
        End If

        ' Row 3: lblVersion
        If lblVersion IsNot Nothing Then
            lblVersion.Location = New Point(pnlHeader.Width - lblVersion.Width - rightMargin, 50)
            lblVersion.ForeColor = Color.FromArgb(120, 140, 170)
            lblVersion.BackColor = Color.Transparent
        End If

        ' Row 4: buttons side by side, right-aligned
        If btnTheme IsNot Nothing Then
            btnTheme.Location = New Point(pnlHeader.Width - 100, 74)
        End If
        If btnLicense IsNot Nothing Then
            btnLicense.Location = New Point(pnlHeader.Width - 195, 74)
        End If
    End Sub

    '========================================================
    ' CENTRE (OLD UI FROM MainForm - old.txt)
    '========================================================
    Private Sub BuildCentre()
        lblSwVersion = New Label() With {
            .Text = "SOLIDWORKS VERSION",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .ForeColor = Color.FromArgb(120, 140, 170),
            .AutoSize = True,
            .Location = New Point(250, 120)
        }
        Me.Controls.Add(lblSwVersion)

        cmbSW = New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(250, 138)
        }
        cmbSW.Items.AddRange(New String() {"2022", "2023", "2024", "2025"})
        cmbSW.SelectedIndex = 0
        Me.Controls.Add(cmbSW)

        lblAssembly = New Label() With {
            .Text = "ASSEMBLY",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .ForeColor = Color.FromArgb(120, 140, 170),
            .AutoSize = True,
            .Location = New Point(400, 120)
        }
        Me.Controls.Add(lblAssembly)

        btnSelectFile = New Button() With {
            .Text = "Select Assembly (.SLDASM)",
            .Size = New Size(260, 38),
            .Location = New Point(400, 134),
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.FromArgb(0, 180, 255),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnSelectFile.FlatAppearance.BorderSize = 0
        AddHandler btnSelectFile.Click, AddressOf SelectAssembly
        Me.Controls.Add(btnSelectFile)

        lblLogHeader = New Label() With {
            .Text = "OUTPUT LOG",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .ForeColor = Color.FromArgb(150, 165, 185),
            .AutoSize = True,
            .Location = New Point(250, 180)
        }
        Me.Controls.Add(lblLogHeader)

        txtLog = New TextBox() With {
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Vertical,
            .BackColor = Color.FromArgb(240, 242, 245),
            .ForeColor = Color.FromArgb(20, 20, 20),
            .Font = New Font("Consolas", 9.5F),
            .BorderStyle = BorderStyle.None,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .Location = New Point(250, 198),
            .Size = New Size(Me.ClientSize.Width - 620, Me.ClientSize.Height - 270)
        }
        Me.Controls.Add(txtLog)
    End Sub

    '========================================================
    ' SIDE PANELS (OLD UI FROM MainForm - old.txt)
    '========================================================
    Private Sub BuildSidePanels()
        pnlDesign = New Panel() With {
            .Width = 210,
            .Location = New Point(20, 198),
            .Height = Me.ClientSize.Height - 270,
            .BorderStyle = BorderStyle.None,
            .Padding = New Padding(8),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left
        }
        Me.Controls.Add(pnlDesign)

        ' Content panel FIRST (dock fill must be added before dock top)
        pnlDesignContent = New Panel() With {.Dock = DockStyle.Fill, .AutoScroll = True}
        pnlDesign.Controls.Add(pnlDesignContent)

        ' Panel header with accent left border — added AFTER content (dock top renders on top)
        Dim pnlDesignHeader As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 40,
            .BackColor = Color.FromArgb(15, 23, 42)
        }
        pnlDesign.Controls.Add(pnlDesignHeader)

        Dim pnlDesignAccent As New Panel() With {
            .Dock = DockStyle.Left,
            .Width = 3,
            .BackColor = Color.FromArgb(0, 180, 255)
        }
        pnlDesignHeader.Controls.Add(pnlDesignAccent)

        ' 1px bottom accent border
        Dim pnlDesignBottomLine As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 1,
            .BackColor = Color.FromArgb(0, 180, 255)
        }
        pnlDesignHeader.Controls.Add(pnlDesignBottomLine)

        pnlDesignHeader.Controls.Add(New Label() With {
            .Text = "DESIGN TOOLS",
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.Transparent,
            .TextAlign = ContentAlignment.MiddleCenter
        })

        pnlEngineering = New Panel() With {
            .Width = 300,
            .Location = New Point(Me.ClientSize.Width - 320, 198),
            .Height = Me.ClientSize.Height - 270,
            .BorderStyle = BorderStyle.None,
            .Padding = New Padding(8),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Right
        }
        Me.Controls.Add(pnlEngineering)

        ' Content panel FIRST (dock fill must be added before dock top)
        pnlEngineeringContent = New Panel() With {.Dock = DockStyle.Fill, .AutoScroll = True}
        pnlEngineering.Controls.Add(pnlEngineeringContent)

        ' Panel header — added AFTER content (dock top renders on top)
        Dim pnlEngHeader As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 40,
            .BackColor = Color.FromArgb(15, 23, 42)
        }
        pnlEngineering.Controls.Add(pnlEngHeader)

        Dim pnlEngAccent As New Panel() With {
            .Dock = DockStyle.Left,
            .Width = 3,
            .BackColor = Color.FromArgb(0, 180, 255)
        }
        pnlEngHeader.Controls.Add(pnlEngAccent)

        ' 1px bottom accent border
        Dim pnlEngBottomLine As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 1,
            .BackColor = Color.FromArgb(0, 180, 255)
        }
        pnlEngHeader.Controls.Add(pnlEngBottomLine)

        pnlEngHeader.Controls.Add(New Label() With {
            .Text = "ENGINEERING TOOLS",
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.Transparent,
            .TextAlign = ContentAlignment.MiddleCenter
        })
    End Sub

    '========================================================
    ' FOOTER (OLD UI FROM MainForm - old.txt)
    '========================================================
    Private Sub BuildFooter()
        pnlFooter = New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 32
        }
        Me.Controls.Add(pnlFooter)

        pnlFooterBorder = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 1,
            .BackColor = Color.FromArgb(60, 0, 180, 255)
        }
        pnlFooter.Controls.Add(pnlFooterBorder)

        lblFooterLeft = New Label() With {
            .Text = "© 2026 MetaMech Solutions. All Rights Reserved.",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .AutoSize = True,
            .Location = New Point(10, 8)
        }
        pnlFooter.Controls.Add(lblFooterLeft)

        lblFooterRight = New Label() With {
            .Text = "MDAT v1.0 | © 2026 MetaMech Solutions",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .ForeColor = Color.FromArgb(120, 140, 170),
            .AutoSize = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        pnlFooter.Controls.Add(lblFooterRight)
        lblFooterRight.Location = New Point(pnlFooter.Width - lblFooterRight.Width - 15, 8)

        ' Website link
        Dim accentCol As Color = Color.FromArgb(0, 180, 255)
        Dim lnkSite As New LinkLabel() With {
            .Text = "metamechsolutions.com",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .LinkColor = accentCol,
            .ActiveLinkColor = accentCol,
            .VisitedLinkColor = accentCol,
            .BackColor = Color.Transparent,
            .AutoSize = True
        }
        AddHandler lnkSite.LinkClicked, Sub(ss As Object, ee As LinkLabelLinkClickedEventArgs)
                                             Try
                                                 Process.Start("https://metamechsolutions.com/")
                                             Catch
                                             End Try
                                         End Sub
        pnlFooter.Controls.Add(lnkSite)

        ' Blog link
        Dim lnkBlog As New LinkLabel() With {
            .Text = "Help & Blog",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .LinkColor = accentCol,
            .ActiveLinkColor = accentCol,
            .VisitedLinkColor = accentCol,
            .BackColor = Color.Transparent,
            .AutoSize = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        AddHandler lnkBlog.LinkClicked, Sub(ss As Object, ee As LinkLabelLinkClickedEventArgs)
                                             Try
                                                 Process.Start("https://metamechsolutions.com/blog/")
                                             Catch
                                             End Try
                                         End Sub
        pnlFooter.Controls.Add(lnkBlog)

        AddHandler pnlFooter.Resize, Sub(s As Object, ev As EventArgs)
                                          Try
                                              If lblFooterRight IsNot Nothing Then
                                                  lblFooterRight.Location = New Point(pnlFooter.Width - lblFooterRight.Width - 15, 8)
                                              End If
                                              lnkSite.Location = New Point((pnlFooter.Width - lnkSite.Width) \ 2, 8)
                                              lnkBlog.Location = New Point(pnlFooter.Width - lnkBlog.Width - lblFooterRight.Width - 30, 8)
                                          Catch
                                          End Try
                                      End Sub

        ' Keep old field for backward compat (hidden)
        lblFooter = lblFooterLeft
    End Sub

    '========================================================
    ' ASSEMBLY SELECTION
    '========================================================
    Private Sub SelectAssembly(sender As Object, e As EventArgs)
        Using ofd As New OpenFileDialog()
            ofd.Title = "Select Top-Level Assembly"
            ofd.Filter = "SolidWorks Assembly (*.sldasm)|*.sldasm"
            ofd.Multiselect = False

            If ofd.ShowDialog() <> DialogResult.OK Then Exit Sub

            selectedAssemblyPath = ofd.FileName
            LogA("📁", "Assembly selected:")
            LogA("🔗", selectedAssemblyPath)
        End Using
    End Sub

    Private Function IsAssemblySelected() As Boolean
        Return Not String.IsNullOrEmpty(selectedAssemblyPath)
    End Function

    '========================================================
    ' SOLIDWORKS BOOTSTRAP
    '========================================================
    Private Function EnsureSolidWorksRunning() As Boolean
        Dim year As String = "2022"
        Try
            If cmbSW IsNot Nothing AndAlso cmbSW.SelectedItem IsNot Nothing Then
                year = cmbSW.SelectedItem.ToString().Trim()
            End If
        Catch
        End Try

        Dim suffix As String = GetSWProgIDSuffix(year)
        Dim progId As String = "SldWorks.Application." & suffix

        If swApp IsNot Nothing Then
            If IsComObjectAlive(swApp) Then
                Return True
            Else
                swApp = Nothing
            End If
        End If

        Try
            swApp = Marshal.GetActiveObject(progId)
            If swApp IsNot Nothing Then
                LogA("🧷", "Attached to SolidWorks " & year & " (" & progId & ").")
                TrySetSwVisible(swApp, True)
                Return True
            End If
        Catch
            swApp = Nothing
        End Try

        Try
            swApp = CreateObject(progId)
            If swApp Is Nothing Then
                LogA("❌", "Failed to create SolidWorks instance: " & progId)
                Return False
            End If

            TrySetSwVisible(swApp, True)
            LogA("✅", "SolidWorks " & year & " started successfully (" & progId & ").")
            Return True

        Catch ex As Exception
            LogA("❌", "Failed to start SolidWorks " & year & ": " & ex.Message)
            swApp = Nothing
            Return False
        End Try
    End Function

    Private Function IsComObjectAlive(comObj As Object) As Boolean
        Try
            Dim v As Object = comObj.GetType().InvokeMember( _
                "Visible", _
                BindingFlags.GetProperty Or BindingFlags.Public Or BindingFlags.Instance, _
                Nothing, _
                comObj, _
                Nothing _
            )
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Sub TrySetSwVisible(appObj As Object, isVisible As Boolean)
        Try
            appObj.GetType().InvokeMember( _
                "Visible", _
                BindingFlags.SetProperty Or BindingFlags.Public Or BindingFlags.Instance, _
                Nothing, _
                appObj, _
                New Object() {isVisible} _
            )
        Catch
        End Try
    End Sub

    '========================================================
    ' OPEN OR REUSE ASSEMBLY (OpenDoc)
    '========================================================
    Private Function EnsureAssemblyOpen() As Boolean
        If swApp Is Nothing Then
            LogA("❌", "SolidWorks not available.")
            Return False
        End If

        If String.IsNullOrEmpty(selectedAssemblyPath) OrElse Not File.Exists(selectedAssemblyPath) Then
            LogA("❌", "Selected assembly path invalid.")
            Return False
        End If

        Try
            Dim existing As Object = Nothing
            Try
                existing = swApp.GetType().InvokeMember( _
                    "GetOpenDocumentByName", _
                    BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance, _
                    Nothing, _
                    swApp, _
                    New Object() {CStr(selectedAssemblyPath)} _
                )
            Catch
                existing = Nothing
            End Try

            If existing IsNot Nothing Then
                swModel = existing
                LogA("📌", "Assembly already open.")
            Else
                Dim opened As Object = Nothing
                Try
                    opened = swApp.GetType().InvokeMember( _
                        "OpenDoc", _
                        BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance, _
                        Nothing, _
                        swApp, _
                        New Object() {CStr(selectedAssemblyPath), CInt(2)} _
                    )
                Catch ex As Exception
                    If ex.InnerException IsNot Nothing Then
                        LogA("❌", "OpenDoc inner error: " & ex.InnerException.Message)
                    End If
                    LogA("❌", "OpenDoc error: " & ex.Message)
                    opened = Nothing
                End Try

                If opened Is Nothing Then
                    LogA("❌", "Failed to open assembly. (OpenDoc returned Nothing)")
                    swModel = Nothing
                    Return False
                End If

                swModel = opened
                LogA("✅", "Assembly opened successfully (OpenDoc).")
            End If

            Try
                swApp.GetType().InvokeMember( _
                    "ActivateDoc3", _
                    BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance, _
                    Nothing, _
                    swApp, _
                    New Object() {Path.GetFileName(selectedAssemblyPath), False, 0, 0} _
                )
            Catch
            End Try

            Try
                CallByName(swModel, "ResolveAllLightWeightComponents", CallType.Method, True)
                LogA("🧩", "Resolved lightweight components.")
            Catch
            End Try

            Try
                CallByName(swModel, "ForceRebuild3", CallType.Method, False)
            Catch
            End Try

            Return True

        Catch ex As Exception
            If ex.InnerException IsNot Nothing Then
                LogA("❌", "SolidWorks inner error: " & ex.InnerException.Message)
            End If
            LogA("❌", "SolidWorks error: " & ex.Message)
            swModel = Nothing
            Return False
        End Try
    End Function

    '========================================================
    ' RUN ACTION (macro) FROM CONFIG
    '========================================================
    Private Function RunMacroSlot(slot As Integer) As Boolean
        If Not macroExecMap.ContainsKey(slot) Then
            LogA("⚠️", "No tool configured for slot " & slot.ToString() & ".")
            Return False
        End If

        Dim execRaw As String = macroExecMap(slot)
        Dim parts() As String = execRaw.Split("|"c)

        If parts.Length < 3 Then
            LogA("❌", "Invalid tool config for slot " & slot.ToString() & ": " & execRaw)
            Return False
        End If

        Dim swpPath As String = parts(0).Trim()
        Dim moduleName As String = parts(1).Trim()
        Dim procName As String = parts(2).Trim()

        Dim toolName As String = ""
        If macroDisplayMap.ContainsKey(slot) Then
            toolName = macroDisplayMap(slot)
        Else
            toolName = Path.GetFileNameWithoutExtension(swpPath)
        End If

        ' ---- AUTO FETCH MACRO (R2 delivery) ----
        If (String.IsNullOrEmpty(swpPath) OrElse Not File.Exists(swpPath)) Then
            Dim cachedPath As String = ""
            If EnsureMacroAvailable(swpPath, cachedPath) Then
                swpPath = cachedPath
                macroExecMap(slot) = swpPath & "|" & moduleName & "|" & procName
            End If
        End If

        LogActionStart(slot, toolName, swpPath, moduleName & "." & procName)

        ' ---- TELEMETRY START ----
        Dim teleStartUtc As DateTime = DateTime.UtcNow
        Dim asmName As String = GetAssemblyNameSafe()
        SendTelemetrySafe("START", slot, toolName, 0, "ASM=" & asmName)

        If String.IsNullOrEmpty(swpPath) OrElse Not File.Exists(swpPath) Then
            LogA("❌", "Tool file not found: " & swpPath)
            LogActionEnd(slot, False, "Failed")
            SendTelemetrySafe("FAIL", slot, toolName, CInt((DateTime.UtcNow - teleStartUtc).TotalMilliseconds), "ASM=" & asmName & " | Missing tool file")
            Return False
        End If

        If swApp Is Nothing Then
            LogA("❌", "SolidWorks not available.")
            LogActionEnd(slot, False, "Failed")
            SendTelemetrySafe("FAIL", slot, toolName, CInt((DateTime.UtcNow - teleStartUtc).TotalMilliseconds), "ASM=" & asmName & " | SolidWorks not available")
            Return False
        End If

        Dim ok As Boolean = False
        Dim detail As String = ""

        ' Try RunMacro2 first
        ok = InvokeRunMacro2(swpPath, moduleName, procName, detail)
        If ok Then
            LogActionEnd(slot, True, "✅ Completed")
            SendTelemetrySafe("END", slot, toolName, CInt((DateTime.UtcNow - teleStartUtc).TotalMilliseconds), "ASM=" & asmName & " | Completed (RunMacro2)")
            LogA("🎉", "Successful — Thanks for using MetaMech Design Automation Tool")

            If autoCleanMacroCacheAfterSuccess Then StartMacroCacheCleanupRetrySilent()
            Return True
        Else
            If detail <> "" Then LogA("⚠️", "RunMacro2: " & detail)
        End If

        ' Fallback RunMacro
        detail = ""
        ok = InvokeRunMacro(swpPath, moduleName, procName, detail)
        If ok Then
            LogActionEnd(slot, True, "✅ Completed")
            SendTelemetrySafe("END", slot, toolName, CInt((DateTime.UtcNow - teleStartUtc).TotalMilliseconds), "ASM=" & asmName & " | Completed (fallback)")
            LogA("🎉", "Successful — Thanks for using MetaMech Design Automation Tool")

            If autoCleanMacroCacheAfterSuccess Then StartMacroCacheCleanupRetrySilent()
            Return True
        Else
            If detail <> "" Then LogA("⚠️", "RunMacro: " & detail)
        End If

        LogActionEnd(slot, False, "Failed")
        SendTelemetrySafe("FAIL", slot, toolName, CInt((DateTime.UtcNow - teleStartUtc).TotalMilliseconds), "ASM=" & asmName & " | Failed")
        Return False
    End Function

    '========================================================
    ' RunMacro2 / RunMacro
    '========================================================
    Private Function InvokeRunMacro2(macroPath As String, moduleName As String, procName As String, ByRef detail As String) As Boolean
        detail = ""
        Dim args5(4) As Object
        args5(0) = CStr(macroPath)
        args5(1) = CStr(moduleName)
        args5(2) = CStr(procName)
        args5(3) = CInt(0)
        args5(4) = CInt(0)

        Try
            Dim result As Object = swApp.GetType().InvokeMember( _
                "RunMacro2", _
                BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance, _
                Nothing, _
                swApp, _
                args5 _
            )
            Return MacroResultIsSuccess(result, detail)
        Catch ex As Exception
            detail = "Exception: " & ex.Message
            Return False
        End Try
    End Function

    Private Function InvokeRunMacro(macroPath As String, moduleName As String, procName As String, ByRef detail As String) As Boolean
        detail = ""
        Dim args3(2) As Object
        args3(0) = CStr(macroPath)
        args3(1) = CStr(moduleName)
        args3(2) = CStr(procName)

        Try
            Dim result As Object = swApp.GetType().InvokeMember( _
                "RunMacro", _
                BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance, _
                Nothing, _
                swApp, _
                args3 _
            )
            Return MacroResultIsSuccess(result, detail)
        Catch ex As Exception
            detail = "Exception: " & ex.Message
            Return False
        End Try
    End Function

    Private Function MacroResultIsSuccess(v As Object, ByRef detail As String) As Boolean
        detail = ""

        If v Is Nothing Then
            detail = "Returned Nothing"
            Return False
        End If

        Try
            If TypeOf v Is Integer Then
                Dim code As Integer = CInt(v)
                detail = "Code=" & code.ToString()
                Return (code = 0)
            End If
            If TypeOf v Is Long Then
                Dim code As Long = CLng(v)
                detail = "Code=" & code.ToString()
                Return (code = 0L)
            End If
            If TypeOf v Is Short Then
                Dim code As Short = CShort(v)
                detail = "Code=" & code.ToString()
                Return (code = 0S)
            End If

            If TypeOf v Is Boolean Then
                Dim b As Boolean = CBool(v)
                detail = "Bool=" & If(b, "True", "False")
                Return b
            End If

            Dim s As String = Convert.ToString(v)
            If s IsNot Nothing Then
                s = s.Trim()
                Dim n As Integer = 0
                If Integer.TryParse(s, n) Then
                    detail = "Code=" & n.ToString()
                    Return (n = 0)
                End If
                Dim bb As Boolean = False
                If Boolean.TryParse(s, bb) Then
                    detail = "Bool=" & If(bb, "True", "False")
                    Return bb
                End If
            End If

            detail = "Unknown return type: " & v.GetType().FullName
            Return False

        Catch ex As Exception
            detail = "Parse error: " & ex.Message
            Return False
        End Try
    End Function

    '========================================================
    ' OPTION C – CLEANUP WITH RETRIES (SILENT)
    '========================================================
    Private Sub StartMacroCacheCleanupRetrySilent()
        Try
            Dim ok As Boolean = TryCleanMacroCacheOnce()
            If ok Then
                Return
            End If

            cleanupRetryCount = 0
            If cleanupTimer Is Nothing Then
                cleanupTimer = New System.Windows.Forms.Timer()
                cleanupTimer.Interval = 2000
                AddHandler cleanupTimer.Tick, AddressOf CleanupTimerTickSilent
            End If

            If Not cleanupTimer.Enabled Then
                cleanupTimer.Start()
            End If
        Catch
        End Try
    End Sub

    Private Sub CleanupTimerTickSilent(sender As Object, e As EventArgs)
        Try
            cleanupRetryCount += 1

            Dim ok As Boolean = TryCleanMacroCacheOnce()
            If ok Then
                Try
                    cleanupTimer.Stop()
                Catch
                End Try
                Return
            End If

            If cleanupRetryCount >= CLEANUP_MAX_RETRIES Then
                Try
                    cleanupTimer.Stop()
                Catch
                End Try
            End If
        Catch
        End Try
    End Sub

    Private Function TryCleanMacroCacheOnce() As Boolean
        Try
            If Not EnsureMacroCacheFolder() Then Return False
            If macroCacheDir Is Nothing OrElse macroCacheDir.Trim() = "" Then Return False

            If Not Directory.Exists(macroCacheDir) Then
                Return True
            End If

            Try
                File.SetAttributes(macroCacheDir, FileAttributes.Normal)
            Catch
            End Try

            Dim anyLocked As Boolean = False

            Dim files() As String = Nothing
            Try
                files = Directory.GetFiles(macroCacheDir, "*", SearchOption.AllDirectories)
            Catch
                files = Nothing
            End Try

            If files IsNot Nothing Then
                For Each f As String In files
                    Dim deleted As Boolean = False
                    For i As Integer = 1 To 6
                        Try
                            File.SetAttributes(f, FileAttributes.Normal)
                        Catch
                        End Try

                        Try
                            File.Delete(f)
                            deleted = True
                            Exit For
                        Catch
                            Try
                                Thread.Sleep(150)
                            Catch
                            End Try
                        End Try
                    Next

                    If Not deleted Then
                        anyLocked = True
                    End If
                Next
            End If

            Try
                RemoveEmptyDirsSafe(macroCacheDir)
            Catch
            End Try

            Try
                If Directory.Exists(macroCacheDir) Then
                    Dim hasFiles As Boolean = (Directory.GetFiles(macroCacheDir, "*", SearchOption.AllDirectories).Length > 0)
                    Dim hasDirs As Boolean = (Directory.GetDirectories(macroCacheDir, "*", SearchOption.AllDirectories).Length > 0)
                    If (Not hasFiles) AndAlso (Not hasDirs) Then
                        Directory.Delete(macroCacheDir, False)
                    End If
                End If
            Catch
                anyLocked = True
            End Try

            If hideMacroCache AndAlso Directory.Exists(macroCacheDir) Then
                Try
                    Dim attr As FileAttributes = File.GetAttributes(macroCacheDir)
                    If (attr And FileAttributes.Hidden) = 0 Then
                        File.SetAttributes(macroCacheDir, attr Or FileAttributes.Hidden)
                    End If
                Catch
                End Try
            End If

            Return Not anyLocked

        Catch
            Return False
        End Try
    End Function

    Private Sub RemoveEmptyDirsSafe(root As String)
        Try
            Dim dirs() As String = Directory.GetDirectories(root, "*", SearchOption.AllDirectories)
            If dirs Is Nothing Then Exit Sub

            Array.Sort(dirs)
            Array.Reverse(dirs)

            For Each d As String In dirs
                Try
                    If Directory.Exists(d) Then
                        Dim hasFiles As Boolean = (Directory.GetFiles(d).Length > 0)
                        Dim hasDirs As Boolean = (Directory.GetDirectories(d).Length > 0)
                        If (Not hasFiles) AndAlso (Not hasDirs) Then
                            Directory.Delete(d, False)
                        End If
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub

    '========================================================
    ' ACTION TIMING HELPERS
    '========================================================
    Private Sub LogActionStart(slot As Integer, toolName As String, toolPath As String, entryPoint As String)
        actionStart(slot) = DateTime.Now

        Log("")
        Log("🧩 ACTION START — Slot " & slot.ToString())
        LogA("🧾", "Tool: " & toolName)
        LogA("📄", "Script: " & toolPath)
        LogA("▶", "Running: " & entryPoint)

        Dim est As String = GetEstimatedTimeString(slot)
        If est <> "" Then
            LogA("⏱️", "Estimated: ~" & est & " (based on previous runs)")
        Else
            LogA("⏱️", "Estimated: learning… (will appear after first run)")
        End If
    End Sub

    Private Sub LogActionEnd(slot As Integer, success As Boolean, note As String)
        Dim elapsedSec As Double = 0
        If actionStart.ContainsKey(slot) Then
            elapsedSec = (DateTime.Now - actionStart(slot)).TotalSeconds
            actionStart.Remove(slot)
        End If

        If elapsedSec > 0 Then
            UpdateActionStats(slot, elapsedSec)
        End If

        Dim elapsedStr As String = FormatDuration(elapsedSec)

        If success Then
            Log(note)
            Log("✅ ACTION END — Slot " & slot.ToString() & " (" & elapsedStr & " elapsed)")
        Else
            LogA("❌", note)
            Log("❌ ACTION END — Slot " & slot.ToString() & " (" & elapsedStr & " elapsed)")
        End If

        Log("")
    End Sub

    Private Function GetEstimatedTimeString(slot As Integer) As String
        If actionStats.ContainsKey(slot) Then
            Dim st As ActionStat = actionStats(slot)
            If st IsNot Nothing AndAlso st.Count > 0 AndAlso st.AvgSeconds > 0 Then
                Return FormatDuration(st.AvgSeconds)
            End If
        End If
        Return ""
    End Function

    Private Sub UpdateActionStats(slot As Integer, lastSeconds As Double)
        If lastSeconds <= 0 Then Exit Sub

        Dim st As ActionStat = Nothing
        If actionStats.ContainsKey(slot) Then
            st = actionStats(slot)
        Else
            st = New ActionStat()
            st.Count = 0
            st.AvgSeconds = 0
            actionStats(slot) = st
        End If

        st.Count += 1
        If st.Count = 1 Then
            st.AvgSeconds = lastSeconds
        Else
            st.AvgSeconds = ((st.AvgSeconds * CDbl(st.Count - 1)) + lastSeconds) / CDbl(st.Count)
        End If

        SaveActionStats()
    End Sub

    Private Function FormatDuration(seconds As Double) As String
        If seconds < 0 Then seconds = 0
        Dim total As Integer = CInt(Math.Floor(seconds))
        Dim mins As Integer = total \ 60
        Dim secs As Integer = total Mod 60
        Return mins.ToString("00") & ":" & secs.ToString("00")
    End Function

    Private Sub LoadActionStats()
        Try
            actionStats.Clear()

            Dim p As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ACTION_TIMES_FILE)
            If Not File.Exists(p) Then Exit Sub

            Dim lines() As String = File.ReadAllLines(p)
            For Each line As String In lines
                Dim t As String = line.Trim()
                If t = "" OrElse t.StartsWith("#") Then Continue For

                Dim parts() As String = t.Split("|"c)
                If parts.Length < 3 Then Continue For

                Dim slot As Integer = 0
                Dim cnt As Integer = 0
                Dim avg As Double = 0

                If Not Integer.TryParse(parts(0).Trim(), slot) Then Continue For
                If Not Integer.TryParse(parts(1).Trim(), cnt) Then Continue For
                If Not Double.TryParse(parts(2).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, avg) Then Continue For

                Dim st As New ActionStat()
                st.Count = Math.Max(0, cnt)
                st.AvgSeconds = Math.Max(0, avg)
                actionStats(slot) = st
            Next
        Catch
        End Try
    End Sub

    Private Sub SaveActionStats()
        Try
            Dim p As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ACTION_TIMES_FILE)

            Using sw As New StreamWriter(p, False)
                sw.WriteLine("# slot|count|avg_seconds")
                For Each kv As KeyValuePair(Of Integer, ActionStat) In actionStats
                    Dim slot As Integer = kv.Key
                    Dim st As ActionStat = kv.Value
                    If st Is Nothing Then Continue For
                    sw.WriteLine(slot.ToString() & "|" & st.Count.ToString() & "|" & st.AvgSeconds.ToString(CultureInfo.InvariantCulture))
                Next
            End Using
        Catch
        End Try
    End Sub

    '========================================================
    ' SolidWorks ProgID suffix mapping
    '========================================================
    Private Function GetSWProgIDSuffix(selectedYear As String) As String
        Select Case selectedYear
            Case "2022" : Return "30"
            Case "2023" : Return "31"
            Case "2024" : Return "32"
            Case "2025" : Return "33"
        End Select
        Return "30"
    End Function

    '========================================================
    ' BUTTON STYLE – LOCKED
    '========================================================
    Private Sub StyleDisabledButton(btn As Button)
        btn.Enabled = False
        btn.BackColor = Color.FromArgb(230, 230, 230)
        btn.ForeColor = Color.FromArgb(140, 140, 140)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200)
        btn.FlatAppearance.BorderSize = 1
        btn.Cursor = Cursors.[Default]

        If Not btn.Text.StartsWith("🔒") Then
            btn.Text = "🔒 " & btn.Text
        End If

        Dim tip As New ToolTip()
        tip.SetToolTip(btn, "Locked — upgrade license to use this tool")
    End Sub

    '========================================================
    ' BUTTON BUILDER
    '========================================================
    Private Sub AddButton(parent As Panel, title As String, slot As Integer, y As Integer, clickHandler As EventHandler)
        Dim b As New Button()
        b.Text = title
        b.Tag = slot
        b.Name = "btnTool_" & slot.ToString()

        b.Location = New Point(8, y)
        b.Size = New Size(parent.Width - 16, 42)

        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderSize = 1
        b.FlatAppearance.BorderColor = If(currentThemeIsDark, Color.FromArgb(30, 50, 75), Color.FromArgb(200, 210, 220))
        b.BackColor = If(currentThemeIsDark, Color.FromArgb(22, 44, 69), Color.FromArgb(232, 236, 240))
        b.ForeColor = If(currentThemeIsDark, Color.FromArgb(234, 244, 255), Color.FromArgb(26, 26, 46))
        b.Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
        b.Cursor = Cursors.Hand

        AddHandler b.Click, clickHandler
        AddHandler b.MouseEnter, AddressOf ToolButton_MouseEnter
        AddHandler b.MouseLeave, AddressOf ToolButton_MouseLeave
        parent.Controls.Add(b)
    End Sub

    Private Sub ToolButton_MouseEnter(sender As Object, e As EventArgs)
        Dim b As Button = TryCast(sender, Button)
        If b Is Nothing Then Return
        If b Is selectedButton Then Return
        If Not b.Enabled Then Return
        If currentThemeIsDark Then
            b.BackColor = Color.FromArgb(30, 58, 95)
        Else
            b.BackColor = Color.FromArgb(208, 228, 244)
        End If
    End Sub

    Private Sub ToolButton_MouseLeave(sender As Object, e As EventArgs)
        Dim b As Button = TryCast(sender, Button)
        If b Is Nothing Then Return
        If b Is selectedButton Then Return
        If Not b.Enabled Then Return
        If currentThemeIsDark Then
            b.BackColor = Color.FromArgb(22, 44, 69)
        Else
            b.BackColor = Color.FromArgb(232, 236, 240)
        End If
    End Sub

    '========================================================
    ' BUILD BUTTONS
    '========================================================
    Private Sub BuildButtons()
        pnlDesignContent.Controls.Clear()
        pnlEngineeringContent.Controls.Clear()

        Dim y As Integer = 10
        For i As Integer = 1 To 10
            If macroDisplayMap.ContainsKey(i) Then
                AddButton(pnlDesignContent, macroDisplayMap(i), i, y, AddressOf RunDesignTool)
                y += 48
            End If
        Next

        y = 10
        AddButton(pnlEngineeringContent, "CONVEYOR CALCULATOR", 11, y, AddressOf RunEngineeringTool)
        y += 48
        AddButton(pnlEngineeringContent, "PNEUMATIC CALCULATOR", 12, y, AddressOf RunEngineeringTool)
        y += 48
        AddButton(pnlEngineeringContent, "AIR CONSUMPTION & COMPRESSOR SIZING", 13, y, AddressOf RunEngineeringTool)
        y += 48
        AddButton(pnlEngineeringContent, "BEAM DEFLECTION & FRAME CHECK", 14, y, AddressOf RunEngineeringTool)
    End Sub

    '========================================================
    ' MACRO LOADER (supports [MACROS] section)
    '========================================================
    Private Sub LoadMacrosFromConfig()
        macroDisplayMap.Clear()
        macroExecMap.Clear()

        Dim cfgPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, CONFIG_FILE)
        If Not File.Exists(cfgPath) Then Return

        Dim inMacrosSection As Boolean = False

        For Each rawLine As String In File.ReadAllLines(cfgPath)
            Dim line As String = rawLine.Trim()
            If line = "" OrElse line.StartsWith("#") Then Continue For

            If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                inMacrosSection = (line.Trim().ToUpperInvariant() = "[MACROS]")
                Continue For
            End If

            If HasMacrosSection(cfgPath) AndAlso Not inMacrosSection Then
                Continue For
            End If

            If Not line.Contains("=") Then Continue For

            Dim parts() As String = line.Split(New Char() {"="c}, 2)
            Dim slot As Integer
            If Not Integer.TryParse(parts(0).Trim(), slot) Then Continue For

            Dim execRaw As String = parts(1).Trim()

            Dim swpPath As String = execRaw
            If swpPath.Contains("|") Then swpPath = swpPath.Split("|"c)(0).Trim()
            If Not swpPath.ToLowerInvariant().EndsWith(".swp") Then Continue For

            macroExecMap(slot) = execRaw
            macroDisplayMap(slot) = Path.GetFileNameWithoutExtension(swpPath)
        Next
    End Sub

    Private Function HasMacrosSection(cfgPath As String) As Boolean
        Try
            Dim lines() As String = File.ReadAllLines(cfgPath)
            For Each l As String In lines
                If l IsNot Nothing Then
                    Dim t As String = l.Trim().ToUpperInvariant()
                    If t = "[MACROS]" Then Return True
                End If
            Next
        Catch
        End Try
        Return False
    End Function

    '========================================================
    ' TIER LOCKS
    '========================================================
    Private Sub ApplyTierLocks()
        For Each c As Control In pnlDesignContent.Controls
            Dim b As Button = TryCast(c, Button)
            If b Is Nothing Then Continue For
            If Not licenseValid OrElse Not TierLocks.CanRunDesignTool(CInt(b.Tag), currentTier) Then
                StyleDisabledButton(b)
            End If
        Next

        For Each c As Control In pnlEngineeringContent.Controls
            Dim b As Button = TryCast(c, Button)
            If b Is Nothing Then Continue For
            If Not licenseValid OrElse Not TierLocks.CanRunEngineeringTool(CInt(b.Tag), currentTier) Then
                StyleDisabledButton(b)
            End If
        Next
    End Sub

    '========================================================
    ' DESIGN TOOLS
    '========================================================
    Private Sub RunDesignTool(sender As Object, e As EventArgs)
        If Not IsAssemblySelected() Then
            LogA("⚠️", "Please select a Top-Level Assembly before running design tools.")
            Return
        End If

        If Not EnsureSeatForAction() Then
            Return
        End If

        Dim slot As Integer = 0
        Dim b As Button = TryCast(sender, Button)
        Try
            If b IsNot Nothing AndAlso b.Tag IsNot Nothing Then slot = CInt(b.Tag)
        Catch
            slot = 0
        End Try

        If slot <= 0 Then
            LogA("❌", "Invalid tool slot.")
            Return
        End If

        If Not EnsureSolidWorksRunning() Then Return
        If Not EnsureAssemblyOpen() Then Return

        RunMacroSlot(slot)
    End Sub

    '========================================================
    ' ENGINEERING TOOLS
    '========================================================
    Private Sub RunEngineeringTool(sender As Object, e As EventArgs)
        Dim slot As Integer = 0
        Dim b As Button = TryCast(sender, Button)

        Try
            If b IsNot Nothing AndAlso b.Tag IsNot Nothing Then
                slot = CInt(b.Tag)
            End If
        Catch
            slot = 0
        End Try

        If slot <= 0 Then
            LogA("❌", "Invalid engineering tool slot.")
            Return
        End If

        Dim frm As Form = Nothing
        Dim toolName As String = ""

        ' FIX: Proper Try/Catch/End Try (your build failed due to missing End Try)
        Try
            If slot = 11 Then
                toolName = "CONVEYOR CONFIGURATOR"
                frm = New ConveyorCalculatorForm()

            ElseIf slot = 12 Then
                toolName = "PNEUMATIC CYLINDER CALCULATOR"
                frm = New PneumaticCylinderCalculatorForm()

            ElseIf slot = 13 Then
                toolName = "AIR CONSUMPTION & COMPRESSOR SIZING"
                frm = New AirConsumptionForm()

            ElseIf slot = 14 Then
                toolName = "BEAM DEFLECTION & FRAME CHECK"
                MessageBox.Show("This tool is reserved (coming soon).", "Engineering Tool", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return

            Else
                LogA("⚠️", "Engineering tool not wired for slot " & slot.ToString() & ".")
                Return
            End If

        Catch ex As Exception
            LogA("❌", "Failed to create tool window: " & ex.Message)
            Return
        End Try

        If frm Is Nothing Then
            LogA("❌", "Tool window is Nothing.")
            Return
        End If

        Try
            ApplyThemeToToolForm(frm)
        Catch
        End Try

        Try
            Log("")
            Log("🧩 ACTION START — " & toolName)
            frm.Show(Me)
            frm.BringToFront()
            Log("✅ ACTION END — " & toolName & " (opened)")
            Log("")
        Catch ex As Exception
            LogA("❌", "Failed to show tool window: " & ex.Message)
        End Try
    End Sub

    Private Sub ApplyThemeToToolForm(frm As Form)
        If frm Is Nothing Then Exit Sub

        Dim bg As Color = BG_LIGHT
        Dim panel As Color = PANEL_LIGHT
        Dim isDark As Boolean = False

        Select Case currentTheme
            Case ThemeMode.PANEL_DARK
                bg = BG_DARK
                panel = PANEL_DARK
                isDark = True
            Case ThemeMode.PANEL_MM
                bg = BG_MM
                panel = PANEL_MM
                isDark = False
            Case ThemeMode.PANEL_CUSTOM
                bg = BG_LIGHT
                panel = PANEL_LIGHT
                isDark = False
            Case Else
                bg = BG_LIGHT
                panel = PANEL_LIGHT
                isDark = False
        End Select

        Try
            ThemeApplier.ApplyTheme(frm, bg, panel, customAccent, isDark)
        Catch
        End Try
    End Sub

    '========================================================
    ' LICENSE
    '========================================================
    Private Sub ResolveTierFromLicense(lic As LicenseInfo)
        licenseValid = False

        If lic Is Nothing OrElse Not lic.IsValid Then
            lblLicence.Text = "LICENCE: INVALID"
            lblLicence.ForeColor = Color.Red

            lblValidity.Text = "Please activate a valid license."
            lblValidity.ForeColor = Color.Red
            Return
        End If

        licenseValid = True
        currentTier = lic.Tier

        lblLicence.ForeColor = If(currentThemeIsDark, Color.White, Color.Black)
        lblLicence.Text = "LICENCE: ACTIVE"

        Dim tierName As String = GetTierNameSafe(lic.Tier)

        Dim seatsVal As Integer = -1
        Try
            seatsVal = lic.Seats
        Catch
            seatsVal = -1
        End Try

        Dim seatsShort As String
        If seatsVal >= 0 Then
            seatsShort = "S:" & seatsVal.ToString()
        Else
            seatsShort = "S:N/A"
        End If

        If lic.ExpiryUtc <= DateTime.UtcNow Then
            lblValidity.Text = tierName & " | " & seatsShort & " | ⛔ EXPIRED"
            lblValidity.ForeColor = Color.Red
            Return
        End If

        Dim daysLeft As Integer = CInt(Math.Floor((lic.ExpiryUtc - DateTime.UtcNow).TotalDays))
        If daysLeft < 0 Then daysLeft = 0

        If lic.Tier = 0 Then
            If daysLeft <= EXPIRY_CRITICAL_DAYS Then
                lblValidity.Text = "TRIAL | " & seatsShort & " | ⛔ ENDING IN " & daysLeft.ToString() & "D"
                lblValidity.ForeColor = Color.Red
            ElseIf daysLeft <= EXPIRY_WARNING_DAYS Then
                lblValidity.Text = "TRIAL | " & seatsShort & " | ⚠ ENDING IN " & daysLeft.ToString() & "D"
                lblValidity.ForeColor = Color.DarkOrange
            Else
                lblValidity.Text = "TRIAL | " & seatsShort & " | " & daysLeft.ToString() & "D LEFT"
                lblValidity.ForeColor = Color.Green
            End If
        Else
            If daysLeft <= EXPIRY_CRITICAL_DAYS Then
                lblValidity.Text = tierName & " | " & seatsShort & " | ⛔ EXPIRING IN " & daysLeft.ToString() & "D"
                lblValidity.ForeColor = Color.Red
            ElseIf daysLeft <= EXPIRY_WARNING_DAYS Then
                lblValidity.Text = tierName & " | " & seatsShort & " | ⚠ EXPIRING IN " & daysLeft.ToString() & "D"
                lblValidity.ForeColor = Color.DarkOrange
            Else
                lblValidity.Text = tierName & " | " & seatsShort & " | " & daysLeft.ToString() & "D LEFT"
                lblValidity.ForeColor = Color.Green
            End If
        End If
    End Sub

    Private Function GetTierNameSafe(tier As Integer) As String
        Select Case tier
            Case 0 : Return "TRIAL"
            Case 1 : Return "STANDARD"
            Case 2 : Return "PREMIUM"
            Case 3 : Return "PREMIUM PLUS"
            Case Else : Return "UNKNOWN"
        End Select
    End Function

    Private Sub ShowLicensePopup(sender As Object, e As EventArgs)
        MessageBox.Show(lblLicence.Text & vbCrLf & lblValidity.Text, "License Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    '========================================================
    ' THEME
    '========================================================
    Private Sub ShowThemeMenu(sender As Object, e As EventArgs)
        ' Cycle between Light and Dark
        If currentTheme = ThemeMode.PANEL_DARK Then
            SetTheme(ThemeMode.PANEL_LIGHT)
        Else
            SetTheme(ThemeMode.PANEL_DARK)
        End If
    End Sub

    Private Sub PickCustomTheme()
        Using d As New ColorDialog()
            If d.ShowDialog() = DialogResult.OK Then
                customAccent = d.Color
                SetTheme(ThemeMode.PANEL_CUSTOM)
            End If
        End Using
    End Sub

    Private Sub SetTheme(m As ThemeMode)
        currentTheme = m

        ' Sync UITheme class state
        If m = ThemeMode.PANEL_DARK Then
            UITheme.SetTheme(UITheme.ThemeKind.Dark)
        Else
            UITheme.SetTheme(UITheme.ThemeKind.Light)
        End If

        ApplyTheme()

        ' Apply to all open child forms
        Try
            ThemeApplier.ApplyToAllOwnedForms(Me)
        Catch
        End Try

        ' Persist preference
        Try
            UISettings.SaveThemeName(If(m = ThemeMode.PANEL_DARK, "Dark", "Light"))
        Catch
        End Try
    End Sub

    Private Sub ApplyTheme()
        Dim bg As Color = BG_LIGHT
        Dim panel As Color = PANEL_LIGHT

        Select Case currentTheme
            Case ThemeMode.PANEL_LIGHT
                bg = BG_LIGHT
                panel = PANEL_LIGHT
                currentThemeIsDark = False

            Case ThemeMode.PANEL_DARK
                bg = BG_DARK
                panel = PANEL_DARK
                currentThemeIsDark = True

            Case ThemeMode.PANEL_MM
                bg = BG_MM
                panel = PANEL_MM
                currentThemeIsDark = False

            Case ThemeMode.PANEL_CUSTOM
                bg = BG_LIGHT
                panel = PANEL_LIGHT
                currentThemeIsDark = False
        End Select

        ' Apply theme recursively to the whole form
        Try
            ThemeApplier.ApplyTheme(Me, bg, panel, customAccent, currentThemeIsDark)
        Catch
        End Try

        ' Override header — ALWAYS dark regardless of theme
        Try
            pnlHeader.BackColor = Color.FromArgb(15, 23, 42)
            pnlHeaderLine.BackColor = Color.FromArgb(0, 180, 255)
            lblTitle.ForeColor = Color.White
            lblTitle.BackColor = Color.Transparent
            lblSubtitle.ForeColor = Color.FromArgb(169, 199, 232)
            lblSubtitle.BackColor = Color.Transparent
            If lblVersion IsNot Nothing Then
                lblVersion.ForeColor = Color.FromArgb(120, 140, 170)
                lblVersion.BackColor = Color.Transparent
            End If

            ' Header buttons always white on dark
            btnTheme.ForeColor = Color.White
            btnTheme.BackColor = Color.FromArgb(15, 23, 42)
            btnTheme.FlatAppearance.BorderSize = 1
            btnTheme.FlatAppearance.BorderColor = Color.FromArgb(60, 80, 110)
            btnLicense.ForeColor = Color.White
            btnLicense.BackColor = Color.FromArgb(15, 23, 42)
            btnLicense.FlatAppearance.BorderSize = 1
            btnLicense.FlatAppearance.BorderColor = Color.FromArgb(60, 80, 110)

            lblLicence.ForeColor = Color.White
            lblLicence.BackColor = Color.Transparent
            lblValidity.BackColor = Color.Transparent

            ' Update theme button text
            If btnTheme IsNot Nothing Then
                btnTheme.Text = If(currentThemeIsDark, "☀ LIGHT", "🌙 DARK")
            End If
        Catch
        End Try

        ' Override log — theme-aware
        Try
            If currentThemeIsDark Then
                txtLog.BackColor = Color.FromArgb(30, 38, 50)
                txtLog.ForeColor = Color.FromArgb(220, 225, 230)
            Else
                txtLog.BackColor = Color.FromArgb(240, 242, 245)
                txtLog.ForeColor = Color.FromArgb(20, 20, 20)
            End If
            If lblLogHeader IsNot Nothing Then
                lblLogHeader.ForeColor = Color.FromArgb(150, 165, 185)
            End If
        Catch
        End Try

        ' Override side panel backgrounds
        Try
            pnlDesign.BackColor = panel
            pnlEngineering.BackColor = panel
        Catch
        End Try

        ' Override footer
        Try
            If pnlFooter IsNot Nothing Then
                pnlFooter.BackColor = If(currentThemeIsDark, Color.FromArgb(10, 25, 41), Color.FromArgb(240, 242, 245))
                If lblFooterLeft IsNot Nothing Then
                    lblFooterLeft.ForeColor = If(currentThemeIsDark, Color.FromArgb(169, 199, 232), Color.FromArgb(80, 90, 100))
                End If
                If lblFooterRight IsNot Nothing Then
                    lblFooterRight.ForeColor = Color.FromArgb(120, 140, 170)
                End If
            End If
        Catch
        End Try

        ' Override button colors for premium look
        Try
            Dim btnBorderCol As Color = If(currentThemeIsDark, Color.FromArgb(30, 50, 75), Color.FromArgb(200, 210, 220))
            For Each c As Control In pnlDesignContent.Controls
                Dim b As Button = TryCast(c, Button)
                If b Is Nothing OrElse Not b.Enabled Then Continue For
                If b Is selectedButton Then
                    b.BackColor = Color.FromArgb(0, 180, 255)
                    b.ForeColor = Color.White
                Else
                    b.BackColor = If(currentThemeIsDark, Color.FromArgb(22, 44, 69), Color.FromArgb(232, 236, 240))
                    b.ForeColor = If(currentThemeIsDark, Color.FromArgb(234, 244, 255), Color.FromArgb(26, 26, 46))
                End If
                b.FlatStyle = FlatStyle.Flat
                b.FlatAppearance.BorderSize = 1
                b.FlatAppearance.BorderColor = btnBorderCol
            Next

            For Each c As Control In pnlEngineeringContent.Controls
                Dim b As Button = TryCast(c, Button)
                If b Is Nothing OrElse Not b.Enabled Then Continue For
                If b Is selectedButton Then
                    b.BackColor = Color.FromArgb(0, 180, 255)
                    b.ForeColor = Color.White
                Else
                    b.BackColor = If(currentThemeIsDark, Color.FromArgb(22, 44, 69), Color.FromArgb(232, 236, 240))
                    b.ForeColor = If(currentThemeIsDark, Color.FromArgb(234, 244, 255), Color.FromArgb(26, 26, 46))
                End If
                b.FlatStyle = FlatStyle.Flat
                b.FlatAppearance.BorderSize = 1
                b.FlatAppearance.BorderColor = btnBorderCol
            Next
        Catch
        End Try

        ' Select assembly button always accent
        Try
            If btnSelectFile IsNot Nothing Then
                btnSelectFile.BackColor = Color.FromArgb(0, 180, 255)
                btnSelectFile.ForeColor = Color.White
                btnSelectFile.FlatStyle = FlatStyle.Flat
                btnSelectFile.FlatAppearance.BorderSize = 0
            End If
        Catch
        End Try

        ' Section labels always muted
        Try
            If lblSwVersion IsNot Nothing Then lblSwVersion.ForeColor = Color.FromArgb(120, 140, 170)
            If lblAssembly IsNot Nothing Then lblAssembly.ForeColor = Color.FromArgb(120, 140, 170)
        Catch
        End Try

        Try
            If licenseValid Then
                lblLicence.ForeColor = Color.White
            End If
        Catch
        End Try

        ' Re-apply license validity colors (ThemeApplier may have overwritten them)
        Try
            If activeLicense IsNot Nothing Then
                ResolveTierFromLicense(activeLicense)
            End If
        Catch
        End Try
    End Sub

    '========================================================
    ' HYBRID SEAT ENFORCEMENT
    '========================================================
    Private Function EnsureSeatForAction() As Boolean
        If Not licenseValid OrElse activeLicense Is Nothing OrElse Not activeLicense.IsValid Then
            LogA("⛔", "License invalid. Please activate a valid license.")
            Return False
        End If

        Try
            If activeLicense.ExpiryUtc <= DateTime.UtcNow Then
                LogA("⛔", "License expired. Please renew your license.")
                Return False
            End If
        Catch
            LogA("⛔", "License expiry check failed. Please re-activate your license.")
            Return False
        End Try

        Dim licenseId As String = GetLicenseIdSafe(activeLicense)
        If licenseId Is Nothing OrElse licenseId.Trim().Length = 0 Then
            LogA("❌", "Seat check failed: License ID not found in license payload.")
            Return False
        End If

        Dim tier As Integer = 0
        Dim seatsMax As Integer = 1

        Try
            tier = activeLicense.Tier
        Catch
            tier = currentTier
        End Try

        Try
            seatsMax = activeLicense.Seats
            If seatsMax <= 0 Then seatsMax = 1
        Catch
            seatsMax = 1
        End Try

        Try
            SeatEnforcer.EnsureSeatOrThrow(licenseId.Trim(), tier, seatsMax)
            Return True
        Catch ex As Exception
            LogA("⛔", "Seat enforcement blocked this action.")
            LogA("🧾", ex.Message)
            MessageBox.Show(ex.Message, "MDAT Seat Enforcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End Try
    End Function

    Private Function GetLicenseIdSafe(lic As LicenseInfo) As String
        If lic Is Nothing Then Return ""

        Try
            Dim t As Type = lic.GetType()

            Dim names() As String = New String() {"LicenseId", "LICENSEID", "LicenseID", "licenseId", "Id", "ID"}
            For i As Integer = 0 To names.Length - 1
                Dim p As PropertyInfo = t.GetProperty(names(i), BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.IgnoreCase)
                If p IsNot Nothing Then
                    Dim v As Object = p.GetValue(lic, Nothing)
                    If v IsNot Nothing Then
                        Dim s As String = Convert.ToString(v)
                        If s IsNot Nothing Then
                            s = s.Trim()
                            If s.Length > 0 Then Return s
                        End If
                    End If
                End If
            Next

            For i As Integer = 0 To names.Length - 1
                Dim f As FieldInfo = t.GetField(names(i), BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.IgnoreCase)
                If f IsNot Nothing Then
                    Dim v As Object = f.GetValue(lic)
                    If v IsNot Nothing Then
                        Dim s As String = Convert.ToString(v)
                        If s IsNot Nothing Then
                            s = s.Trim()
                            If s.Length > 0 Then Return s
                        End If
                    End If
                End If
            Next
        Catch
        End Try

        Return ""
    End Function

    '========================================================
    ' CONFIG LOAD (seat + telemetry + macro delivery)
    '========================================================
    Private Sub LoadSeatConfigFromConfig()
        Try
            Dim cfgPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, CONFIG_FILE)
            If Not File.Exists(cfgPath) Then Exit Sub

            Dim seatServer As String = ""
            Dim seatToken As String = ""

            Dim localSyncUrl As String = ""
            Dim localMacroUrl As String = ""
            Dim localMacroToken As String = ""
            Dim localMacroCacheDir As String = ""
            Dim localHideCache As String = ""
            Dim localAutoClean As String = ""
            Dim localTelemetry As String = ""

            For Each rawLine As String In File.ReadAllLines(cfgPath)
                Dim line As String = rawLine.Trim()
                If line = "" OrElse line.StartsWith("#") OrElse line.StartsWith("[") Then Continue For

                Dim eq As Integer = line.IndexOf("="c)
                If eq <= 0 Then Continue For

                Dim k As String = line.Substring(0, eq).Trim().ToUpperInvariant()
                k = k.Replace(ChrW(&HFEFF), "")
                Dim v As String = line.Substring(eq + 1).Trim()

                If k = "SEAT_SERVER" Then
                    seatServer = v
                ElseIf k = "SEAT_TOKEN" Then
                    seatToken = v
                ElseIf k = "CLIENT_TOKEN" Then
                    If seatToken = "" Then seatToken = v
                ElseIf k = "SYNC_URL" Then
                    localSyncUrl = v
                ElseIf k = "MACRO_DELIVERY_URL" Then
                    localMacroUrl = v
                ElseIf k = "MACRO_TOKEN" Then
                    localMacroToken = v
                ElseIf k = "MACRO_CACHE_DIR" Then
                    localMacroCacheDir = v
                ElseIf k = "HIDE_MACRO_CACHE" Then
                    localHideCache = v
                ElseIf k = "AUTO_CLEAN_MACRO_CACHE" Then
                    localAutoClean = v
                ElseIf k = "TELEMETRY" Then
                    localTelemetry = v
                End If
            Next

            If localTelemetry <> "" Then
                Dim t As String = localTelemetry.Trim().ToUpperInvariant()
                telemetryEnabled = Not (t = "0" OrElse t = "FALSE" OrElse t = "NO" OrElse t = "OFF")
            End If

            If seatServer <> "" Then
                Try
                    SeatServerClient.ServerBaseUrl = seatServer
                    LogA("🔗", "Seat server loaded from config.")
                Catch
                End Try
            End If

            If seatToken <> "" Then
                Try
                    SeatServerClient.ClientToken = seatToken
                    LogA("🔐", "Client token loaded from config.")
                Catch
                End Try
                telemetryToken = seatToken
            End If

            If localSyncUrl <> "" Then syncUrl = localSyncUrl

            If localMacroUrl <> "" Then macroDeliveryUrl = localMacroUrl.Trim()
            If localMacroToken <> "" Then macroToken = localMacroToken.Trim()
            If localMacroCacheDir <> "" Then macroCacheDir = localMacroCacheDir.Trim()

            If localHideCache <> "" Then
                Dim t As String = localHideCache.Trim().ToUpperInvariant()
                hideMacroCache = (t = "1" OrElse t = "TRUE" OrElse t = "YES" OrElse t = "ON")
            End If

            If localAutoClean <> "" Then
                Dim t As String = localAutoClean.Trim().ToUpperInvariant()
                autoCleanMacroCacheAfterSuccess = (t = "1" OrElse t = "TRUE" OrElse t = "YES" OrElse t = "ON")
            End If
        Catch
        End Try
    End Sub

    '========================================================
    ' MACRO DELIVERY HELPERS
    '========================================================
    Private Sub EnsureTls12()
        Try
            ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
        Catch
        End Try
    End Sub

    Private Function EnsureMacroCacheFolder() As Boolean
        Try
            Dim baseDirRaw As String = AppDomain.CurrentDomain.BaseDirectory
            Dim baseDir As String = baseDirRaw
            Try
                baseDir = baseDir.TrimEnd("\"c)
            Catch
            End Try

            Dim baseName As String = ""
            Try
                baseName = Path.GetFileName(baseDir).ToLowerInvariant()
            Catch
                baseName = ""
            End Try

            If macroCacheDir Is Nothing OrElse macroCacheDir.Trim() = "" Then
                If baseName = "output" Then
                    macroCacheDir = Path.Combine(baseDirRaw, "macros")
                Else
                    macroCacheDir = Path.Combine(baseDirRaw, "output\macros")
                End If
            Else
                Dim p As String = macroCacheDir.Trim()

                If Not Path.IsPathRooted(p) Then
                    If baseName = "output" Then
                        If p.ToLowerInvariant().StartsWith("output\") Then
                            p = p.Substring(7)
                        End If
                    End If
                    macroCacheDir = Path.Combine(baseDirRaw, p)
                Else
                    macroCacheDir = p
                End If
            End If

            If Not Directory.Exists(macroCacheDir) Then
                Directory.CreateDirectory(macroCacheDir)
            End If

            If hideMacroCache Then
                Try
                    Dim attr As FileAttributes = File.GetAttributes(macroCacheDir)
                    If (attr And FileAttributes.Hidden) = 0 Then
                        File.SetAttributes(macroCacheDir, attr Or FileAttributes.Hidden)
                    End If
                Catch
                End Try
            End If

            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function EnsureMacroAvailable(originalPathOrKey As String, ByRef cachedLocalPath As String) As Boolean
        cachedLocalPath = ""

        Try
            If String.IsNullOrEmpty(originalPathOrKey) Then Return False

            If File.Exists(originalPathOrKey) Then
                cachedLocalPath = originalPathOrKey
                Return True
            End If

            If macroDeliveryUrl Is Nothing OrElse macroDeliveryUrl.Trim() = "" Then Return False
            If macroToken Is Nothing OrElse macroToken.Trim() = "" Then Return False

            If Not EnsureMacroCacheFolder() Then Return False

            Dim key As String = originalPathOrKey.Trim()
            Try
                If key.Contains("\") OrElse key.Contains("/") Then
                    key = Path.GetFileName(key)
                End If
            Catch
            End Try

            If key = "" Then Return False

            Dim dest As String = Path.Combine(macroCacheDir, key)
            If File.Exists(dest) Then
                cachedLocalPath = dest
                Return True
            End If

            EnsureTls12()

            Dim baseUrl As String = macroDeliveryUrl.Trim().TrimEnd("/"c)
            Dim url As String = baseUrl & "/api/v1/macro?key=" & Uri.EscapeDataString(key)

            LogA("☁️", "Fetching tool from server: " & key)

            Dim req As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            req.Method = "GET"
            req.UserAgent = "MDAT-MacroDelivery/1.0"
            req.Timeout = 20000
            req.ReadWriteTimeout = 20000
            req.Headers(HttpRequestHeader.Authorization) = "Bearer " & macroToken.Trim()

            Dim resp As HttpWebResponse = Nothing
            Try
                resp = CType(req.GetResponse(), HttpWebResponse)
            Catch ex As WebException
                Dim code As Integer = 0
                Try
                    If ex.Response IsNot Nothing Then
                        Dim r As HttpWebResponse = TryCast(ex.Response, HttpWebResponse)
                        If r IsNot Nothing Then code = CInt(r.StatusCode)
                    End If
                Catch
                End Try

                LogA("❌", "Macro fetch failed (" & code.ToString() & "): " & key)
                Return False
            End Try

            If resp Is Nothing Then
                LogA("❌", "Macro fetch failed: no response")
                Return False
            End If

            If resp.StatusCode <> HttpStatusCode.OK Then
                LogA("❌", "Macro fetch failed: HTTP " & CInt(resp.StatusCode).ToString())
                Try
                    resp.Close()
                Catch
                End Try
                Return False
            End If

            Try
                Using rs As Stream = resp.GetResponseStream()
                    Using fs As New FileStream(dest, FileMode.Create, FileAccess.Write, FileShare.Read)
                        Dim buf(8191) As Byte
                        While True
                            Dim n As Integer = rs.Read(buf, 0, buf.Length)
                            If n <= 0 Then Exit While
                            fs.Write(buf, 0, n)
                        End While
                    End Using
                End Using
            Finally
                Try
                    resp.Close()
                Catch
                End Try
            End Try

            If File.Exists(dest) Then
                cachedLocalPath = dest
                LogA("✅", "Tool cached: " & key)
                Return True
            End If

            Return False

        Catch ex As Exception
            LogA("❌", "Macro fetch error: " & ex.Message)
            Return False
        End Try
    End Function

    '========================================================
    ' TELEMETRY HELPERS
    '========================================================
    Private Sub SendTelemetrySafe(status As String, slot As Integer, actionName As String, durationMs As Integer, logText As String)
        Try
            If Not telemetryEnabled Then Exit Sub
            If syncUrl Is Nothing OrElse syncUrl.Trim() = "" Then Exit Sub
            If telemetryToken Is Nothing OrElse telemetryToken.Trim() = "" Then Exit Sub

            Dim exeVersion As String = GetExeVersionSafe()

            Dim licenseId As String = ""
            Try
                licenseId = GetLicenseIdSafe(activeLicense)
            Catch
                licenseId = ""
            End Try

            Dim machineId As String = ""
            Try
                machineId = Environment.MachineName
            Catch
                machineId = ""
            End Try

            TelemetryService.SendEvent(syncUrl, telemetryToken, status, slot, actionName, exeVersion, licenseId, machineId, durationMs, logText)
        Catch
        End Try
    End Sub

    Private Function GetExeVersionSafe() As String
        Try
            Dim v As Version = Assembly.GetExecutingAssembly().GetName().Version
            If v IsNot Nothing Then Return v.ToString()
        Catch
        End Try
        Return "1.0"
    End Function

    Private Function GetAssemblyNameSafe() As String
        Try
            If selectedAssemblyPath Is Nothing Then Return ""
            Dim p As String = selectedAssemblyPath.Trim()
            If p = "" Then Return ""
            Return Path.GetFileName(p)
        Catch
            Return ""
        End Try
    End Function

    '========================================================
    ' LOG (Option A)
    '========================================================
    Private Sub Log(msg As String)
        If txtLog Is Nothing Then Return
        txtLog.AppendText(DateTime.Now.ToString("HH:mm:ss") & "  " & msg & vbCrLf)
    End Sub

    Private Sub LogA(icon As String, msg As String)
        If String.IsNullOrEmpty(icon) Then
            Log(msg)
        Else
            Log(icon & " " & msg)
        End If
    End Sub

End Class
