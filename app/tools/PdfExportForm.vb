Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Collections.Generic

Public Class PdfExportForm
    Inherits Form

    Private lstFiles As ListBox
    Private btnAddFiles As Button
    Private btnAddFolder As Button
    Private btnRemove As Button
    Private btnMoveUp As Button
    Private btnMoveDown As Button
    Private btnMerge As Button
    Private btnMergeNoIndex As Button
    Private chkOpenAfter As CheckBox
    Private lblStatus As Label

    Public Sub New()
        Me.Text = "PDF Merge Tool - MetaMech"
        Me.Size = New Size(700, 520)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MinimumSize = New Size(550, 400)
        Me.BackColor = Color.FromArgb(24, 30, 40)
        Me.ForeColor = Color.White
        Me.Font = New Font("Segoe UI", 9)

        FormHeader.AddPremiumHeader(Me, "PDF Merge Tool", "MetaMech Engineering Tools")

        ' File list
        lstFiles = New ListBox()
        lstFiles.Location = New Point(20, 70)
        lstFiles.Size = New Size(440, 340)
        lstFiles.BackColor = Color.FromArgb(30, 38, 50)
        lstFiles.ForeColor = Color.White
        lstFiles.Font = New Font("Consolas", 9)
        lstFiles.HorizontalScrollbar = True
        lstFiles.SelectionMode = SelectionMode.ExtendedSimple
        lstFiles.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Me.Controls.Add(lstFiles)

        ' Right-side buttons
        Dim btnX As Integer = 475
        Dim btnW As Integer = 190

        btnAddFiles = New Button()
        btnAddFiles.Text = "📄 Add PDF Files..."
        btnAddFiles.Location = New Point(btnX, 70)
        btnAddFiles.Size = New Size(btnW, 32)
        btnAddFiles.FlatStyle = FlatStyle.Flat
        btnAddFiles.BackColor = Color.FromArgb(45, 55, 72)
        btnAddFiles.ForeColor = Color.White
        btnAddFiles.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnAddFiles.Click, AddressOf AddFiles_Click
        Me.Controls.Add(btnAddFiles)

        btnAddFolder = New Button()
        btnAddFolder.Text = "📁 Add Folder..."
        btnAddFolder.Location = New Point(btnX, 110)
        btnAddFolder.Size = New Size(btnW, 32)
        btnAddFolder.FlatStyle = FlatStyle.Flat
        btnAddFolder.BackColor = Color.FromArgb(45, 55, 72)
        btnAddFolder.ForeColor = Color.White
        btnAddFolder.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnAddFolder.Click, AddressOf AddFolder_Click
        Me.Controls.Add(btnAddFolder)

        btnRemove = New Button()
        btnRemove.Text = "❌ Remove Selected"
        btnRemove.Location = New Point(btnX, 160)
        btnRemove.Size = New Size(btnW, 32)
        btnRemove.FlatStyle = FlatStyle.Flat
        btnRemove.BackColor = Color.FromArgb(45, 55, 72)
        btnRemove.ForeColor = Color.White
        btnRemove.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnRemove.Click, AddressOf Remove_Click
        Me.Controls.Add(btnRemove)

        btnMoveUp = New Button()
        btnMoveUp.Text = "⬆ Move Up"
        btnMoveUp.Location = New Point(btnX, 210)
        btnMoveUp.Size = New Size(90, 32)
        btnMoveUp.FlatStyle = FlatStyle.Flat
        btnMoveUp.BackColor = Color.FromArgb(45, 55, 72)
        btnMoveUp.ForeColor = Color.White
        btnMoveUp.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnMoveUp.Click, AddressOf MoveUp_Click
        Me.Controls.Add(btnMoveUp)

        btnMoveDown = New Button()
        btnMoveDown.Text = "⬇ Move Down"
        btnMoveDown.Location = New Point(btnX + 95, 210)
        btnMoveDown.Size = New Size(95, 32)
        btnMoveDown.FlatStyle = FlatStyle.Flat
        btnMoveDown.BackColor = Color.FromArgb(45, 55, 72)
        btnMoveDown.ForeColor = Color.White
        btnMoveDown.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnMoveDown.Click, AddressOf MoveDown_Click
        Me.Controls.Add(btnMoveDown)

        chkOpenAfter = New CheckBox()
        chkOpenAfter.Text = "Open PDF after merge"
        chkOpenAfter.Location = New Point(btnX, 260)
        chkOpenAfter.AutoSize = True
        chkOpenAfter.Checked = True
        chkOpenAfter.ForeColor = Color.White
        chkOpenAfter.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Me.Controls.Add(chkOpenAfter)

        btnMerge = New Button()
        btnMerge.Text = "🔗 Merge with Index"
        btnMerge.Location = New Point(btnX, 300)
        btnMerge.Size = New Size(btnW, 40)
        btnMerge.FlatStyle = FlatStyle.Flat
        btnMerge.BackColor = Color.FromArgb(0, 120, 215)
        btnMerge.ForeColor = Color.White
        btnMerge.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnMerge.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnMerge.Click, AddressOf Merge_Click
        Me.Controls.Add(btnMerge)

        btnMergeNoIndex = New Button()
        btnMergeNoIndex.Text = "📋 Merge (No Index)"
        btnMergeNoIndex.Location = New Point(btnX, 348)
        btnMergeNoIndex.Size = New Size(btnW, 32)
        btnMergeNoIndex.FlatStyle = FlatStyle.Flat
        btnMergeNoIndex.BackColor = Color.FromArgb(60, 70, 85)
        btnMergeNoIndex.ForeColor = Color.White
        btnMergeNoIndex.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnMergeNoIndex.Click, AddressOf MergeNoIndex_Click
        Me.Controls.Add(btnMergeNoIndex)

        lblStatus = New Label()
        lblStatus.Text = "Add PDF files to merge."
        lblStatus.Location = New Point(20, 420)
        lblStatus.Size = New Size(640, 40)
        lblStatus.ForeColor = Color.FromArgb(160, 170, 180)
        lblStatus.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Me.Controls.Add(lblStatus)
    End Sub

    Private Sub AddFiles_Click(sender As Object, e As EventArgs)
        Using ofd As New OpenFileDialog()
            ofd.Title = "Select PDF files to merge"
            ofd.Filter = "PDF Files (*.pdf)|*.pdf"
            ofd.Multiselect = True
            If ofd.ShowDialog() = DialogResult.OK Then
                For Each f As String In ofd.FileNames
                    If Not lstFiles.Items.Contains(f) Then
                        lstFiles.Items.Add(f)
                    End If
                Next
                UpdateStatus()
            End If
        End Using
    End Sub

    Private Sub AddFolder_Click(sender As Object, e As EventArgs)
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select folder containing PDF files"
            If fbd.ShowDialog() = DialogResult.OK Then
                Dim files() As String = Directory.GetFiles(fbd.SelectedPath, "*.pdf", SearchOption.TopDirectoryOnly)
                Array.Sort(files, StringComparer.OrdinalIgnoreCase)
                For Each f As String In files
                    If Not lstFiles.Items.Contains(f) Then
                        lstFiles.Items.Add(f)
                    End If
                Next
                UpdateStatus()
            End If
        End Using
    End Sub

    Private Sub Remove_Click(sender As Object, e As EventArgs)
        Dim selected As New List(Of Object)()
        For Each item As Object In lstFiles.SelectedItems
            selected.Add(item)
        Next
        For Each item As Object In selected
            lstFiles.Items.Remove(item)
        Next
        UpdateStatus()
    End Sub

    Private Sub MoveUp_Click(sender As Object, e As EventArgs)
        If lstFiles.SelectedIndex > 0 Then
            Dim idx As Integer = lstFiles.SelectedIndex
            Dim item As Object = lstFiles.Items(idx)
            lstFiles.Items.RemoveAt(idx)
            lstFiles.Items.Insert(idx - 1, item)
            lstFiles.SelectedIndex = idx - 1
        End If
    End Sub

    Private Sub MoveDown_Click(sender As Object, e As EventArgs)
        If lstFiles.SelectedIndex >= 0 AndAlso lstFiles.SelectedIndex < lstFiles.Items.Count - 1 Then
            Dim idx As Integer = lstFiles.SelectedIndex
            Dim item As Object = lstFiles.Items(idx)
            lstFiles.Items.RemoveAt(idx)
            lstFiles.Items.Insert(idx + 1, item)
            lstFiles.SelectedIndex = idx + 1
        End If
    End Sub

    Private Sub Merge_Click(sender As Object, e As EventArgs)
        RunMerge(False)
    End Sub

    Private Sub MergeNoIndex_Click(sender As Object, e As EventArgs)
        RunMerge(True)
    End Sub

    Private Sub RunMerge(noIndex As Boolean)
        If lstFiles.Items.Count < 2 Then
            MessageBox.Show("Add at least 2 PDF files to merge.", "PDF Merge", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Ask where to save
        Dim outputPath As String = ""
        Using sfd As New SaveFileDialog()
            sfd.Title = "Save merged PDF as"
            sfd.Filter = "PDF Files (*.pdf)|*.pdf"
            sfd.FileName = "Combined_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"
            If sfd.ShowDialog() <> DialogResult.OK Then Return
            outputPath = sfd.FileName
        End Using

        ' Find PdfMergeTool.exe beside the main EXE
        Dim mergeToolPath As String = Path.Combine(Application.StartupPath, "PdfMergeTool.exe")
        If Not File.Exists(mergeToolPath) Then
            MessageBox.Show("PdfMergeTool.exe not found beside MDAT.exe." & vbCrLf & _
                            "Expected: " & mergeToolPath, _
                            "PDF Merge Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return
        End If

        ' Write merge_order.txt in the output folder
        Dim orderFile As String = Path.Combine(Path.GetDirectoryName(outputPath), "merge_order.txt")
        Dim lines As New List(Of String)()
        For Each item As Object In lstFiles.Items
            lines.Add(CStr(item))
        Next
        File.WriteAllLines(orderFile, lines.ToArray())

        ' Build arguments
        Dim args As String = ""
        If noIndex Then
            args = "-noindex "
        End If
        args &= "-out """ & outputPath & """ -list """ & orderFile & """"

        ' Run PdfMergeTool.exe
        lblStatus.Text = "Merging " & lstFiles.Items.Count.ToString() & " PDFs..."
        lblStatus.ForeColor = Color.FromArgb(100, 200, 255)
        Me.Refresh()

        Try
            Dim psi As New ProcessStartInfo()
            psi.FileName = mergeToolPath
            psi.Arguments = args
            psi.UseShellExecute = False
            psi.RedirectStandardOutput = True
            psi.RedirectStandardError = True
            psi.CreateNoWindow = True
            psi.WorkingDirectory = Path.GetDirectoryName(outputPath)

            Dim proc As Process = Process.Start(psi)
            Dim stdout As String = proc.StandardOutput.ReadToEnd()
            Dim stderr As String = proc.StandardError.ReadToEnd()
            proc.WaitForExit()

            If proc.ExitCode = 0 AndAlso File.Exists(outputPath) Then
                lblStatus.Text = "✅ Merged " & lstFiles.Items.Count.ToString() & " PDFs → " & Path.GetFileName(outputPath)
                lblStatus.ForeColor = Color.FromArgb(100, 255, 100)

                MessageBox.Show("PDF merge complete!" & vbCrLf & vbCrLf & _
                                "Output: " & outputPath & vbCrLf & _
                                "Files merged: " & lstFiles.Items.Count.ToString(), _
                                "PDF Merge", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If chkOpenAfter.Checked Then
                    Process.Start(outputPath)
                End If
            Else
                Dim errMsg As String = "PdfMergeTool failed." & vbCrLf & vbCrLf
                If stdout.Trim().Length > 0 Then errMsg &= "Output: " & stdout.Trim() & vbCrLf
                If stderr.Trim().Length > 0 Then errMsg &= "Error: " & stderr.Trim() & vbCrLf

                lblStatus.Text = "❌ Merge failed. See error."
                lblStatus.ForeColor = Color.FromArgb(255, 100, 100)
                MessageBox.Show(errMsg, "PDF Merge Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End If

        Catch ex As Exception
            lblStatus.Text = "❌ Error: " & ex.Message
            lblStatus.ForeColor = Color.FromArgb(255, 100, 100)
            MessageBox.Show("Failed to run PdfMergeTool:" & vbCrLf & ex.Message, _
                            "PDF Merge Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub UpdateStatus()
        lblStatus.Text = lstFiles.Items.Count.ToString() & " PDF file(s) ready to merge."
        lblStatus.ForeColor = Color.FromArgb(160, 170, 180)
    End Sub

End Class
