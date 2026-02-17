Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Collections.Generic

Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Imports PdfSharp.Drawing

Module PdfMergeTool

    Private Const INDEX_TITLE As String = "INDEX OF DRAWINGS"

    ' Page number placement (lower)
    Private Const PAGE_NUM_FONT_SIZE As Double = 8.0
    Private Const PAGE_NUM_RIGHT_MARGIN As Double = 8.0
    Private Const PAGE_NUM_BOTTOM_MARGIN As Double = 3.0

    ' Index layout
    Private Const INDEX_LINES_PER_PAGE As Integer = 38
    Private Const INDEX_INDENT_STEP As Double = 14.0
    Private Const INDEX_LEFT_MARGIN As Double = 45.0

    ' Columns
    Private Const COL_NO_WIDTH As Double = 70.0
    Private Const COL_NO_PAD As Double = 4.0
    Private Const COL_DRAWING_X As Double = INDEX_LEFT_MARGIN + COL_NO_WIDTH + 10.0

    Private Class RawLine
        Public IndentLevel As Integer
        Public Path As String
        Public Display As String ' NEW: description / display text
    End Class

    Private Class IndexEntry
        Public FilePath As String
        Public FileName As String
        Public DisplayName As String ' NEW: what we show in index
        Public PageCount As Integer
        Public StartPageInFinal As Integer
        Public IndentLevel As Integer
        Public HierNo As String
        Public TreePrefix As String
    End Class

    Private Class Node
        Public Entry As IndexEntry
        Public Children As New List(Of Node)()
    End Class

    Sub Main()
        Dim args() As String = Environment.GetCommandLineArgs()

        If args Is Nothing OrElse args.Length < 2 Then
            PrintUsage()
            Environment.ExitCode = 2
            Return
        End If

        Try
            ' Check DLL beside EXE
            Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
            Dim dllPathGdi As String = Path.Combine(baseDir, "PdfSharp-gdi.dll")
            Dim dllPathStd As String = Path.Combine(baseDir, "PdfSharp.dll")
            If Not File.Exists(dllPathGdi) AndAlso Not File.Exists(dllPathStd) Then
                Throw New FileNotFoundException("Missing PDFsharp DLL. Place PdfSharp-gdi.dll beside PdfMergeTool.exe", dllPathGdi)
            End If

            Dim outPath As String = ""
            Dim listPath As String = ""
            Dim inputs As New List(Of String)()

            Dim i As Integer = 1
            While i < args.Length
                Dim a As String = args(i)

                If String.Equals(a, "-out", StringComparison.OrdinalIgnoreCase) Then
                    i += 1
                    If i >= args.Length Then Throw New ArgumentException("Missing value for -out")
                    outPath = args(i)

                ElseIf String.Equals(a, "-list", StringComparison.OrdinalIgnoreCase) Then
                    i += 1
                    If i >= args.Length Then Throw New ArgumentException("Missing value for -list")
                    listPath = args(i)

                Else
                    inputs.Add(a)
                End If

                i += 1
            End While

            If outPath.Trim().Length = 0 Then
                Throw New ArgumentException("Output not specified. Use -out ""fullpath.pdf""")
            End If
            outPath = Path.GetFullPath(outPath)

            Dim raw As New List(Of RawLine)()

            If listPath.Trim().Length > 0 Then
                listPath = Path.GetFullPath(listPath)
                If Not File.Exists(listPath) Then Throw New FileNotFoundException("List file not found: " & listPath)

                Dim lines() As String = File.ReadAllLines(listPath)
                For Each line As String In lines
                    Dim rl As RawLine = ParseIndentedLine_WithOptionalDisplay(line)
                    If rl.Path.Trim().Length > 0 Then raw.Add(rl)
                Next
            End If

            If inputs.Count > 0 Then
                For Each p As String In inputs
                    Dim rl As New RawLine()
                    rl.IndentLevel = 0
                    rl.Path = p
                    rl.Display = ""
                    raw.Add(rl)
                Next
            End If

            If raw.Count = 0 Then Throw New ArgumentException("No input PDFs provided. Use -list or pass PDF paths.")

            raw = ExpandWildcards(raw)

            ' Validate inputs
            Dim cleaned As New List(Of RawLine)()
            For Each rl As RawLine In raw
                Dim full As String = Path.GetFullPath(rl.Path)
                If Not File.Exists(full) Then Throw New FileNotFoundException("Input PDF not found: " & full)
                If Not full.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then Throw New ArgumentException("Not a PDF: " & full)

                Dim x As New RawLine()
                x.IndentLevel = rl.IndentLevel
                x.Path = full
                x.Display = SafeDisplayText(rl.Display)
                cleaned.Add(x)
            Next

            Dim outDir As String = Path.GetDirectoryName(outPath)
            If outDir Is Nothing OrElse outDir.Trim().Length = 0 Then Throw New ArgumentException("Invalid output path: " & outPath)
            If Not Directory.Exists(outDir) Then Directory.CreateDirectory(outDir)

            Dim entries As List(Of IndexEntry) = BuildIndexEntries(cleaned)

            entries = SortChildrenAlphabeticallyWithinParents(entries)
            AssignHierarchicalNumbers(entries)
            AssignTreePrefixes(entries)

            CreateMergedWithIndexAndPageNumbers(entries, outPath)

            Console.WriteLine("OK: Merged " & entries.Count.ToString() & " PDFs -> " & outPath)
            Environment.ExitCode = 0

        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Environment.ExitCode = 1
        End Try
    End Sub

    ' ============================================================
    ' NEW: parse line with optional 2nd column:
    '   [tabs/spaces indent] path [TAB display text]
    ' ============================================================
    Private Function ParseIndentedLine_WithOptionalDisplay(ByVal line As String) As RawLine
        Dim rl As New RawLine()
        rl.IndentLevel = 0
        rl.Path = ""
        rl.Display = ""

        If line Is Nothing Then Return rl
        If line.Trim().Length = 0 Then Return rl

        Dim s As String = line

        ' Count leading indent
        Dim tabs As Integer = 0
        Dim spaces As Integer = 0
        Dim idx As Integer = 0
        While idx < s.Length
            Dim ch As Char = s.Chars(idx)
            If ch = ControlChars.Tab Then
                tabs += 1
                idx += 1
            ElseIf ch = " "c Then
                spaces += 1
                idx += 1
            Else
                Exit While
            End If
        End While

        Dim spaceIndents As Integer = spaces \ 4
        rl.IndentLevel = tabs + spaceIndents

        ' Remove indent only (don’t Trim the whole line or we lose ability to split cleanly)
        Dim rest As String = s.Substring(idx)

        ' Split into columns by TAB
        ' We support:
        '   path
        '   path<TAB>display
        Dim parts() As String = rest.Split(ControlChars.Tab)

        If parts.Length >= 1 Then
            rl.Path = parts(0).Trim()
        End If
        If parts.Length >= 2 Then
            rl.Display = parts(1).Trim()
        Else
            rl.Display = ""
        End If

        Return rl
    End Function

    Private Function SafeDisplayText(ByVal s As String) As String
        If s Is Nothing Then Return ""
        Dim t As String = s.Trim()
        If t.Length = 0 Then Return ""
        t = t.Replace(ControlChars.Tab, " "c)
        t = t.Replace(ControlChars.Cr, " "c)
        t = t.Replace(ControlChars.Lf, " "c)
        While t.Contains("  ")
            t = t.Replace("  ", " ")
        End While
        Return t.Trim()
    End Function

    Private Function ExpandWildcards(ByVal lines As List(Of RawLine)) As List(Of RawLine)
        Dim result As New List(Of RawLine)()

        For Each rl As RawLine In lines
            Dim p As String = rl.Path

            If p.IndexOfAny(New Char() {"*"c, "?"c}) >= 0 Then
                Dim dir As String = Path.GetDirectoryName(p)
                Dim pattern As String = Path.GetFileName(p)
                If dir Is Nothing OrElse dir.Trim().Length = 0 Then dir = Directory.GetCurrentDirectory()

                If Directory.Exists(dir) Then
                    Dim files() As String = Directory.GetFiles(dir, pattern)
                    Array.Sort(files, StringComparer.OrdinalIgnoreCase)
                    For Each f As String In files
                        Dim x As New RawLine()
                        x.IndentLevel = rl.IndentLevel
                        x.Path = f
                        x.Display = rl.Display
                        result.Add(x)
                    Next
                End If
            Else
                result.Add(rl)
            End If
        Next

        Return result
    End Function

    Private Function BuildIndexEntries(ByVal inputFiles As List(Of RawLine)) As List(Of IndexEntry)
        Dim entries As New List(Of IndexEntry)()

        For Each rl As RawLine In inputFiles
            Dim pc As Integer
            Using inDoc As PdfDocument = PdfReader.Open(rl.Path, PdfDocumentOpenMode.Import)
                pc = inDoc.PageCount
            End Using

            Dim e As New IndexEntry()
            e.FilePath = rl.Path
            e.FileName = Path.GetFileName(rl.Path)
            e.DisplayName = If(rl.Display IsNot Nothing AndAlso rl.Display.Trim().Length > 0, rl.Display.Trim(), e.FileName)
            e.PageCount = pc
            e.StartPageInFinal = 0
            e.IndentLevel = rl.IndentLevel
            e.HierNo = ""
            e.TreePrefix = ""
            entries.Add(e)
        Next

        Return entries
    End Function

    ' ============================================================
    ' SORT CHILDREN ALPHABETICALLY WITHIN EACH PARENT
    ' ============================================================
    Private Function SortChildrenAlphabeticallyWithinParents(ByVal entries As List(Of IndexEntry)) As List(Of IndexEntry)
        Dim roots As New List(Of Node)()
        Dim stack As New List(Of Node)()

        For i As Integer = 0 To entries.Count - 1
            Dim e As IndexEntry = entries(i)

            Dim d As Integer = e.IndentLevel
            If d < 0 Then d = 0
            If i > 0 Then
                Dim prevD As Integer = entries(i - 1).IndentLevel
                If d > prevD + 1 Then d = prevD + 1
            End If
            e.IndentLevel = d

            Dim n As New Node()
            n.Entry = e

            While stack.Count > d
                stack.RemoveAt(stack.Count - 1)
            End While

            If d = 0 OrElse stack.Count = 0 Then
                roots.Add(n)
            Else
                stack(stack.Count - 1).Children.Add(n)
            End If

            stack.Add(n)
        Next

        For Each r As Node In roots
            SortChildrenRecursive(r)
        Next

        Dim flat As New List(Of IndexEntry)()
        For Each r As Node In roots
            Flatten(r, flat)
        Next

        Return flat
    End Function

    Private Sub SortChildrenRecursive(ByVal node As Node)
        If node Is Nothing Then Return
        If node.Children Is Nothing OrElse node.Children.Count = 0 Then Return

        node.Children.Sort(Function(a As Node, b As Node)
                               Return StringComparer.OrdinalIgnoreCase.Compare(a.Entry.FileName, b.Entry.FileName)
                           End Function)

        For Each c As Node In node.Children
            SortChildrenRecursive(c)
        Next
    End Sub

    Private Sub Flatten(ByVal node As Node, ByVal list As List(Of IndexEntry))
        list.Add(node.Entry)
        For Each c As Node In node.Children
            Flatten(c, list)
        Next
    End Sub

    ' ============================================================
    ' Hierarchical numbering
    ' ============================================================
    Private Sub AssignHierarchicalNumbers(ByVal entries As List(Of IndexEntry))
        Dim counters As New List(Of Integer)()
        Dim prevDepth As Integer = 0

        For i As Integer = 0 To entries.Count - 1
            Dim d As Integer = entries(i).IndentLevel
            If d < 0 Then d = 0
            If i > 0 AndAlso d > prevDepth + 1 Then d = prevDepth + 1
            entries(i).IndentLevel = d

            While counters.Count < d + 1
                counters.Add(0)
            End While

            If counters.Count > d + 1 Then
                counters.RemoveRange(d + 1, counters.Count - (d + 1))
            End If

            counters(d) = counters(d) + 1
            entries(i).HierNo = BuildHierNo(counters)

            prevDepth = d
        Next
    End Sub

    Private Function BuildHierNo(ByVal counters As List(Of Integer)) As String
        Dim parts As New List(Of String)()
        For Each n As Integer In counters
            If n <= 0 Then Exit For
            parts.Add(n.ToString())
        Next
        If parts.Count = 0 Then Return "1."
        Return String.Join(".", parts.ToArray()) & "."
    End Function

    ' ============================================================
    ' Tree prefix builder
    ' ============================================================
    Private Sub AssignTreePrefixes(ByVal entries As List(Of IndexEntry))
        Dim depthHasMore As New List(Of Boolean)()

        For i As Integer = 0 To entries.Count - 1
            Dim d As Integer = entries(i).IndentLevel
            If d < 0 Then d = 0

            While depthHasMore.Count < d
                depthHasMore.Add(False)
            End While
            If depthHasMore.Count > d Then
                depthHasMore.RemoveRange(d, depthHasMore.Count - d)
            End If

            Dim isLastSibling As Boolean = True
            If i < entries.Count - 1 Then
                Dim nextD As Integer = entries(i + 1).IndentLevel
                If nextD = d Then
                    isLastSibling = False
                Else
                    isLastSibling = True
                End If

                If d > 0 Then
                    Dim j As Integer = i + 1
                    Dim foundSame As Boolean = False
                    While j < entries.Count
                        Dim dj As Integer = entries(j).IndentLevel
                        If dj < d Then Exit While
                        If dj = d Then
                            foundSame = True
                            Exit While
                        End If
                        j += 1
                    End While
                    isLastSibling = Not foundSame
                End If
            End If

            If d > 0 Then
                While depthHasMore.Count < d
                    depthHasMore.Add(False)
                End While
                depthHasMore(d - 1) = Not isLastSibling
            End If

            Dim prefix As String = ""

            If d > 0 Then
                For level As Integer = 0 To d - 2
                    If level < depthHasMore.Count AndAlso depthHasMore(level) Then
                        prefix &= "│  "
                    Else
                        prefix &= "   "
                    End If
                Next

                If isLastSibling Then
                    prefix &= "└─ "
                Else
                    prefix &= "├─ "
                End If
            End If

            entries(i).TreePrefix = prefix
        Next
    End Sub

    Private Sub CreateMergedWithIndexAndPageNumbers(ByVal entries As List(Of IndexEntry), ByVal outputFile As String)
        Using outDoc As New PdfDocument()

            Dim indexPageCount As Integer = GetIndexPageCount(entries.Count, INDEX_LINES_PER_PAGE)

            Dim running As Integer = indexPageCount + 1
            For i As Integer = 0 To entries.Count - 1
                entries(i).StartPageInFinal = running
                running += entries(i).PageCount
            Next

            CreateIndexPages(outDoc, entries, INDEX_LINES_PER_PAGE, indexPageCount)

            For Each e As IndexEntry In entries
                Using inDoc As PdfDocument = PdfReader.Open(e.FilePath, PdfDocumentOpenMode.Import)
                    For p As Integer = 0 To inDoc.PageCount - 1
                        outDoc.AddPage(inDoc.Pages(p))
                    Next
                End Using
            Next

            StampPageNumbers(outDoc)
            outDoc.Save(outputFile)
        End Using
    End Sub

    Private Function GetIndexPageCount(ByVal totalEntries As Integer, ByVal linesPerPage As Integer) As Integer
        If totalEntries <= 0 Then Return 1
        Dim pages As Integer = CInt(Math.Ceiling(totalEntries / CDbl(linesPerPage)))
        If pages < 1 Then pages = 1
        Return pages
    End Function

    Private Sub CreateIndexPages(ByVal outDoc As PdfDocument, ByVal entries As List(Of IndexEntry), ByVal linesPerPage As Integer, ByVal indexPageCount As Integer)

        Dim titleFont As New XFont("Arial", 16, XFontStyle.Bold)
        Dim headerFont As New XFont("Arial", 10, XFontStyle.Bold)
        Dim rowFont As New XFont("Arial", 10, XFontStyle.Regular)

        Dim leftMargin As Double = INDEX_LEFT_MARGIN
        Dim topMargin As Double = 60.0
        Dim lineH As Double = 14.0

        Dim entryIndex As Integer = 0

        For pageNum As Integer = 1 To indexPageCount
            Dim page As PdfPage = outDoc.AddPage()
            page.Size = PdfSharp.PageSize.A4
            page.Orientation = PdfSharp.PageOrientation.Portrait

            Dim pageW As Double = page.Width.Point

            Using gfx As XGraphics = XGraphics.FromPdfPage(page)

                gfx.DrawString(INDEX_TITLE, titleFont, XBrushes.Black, New XRect(0, 25, pageW, 30), XStringFormats.TopCenter)

                Dim subText As String = "Index Page " & pageNum.ToString() & " of " & indexPageCount.ToString()
                gfx.DrawString(subText, rowFont, XBrushes.Black, New XRect(leftMargin, 45, pageW - (leftMargin * 2.0), 15), XStringFormats.TopLeft)

                Dim y As Double = topMargin

                gfx.DrawString("No.", headerFont, XBrushes.Black, New XRect(leftMargin, y, COL_NO_WIDTH, lineH), XStringFormats.TopLeft)
                gfx.DrawString("Drawing", headerFont, XBrushes.Black, New XRect(COL_DRAWING_X, y, pageW - 235, lineH), XStringFormats.TopLeft)
                gfx.DrawString("Start Pg", headerFont, XBrushes.Black, New XRect(pageW - 140, y, 80, lineH), XStringFormats.TopLeft)
                gfx.DrawString("Pages", headerFont, XBrushes.Black, New XRect(pageW - 70, y, 60, lineH), XStringFormats.TopLeft)

                y += (lineH + 6)

                Dim lineOnPage As Integer = 0
                While entryIndex < entries.Count AndAlso lineOnPage < linesPerPage
                    Dim e As IndexEntry = entries(entryIndex)

                    Dim noIndent As Double = (e.IndentLevel * INDEX_INDENT_STEP)
                    Dim noX As Double = leftMargin + noIndent
                    gfx.DrawString(e.HierNo, rowFont, XBrushes.Black, New XRect(noX, y, COL_NO_WIDTH - noIndent - COL_NO_PAD, lineH), XStringFormats.TopLeft)

                    ' NEW: use DisplayName (from 2nd column), not file name
                    Dim dn As String = e.DisplayName
                    If dn Is Nothing Then dn = ""
                    dn = dn.Trim()
                    If dn.Length = 0 Then dn = e.FileName

                    If dn.Length > 110 Then dn = dn.Substring(0, 107) & "..."

                    Dim displayLine As String = e.TreePrefix & dn
                    gfx.DrawString(displayLine, rowFont, XBrushes.Black, New XRect(COL_DRAWING_X, y, pageW - 235, lineH), XStringFormats.TopLeft)

                    gfx.DrawString(e.StartPageInFinal.ToString(), rowFont, XBrushes.Black, New XRect(pageW - 140, y, 80, lineH), XStringFormats.TopLeft)
                    gfx.DrawString(e.PageCount.ToString(), rowFont, XBrushes.Black, New XRect(pageW - 70, y, 60, lineH), XStringFormats.TopLeft)

                    y += lineH
                    entryIndex += 1
                    lineOnPage += 1
                End While

            End Using
        Next
    End Sub

    Private Sub StampPageNumbers(ByVal doc As PdfDocument)
        Dim total As Integer = doc.PageCount
        Dim font As New XFont("Arial", PAGE_NUM_FONT_SIZE, XFontStyle.Regular)

        For i As Integer = 0 To total - 1
            Dim page As PdfPage = doc.Pages(i)
            Dim pageW As Double = page.Width.Point
            Dim pageH As Double = page.Height.Point

            Using gfx As XGraphics = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append)
                Dim text As String = "Page " & (i + 1).ToString() & " of " & total.ToString()
                Dim y As Double = pageH - PAGE_NUM_BOTTOM_MARGIN - 10.0
                gfx.DrawString(text, font, XBrushes.Black, New XRect(0, y, pageW - PAGE_NUM_RIGHT_MARGIN, 12.0), XStringFormats.TopRight)
            End Using
        Next
    End Sub

    Private Sub PrintUsage()
        Console.WriteLine("PdfMergeTool (VB.NET 4.0 + PDFsharp-gdi)")
        Console.WriteLine("Usage:")
        Console.WriteLine("  PdfMergeTool.exe -out ""C:\out\Combined.pdf"" -list ""C:\out\merge_order.txt""")
        Console.WriteLine("  PdfMergeTool.exe -out ""C:\out\Combined.pdf"" -list ""C:\out\merge_order_with_desc.txt""")
    End Sub

End Module
