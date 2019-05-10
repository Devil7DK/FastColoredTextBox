Imports System
Imports System.Text
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Namespace FastColoredTextBoxNS
    Public Class Range
        Implements IEnumerable(Of Place)

        Private start As Place
        Private [end] As Place
        Public ReadOnly tb As FastColoredTextBox
        Private preferedPos As Integer = -1
        Private updating As Integer = 0
        Private cachedText As String
        Private cachedCharIndexToPlace As List(Of Place)
        Private cachedTextVersion As Integer = -1

        Public Sub New(ByVal tb As FastColoredTextBox)
            Me.tb = tb
        End Sub

        Public Overridable ReadOnly Property IsEmpty As Boolean
            Get
                If columnSelectionMode Then Return start.iChar = [end].iChar
                Return start = [end]
            End Get
        End Property

        Private columnSelectionMode As Boolean

        Public Property ColumnSelectionMode As Boolean
            Get
                Return ColumnSelectionMode
            End Get
            Set(ByVal value As Boolean)
                columnSelectionMode = value
            End Set
        End Property

        Public Sub New(ByVal tb As FastColoredTextBox, ByVal iStartChar As Integer, ByVal iStartLine As Integer, ByVal iEndChar As Integer, ByVal iEndLine As Integer)
            Me.New(tb)
            start = New Place(iStartChar, iStartLine)
            [end] = New Place(iEndChar, iEndLine)
        End Sub

        Public Sub New(ByVal tb As FastColoredTextBox, ByVal start As Place, ByVal [end] As Place)
            Me.New(tb)
            Me.start = start
            Me.[end] = [end]
        End Sub

        Public Sub New(ByVal tb As FastColoredTextBox, ByVal iLine As Integer)
            Me.New(tb)
            start = New Place(0, iLine)
            [end] = New Place(tb(iLine).Count, iLine)
        End Sub

        Public Function Contains(ByVal place As Place) As Boolean
            If place.iLine < Math.Min(start.iLine, [end].iLine) Then Return False
            If place.iLine > Math.Max(start.iLine, [end].iLine) Then Return False
            Dim s As Place = start
            Dim e As Place = [end]

            If s.iLine > e.iLine OrElse (s.iLine = e.iLine AndAlso s.iChar > e.iChar) Then
                Dim temp = s
                s = e
                e = temp
            End If

            If columnSelectionMode Then
                If place.iChar < s.iChar OrElse place.iChar > e.iChar Then Return False
            Else
                If place.iLine = s.iLine AndAlso place.iChar < s.iChar Then Return False
                If place.iLine = e.iLine AndAlso place.iChar > e.iChar Then Return False
            End If

            Return True
        End Function

        Public Overridable Function GetIntersectionWith(ByVal range As Range) As Range
            If columnSelectionMode Then Return GetIntersectionWith_ColumnSelectionMode(range)
            Dim r1 As Range = Me.Clone()
            Dim r2 As Range = range.Clone()
            r1.Normalize()
            r2.Normalize()
            Dim newStart As Place = If(r1.start > r2.start, r1.start, r2.start)
            Dim newEnd As Place = If(r1.[end] < r2.[end], r1.[end], r2.[end])
            If newEnd < newStart Then Return New Range(tb, start, start)
            Return tb.GetRange(newStart, newEnd)
        End Function

        Public Function GetUnionWith(ByVal range As Range) As Range
            Dim r1 As Range = Me.Clone()
            Dim r2 As Range = range.Clone()
            r1.Normalize()
            r2.Normalize()
            Dim newStart As Place = If(r1.start < r2.start, r1.start, r2.start)
            Dim newEnd As Place = If(r1.[end] > r2.[end], r1.[end], r2.[end])
            Return tb.GetRange(newStart, newEnd)
        End Function

        Public Sub SelectAll()
            columnSelectionMode = False
            start = New Place(0, 0)

            If tb.LinesCount = 0 Then
                start = New Place(0, 0)
            Else
                [end] = New Place(0, 0)
                start = New Place(tb(tb.LinesCount - 1).Count, tb.LinesCount - 1)
            End If

            If Me = tb.Selection Then tb.Invalidate()
        End Sub

        Public Property Start As Place
            Get
                Return Start
            End Get
            Set(ByVal value As Place)
                [end] = CSharpImpl.__Assign(start, value)
                preferedPos = -1
                OnSelectionChanged()
            End Set
        End Property

        Public Property [End] As Place
            Get
                Return [End]
            End Get
            Set(ByVal value As Place)
                [end] = value
                OnSelectionChanged()
            End Set
        End Property

        Public Overridable ReadOnly Property Text As String
            Get
                If columnSelectionMode Then Return Text_ColumnSelectionMode
                Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
                Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
                Dim fromChar As Integer = FromX
                Dim toChar As Integer = ToX
                If fromLine < 0 Then Return Nothing
                Dim sb As StringBuilder = New StringBuilder()

                For y As Integer = fromLine To toLine
                    Dim fromX As Integer = If(y = fromLine, fromChar, 0)
                    Dim toX As Integer = If(y = toLine, Math.Min(tb(y).Count - 1, toChar - 1), tb(y).Count - 1)

                    For x As Integer = fromX To toX
                        sb.Append(tb(y)(x).c)
                    Next

                    If y <> toLine AndAlso fromLine <> toLine Then sb.AppendLine()
                Next

                Return sb.ToString()
            End Get
        End Property

        Public ReadOnly Property Length As Integer
            Get
                If columnSelectionMode Then Return Length_ColumnSelectionMode(False)
                Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
                Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
                Dim cnt As Integer = 0
                If fromLine < 0 Then Return 0

                For y As Integer = fromLine To toLine
                    Dim fromX As Integer = If(y = fromLine, fromX, 0)
                    Dim toX As Integer = If(y = toLine, Math.Min(tb(y).Count - 1, toX - 1), tb(y).Count - 1)
                    cnt += toX - fromX + 1
                    If y <> toLine AndAlso fromLine <> toLine Then cnt += Environment.NewLine.Length
                Next

                Return cnt
            End Get
        End Property

        Public ReadOnly Property TextLength As Integer
            Get

                If columnSelectionMode Then
                    Return Length_ColumnSelectionMode(True)
                Else
                    Return Length
                End If
            End Get
        End Property

        Friend Sub GetText(<Out> ByRef text As String, <Out> ByRef charIndexToPlace As List(Of Place))
            If tb.TextVersion = cachedTextVersion Then
                text = cachedText
                charIndexToPlace = cachedCharIndexToPlace
                Return
            End If

            Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
            Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
            Dim fromChar As Integer = FromX
            Dim toChar As Integer = ToX
            Dim sb As StringBuilder = New StringBuilder((toLine - fromLine) * 50)
            charIndexToPlace = New List(Of Place)(sb.Capacity)

            If fromLine >= 0 Then

                For y As Integer = fromLine To toLine
                    Dim fromX As Integer = If(y = fromLine, fromChar, 0)
                    Dim toX As Integer = If(y = toLine, Math.Min(toChar - 1, tb(y).Count - 1), tb(y).Count - 1)

                    For x As Integer = fromX To toX
                        sb.Append(tb(y)(x).c)
                        charIndexToPlace.Add(New Place(x, y))
                    Next

                    If y <> toLine AndAlso fromLine <> toLine Then

                        For Each c As Char In Environment.NewLine
                            sb.Append(c)
                            charIndexToPlace.Add(New Place(tb(y).Count, y))
                        Next
                    End If
                Next
            End If

            text = sb.ToString()
            charIndexToPlace.Add(If([end] > start, [end], start))
            cachedText = text
            cachedCharIndexToPlace = charIndexToPlace
            cachedTextVersion = tb.TextVersion
        End Sub

        Public ReadOnly Property CharAfterStart As Char
            Get
                If start.iChar >= tb(start.iLine).Count Then Return vbLf
                Return tb(start.iLine)(start.iChar).c
            End Get
        End Property

        Public ReadOnly Property CharBeforeStart As Char
            Get
                If start.iChar > tb(start.iLine).Count Then Return vbLf
                If start.iChar <= 0 Then Return vbLf
                Return tb(start.iLine)(start.iChar - 1).c
            End Get
        End Property

        Public Function GetCharsBeforeStart(ByVal charsCount As Integer) As String
            Dim pos = tb.PlaceToPosition(start) - charsCount
            If pos < 0 Then pos = 0
            Return New Range(tb, tb.PositionToPlace(pos), start).Text
        End Function

        Public Function GetCharsAfterStart(ByVal charsCount As Integer) As String
            Return GetCharsBeforeStart(-charsCount)
        End Function

        Public Function Clone() As Range
            Return CType(MemberwiseClone(), Range)
        End Function

        Friend ReadOnly Property FromX As Integer
            Get
                If [end].iLine < start.iLine Then Return [end].iChar
                If [end].iLine > start.iLine Then Return start.iChar
                Return Math.Min([end].iChar, start.iChar)
            End Get
        End Property

        Friend ReadOnly Property ToX As Integer
            Get
                If [end].iLine < start.iLine Then Return start.iChar
                If [end].iLine > start.iLine Then Return [end].iChar
                Return Math.Max([end].iChar, start.iChar)
            End Get
        End Property

        Public ReadOnly Property FromLine As Integer
            Get
                Return Math.Min(start.iLine, [end].iLine)
            End Get
        End Property

        Public ReadOnly Property ToLine As Integer
            Get
                Return Math.Max(start.iLine, [end].iLine)
            End Get
        End Property

        Public Function GoRight() As Boolean
            Dim prevStart As Place = start
            GoRight(False)
            Return prevStart <> start
        End Function

        Public Overridable Function GoRightThroughFolded() As Boolean
            If columnSelectionMode Then Return GoRightThroughFolded_ColumnSelectionMode()
            If start.iLine >= tb.LinesCount - 1 AndAlso start.iChar >= tb(tb.LinesCount - 1).Count Then Return False

            If start.iChar < tb(start.iLine).Count Then
                start.Offset(1, 0)
            Else
                start = New Place(0, start.iLine + 1)
            End If

            preferedPos = -1
            [end] = start
            OnSelectionChanged()
            Return True
        End Function

        Public Function GoLeft() As Boolean
            columnSelectionMode = False
            Dim prevStart As Place = start
            GoLeft(False)
            Return prevStart <> start
        End Function

        Public Function GoLeftThroughFolded() As Boolean
            columnSelectionMode = False
            If start.iChar = 0 AndAlso start.iLine = 0 Then Return False

            If start.iChar > 0 Then
                start.Offset(-1, 0)
            Else
                start = New Place(tb(start.iLine - 1).Count, start.iLine - 1)
            End If

            preferedPos = -1
            [end] = start
            OnSelectionChanged()
            Return True
        End Function

        Public Sub GoLeft(ByVal shift As Boolean)
            columnSelectionMode = False

            If Not shift Then

                If start > [end] Then
                    start = [end]
                    Return
                End If
            End If

            If start.iChar <> 0 OrElse start.iLine <> 0 Then

                If start.iChar > 0 AndAlso tb.LineInfos(start.iLine).VisibleState = VisibleState.Visible Then
                    start.Offset(-1, 0)
                Else
                    Dim i As Integer = tb.FindPrevVisibleLine(start.iLine)
                    If i = start.iLine Then Return
                    start = New Place(tb(i).Count, i)
                End If
            End If

            If Not shift Then [end] = start
            OnSelectionChanged()
            preferedPos = -1
        End Sub

        Public Sub GoRight(ByVal shift As Boolean)
            columnSelectionMode = False

            If Not shift Then

                If start < [end] Then
                    start = [end]
                    Return
                End If
            End If

            If start.iLine < tb.LinesCount - 1 OrElse start.iChar < tb(tb.LinesCount - 1).Count Then

                If start.iChar < tb(start.iLine).Count AndAlso tb.LineInfos(start.iLine).VisibleState = VisibleState.Visible Then
                    start.Offset(1, 0)
                Else
                    Dim i As Integer = tb.FindNextVisibleLine(start.iLine)
                    If i = start.iLine Then Return
                    start = New Place(0, i)
                End If
            End If

            If Not shift Then [end] = start
            OnSelectionChanged()
            preferedPos = -1
        End Sub

        Friend Sub GoUp(ByVal shift As Boolean)
            columnSelectionMode = False

            If Not shift Then

                If start.iLine > [end].iLine Then
                    start = [end]
                    Return
                End If
            End If

            If preferedPos < 0 Then preferedPos = start.iChar - tb.LineInfos(start.iLine).GetWordWrapStringStartPosition(tb.LineInfos(start.iLine).GetWordWrapStringIndex(start.iChar))
            Dim iWW As Integer = tb.LineInfos(start.iLine).GetWordWrapStringIndex(start.iChar)

            If iWW = 0 Then
                If start.iLine <= 0 Then Return
                Dim i As Integer = tb.FindPrevVisibleLine(start.iLine)
                If i = start.iLine Then Return
                start.iLine = i
                iWW = tb.LineInfos(start.iLine).WordWrapStringsCount
            End If

            If iWW > 0 Then
                Dim finish As Integer = tb.LineInfos(start.iLine).GetWordWrapStringFinishPosition(iWW - 1, tb(start.iLine))
                start.iChar = tb.LineInfos(start.iLine).GetWordWrapStringStartPosition(iWW - 1) + preferedPos
                If start.iChar > finish + 1 Then start.iChar = finish + 1
            End If

            If Not shift Then [end] = start
            OnSelectionChanged()
        End Sub

        Friend Sub GoPageUp(ByVal shift As Boolean)
            columnSelectionMode = False
            If preferedPos < 0 Then preferedPos = start.iChar - tb.LineInfos(start.iLine).GetWordWrapStringStartPosition(tb.LineInfos(start.iLine).GetWordWrapStringIndex(start.iChar))
            Dim pageHeight As Integer = tb.ClientRectangle.Height / tb.CharHeight - 1

            For i As Integer = 0 To pageHeight - 1
                Dim iWW As Integer = tb.LineInfos(start.iLine).GetWordWrapStringIndex(start.iChar)

                If iWW = 0 Then
                    If start.iLine <= 0 Then Exit For
                    Dim newLine As Integer = tb.FindPrevVisibleLine(start.iLine)
                    If newLine = start.iLine Then Exit For
                    start.iLine = newLine
                    iWW = tb.LineInfos(start.iLine).WordWrapStringsCount
                End If

                If iWW > 0 Then
                    Dim finish As Integer = tb.LineInfos(start.iLine).GetWordWrapStringFinishPosition(iWW - 1, tb(start.iLine))
                    start.iChar = tb.LineInfos(start.iLine).GetWordWrapStringStartPosition(iWW - 1) + preferedPos
                    If start.iChar > finish + 1 Then start.iChar = finish + 1
                End If
            Next

            If Not shift Then [end] = start
            OnSelectionChanged()
        End Sub

        Friend Sub GoDown(ByVal shift As Boolean)
            columnSelectionMode = False

            If Not shift Then

                If start.iLine < [end].iLine Then
                    start = [end]
                    Return
                End If
            End If

            If preferedPos < 0 Then preferedPos = start.iChar - tb.LineInfos(start.iLine).GetWordWrapStringStartPosition(tb.LineInfos(start.iLine).GetWordWrapStringIndex(start.iChar))
            Dim iWW As Integer = tb.LineInfos(start.iLine).GetWordWrapStringIndex(start.iChar)

            If iWW >= tb.LineInfos(start.iLine).WordWrapStringsCount - 1 Then
                If start.iLine >= tb.LinesCount - 1 Then Return
                Dim i As Integer = tb.FindNextVisibleLine(start.iLine)
                If i = start.iLine Then Return
                start.iLine = i
                iWW = -1
            End If

            If iWW < tb.LineInfos(start.iLine).WordWrapStringsCount - 1 Then
                Dim finish As Integer = tb.LineInfos(start.iLine).GetWordWrapStringFinishPosition(iWW + 1, tb(start.iLine))
                start.iChar = tb.LineInfos(start.iLine).GetWordWrapStringStartPosition(iWW + 1) + preferedPos
                If start.iChar > finish + 1 Then start.iChar = finish + 1
            End If

            If Not shift Then [end] = start
            OnSelectionChanged()
        End Sub

        Friend Sub GoPageDown(ByVal shift As Boolean)
            columnSelectionMode = False
            If preferedPos < 0 Then preferedPos = start.iChar - tb.LineInfos(start.iLine).GetWordWrapStringStartPosition(tb.LineInfos(start.iLine).GetWordWrapStringIndex(start.iChar))
            Dim pageHeight As Integer = tb.ClientRectangle.Height / tb.CharHeight - 1

            For i As Integer = 0 To pageHeight - 1
                Dim iWW As Integer = tb.LineInfos(start.iLine).GetWordWrapStringIndex(start.iChar)

                If iWW >= tb.LineInfos(start.iLine).WordWrapStringsCount - 1 Then
                    If start.iLine >= tb.LinesCount - 1 Then Exit For
                    Dim newLine As Integer = tb.FindNextVisibleLine(start.iLine)
                    If newLine = start.iLine Then Exit For
                    start.iLine = newLine
                    iWW = -1
                End If

                If iWW < tb.LineInfos(start.iLine).WordWrapStringsCount - 1 Then
                    Dim finish As Integer = tb.LineInfos(start.iLine).GetWordWrapStringFinishPosition(iWW + 1, tb(start.iLine))
                    start.iChar = tb.LineInfos(start.iLine).GetWordWrapStringStartPosition(iWW + 1) + preferedPos
                    If start.iChar > finish + 1 Then start.iChar = finish + 1
                End If
            Next

            If Not shift Then [end] = start
            OnSelectionChanged()
        End Sub

        Friend Sub GoHome(ByVal shift As Boolean)
            columnSelectionMode = False
            If start.iLine < 0 Then Return
            If tb.LineInfos(start.iLine).VisibleState <> VisibleState.Visible Then Return
            start = New Place(0, start.iLine)
            If Not shift Then [end] = start
            OnSelectionChanged()
            preferedPos = -1
        End Sub

        Friend Sub GoEnd(ByVal shift As Boolean)
            columnSelectionMode = False
            If start.iLine < 0 Then Return
            If tb.LineInfos(start.iLine).VisibleState <> VisibleState.Visible Then Return
            start = New Place(tb(start.iLine).Count, start.iLine)
            If Not shift Then [end] = start
            OnSelectionChanged()
            preferedPos = -1
        End Sub

        Public Sub SetStyle(ByVal style As Style)
            Dim code As Integer = tb.GetOrSetStyleLayerIndex(style)
            SetStyle(ToStyleIndex(code))
            tb.Invalidate()
        End Sub

        Public Sub SetStyle(ByVal style As Style, ByVal regexPattern As String)
            Dim layer As StyleIndex = ToStyleIndex(tb.GetOrSetStyleLayerIndex(style))
            SetStyle(layer, regexPattern, RegexOptions.None)
        End Sub

        Public Sub SetStyle(ByVal style As Style, ByVal regex As Regex)
            Dim layer As StyleIndex = ToStyleIndex(tb.GetOrSetStyleLayerIndex(style))
            SetStyle(layer, regex)
        End Sub

        Public Sub SetStyle(ByVal style As Style, ByVal regexPattern As String, ByVal options As RegexOptions)
            Dim layer As StyleIndex = ToStyleIndex(tb.GetOrSetStyleLayerIndex(style))
            SetStyle(layer, regexPattern, options)
        End Sub

        Public Sub SetStyle(ByVal styleLayer As StyleIndex, ByVal regexPattern As String, ByVal options As RegexOptions)
            If Math.Abs(start.iLine - [end].iLine) > 1000 Then options = options Or SyntaxHighlighter.RegexCompiledOption

            For Each range In GetRanges(regexPattern, options)
                range.SetStyle(styleLayer)
            Next

            tb.Invalidate()
        End Sub

        Public Sub SetStyle(ByVal styleLayer As StyleIndex, ByVal regex As Regex)
            For Each range In GetRanges(regex)
                range.SetStyle(styleLayer)
            Next

            tb.Invalidate()
        End Sub

        Public Sub SetStyle(ByVal styleIndex As StyleIndex)
            Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
            Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
            Dim fromChar As Integer = FromX
            Dim toChar As Integer = ToX
            If fromLine < 0 Then Return

            For y As Integer = fromLine To toLine
                Dim fromX As Integer = If(y = fromLine, fromChar, 0)
                Dim toX As Integer = If(y = toLine, Math.Min(toChar - 1, tb(y).Count - 1), tb(y).Count - 1)

                For x As Integer = fromX To toX
                    Dim c As Char = tb(y)(x)
                    c.style = c.style Or styleIndex
                    tb(y)(x) = c
                Next
            Next
        End Sub

        Public Sub SetFoldingMarkers(ByVal startFoldingPattern As String, ByVal finishFoldingPattern As String)
            SetFoldingMarkers(startFoldingPattern, finishFoldingPattern, SyntaxHighlighter.RegexCompiledOption)
        End Sub

        Public Sub SetFoldingMarkers(ByVal startFoldingPattern As String, ByVal finishFoldingPattern As String, ByVal options As RegexOptions)
            If startFoldingPattern = finishFoldingPattern Then
                SetFoldingMarkers(startFoldingPattern, options)
                Return
            End If

            For Each range In GetRanges(startFoldingPattern, options)
                tb(range.start.iLine).FoldingStartMarker = startFoldingPattern
            Next

            For Each range In GetRanges(finishFoldingPattern, options)
                tb(range.start.iLine).FoldingEndMarker = startFoldingPattern
            Next

            tb.Invalidate()
        End Sub

        Public Sub SetFoldingMarkers(ByVal foldingPattern As String, ByVal options As RegexOptions)
            For Each range In GetRanges(foldingPattern, options)
                If range.start.iLine > 0 Then tb(range.start.iLine - 1).FoldingEndMarker = foldingPattern
                tb(range.start.iLine).FoldingStartMarker = foldingPattern
            Next

            tb.Invalidate()
        End Sub

        Public Function GetRanges(ByVal regexPattern As String) As IEnumerable(Of Range)
            Return GetRanges(regexPattern, RegexOptions.None)
        End Function

        Public Iterator Function GetRanges(ByVal regexPattern As String, ByVal options As RegexOptions) As IEnumerable(Of Range)
            Dim text As String
            Dim charIndexToPlace As List(Of Place)
            GetText(text, charIndexToPlace)
            Dim regex As Regex = New Regex(regexPattern, options)

            For Each m As Match In regex.Matches(text)
                Dim r As Range = New Range(Me.tb)
                Dim group As Group = m.Groups("range")
                If Not group.Success Then group = m.Groups(0)
                r.start = charIndexToPlace(group.Index)
                r.[end] = charIndexToPlace(group.Index + group.Length)
                Yield r
            Next
        End Function

        Public Iterator Function GetRangesByLines(ByVal regexPattern As String, ByVal options As RegexOptions) As IEnumerable(Of Range)
            Dim regex = New Regex(regexPattern, options)

            For Each r In GetRangesByLines(regex)
                Yield r
            Next
        End Function

        Public Iterator Function GetRangesByLines(ByVal regex As Regex) As IEnumerable(Of Range)
            Normalize()
            Dim fts = TryCast(tb.TextSource, FileTextSource)

            For iLine As Integer = start.iLine To [end].iLine
                Dim isLineLoaded As Boolean = If(fts IsNot Nothing, fts.IsLineLoaded(iLine), True)
                Dim r = New Range(tb, New Place(0, iLine), New Place(tb(iLine).Count, iLine))
                If iLine = start.iLine OrElse iLine = [end].iLine Then r = r.GetIntersectionWith(Me)

                For Each foundRange In r.GetRanges(regex)
                    Yield foundRange
                Next

                If Not isLineLoaded Then fts.UnloadLine(iLine)
            Next
        End Function

        Public Iterator Function GetRangesByLinesReversed(ByVal regexPattern As String, ByVal options As RegexOptions) As IEnumerable(Of Range)
            Normalize()
            Dim regex As Regex = New Regex(regexPattern, options)
            Dim fts = TryCast(tb.TextSource, FileTextSource)

            For iLine As Integer = [end].iLine To start.iLine
                Dim isLineLoaded As Boolean = If(fts IsNot Nothing, fts.IsLineLoaded(iLine), True)
                Dim r = New Range(tb, New Place(0, iLine), New Place(tb(iLine).Count, iLine))
                If iLine = start.iLine OrElse iLine = [end].iLine Then r = r.GetIntersectionWith(Me)
                Dim list = New List(Of Range)()

                For Each foundRange In r.GetRanges(regex)
                    list.Add(foundRange)
                Next

                For i As Integer = list.Count - 1 To 0
                    Yield list(i)
                Next

                If Not isLineLoaded Then fts.UnloadLine(iLine)
            Next
        End Function

        Public Iterator Function GetRanges(ByVal regex As Regex) As IEnumerable(Of Range)
            Dim text As String
            Dim charIndexToPlace As List(Of Place)
            GetText(text, charIndexToPlace)

            For Each m As Match In regex.Matches(text)
                Dim r As Range = New Range(Me.tb)
                Dim group As Group = m.Groups("range")
                If Not group.Success Then group = m.Groups(0)
                r.start = charIndexToPlace(group.Index)
                r.[end] = charIndexToPlace(group.Index + group.Length)
                Yield r
            Next
        End Function

        Public Sub ClearStyle(ParamArray styles As Style())
            Try
                ClearStyle(tb.GetStyleIndexMask(styles))
            Catch
            End Try
        End Sub

        Public Sub ClearStyle(ByVal styleIndex As StyleIndex)
            Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
            Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
            Dim fromChar As Integer = FromX
            Dim toChar As Integer = ToX
            If fromLine < 0 Then Return

            For y As Integer = fromLine To toLine
                Dim fromX As Integer = If(y = fromLine, fromChar, 0)
                Dim toX As Integer = If(y = toLine, Math.Min(toChar - 1, tb(y).Count - 1), tb(y).Count - 1)

                For x As Integer = fromX To toX
                    Dim c As Char = tb(y)(x)
                    c.style = c.style And Not styleIndex
                    tb(y)(x) = c
                Next
            Next

            tb.Invalidate()
        End Sub

        Public Sub ClearFoldingMarkers()
            Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
            Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
            If fromLine < 0 Then Return

            For y As Integer = fromLine To toLine
                tb(y).ClearFoldingMarkers()
            Next

            tb.Invalidate()
        End Sub

        Private Sub OnSelectionChanged()
            cachedTextVersion = -1
            cachedText = Nothing
            cachedCharIndexToPlace = Nothing

            If tb.Selection = Me Then
                If updating = 0 Then tb.OnSelectionChanged()
            End If
        End Sub

        Public Sub BeginUpdate()
            updating += 1
        End Sub

        Public Sub EndUpdate()
            updating -= 1
            If updating = 0 Then OnSelectionChanged()
        End Sub

        Public Overrides Function ToString() As String
            Return "Start: " & start & " End: " + [end]
        End Function

        Public Sub Normalize()
            If start > [end] Then Inverse()
        End Sub

        Public Sub Inverse()
            Dim temp = start
            start = [end]
            [end] = temp
        End Sub

        Public Sub Expand()
            Normalize()
            start = New Place(0, start.iLine)
            [end] = New Place(tb.GetLineLength([end].iLine), [end].iLine)
        End Sub

        Private Iterator Function GetEnumerator() As IEnumerator(Of Place)
            If columnSelectionMode Then

                For Each p In GetEnumerator_ColumnSelectionMode()
                    Yield p
                Next

                Return
            End If

            Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
            Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
            Dim fromChar As Integer = FromX
            Dim toChar As Integer = ToX
            If fromLine < 0 Then Return

            For y As Integer = fromLine To toLine
                Dim fromX As Integer = If(y = fromLine, fromChar, 0)
                Dim toX As Integer = If(y = toLine, Math.Min(toChar - 1, tb(y).Count - 1), tb(y).Count - 1)

                For x As Integer = fromX To toX
                    Yield New Place(x, y)
                Next
            Next
        End Function

        Private Function GetEnumerator() As System.Collections.IEnumerator
            Return (TryCast(Me, IEnumerable(Of Place))).GetEnumerator()
        End Function

        Public ReadOnly Iterator Property Chars As IEnumerable(Of Char)
            Get

                If columnSelectionMode Then

                    For Each p In GetEnumerator_ColumnSelectionMode()
                        Yield tb(p)
                    Next

                    Return
                End If

                Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
                Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
                Dim fromChar As Integer = FromX
                Dim toChar As Integer = ToX
                If fromLine < 0 Then Return

                For y As Integer = fromLine To toLine
                    Dim fromX As Integer = If(y = fromLine, fromChar, 0)
                    Dim toX As Integer = If(y = toLine, Math.Min(toChar - 1, tb(y).Count - 1), tb(y).Count - 1)
                    Dim line = tb(y)

                    For x As Integer = fromX To toX
                        Yield line(x)
                    Next
                Next
            End Get
        End Property

        Public Function GetFragment(ByVal allowedSymbolsPattern As String) As Range
            Return GetFragment(allowedSymbolsPattern, RegexOptions.None)
        End Function

        Public Function GetFragment(ByVal style As Style, ByVal allowLineBreaks As Boolean) As Range
            Dim mask = tb.GetStyleIndexMask(New Style() {style})
            Dim r As Range = New Range(tb)
            r.start = start

            While r.GoLeftThroughFolded()
                If Not allowLineBreaks AndAlso r.CharAfterStart = vbLf Then Exit While

                If r.start.iChar < tb.GetLineLength(r.start.iLine) Then

                    If (tb(r.start).style And mask) = 0 Then
                        r.GoRightThroughFolded()
                        Exit While
                    End If
                End If
            End While

            Dim startFragment As Place = r.start
            r.start = start

            Do
                If Not allowLineBreaks AndAlso r.CharAfterStart = vbLf Then Exit Do

                If r.start.iChar < tb.GetLineLength(r.start.iLine) Then
                    If (tb(r.start).style And mask) = 0 Then Exit Do
                End If
            Loop While r.GoRightThroughFolded()

            Dim endFragment As Place = r.start
            Return New Range(tb, startFragment, endFragment)
        End Function

        Public Function GetFragment(ByVal allowedSymbolsPattern As String, ByVal options As RegexOptions) As Range
            Dim r As Range = New Range(tb)
            r.start = start
            Dim regex As Regex = New Regex(allowedSymbolsPattern, options)

            While r.GoLeftThroughFolded()

                If Not regex.IsMatch(r.CharAfterStart.ToString()) Then
                    r.GoRightThroughFolded()
                    Exit While
                End If
            End While

            Dim startFragment As Place = r.start
            r.start = start

            Do
                If Not regex.IsMatch(r.CharAfterStart.ToString()) Then Exit Do
            Loop While r.GoRightThroughFolded()

            Dim endFragment As Place = r.start
            Return New Range(tb, startFragment, endFragment)
        End Function

        Private Function IsIdentifierChar(ByVal c As Char) As Boolean
            Return Char.IsLetterOrDigit(c) OrElse c = "_"c
        End Function

        Private Function IsSpaceChar(ByVal c As Char) As Boolean
            Return c = " "c OrElse c = vbTab
        End Function

        Public Sub GoWordLeft(ByVal shift As Boolean)
            columnSelectionMode = False

            If Not shift AndAlso start > [end] Then
                start = [end]
                Return
            End If

            Dim range As Range = Me.Clone()
            Dim wasSpace As Boolean = False

            While IsSpaceChar(range.CharBeforeStart)
                wasSpace = True
                range.GoLeft(shift)
            End While

            Dim wasIdentifier As Boolean = False

            While IsIdentifierChar(range.CharBeforeStart)
                wasIdentifier = True
                range.GoLeft(shift)
            End While

            If Not wasIdentifier AndAlso (Not wasSpace OrElse range.CharBeforeStart <> vbLf) Then range.GoLeft(shift)
            Me.start = range.start
            Me.[end] = range.[end]
            If tb.LineInfos(start.iLine).VisibleState <> VisibleState.Visible Then GoRight(shift)
        End Sub

        Public Sub GoWordRight(ByVal shift As Boolean, ByVal Optional goToStartOfNextWord As Boolean = False)
            columnSelectionMode = False

            If Not shift AndAlso start < [end] Then
                start = [end]
                Return
            End If

            Dim range As Range = Me.Clone()
            Dim wasNewLine As Boolean = False

            If range.CharAfterStart = vbLf Then
                range.GoRight(shift)
                wasNewLine = True
            End If

            Dim wasSpace As Boolean = False

            While IsSpaceChar(range.CharAfterStart)
                wasSpace = True
                range.GoRight(shift)
            End While

            If Not ((wasSpace OrElse wasNewLine) AndAlso goToStartOfNextWord) Then
                Dim wasIdentifier As Boolean = False

                While IsIdentifierChar(range.CharAfterStart)
                    wasIdentifier = True
                    range.GoRight(shift)
                End While

                If Not wasIdentifier Then range.GoRight(shift)

                If goToStartOfNextWord AndAlso Not wasSpace Then

                    While IsSpaceChar(range.CharAfterStart)
                        range.GoRight(shift)
                    End While
                End If
            End If

            Me.start = range.start
            Me.[end] = range.[end]
            If tb.LineInfos(start.iLine).VisibleState <> VisibleState.Visible Then GoLeft(shift)
        End Sub

        Friend Sub GoFirst(ByVal shift As Boolean)
            columnSelectionMode = False
            start = New Place(0, 0)
            If tb.LineInfos(start.iLine).VisibleState <> VisibleState.Visible Then tb.ExpandBlock(start.iLine)
            If Not shift Then [end] = start
            OnSelectionChanged()
        End Sub

        Friend Sub GoLast(ByVal shift As Boolean)
            columnSelectionMode = False
            start = New Place(tb(tb.LinesCount - 1).Count, tb.LinesCount - 1)
            If tb.LineInfos(start.iLine).VisibleState <> VisibleState.Visible Then tb.ExpandBlock(start.iLine)
            If Not shift Then [end] = start
            OnSelectionChanged()
        End Sub

        Public Shared Function ToStyleIndex(ByVal i As Integer) As StyleIndex
            ''' Cannot convert ReturnStatementSyntax, System.ArgumentOutOfRangeException: Exception of type 'System.ArgumentOutOfRangeException' was thrown.
            ''' Parameter name: op
            ''' Actual value was LeftShiftExpression.
            '''    at ICSharpCode.CodeConverter.Util.VBUtil.GetExpressionOperatorTokenKind(SyntaxKind op)
            '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node)
            '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
            '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
            '''    at ICSharpCode.CodeConverter.VB.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
            '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.VisitBinaryExpression(BinaryExpressionSyntax node)
            '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
            '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitParenthesizedExpression(ParenthesizedExpressionSyntax node)
            '''    at Microsoft.CodeAnalysis.CSharp.Syntax.ParenthesizedExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
            '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
            '''    at ICSharpCode.CodeConverter.VB.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
            '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.VisitParenthesizedExpression(ParenthesizedExpressionSyntax node)
            '''    at Microsoft.CodeAnalysis.CSharp.Syntax.ParenthesizedExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
            '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitCastExpression(CastExpressionSyntax node)
            '''    at Microsoft.CodeAnalysis.CSharp.Syntax.CastExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
            '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
            '''    at ICSharpCode.CodeConverter.VB.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
            '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.VisitCastExpression(CastExpressionSyntax node)
            '''    at Microsoft.CodeAnalysis.CSharp.Syntax.CastExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
            '''    at ICSharpCode.CodeConverter.VB.MethodBodyVisitor.VisitReturnStatement(ReturnStatementSyntax node)
            '''    at Microsoft.CodeAnalysis.CSharp.Syntax.ReturnStatementSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
            '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
            '''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
            '''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)
            ''' 
            ''' Input: 
            '''             return (StyleIndex)(1 << i);

            ''' 
        End Function

        Public ReadOnly Property Bounds As RangeRect
            Get
                Dim minX As Integer = Math.Min(start.iChar, [end].iChar)
                Dim minY As Integer = Math.Min(start.iLine, [end].iLine)
                Dim maxX As Integer = Math.Max(start.iChar, [end].iChar)
                Dim maxY As Integer = Math.Max(start.iLine, [end].iLine)
                Return New RangeRect(minY, minX, maxY, maxX)
            End Get
        End Property

        Public Iterator Function GetSubRanges(ByVal includeEmpty As Boolean) As IEnumerable(Of Range)
            If Not columnSelectionMode Then
                Yield Me
                Return
            End If

            Dim rect = Bounds

            For y As Integer = rect.iStartLine To rect.iEndLine
                If rect.iStartChar > tb(y).Count AndAlso Not includeEmpty Then Continue For
                Dim r = New Range(tb, rect.iStartChar, y, Math.Min(rect.iEndChar, tb(y).Count), y)
                Yield r
            Next
        End Function

        Public Property [ReadOnly] As Boolean
            Get
                If tb.[ReadOnly] Then Return True
                Dim readonlyStyle As ReadOnlyStyle = Nothing

                For Each style In tb.Styles

                    If TypeOf style Is ReadOnlyStyle Then
                        readonlyStyle = CType(style, ReadOnlyStyle)
                        Exit For
                    End If
                Next

                If readonlyStyle IsNot Nothing Then
                    Dim si = ToStyleIndex(tb.GetStyleIndex(readonlyStyle))

                    If IsEmpty Then
                        Dim line = tb(start.iLine)

                        If columnSelectionMode Then

                            For Each sr In GetSubRanges(False)
                                line = tb(sr.start.iLine)

                                If sr.start.iChar < line.Count AndAlso sr.start.iChar > 0 Then
                                    Dim left = line(sr.start.iChar - 1)
                                    Dim right = line(sr.start.iChar)
                                    If (left.style And si) <> 0 AndAlso (right.style And si) <> 0 Then Return True
                                End If
                            Next
                        ElseIf start.iChar < line.Count AndAlso start.iChar > 0 Then
                            Dim left = line(start.iChar - 1)
                            Dim right = line(start.iChar)
                            If (left.style And si) <> 0 AndAlso (right.style And si) <> 0 Then Return True
                        End If
                    Else

                        For Each c As Char In Chars
                            If (c.style And si) <> 0 Then Return True
                        Next
                    End If
                End If

                Return False
            End Get
            Set(ByVal value As Boolean)
                Dim readonlyStyle As ReadOnlyStyle = Nothing

                For Each style In tb.Styles

                    If TypeOf style Is ReadOnlyStyle Then
                        readonlyStyle = CType(style, ReadOnlyStyle)
                        Exit For
                    End If
                Next

                If readonlyStyle Is Nothing Then readonlyStyle = New ReadOnlyStyle()

                If value Then
                    SetStyle(readonlyStyle)
                Else
                    ClearStyle(readonlyStyle)
                End If
            End Set
        End Property

        Public Function IsReadOnlyLeftChar() As Boolean
            If tb.[ReadOnly] Then Return True
            Dim r = Clone()
            r.Normalize()
            If r.start.iChar = 0 Then Return False

            If columnSelectionMode Then
                r.GoLeft_ColumnSelectionMode()
            Else
                r.GoLeft(True)
            End If

            Return r.[ReadOnly]
        End Function

        Public Function IsReadOnlyRightChar() As Boolean
            If tb.[ReadOnly] Then Return True
            Dim r = Clone()
            r.Normalize()
            If r.[end].iChar >= tb([end].iLine).Count Then Return False

            If columnSelectionMode Then
                r.GoRight_ColumnSelectionMode()
            Else
                r.GoRight(True)
            End If

            Return r.[ReadOnly]
        End Function

        Public Iterator Function GetPlacesCyclic(ByVal startPlace As Place, ByVal Optional backward As Boolean = False) As IEnumerable(Of Place)
            If backward Then
                Dim r = New Range(Me.tb, startPlace, startPlace)

                While r.GoLeft() AndAlso r.start >= start
                    If r.start.iChar < tb(r.start.iLine).Count Then Yield r.start
                End While

                r = New Range(Me.tb, [end], [end])

                While r.GoLeft() AndAlso r.start >= startPlace
                    If r.start.iChar < tb(r.start.iLine).Count Then Yield r.start
                End While
            Else
                Dim r = New Range(Me.tb, startPlace, startPlace)

                If startPlace < [end] Then

                    Do
                        If r.start.iChar < tb(r.start.iLine).Count Then Yield r.start
                    Loop While r.GoRight()
                End If

                r = New Range(Me.tb, start, start)

                If r.start < startPlace Then

                    Do
                        If r.start.iChar < tb(r.start.iLine).Count Then Yield r.start
                    Loop While r.GoRight() AndAlso r.start < startPlace
                End If
            End If
        End Function

        Private Function GetIntersectionWith_ColumnSelectionMode(ByVal range As Range) As Range
            If range.start.iLine <> range.[end].iLine Then Return New Range(tb, start, start)
            Dim rect = Bounds
            If range.start.iLine < rect.iStartLine OrElse range.start.iLine > rect.iEndLine Then Return New Range(tb, start, start)
            Return New Range(tb, rect.iStartChar, range.start.iLine, rect.iEndChar, range.start.iLine).GetIntersectionWith(range)
        End Function

        Private Function GoRightThroughFolded_ColumnSelectionMode() As Boolean
            Dim boundes = Bounds
            Dim endOfLines = True

            For iLine As Integer = boundes.iStartLine To boundes.iEndLine

                If boundes.iEndChar < tb(iLine).Count Then
                    endOfLines = False
                    Exit For
                End If
            Next

            If endOfLines Then Return False
            Dim start = start
            Dim [end] = [end]
            start.Offset(1, 0)
            [end].Offset(1, 0)
            BeginUpdate()
            start = start
            [end] = [end]
            EndUpdate()
            Return True
        End Function

        Private Iterator Function GetEnumerator_ColumnSelectionMode() As IEnumerable(Of Place)
            Dim bounds = bounds
            If bounds.iStartLine < 0 Then Return

            For y As Integer = bounds.iStartLine To bounds.iEndLine

                For x As Integer = bounds.iStartChar To bounds.iEndChar - 1
                    If x < tb(y).Count Then Yield New Place(x, y)
                Next
            Next
        End Function

        Private ReadOnly Property Text_ColumnSelectionMode As String
            Get
                Dim sb As StringBuilder = New StringBuilder()
                Dim bounds = bounds
                If bounds.iStartLine < 0 Then Return ""

                For y As Integer = bounds.iStartLine To bounds.iEndLine

                    For x As Integer = bounds.iStartChar To bounds.iEndChar - 1
                        If x < tb(y).Count Then sb.Append(tb(y)(x).c)
                    Next

                    If bounds.iEndLine <> bounds.iStartLine AndAlso y <> bounds.iEndLine Then sb.AppendLine()
                Next

                Return sb.ToString()
            End Get
        End Property

        Private Function Length_ColumnSelectionMode(ByVal withNewLines As Boolean) As Integer
            Dim bounds = bounds
            If bounds.iStartLine < 0 Then Return 0
            Dim cnt As Integer = 0

            For y As Integer = bounds.iStartLine To bounds.iEndLine

                For x As Integer = bounds.iStartChar To bounds.iEndChar - 1
                    If x < tb(y).Count Then cnt += 1
                Next

                If withNewLines AndAlso bounds.iEndLine <> bounds.iStartLine AndAlso y <> bounds.iEndLine Then cnt += Environment.NewLine.Length
            Next

            Return cnt
        End Function

        Friend Sub GoDown_ColumnSelectionMode()
            Dim iLine = tb.FindNextVisibleLine([end].iLine)
            [end] = New Place([end].iChar, iLine)
        End Sub

        Friend Sub GoUp_ColumnSelectionMode()
            Dim iLine = tb.FindPrevVisibleLine([end].iLine)
            [end] = New Place([end].iChar, iLine)
        End Sub

        Friend Sub GoRight_ColumnSelectionMode()
            [end] = New Place([end].iChar + 1, [end].iLine)
        End Sub

        Friend Sub GoLeft_ColumnSelectionMode()
            If [end].iChar > 0 Then [end] = New Place([end].iChar - 1, [end].iLine)
        End Sub

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class

    Public Structure RangeRect
        Public Sub New(ByVal iStartLine As Integer, ByVal iStartChar As Integer, ByVal iEndLine As Integer, ByVal iEndChar As Integer)
            Me.iStartLine = iStartLine
            Me.iStartChar = iStartChar
            Me.iEndLine = iEndLine
            Me.iEndChar = iEndChar
        End Sub

        Public iStartLine As Integer
        Public iStartChar As Integer
        Public iEndLine As Integer
        Public iEndChar As Integer

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Structure
End Namespace