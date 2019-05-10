Imports System
Imports System.Collections.Generic

Namespace FastColoredTextBoxNS
    Public Class InsertCharCommand
        Inherits UndoableCommand

        Public c As Char
        Private deletedChar As Char = vbNullChar

        Public Sub New(ByVal ts As TextSource, ByVal c As Char)
            MyBase.New(ts)
            Me.c = c
        End Sub

        Public Overrides Sub Undo()
            ts.OnTextChanging()

            Select Case c
                Case vbLf
                    MergeLines(sel.Start.iLine, ts)
                Case vbCr
                Case vbBack
                    ts.CurrentTB.Selection.Start = lastSel.Start
                    Dim cc As Char = vbNullChar

                    If deletedChar <> vbNullChar Then
                        ts.CurrentTB.ExpandBlock(ts.CurrentTB.Selection.Start.iLine)
                        InsertChar(deletedChar, cc, ts)
                    End If

                Case vbTab
                    ts.CurrentTB.ExpandBlock(sel.Start.iLine)

                    For i As Integer = sel.FromX To lastSel.FromX - 1
                        ts(sel.Start.iLine).RemoveAt(sel.Start.iChar)
                    Next

                    ts.CurrentTB.Selection.Start = sel.Start
                Case Else
                    ts.CurrentTB.ExpandBlock(sel.Start.iLine)
                    ts(sel.Start.iLine).RemoveAt(sel.Start.iChar)
                    ts.CurrentTB.Selection.Start = sel.Start
            End Select

            ts.NeedRecalc(New TextSource.TextChangedEventArgs(sel.Start.iLine, sel.Start.iLine))
            MyBase.Undo()
        End Sub

        Public Overrides Sub Execute()
            ts.CurrentTB.ExpandBlock(ts.CurrentTB.Selection.Start.iLine)
            Dim s As String = c.ToString()
            ts.OnTextChanging(s)
            If s.Length = 1 Then c = s(0)
            If String.IsNullOrEmpty(s) Then Throw New ArgumentOutOfRangeException()
            If ts.Count = 0 Then InsertLine(ts)
            InsertChar(c, deletedChar, ts)
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(ts.CurrentTB.Selection.Start.iLine, ts.CurrentTB.Selection.Start.iLine))
            MyBase.Execute()
        End Sub

        Friend Shared Sub InsertChar(ByVal c As Char, ByRef deletedChar As Char, ByVal ts As TextSource)
            Dim tb = ts.CurrentTB

            Select Case c
                Case vbLf
                    If Not ts.CurrentTB.AllowInsertRemoveLines Then Throw New ArgumentOutOfRangeException("Cant insert this char in ColumnRange mode")
                    If ts.Count = 0 Then InsertLine(ts)
                    InsertLine(ts)
                Case vbCr
                Case vbBack
                    If tb.Selection.Start.iChar = 0 AndAlso tb.Selection.Start.iLine = 0 Then Return

                    If tb.Selection.Start.iChar = 0 Then
                        If Not ts.CurrentTB.AllowInsertRemoveLines Then Throw New ArgumentOutOfRangeException("Cant insert this char in ColumnRange mode")
                        If tb.LineInfos(tb.Selection.Start.iLine - 1).VisibleState <> VisibleState.Visible Then tb.ExpandBlock(tb.Selection.Start.iLine - 1)
                        deletedChar = vbLf
                        MergeLines(tb.Selection.Start.iLine - 1, ts)
                    Else
                        deletedChar = ts(tb.Selection.Start.iLine)(tb.Selection.Start.iChar - 1).c
                        ts(tb.Selection.Start.iLine).RemoveAt(tb.Selection.Start.iChar - 1)
                        tb.Selection.Start = New Place(tb.Selection.Start.iChar - 1, tb.Selection.Start.iLine)
                    End If

                Case vbTab
                    Dim spaceCountNextTabStop As Integer = tb.TabLength - (tb.Selection.Start.iChar Mod tb.TabLength)
                    If spaceCountNextTabStop = 0 Then spaceCountNextTabStop = tb.TabLength

                    For i As Integer = 0 To spaceCountNextTabStop - 1
                        ts(tb.Selection.Start.iLine).Insert(tb.Selection.Start.iChar, New Char(" "c))
                    Next

                    tb.Selection.Start = New Place(tb.Selection.Start.iChar + spaceCountNextTabStop, tb.Selection.Start.iLine)
                Case Else
                    ts(tb.Selection.Start.iLine).Insert(tb.Selection.Start.iChar, New Char(c))
                    tb.Selection.Start = New Place(tb.Selection.Start.iChar + 1, tb.Selection.Start.iLine)
            End Select
        End Sub

        Friend Shared Sub InsertLine(ByVal ts As TextSource)
            Dim tb = ts.CurrentTB
            If Not tb.Multiline AndAlso tb.LinesCount > 0 Then Return

            If ts.Count = 0 Then
                ts.InsertLine(0, ts.CreateLine())
            Else
                BreakLines(tb.Selection.Start.iLine, tb.Selection.Start.iChar, ts)
            End If

            tb.Selection.Start = New Place(0, tb.Selection.Start.iLine + 1)
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(0, 1))
        End Sub

        Friend Shared Sub MergeLines(ByVal i As Integer, ByVal ts As TextSource)
            Dim tb = ts.CurrentTB
            If i + 1 >= ts.Count Then Return
            tb.ExpandBlock(i)
            tb.ExpandBlock(i + 1)
            Dim pos As Integer = ts(i).Count

            If ts(i + 1).Count = 0 Then
                ts.RemoveLine(i + 1)
            Else
                ts(i).AddRange(ts(i + 1))
                ts.RemoveLine(i + 1)
            End If

            tb.Selection.Start = New Place(pos, i)
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(0, 1))
        End Sub

        Friend Shared Sub BreakLines(ByVal iLine As Integer, ByVal pos As Integer, ByVal ts As TextSource)
            Dim newLine As Line = ts.CreateLine()

            For i As Integer = pos To ts(iLine).Count - 1
                newLine.Add(ts(iLine)(i))
            Next

            ts(iLine).RemoveRange(pos, ts(iLine).Count - pos)
            ts.InsertLine(iLine + 1, newLine)
        End Sub

        Public Overrides Function Clone() As UndoableCommand
            Return New InsertCharCommand(ts, c)
        End Function
    End Class

    Public Class InsertTextCommand
        Inherits UndoableCommand

        Public InsertedText As String

        Public Sub New(ByVal ts As TextSource, ByVal insertedText As String)
            MyBase.New(ts)
            Me.InsertedText = insertedText
        End Sub

        Public Overrides Sub Undo()
            ts.CurrentTB.Selection.Start = sel.Start
            ts.CurrentTB.Selection.[End] = lastSel.Start
            ts.OnTextChanging()
            ClearSelectedCommand.ClearSelected(ts)
            MyBase.Undo()
        End Sub

        Public Overrides Sub Execute()
            ts.OnTextChanging(InsertedText)
            InsertText(InsertedText, ts)
            MyBase.Execute()
        End Sub

        Friend Shared Sub InsertText(ByVal insertedText As String, ByVal ts As TextSource)
            Dim tb = ts.CurrentTB

            Try
                tb.Selection.BeginUpdate()
                Dim cc As Char = vbNullChar

                If ts.Count = 0 Then
                    InsertCharCommand.InsertLine(ts)
                    tb.Selection.Start = Place.Empty
                End If

                tb.ExpandBlock(tb.Selection.Start.iLine)
                Dim len = insertedText.Length

                For i As Integer = 0 To len - 1
                    Dim c = insertedText(i)

                    If c = vbCr AndAlso (i >= len - 1 OrElse insertedText(i + 1) <> vbLf) Then
                        InsertCharCommand.InsertChar(vbLf, cc, ts)
                    Else
                        InsertCharCommand.InsertChar(c, cc, ts)
                    End If
                Next

                ts.NeedRecalc(New TextSource.TextChangedEventArgs(0, 1))
            Finally
                tb.Selection.EndUpdate()
            End Try
        End Sub

        Public Overrides Function Clone() As UndoableCommand
            Return New InsertTextCommand(ts, InsertedText)
        End Function
    End Class

    Public Class ReplaceTextCommand
        Inherits UndoableCommand

        Private insertedText As String
        Private ranges As List(Of Range)
        Private prevText As List(Of String) = New List(Of String)()

        Public Sub New(ByVal ts As TextSource, ByVal ranges As List(Of Range), ByVal insertedText As String)
            MyBase.New(ts)
            ranges.Sort(Function(r1, r2)
                            If r1.Start.iLine = r2.Start.iLine Then Return r1.Start.iChar.CompareTo(r2.Start.iChar)
                            Return r1.Start.iLine.CompareTo(r2.Start.iLine)
                        End Function)
            Me.ranges = ranges
            Me.insertedText = insertedText
            lastSel = CSharpImpl.__Assign(sel, New RangeInfo(ts.CurrentTB.Selection))
        End Sub

        Public Overrides Sub Undo()
            Dim tb = ts.CurrentTB
            ts.OnTextChanging()
            tb.BeginUpdate()
            tb.Selection.BeginUpdate()

            For i As Integer = 0 To ranges.Count - 1
                tb.Selection.Start = ranges(i).Start

                For j As Integer = 0 To insertedText.Length - 1
                    tb.Selection.GoRight(True)
                Next

                ClearSelected(ts)
                InsertTextCommand.InsertText(prevText(prevText.Count - i - 1), ts)
            Next

            tb.Selection.EndUpdate()
            tb.EndUpdate()
            If ranges.Count > 0 Then ts.OnTextChanged(ranges(0).Start.iLine, ranges(ranges.Count - 1).[End].iLine)
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(0, 1))
        End Sub

        Public Overrides Sub Execute()
            Dim tb = ts.CurrentTB
            prevText.Clear()
            ts.OnTextChanging(insertedText)
            tb.Selection.BeginUpdate()
            tb.BeginUpdate()

            For i As Integer = ranges.Count - 1 To 0
                tb.Selection.Start = ranges(i).Start
                tb.Selection.[End] = ranges(i).[End]
                prevText.Add(tb.Selection.Text)
                ClearSelected(ts)
                If insertedText <> "" Then InsertTextCommand.InsertText(insertedText, ts)
            Next

            If ranges.Count > 0 Then ts.OnTextChanged(ranges(0).Start.iLine, ranges(ranges.Count - 1).[End].iLine)
            tb.EndUpdate()
            tb.Selection.EndUpdate()
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(0, 1))
            lastSel = New RangeInfo(tb.Selection)
        End Sub

        Public Overrides Function Clone() As UndoableCommand
            Return New ReplaceTextCommand(ts, New List(Of Range)(ranges), insertedText)
        End Function

        Friend Shared Sub ClearSelected(ByVal ts As TextSource)
            Dim tb = ts.CurrentTB
            tb.Selection.Normalize()
            Dim start As Place = tb.Selection.Start
            Dim [end] As Place = tb.Selection.[End]
            Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
            Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
            Dim fromChar As Integer = tb.Selection.FromX
            Dim toChar As Integer = tb.Selection.ToX
            If fromLine < 0 Then Return

            If fromLine = toLine Then
                ts(fromLine).RemoveRange(fromChar, toChar - fromChar)
            Else
                ts(fromLine).RemoveRange(fromChar, ts(fromLine).Count - fromChar)
                ts(toLine).RemoveRange(0, toChar)
                ts.RemoveLine(fromLine + 1, toLine - fromLine - 1)
                InsertCharCommand.MergeLines(fromLine, ts)
            End If
        End Sub

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class

    Public Class ClearSelectedCommand
        Inherits UndoableCommand

        Private deletedText As String

        Public Sub New(ByVal ts As TextSource)
            MyBase.New(ts)
        End Sub

        Public Overrides Sub Undo()
            ts.CurrentTB.Selection.Start = New Place(sel.FromX, Math.Min(sel.Start.iLine, sel.[End].iLine))
            ts.OnTextChanging()
            InsertTextCommand.InsertText(deletedText, ts)
            ts.OnTextChanged(sel.Start.iLine, sel.[End].iLine)
            ts.CurrentTB.Selection.Start = sel.Start
            ts.CurrentTB.Selection.[End] = sel.[End]
        End Sub

        Public Overrides Sub Execute()
            Dim tb = ts.CurrentTB
            Dim temp As String = Nothing
            ts.OnTextChanging(temp)
            If temp = "" Then Throw New ArgumentOutOfRangeException()
            deletedText = tb.Selection.Text
            ClearSelected(ts)
            lastSel = New RangeInfo(tb.Selection)
            ts.OnTextChanged(lastSel.Start.iLine, lastSel.Start.iLine)
        End Sub

        Friend Shared Sub ClearSelected(ByVal ts As TextSource)
            Dim tb = ts.CurrentTB
            Dim start As Place = tb.Selection.Start
            Dim [end] As Place = tb.Selection.[End]
            Dim fromLine As Integer = Math.Min([end].iLine, start.iLine)
            Dim toLine As Integer = Math.Max([end].iLine, start.iLine)
            Dim fromChar As Integer = tb.Selection.FromX
            Dim toChar As Integer = tb.Selection.ToX
            If fromLine < 0 Then Return

            If fromLine = toLine Then
                ts(fromLine).RemoveRange(fromChar, toChar - fromChar)
            Else
                ts(fromLine).RemoveRange(fromChar, ts(fromLine).Count - fromChar)
                ts(toLine).RemoveRange(0, toChar)
                ts.RemoveLine(fromLine + 1, toLine - fromLine - 1)
                InsertCharCommand.MergeLines(fromLine, ts)
            End If

            tb.Selection.Start = New Place(fromChar, fromLine)
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(fromLine, toLine))
        End Sub

        Public Overrides Function Clone() As UndoableCommand
            Return New ClearSelectedCommand(ts)
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class

    Public Class ReplaceMultipleTextCommand
        Inherits UndoableCommand

        Private ranges As List(Of ReplaceRange)
        Private prevText As List(Of String) = New List(Of String)()

        Public Class ReplaceRange
            Public Property ReplacedRange As Range
            Public Property ReplaceText As String

            Private Class CSharpImpl
                <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
                Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                    target = value
                    Return value
                End Function
            End Class
        End Class

        Public Sub New(ByVal ts As TextSource, ByVal ranges As List(Of ReplaceRange))
            MyBase.New(ts)
            ranges.Sort(Function(r1, r2)
                            If r1.ReplacedRange.Start.iLine = r2.ReplacedRange.Start.iLine Then Return r1.ReplacedRange.Start.iChar.CompareTo(r2.ReplacedRange.Start.iChar)
                            Return r1.ReplacedRange.Start.iLine.CompareTo(r2.ReplacedRange.Start.iLine)
                        End Function)
            Me.ranges = ranges
            lastSel = CSharpImpl.__Assign(sel, New RangeInfo(ts.CurrentTB.Selection))
        End Sub

        Public Overrides Sub Undo()
            Dim tb = ts.CurrentTB
            ts.OnTextChanging()
            tb.Selection.BeginUpdate()

            For i As Integer = 0 To ranges.Count - 1
                tb.Selection.Start = ranges(i).ReplacedRange.Start

                For j As Integer = 0 To ranges(i).ReplaceText.Length - 1
                    tb.Selection.GoRight(True)
                Next

                ClearSelectedCommand.ClearSelected(ts)
                Dim prevTextIndex = ranges.Count - 1 - i
                InsertTextCommand.InsertText(prevText(prevTextIndex), ts)
                ts.OnTextChanged(ranges(i).ReplacedRange.Start.iLine, ranges(i).ReplacedRange.Start.iLine)
            Next

            tb.Selection.EndUpdate()
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(0, 1))
        End Sub

        Public Overrides Sub Execute()
            Dim tb = ts.CurrentTB
            prevText.Clear()
            ts.OnTextChanging()
            tb.Selection.BeginUpdate()

            For i As Integer = ranges.Count - 1 To 0
                tb.Selection.Start = ranges(i).ReplacedRange.Start
                tb.Selection.[End] = ranges(i).ReplacedRange.[End]
                prevText.Add(tb.Selection.Text)
                ClearSelectedCommand.ClearSelected(ts)
                InsertTextCommand.InsertText(ranges(i).ReplaceText, ts)
                ts.OnTextChanged(ranges(i).ReplacedRange.Start.iLine, ranges(i).ReplacedRange.[End].iLine)
            Next

            tb.Selection.EndUpdate()
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(0, 1))
            lastSel = New RangeInfo(tb.Selection)
        End Sub

        Public Overrides Function Clone() As UndoableCommand
            Return New ReplaceMultipleTextCommand(ts, New List(Of ReplaceRange)(ranges))
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class

    Public Class RemoveLinesCommand
        Inherits UndoableCommand

        Private iLines As List(Of Integer)
        Private prevText As List(Of String) = New List(Of String)()

        Public Sub New(ByVal ts As TextSource, ByVal iLines As List(Of Integer))
            MyBase.New(ts)
            iLines.Sort()
            Me.iLines = iLines
            lastSel = CSharpImpl.__Assign(sel, New RangeInfo(ts.CurrentTB.Selection))
        End Sub

        Public Overrides Sub Undo()
            Dim tb = ts.CurrentTB
            ts.OnTextChanging()
            tb.Selection.BeginUpdate()

            For i As Integer = 0 To iLines.Count - 1
                Dim iLine = iLines(i)

                If iLine < ts.Count Then
                    tb.Selection.Start = New Place(0, iLine)
                Else
                    tb.Selection.Start = New Place(ts(ts.Count - 1).Count, ts.Count - 1)
                End If

                InsertCharCommand.InsertLine(ts)
                tb.Selection.Start = New Place(0, iLine)
                Dim text = prevText(prevText.Count - i - 1)
                InsertTextCommand.InsertText(text, ts)
                ts(iLine).IsChanged = True

                If iLine < ts.Count - 1 Then
                    ts(iLine + 1).IsChanged = True
                Else
                    ts(iLine - 1).IsChanged = True
                End If

                If text.Trim() <> String.Empty Then ts.OnTextChanged(iLine, iLine)
            Next

            tb.Selection.EndUpdate()
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(0, 1))
        End Sub

        Public Overrides Sub Execute()
            Dim tb = ts.CurrentTB
            prevText.Clear()
            ts.OnTextChanging()
            tb.Selection.BeginUpdate()

            For i As Integer = iLines.Count - 1 To 0
                Dim iLine = iLines(i)
                prevText.Add(ts(iLine).Text)
                ts.RemoveLine(iLine)
            Next

            tb.Selection.Start = New Place(0, 0)
            tb.Selection.EndUpdate()
            ts.NeedRecalc(New TextSource.TextChangedEventArgs(0, 1))
            lastSel = New RangeInfo(tb.Selection)
        End Sub

        Public Overrides Function Clone() As UndoableCommand
            Return New RemoveLinesCommand(ts, New List(Of Integer)(iLines))
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class

    Public Class MultiRangeCommand
        Inherits UndoableCommand

        Private cmd As UndoableCommand
        Private range As Range
        Private commandsByRanges As List(Of UndoableCommand) = New List(Of UndoableCommand)()

        Public Sub New(ByVal command As UndoableCommand)
            MyBase.New(command.ts)
            Me.cmd = command
            range = ts.CurrentTB.Selection.Clone()
        End Sub

        Public Overrides Sub Execute()
            commandsByRanges.Clear()
            Dim prevSelection = range.Clone()
            Dim iChar = -1
            Dim iStartLine = prevSelection.Start.iLine
            Dim iEndLine = prevSelection.[End].iLine
            ts.CurrentTB.Selection.ColumnSelectionMode = False
            ts.CurrentTB.Selection.BeginUpdate()
            ts.CurrentTB.BeginUpdate()
            ts.CurrentTB.AllowInsertRemoveLines = False

            Try

                If TypeOf cmd Is InsertTextCommand Then
                    ExecuteInsertTextCommand(iChar, (TryCast(cmd, InsertTextCommand)).InsertedText)
                ElseIf TypeOf cmd Is InsertCharCommand AndAlso (TryCast(cmd, InsertCharCommand)).c <> vbNullChar AndAlso (TryCast(cmd, InsertCharCommand)).c <> vbBack Then
                    ExecuteInsertTextCommand(iChar, (TryCast(cmd, InsertCharCommand)).c.ToString())
                Else
                    ExecuteCommand(iChar)
                End If

            Catch __unusedArgumentOutOfRangeException1__ As ArgumentOutOfRangeException
            Finally
                ts.CurrentTB.AllowInsertRemoveLines = True
                ts.CurrentTB.EndUpdate()
                ts.CurrentTB.Selection = range

                If iChar >= 0 Then
                    ts.CurrentTB.Selection.Start = New Place(iChar, iStartLine)
                    ts.CurrentTB.Selection.[End] = New Place(iChar, iEndLine)
                End If

                ts.CurrentTB.Selection.ColumnSelectionMode = True
                ts.CurrentTB.Selection.EndUpdate()
            End Try
        End Sub

        Private Sub ExecuteInsertTextCommand(ByRef iChar As Integer, ByVal text As String)
            Dim lines = text.Split(vbLf)
            Dim iLine = 0

            For Each r In range.GetSubRanges(True)
                Dim line = ts.CurrentTB(r.Start.iLine)
                Dim lineIsEmpty = r.[End] < r.Start AndAlso line.StartSpacesCount = line.Count

                If Not lineIsEmpty Then
                    Dim insertedText = lines(iLine Mod lines.Length)

                    If r.[End] < r.Start AndAlso insertedText <> "" Then
                        insertedText = New String(" "c, r.Start.iChar - r.[End].iChar) & insertedText
                        r.Start = r.[End]
                    End If

                    ts.CurrentTB.Selection = r
                    Dim c = New InsertTextCommand(ts, insertedText)
                    c.Execute()
                    If ts.CurrentTB.Selection.[End].iChar > iChar Then iChar = ts.CurrentTB.Selection.[End].iChar
                    commandsByRanges.Add(c)
                End If

                iLine += 1
            Next
        End Sub

        Private Sub ExecuteCommand(ByRef iChar As Integer)
            For Each r In range.GetSubRanges(False)
                ts.CurrentTB.Selection = r
                Dim c = cmd.Clone()
                c.Execute()
                If ts.CurrentTB.Selection.[End].iChar > iChar Then iChar = ts.CurrentTB.Selection.[End].iChar
                commandsByRanges.Add(c)
            Next
        End Sub

        Public Overrides Sub Undo()
            ts.CurrentTB.BeginUpdate()
            ts.CurrentTB.Selection.BeginUpdate()

            Try

                For i As Integer = commandsByRanges.Count - 1 To 0
                    commandsByRanges(i).Undo()
                Next

            Finally
                ts.CurrentTB.Selection.EndUpdate()
                ts.CurrentTB.EndUpdate()
            End Try

            ts.CurrentTB.Selection = range.Clone()
            ts.CurrentTB.OnTextChanged(range)
            ts.CurrentTB.OnSelectionChanged()
            ts.CurrentTB.Selection.ColumnSelectionMode = True
        End Sub

        Public Overrides Function Clone() As UndoableCommand
            Throw New NotImplementedException()
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class

    Public Class SelectCommand
        Inherits UndoableCommand

        Public Sub New(ByVal ts As TextSource)
            MyBase.New(ts)
        End Sub

        Public Overrides Sub Execute()
            lastSel = New RangeInfo(ts.CurrentTB.Selection)
        End Sub

        Protected Overrides Sub OnTextChanged(ByVal invert As Boolean)
        End Sub

        Public Overrides Sub Undo()
            ts.CurrentTB.Selection = New Range(ts.CurrentTB, lastSel.Start, lastSel.[End])
        End Sub

        Public Overrides Function Clone() As UndoableCommand
            Dim result = New SelectCommand(ts)
            If lastSel IsNot Nothing Then result.lastSel = New RangeInfo(New Range(ts.CurrentTB, lastSel.Start, lastSel.[End]))
            Return result
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace