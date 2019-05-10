Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Collections
Imports System.Drawing
Imports System.IO

Namespace FastColoredTextBoxNS
    Public Class TextSource
        Implements IList(Of Line), IDisposable

        Protected ReadOnly lines As List(Of Line) = New List(Of Line)()
        Protected linesAccessor As LinesAccessor
        Private lastLineUniqueId As Integer
        Public Property Manager As CommandManager
        Private currentTB As FastColoredTextBox
        Public ReadOnly Styles As Style()
        Public Event LineInserted As EventHandler(Of LineInsertedEventArgs)
        Public Event LineRemoved As EventHandler(Of LineRemovedEventArgs)
        Public Event TextChanged As EventHandler(Of TextChangedEventArgs)
        Public Event RecalcNeeded As EventHandler(Of TextChangedEventArgs)
        Public Event RecalcWordWrap As EventHandler(Of TextChangedEventArgs)
        Public Event TextChanging As EventHandler(Of TextChangingEventArgs)
        Public Event CurrentTBChanged As EventHandler

        Public Property CurrentTB As FastColoredTextBox
            Get
                Return CurrentTB
            End Get
            Set(ByVal value As FastColoredTextBox)
                If currentTB = value Then Return
                currentTB = value
                OnCurrentTBChanged()
            End Set
        End Property

        Public Overridable Sub ClearIsChanged()
            For Each line In lines
                line.IsChanged = False
            Next
        End Sub

        Public Overridable Function CreateLine() As Line
            Return New Line(GenerateUniqueLineId())
        End Function

        Private Sub OnCurrentTBChanged()
            RaiseEvent CurrentTBChanged(Me, EventArgs.Empty)
        End Sub

        Public Property DefaultStyle As TextStyle

        Public Sub New(ByVal currentTB As FastColoredTextBox)
            Me.currentTB = currentTB
            linesAccessor = New LinesAccessor(Me)
            Manager = New CommandManager(Me)

            If [Enum].GetUnderlyingType(GetType(StyleIndex)) = GetType(UInt32) Then
                Styles = New Style(31) {}
            Else
                Styles = New Style(15) {}
            End If

            InitDefaultStyle()
        End Sub

        Public Overridable Sub InitDefaultStyle()
            DefaultStyle = New TextStyle(Nothing, Nothing, FontStyle.Regular)
        End Sub

        Default Public Overridable Property Item(ByVal i As Integer) As Line
            Get
                Return lines(i)
            End Get
            Set(ByVal value As Line)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Overridable Function IsLineLoaded(ByVal iLine As Integer) As Boolean
            Return lines(iLine) IsNot Nothing
        End Function

        Public Overridable Function GetLines() As IList(Of String)
            Return linesAccessor
        End Function

        Public Function GetEnumerator() As IEnumerator(Of Line)
            Return lines.GetEnumerator()
        End Function

        Private Function GetEnumerator() As IEnumerator
            Return (TryCast(lines, IEnumerator))
        End Function

        Public Overridable Function BinarySearch(ByVal item As Line, ByVal comparer As IComparer(Of Line)) As Integer
            Return lines.BinarySearch(item, comparer)
        End Function

        Public Overridable Function GenerateUniqueLineId() As Integer
            Return Math.Min(System.Threading.Interlocked.Increment(lastLineUniqueId), lastLineUniqueId - 1)
        End Function

        Public Overridable Sub InsertLine(ByVal index As Integer, ByVal line As Line)
            lines.Insert(index, line)
            OnLineInserted(index)
        End Sub

        Public Overridable Sub OnLineInserted(ByVal index As Integer)
            OnLineInserted(index, 1)
        End Sub

        Public Overridable Sub OnLineInserted(ByVal index As Integer, ByVal count As Integer)
            RaiseEvent LineInserted(Me, New LineInsertedEventArgs(index, count))
        End Sub

        Public Overridable Sub RemoveLine(ByVal index As Integer)
            RemoveLine(index, 1)
        End Sub

        Public Overridable ReadOnly Property IsNeedBuildRemovedLineIds As Boolean
            Get
                Return LineRemoved IsNot Nothing
            End Get
        End Property

        Public Overridable Sub RemoveLine(ByVal index As Integer, ByVal count As Integer)
            Dim removedLineIds As List(Of Integer) = New List(Of Integer)()

            If count > 0 Then

                If IsNeedBuildRemovedLineIds Then

                    For i As Integer = 0 To count - 1
                        removedLineIds.Add(Me(index + i).UniqueId)
                    Next
                End If
            End If

            lines.RemoveRange(index, count)
            OnLineRemoved(index, count, removedLineIds)
        End Sub

        Public Overridable Sub OnLineRemoved(ByVal index As Integer, ByVal count As Integer, ByVal removedLineIds As List(Of Integer))
            If count > 0 Then
                RaiseEvent LineRemoved(Me, New LineRemovedEventArgs(index, count, removedLineIds))
            End If
        End Sub

        Public Overridable Sub OnTextChanged(ByVal fromLine As Integer, ByVal toLine As Integer)
            RaiseEvent TextChanged(Me, New TextChangedEventArgs(Math.Min(fromLine, toLine), Math.Max(fromLine, toLine)))
        End Sub

        Public Class TextChangedEventArgs
            Inherits EventArgs

            Public iFromLine As Integer
            Public iToLine As Integer

            Public Sub New(ByVal iFromLine As Integer, ByVal iToLine As Integer)
                Me.iFromLine = iFromLine
                Me.iToLine = iToLine
            End Sub
        End Class

        Public Overridable Function IndexOf(ByVal item As Line) As Integer
            Return lines.IndexOf(item)
        End Function

        Public Overridable Sub Insert(ByVal index As Integer, ByVal item As Line)
            InsertLine(index, item)
        End Sub

        Public Overridable Sub RemoveAt(ByVal index As Integer)
            RemoveLine(index)
        End Sub

        Public Overridable Sub Add(ByVal item As Line)
            InsertLine(Count, item)
        End Sub

        Public Overridable Sub Clear()
            RemoveLine(0, Count)
        End Sub

        Public Overridable Function Contains(ByVal item As Line) As Boolean
            Return lines.Contains(item)
        End Function

        Public Overridable Sub CopyTo(ByVal array As Line(), ByVal arrayIndex As Integer)
            lines.CopyTo(array, arrayIndex)
        End Sub

        Public Overridable ReadOnly Property Count As Integer
            Get
                Return lines.Count
            End Get
        End Property

        Public Overridable ReadOnly Property IsReadOnly As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overridable Function Remove(ByVal item As Line) As Boolean
            Dim i As Integer = IndexOf(item)

            If i >= 0 Then
                RemoveLine(i)
                Return True
            Else
                Return False
            End If
        End Function

        Public Overridable Sub NeedRecalc(ByVal args As TextChangedEventArgs)
            RaiseEvent RecalcNeeded(Me, args)
        End Sub

        Public Overridable Sub OnRecalcWordWrap(ByVal args As TextChangedEventArgs)
            RaiseEvent RecalcWordWrap(Me, args)
        End Sub

        Public Overridable Sub OnTextChanging()
            Dim temp As String = Nothing
            OnTextChanging(temp)
        End Sub

        Public Overridable Sub OnTextChanging(ByRef text As String)
            If TextChanging IsNot Nothing Then
                Dim args = New TextChangingEventArgs() With {
                    .InsertingText = text
                }
                TextChanging(Me, args)
                text = args.InsertingText
                If args.Cancel Then text = String.Empty
            End If
        End Sub

        Public Overridable Function GetLineLength(ByVal i As Integer) As Integer
            Return lines(i).Count
        End Function

        Public Overridable Function LineHasFoldingStartMarker(ByVal iLine As Integer) As Boolean
            Return Not String.IsNullOrEmpty(lines(iLine).FoldingStartMarker)
        End Function

        Public Overridable Function LineHasFoldingEndMarker(ByVal iLine As Integer) As Boolean
            Return Not String.IsNullOrEmpty(lines(iLine).FoldingEndMarker)
        End Function

        Public Overridable Sub Dispose()
        End Sub

        Public Overridable Sub SaveToFile(ByVal fileName As String, ByVal enc As Encoding)
            Using sw As StreamWriter = New StreamWriter(fileName, False, enc)

                For i As Integer = 0 To Count - 1 - 1
                    sw.WriteLine(lines(i).Text)
                Next

                sw.Write(lines(Count - 1).Text)
            End Using
        End Sub
    End Class
End Namespace