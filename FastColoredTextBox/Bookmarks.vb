Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Text

Namespace FastColoredTextBoxNS
    Public MustInherit Class BaseBookmarks
        Implements ICollection(Of Bookmark), IDisposable

        Public MustOverride Sub Add(ByVal item As Bookmark)
        Public MustOverride Sub Clear()
        Public MustOverride Function Contains(ByVal item As Bookmark) As Boolean
        Public MustOverride Sub CopyTo(ByVal array As Bookmark(), ByVal arrayIndex As Integer)
        Public MustOverride ReadOnly Property Count As Integer
        Public MustOverride ReadOnly Property IsReadOnly As Boolean
        Public MustOverride Function Remove(ByVal item As Bookmark) As Boolean

        Private Function GetEnumerator() As System.Collections.IEnumerator Implements ICollection(Of Bookmark).GetEnumerator
            Return GetEnumerator()
        End Function

        Public MustOverride Sub Dispose()
        Public MustOverride Sub Add(ByVal lineIndex As Integer, ByVal bookmarkName As String)
        Public MustOverride Sub Add(ByVal lineIndex As Integer)
        Public MustOverride Function Contains(ByVal lineIndex As Integer) As Boolean
        Public MustOverride Function Remove(ByVal lineIndex As Integer) As Boolean
        Public MustOverride Function GetBookmark(ByVal i As Integer) As Bookmark
    End Class

    Public Class Bookmarks
        Inherits BaseBookmarks

        Protected tb As FastColoredTextBox
        Protected items As List(Of Bookmark) = New List(Of Bookmark)()
        Protected counter As Integer

        Public Sub New(ByVal tb As FastColoredTextBox)
            Me.tb = tb
            tb.LineInserted += AddressOf tb_LineInserted
            tb.LineRemoved += AddressOf tb_LineRemoved
        End Sub

        Protected Overridable Sub tb_LineRemoved(ByVal sender As Object, ByVal e As LineRemovedEventArgs)
            For i As Integer = 0 To Count - 1

                If items(i).LineIndex >= e.Index Then

                    If items(i).LineIndex >= e.Index + e.Count Then
                        items(i).LineIndex = items(i).LineIndex - e.Count
                        Continue For
                    End If

                    Dim was = e.Index <= 0

                    For Each b In items
                        If b.LineIndex = e.Index - 1 Then was = True
                    Next

                    If was Then
                        items.RemoveAt(i)
                        i -= 1
                    Else
                        items(i).LineIndex = e.Index - 1
                    End If
                End If
            Next
        End Sub

        Protected Overridable Sub tb_LineInserted(ByVal sender As Object, ByVal e As LineInsertedEventArgs)
            For i As Integer = 0 To Count - 1

                If items(i).LineIndex >= e.Index Then
                    items(i).LineIndex = items(i).LineIndex + e.Count
                ElseIf items(i).LineIndex = e.Index - 1 AndAlso e.Count = 1 Then
                    If tb(e.Index - 1).StartSpacesCount = tb(e.Index - 1).Count Then items(i).LineIndex = items(i).LineIndex + e.Count
                End If
            Next
        End Sub

        Public Overrides Sub Dispose()
            tb.LineInserted -= AddressOf tb_LineInserted
            tb.LineRemoved -= AddressOf tb_LineRemoved
        End Sub

        Public Overrides Iterator Function GetEnumerator() As IEnumerator(Of Bookmark)
            For Each item In items
                Yield item
            Next
        End Function

        Public Overrides Sub Add(ByVal lineIndex As Integer, ByVal bookmarkName As String)
            Add(New Bookmark(tb, If(bookmarkName, "Bookmark " & counter), lineIndex))
        End Sub

        Public Overrides Sub Add(ByVal lineIndex As Integer)
            Add(New Bookmark(tb, "Bookmark " & counter, lineIndex))
        End Sub

        Public Overrides Sub Clear()
            items.Clear()
            counter = 0
        End Sub

        Public Overrides Sub Add(ByVal bookmark As Bookmark)
            For Each bm In items
                If bm.LineIndex = bookmark.LineIndex Then Return
            Next

            items.Add(bookmark)
            counter += 1
            tb.Invalidate()
        End Sub

        Public Overrides Function Contains(ByVal item As Bookmark) As Boolean
            Return items.Contains(item)
        End Function

        Public Overrides Function Contains(ByVal lineIndex As Integer) As Boolean
            For Each item In items
                If item.LineIndex = lineIndex Then Return True
            Next

            Return False
        End Function

        Public Overrides Sub CopyTo(ByVal array As Bookmark(), ByVal arrayIndex As Integer)
            items.CopyTo(array, arrayIndex)
        End Sub

        Public Overrides ReadOnly Property Count As Integer
            Get
                Return items.Count
            End Get
        End Property

        Public Overrides ReadOnly Property IsReadOnly As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides Function Remove(ByVal item As Bookmark) As Boolean
            tb.Invalidate()
            Return items.Remove(item)
        End Function

        Public Overrides Function Remove(ByVal lineIndex As Integer) As Boolean
            Dim was As Boolean = False

            For i As Integer = 0 To Count - 1

                If items(i).LineIndex = lineIndex Then
                    items.RemoveAt(i)
                    i -= 1
                    was = True
                End If
            Next

            tb.Invalidate()
            Return was
        End Function

        Public Overrides Function GetBookmark(ByVal i As Integer) As Bookmark
            Return items(i)
        End Function
    End Class

    Public Class Bookmark
        Public Property TB As FastColoredTextBox
        Public Property Name As String
        Public Property LineIndex As Integer
        Public Property Color As Color

        Public Overridable Sub DoVisible()
            TB.Selection.Start = New Place(0, LineIndex)
            TB.DoRangeVisible(TB.Selection, True)
            TB.Invalidate()
        End Sub

        Public Sub New(ByVal tb As FastColoredTextBox, ByVal name As String, ByVal lineIndex As Integer)
            Me.TB = tb
            Me.Name = name
            Me.LineIndex = lineIndex
            Color = tb.BookmarkColor
        End Sub

        Public Overridable Sub Paint(ByVal gr As Graphics, ByVal lineRect As Rectangle)
            Dim size = TB.CharHeight - 1

            Using brush = New LinearGradientBrush(New Rectangle(0, lineRect.Top, size, size), Color.White, Color, 45)
                gr.FillEllipse(brush, 0, lineRect.Top, size, size)
            End Using

            Using pen = New Pen(Color)
                gr.DrawEllipse(pen, 0, lineRect.Top, size, size)
            End Using
        End Sub
    End Class
End Namespace