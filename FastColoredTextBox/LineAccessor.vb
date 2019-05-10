Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace FastColoredTextBoxNS
    Public Class LinesAccessor
        Implements IList(Of String)

        Private ts As IList(Of Line)

        Public Sub New(ByVal ts As IList(Of Line))
            Me.ts = ts
        End Sub

        Public Function IndexOf(ByVal item As String) As Integer
            For i As Integer = 0 To ts.Count - 1
                If ts(i).Text = item Then Return i
            Next

            Return -1
        End Function

        Public Sub Insert(ByVal index As Integer, ByVal item As String)
            Throw New NotImplementedException()
        End Sub

        Public Sub RemoveAt(ByVal index As Integer)
            Throw New NotImplementedException()
        End Sub

        Default Public Property Item(ByVal index As Integer) As String
            Get
                Return ts(index).Text
            End Get
            Set(ByVal value As String)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Sub Add(ByVal item As String)
            Throw New NotImplementedException()
        End Sub

        Public Sub Clear()
            Throw New NotImplementedException()
        End Sub

        Public Function Contains(ByVal item As String) As Boolean
            For i As Integer = 0 To ts.Count - 1
                If ts(i).Text = item Then Return True
            Next

            Return False
        End Function

        Public Sub CopyTo(ByVal array As String(), ByVal arrayIndex As Integer)
            For i As Integer = 0 To ts.Count - 1
                array(i + arrayIndex) = ts(i).Text
            Next
        End Sub

        Public ReadOnly Property Count As Integer
            Get
                Return ts.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean
            Get
                Return True
            End Get
        End Property

        Public Function Remove(ByVal item As String) As Boolean
            Throw New NotImplementedException()
        End Function

        Public Iterator Function GetEnumerator() As IEnumerator(Of String)
            For i As Integer = 0 To ts.Count - 1
                Yield ts(i).Text
            Next
        End Function

        Private Function GetEnumerator() As System.Collections.IEnumerator
            Return GetEnumerator()
        End Function
    End Class
End Namespace