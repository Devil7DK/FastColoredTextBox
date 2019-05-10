Imports System

Namespace FastColoredTextBoxNS
    Public Class LimitedStack(Of T)
        Private items As T()
        Private count As Integer
        Private start As Integer

        Public ReadOnly Property MaxItemCount As Integer
            Get
                Return items.Length
            End Get
        End Property

        Public ReadOnly Property Count As Integer
            Get
                Return Count
            End Get
        End Property

        Public Sub New(ByVal maxItemCount As Integer)
            items = New T(maxItemCount - 1) {}
            count = 0
            start = 0
        End Sub

        Public Function Pop() As T
            If count = 0 Then Throw New Exception("Stack is empty")
            Dim i As Integer = LastIndex
            Dim item As T = items(i)
            items(i) = Nothing
            count -= 1
            Return item
        End Function

        Private ReadOnly Property LastIndex As Integer
            Get
                Return (start + count - 1) Mod items.Length
            End Get
        End Property

        Public Function Peek() As T
            If count = 0 Then Return Nothing
            Return items(LastIndex)
        End Function

        Public Sub Push(ByVal item As T)
            If count = items.Length Then
                start = (start + 1) Mod items.Length
            Else
                count += 1
            End If

            items(LastIndex) = item
        End Sub

        Public Sub Clear()
            items = New T(items.Length - 1) {}
            count = 0
            start = 0
        End Sub

        Public Function ToArray() As T()
            Dim result As T() = New T(count - 1) {}

            For i As Integer = 0 To count - 1
                result(i) = items((start + i) Mod items.Length)
            Next

            Return result
        End Function
    End Class
End Namespace