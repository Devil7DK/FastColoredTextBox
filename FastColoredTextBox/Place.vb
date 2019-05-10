Imports System

Namespace FastColoredTextBoxNS

    ''' <summary>
    ''' Line index and char index
    ''' </summary>
    Public Structure Place
        Implements IEquatable(Of Place)

        Public iChar As Integer

        Public iLine As Integer

        Public Sub New(ByVal iChar As Integer, ByVal iLine As Integer)
            MyBase.New
            Me.iChar = Me.iChar
            Me.iLine = Me.iLine
        End Sub

        Public Sub Offset(ByVal dx As Integer, ByVal dy As Integer)
            Me.iChar = (Me.iChar + dx)
            Me.iLine = (Me.iLine + dy)
        End Sub

        Public Overloads Function Equals(ByVal other As Place) As Boolean
            Return ((Me.iChar = other.iChar) _
                        AndAlso (Me.iLine = other.iLine))
        End Function

        Public Overloads Overrides Function Equals(ByVal obj As Object) As Boolean
            Return ((TypeOf obj Is Place) _
                        AndAlso Me.Equals(CType(obj, Place)))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return (Me.iChar.GetHashCode Or Me.iLine.GetHashCode)
            'The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
        End Function

        Public Overloads Shared Function Operator(ByVal p1 As Place, ByVal p2 As Place) As Boolean
            Return Not p1.Equals(p2)
        End Function

        Public Overloads Shared Function Operator(ByVal p1 As Place, ByVal p2 As Place) As Boolean
            Return p1.Equals(p2)
        End Function

        Public Overloads Shared Function Operator(ByVal p1 As Place, ByVal p2 As Place) As Boolean
            If (p1.iLine < p2.iLine) Then
                Return True
            End If

            If (p1.iLine > p2.iLine) Then
                Return False
            End If

            If (p1.iChar < p2.iChar) Then
                Return True
            End If

            Return False
        End Function

        Public Overloads Shared Function Operator(ByVal p1 As Place, ByVal p2 As Place) As Boolean
            If p1.Equals(p2) Then
                Return True
            End If

            If (p1.iLine < p2.iLine) Then
                Return True
            End If

            If (p1.iLine > p2.iLine) Then
                Return False
            End If

            If (p1.iChar < p2.iChar) Then
                Return True
            End If

            Return False
        End Function

        Public Overloads Shared Function Operator(ByVal p1 As Place, ByVal p2 As Place) As Boolean
            If (p1.iLine > p2.iLine) Then
                Return True
            End If

            If (p1.iLine < p2.iLine) Then
                Return False
            End If

            If (p1.iChar > p2.iChar) Then
                Return True
            End If

            Return False
        End Function

        Public Overloads Shared Function Operator(ByVal p1 As Place, ByVal p2 As Place) As Boolean
            If p1.Equals(p2) Then
                Return True
            End If

            If (p1.iLine > p2.iLine) Then
                Return True
            End If

            If (p1.iLine < p2.iLine) Then
                Return False
            End If

            If (p1.iChar > p2.iChar) Then
                Return True
            End If

            Return False
        End Function

        Public Overloads Shared Function Operator(ByVal p1 As Place, ByVal p2 As Place) As Place
            Return New Place((p1.iChar + p2.iChar), (p1.iLine + p2.iLine))
        End Function

        Public Shared ReadOnly Property Empty As Place
            Get
                Return New Place
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return ("(" _
                        + (Me.iChar + ("," _
                        + (Me.iLine + ")"))))
        End Function
    End Structure
End Namespace