Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System

Namespace FastColoredTextBoxNS
    Public Class SyntaxDescriptor
        Implements IDisposable

        Public leftBracket As Char = "("c
        Public rightBracket As Char = ")"c
        Public leftBracket2 As Char = "{"c
        Public rightBracket2 As Char = "}"c
        Public bracketsHighlightStrategy As BracketsHighlightStrategy = bracketsHighlightStrategy.Strategy2
        Public ReadOnly styles As List(Of Style) = New List(Of Style)()
        Public ReadOnly rules As List(Of RuleDesc) = New List(Of RuleDesc)()
        Public ReadOnly foldings As List(Of FoldingDesc) = New List(Of FoldingDesc)()

        Public Sub Dispose()
            For Each style In styles
                style.Dispose()
            Next
        End Sub
    End Class

    Public Class RuleDesc
        Private regex As Regex
        Public pattern As String
        Public options As RegexOptions = RegexOptions.None
        Public style As Style

        Public ReadOnly Property Regex As Regex
            Get

                If Regex Is Nothing Then
                    Regex = New Regex(pattern, SyntaxHighlighter.RegexCompiledOption Or options)
                End If

                Return Regex
            End Get
        End Property
    End Class

    Public Class FoldingDesc
        Public startMarkerRegex As String
        Public finishMarkerRegex As String
        Public options As RegexOptions = RegexOptions.None
    End Class
End Namespace