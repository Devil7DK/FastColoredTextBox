Imports System

Namespace FastColoredTextBoxNS
    Public Structure [Char]
        Public c As Char
        Public style As StyleIndex

        Public Sub New(ByVal c As Char)
            Me.c = c
            style = StyleIndex.None
        End Sub
    End Structure
End Namespace
