Imports System
Imports System.Drawing
Imports System.Drawing.Printing

Namespace FastColoredTextBoxNS
    Public Class AutocompleteItem
        Public Text As String
        Public ImageIndex As Integer = -1
        Public Tag As Object
        Private toolTipTitle_ As String
        Private toolTipText_ As String
        Private menuText_ As String
        Public Property Parent As AutocompleteMenu

        Public Sub New()
        End Sub

        Public Sub New(ByVal text As String)
            text = text
        End Sub

        Public Sub New(ByVal text As String, ByVal imageIndex As Integer)
            Me.New(text)
            Me.ImageIndex = imageIndex
        End Sub

        Public Sub New(ByVal text As String, ByVal imageIndex As Integer, ByVal menuText As String)
            Me.New(text, imageIndex)
            Me.menuText = menuText
        End Sub

        Public Sub New(ByVal text As String, ByVal imageIndex As Integer, ByVal menuText As String, ByVal toolTipTitle As String, ByVal toolTipText_ As String)
            Me.New(text, imageIndex, menuText)
            Me.toolTipTitle = toolTipTitle
            Me.toolTipText_ = toolTipText_
        End Sub

        Public Overridable Function GetTextForReplace() As String
            Return Text
        End Function

        Public Overridable Function Compare(ByVal fragmentText As String) As CompareResult
            If Text.StartsWith(fragmentText, StringComparison.InvariantCultureIgnoreCase) AndAlso Text <> fragmentText Then Return CompareResult.VisibleAndSelected
            Return CompareResult.Hidden
        End Function

        Public Overrides Function ToString() As String
            Return If(menuText_, Text)
        End Function

        Public Overridable Sub OnSelected(ByVal popupMenu As AutocompleteMenu, ByVal e As SelectedEventArgs)
        End Sub

        Public Overridable Property ToolTipTitle As String
            Get
                Return ToolTipTitle_
            End Get
            Set(ByVal value As String)
                toolTipTitle_ = value
            End Set
        End Property

        Public Overridable Property ToolTipText As String
            Get
                Return toolTipText_
            End Get
            Set(ByVal value As String)
                toolTipText_ = value
            End Set
        End Property

        Public Overridable Property MenuText As String
            Get
                Return MenuText_
            End Get
            Set(ByVal value As String)
                menuText_ = value
            End Set
        End Property

        Public Overridable Property ForeColor As Color
            Get
                Return Color.Transparent
            End Get
            Set(ByVal value As Color)
                Throw New NotImplementedException("Override this property to change color")
            End Set
        End Property

        Public Overridable Property BackColor As Color
            Get
                Return Color.Transparent
            End Get
            Set(ByVal value As Color)
                Throw New NotImplementedException("Override this property to change color")
            End Set
        End Property
    End Class

    Public Enum CompareResult
        Hidden
        Visible
        VisibleAndSelected
    End Enum

    Public Class SnippetAutocompleteItem
        Inherits AutocompleteItem

        Public Sub New(ByVal snippet As String)
            Text = snippet.Replace(vbCr, "")
            ToolTipTitle = "Code snippet:"
            ToolTipText = Text
        End Sub

        Public Overrides Function ToString() As String
            Return If(MenuText, Text.Replace(vbLf, " ").Replace("^", ""))
        End Function

        Public Overrides Function GetTextForReplace() As String
            Return Text
        End Function

        Public Overrides Sub OnSelected(ByVal popupMenu As AutocompleteMenu, ByVal e As SelectedEventArgs)
            e.Tb.BeginUpdate()
            e.Tb.Selection.BeginUpdate()
            Dim p1 = popupMenu.Fragment.Start
            Dim p2 = e.Tb.Selection.Start

            If e.Tb.AutoIndent Then

                For iLine As Integer = p1.iLine + 1 To p2.iLine
                    e.Tb.Selection.Start = New Place(0, iLine)
                    e.Tb.DoAutoIndent(iLine)
                Next
            End If

            e.Tb.Selection.Start = p1

            While e.Tb.Selection.CharBeforeStart <> "^"c
                If Not e.Tb.Selection.GoRightThroughFolded() Then Exit While
            End While

            e.Tb.Selection.GoLeft(True)
            e.Tb.InsertText("")
            e.Tb.Selection.EndUpdate()
            e.Tb.EndUpdate()
        End Sub

        Public Overrides Function Compare(ByVal fragmentText As String) As CompareResult
            If Text.StartsWith(fragmentText, StringComparison.InvariantCultureIgnoreCase) AndAlso Text <> fragmentText Then Return CompareResult.Visible
            Return CompareResult.Hidden
        End Function
    End Class

    Public Class MethodAutocompleteItem
        Inherits AutocompleteItem

        Private firstPart As String
        Private lowercaseText As String

        Public Sub New(ByVal text As String)
            MyBase.New(text)
            lowercaseText = text.ToLower()
        End Sub

        Public Overrides Function Compare(ByVal fragmentText As String) As CompareResult
            Dim i As Integer = fragmentText.LastIndexOf("."c)
            If i < 0 Then Return CompareResult.Hidden
            Dim lastPart As String = fragmentText.Substring(i + 1)
            firstPart = fragmentText.Substring(0, i)
            If lastPart = "" Then Return CompareResult.Visible
            If Text.StartsWith(lastPart, StringComparison.InvariantCultureIgnoreCase) Then Return CompareResult.VisibleAndSelected
            If lowercaseText.Contains(lastPart.ToLower()) Then Return CompareResult.Visible
            Return CompareResult.Hidden
        End Function

        Public Overrides Function GetTextForReplace() As String
            Return firstPart & "." & Text
        End Function
    End Class

    Public Class SuggestItem
        Inherits AutocompleteItem

        Public Sub New(ByVal text As String, ByVal imageIndex As Integer)
            MyBase.New(text, imageIndex)
        End Sub

        Public Overrides Function Compare(ByVal fragmentText As String) As CompareResult
            Return CompareResult.Visible
        End Function
    End Class
End Namespace