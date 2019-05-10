Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Drawing
Imports System.Drawing.Design
Imports System.Globalization
Imports System.Reflection
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Forms.Design
Imports KEYS = System.Windows.Forms.Keys

Namespace FastColoredTextBoxNS
    Public Class HotkeysMapping
        Inherits SortedDictionary(Of KEYS, FCTBAction)

        Public Overridable Sub InitDefault()
            Me(Keys.Control Or Keys.G) = FCTBAction.GoToDialog
            Me(Keys.Control Or Keys.F) = FCTBAction.FindDialog
            Me(Keys.Alt Or Keys.F) = FCTBAction.FindChar
            Me(Keys.F3) = FCTBAction.FindNext
            Me(Keys.Control Or Keys.H) = FCTBAction.ReplaceDialog
            Me(Keys.Control Or Keys.C) = FCTBAction.Copy
            Me(Keys.Control Or Keys.Shift Or Keys.C) = FCTBAction.CommentSelected
            Me(Keys.Control Or Keys.X) = FCTBAction.Cut
            Me(Keys.Control Or Keys.V) = FCTBAction.Paste
            Me(Keys.Control Or Keys.A) = FCTBAction.SelectAll
            Me(Keys.Control Or Keys.Z) = FCTBAction.Undo
            Me(Keys.Control Or Keys.R) = FCTBAction.Redo
            Me(Keys.Control Or Keys.U) = FCTBAction.UpperCase
            Me(Keys.Shift Or Keys.Control Or Keys.U) = FCTBAction.LowerCase
            Me(Keys.Control Or Keys.OemMinus) = FCTBAction.NavigateBackward
            Me(Keys.Control Or Keys.Shift Or Keys.OemMinus) = FCTBAction.NavigateForward
            Me(Keys.Control Or Keys.B) = FCTBAction.BookmarkLine
            Me(Keys.Control Or Keys.Shift Or Keys.B) = FCTBAction.UnbookmarkLine
            Me(Keys.Control Or Keys.N) = FCTBAction.GoNextBookmark
            Me(Keys.Control Or Keys.Shift Or Keys.N) = FCTBAction.GoPrevBookmark
            Me(Keys.Alt Or Keys.Back) = FCTBAction.Undo
            Me(Keys.Control Or Keys.Back) = FCTBAction.ClearWordLeft
            Me(Keys.Insert) = FCTBAction.ReplaceMode
            Me(Keys.Control Or Keys.Insert) = FCTBAction.Copy
            Me(Keys.Shift Or Keys.Insert) = FCTBAction.Paste
            Me(Keys.Delete) = FCTBAction.DeleteCharRight
            Me(Keys.Control Or Keys.Delete) = FCTBAction.ClearWordRight
            Me(Keys.Shift Or Keys.Delete) = FCTBAction.Cut
            Me(Keys.Left) = FCTBAction.GoLeft
            Me(Keys.Shift Or Keys.Left) = FCTBAction.GoLeftWithSelection
            Me(Keys.Control Or Keys.Left) = FCTBAction.GoWordLeft
            Me(Keys.Control Or Keys.Shift Or Keys.Left) = FCTBAction.GoWordLeftWithSelection
            Me(Keys.Alt Or Keys.Shift Or Keys.Left) = FCTBAction.GoLeft_ColumnSelectionMode
            Me(Keys.Right) = FCTBAction.GoRight
            Me(Keys.Shift Or Keys.Right) = FCTBAction.GoRightWithSelection
            Me(Keys.Control Or Keys.Right) = FCTBAction.GoWordRight
            Me(Keys.Control Or Keys.Shift Or Keys.Right) = FCTBAction.GoWordRightWithSelection
            Me(Keys.Alt Or Keys.Shift Or Keys.Right) = FCTBAction.GoRight_ColumnSelectionMode
            Me(Keys.Up) = FCTBAction.GoUp
            Me(Keys.Shift Or Keys.Up) = FCTBAction.GoUpWithSelection
            Me(Keys.Alt Or Keys.Shift Or Keys.Up) = FCTBAction.GoUp_ColumnSelectionMode
            Me(Keys.Alt Or Keys.Up) = FCTBAction.MoveSelectedLinesUp
            Me(Keys.Control Or Keys.Up) = FCTBAction.ScrollUp
            Me(Keys.Down) = FCTBAction.GoDown
            Me(Keys.Shift Or Keys.Down) = FCTBAction.GoDownWithSelection
            Me(Keys.Alt Or Keys.Shift Or Keys.Down) = FCTBAction.GoDown_ColumnSelectionMode
            Me(Keys.Alt Or Keys.Down) = FCTBAction.MoveSelectedLinesDown
            Me(Keys.Control Or Keys.Down) = FCTBAction.ScrollDown
            Me(Keys.PageUp) = FCTBAction.GoPageUp
            Me(Keys.Shift Or Keys.PageUp) = FCTBAction.GoPageUpWithSelection
            Me(Keys.PageDown) = FCTBAction.GoPageDown
            Me(Keys.Shift Or Keys.PageDown) = FCTBAction.GoPageDownWithSelection
            Me(Keys.Home) = FCTBAction.GoHome
            Me(Keys.Shift Or Keys.Home) = FCTBAction.GoHomeWithSelection
            Me(Keys.Control Or Keys.Home) = FCTBAction.GoFirstLine
            Me(Keys.Control Or Keys.Shift Or Keys.Home) = FCTBAction.GoFirstLineWithSelection
            Me(Keys.[End]) = FCTBAction.GoEnd
            Me(Keys.Shift Or Keys.[End]) = FCTBAction.GoEndWithSelection
            Me(Keys.Control Or Keys.[End]) = FCTBAction.GoLastLine
            Me(Keys.Control Or Keys.Shift Or Keys.[End]) = FCTBAction.GoLastLineWithSelection
            Me(Keys.Escape) = FCTBAction.ClearHints
            Me(Keys.Control Or Keys.M) = FCTBAction.MacroRecord
            Me(Keys.Control Or Keys.E) = FCTBAction.MacroExecute
            Me(Keys.Control Or Keys.Space) = FCTBAction.AutocompleteMenu
            Me(Keys.Tab) = FCTBAction.IndentIncrease
            Me(Keys.Shift Or Keys.Tab) = FCTBAction.IndentDecrease
            Me(Keys.Control Or Keys.Subtract) = FCTBAction.ZoomOut
            Me(Keys.Control Or Keys.Add) = FCTBAction.ZoomIn
            Me(Keys.Control Or Keys.D0) = FCTBAction.ZoomNormal
            Me(Keys.Control Or Keys.I) = FCTBAction.AutoIndentChars
        End Sub

        Public Overrides Function ToString() As String
            Dim cult = Thread.CurrentThread.CurrentUICulture
            Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture
            Dim sb As StringBuilder = New StringBuilder()
            Dim kc = New KeysConverter()

            For Each pair In Me
                sb.AppendFormat("{0}={1}, ", kc.ConvertToString(pair.Key), pair.Value)
            Next

            If sb.Length > 1 Then sb.Remove(sb.Length - 2, 2)
            Thread.CurrentThread.CurrentUICulture = cult
            Return sb.ToString()
        End Function

        Public Shared Function Parse(ByVal s As String) As HotkeysMapping
            Dim result = New HotkeysMapping()
            result.Clear()
            Dim cult = Thread.CurrentThread.CurrentUICulture
            Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture
            Dim kc = New KeysConverter()

            For Each p In s.Split(","c)
                Dim pp = p.Split("="c)
                Dim k = CType(kc.ConvertFromString(pp(0).Trim()), KEYS)
                Dim a = CType([Enum].Parse(GetType(FCTBAction), pp(1).Trim()), FCTBAction)
                result(k) = a
            Next

            Thread.CurrentThread.CurrentUICulture = cult
            Return result
        End Function
    End Class

    Public Enum FCTBAction
        None
        AutocompleteMenu
        AutoIndentChars
        BookmarkLine
        ClearHints
        ClearWordLeft
        ClearWordRight
        CommentSelected
        Copy
        Cut
        DeleteCharRight
        FindChar
        FindDialog
        FindNext
        GoDown
        GoDownWithSelection
        GoDown_ColumnSelectionMode
        GoEnd
        GoEndWithSelection
        GoFirstLine
        GoFirstLineWithSelection
        GoHome
        GoHomeWithSelection
        GoLastLine
        GoLastLineWithSelection
        GoLeft
        GoLeftWithSelection
        GoLeft_ColumnSelectionMode
        GoPageDown
        GoPageDownWithSelection
        GoPageUp
        GoPageUpWithSelection
        GoRight
        GoRightWithSelection
        GoRight_ColumnSelectionMode
        GoToDialog
        GoNextBookmark
        GoPrevBookmark
        GoUp
        GoUpWithSelection
        GoUp_ColumnSelectionMode
        GoWordLeft
        GoWordLeftWithSelection
        GoWordRight
        GoWordRightWithSelection
        IndentIncrease
        IndentDecrease
        LowerCase
        MacroExecute
        MacroRecord
        MoveSelectedLinesDown
        MoveSelectedLinesUp
        NavigateBackward
        NavigateForward
        Paste
        Redo
        ReplaceDialog
        ReplaceMode
        ScrollDown
        ScrollUp
        SelectAll
        UnbookmarkLine
        Undo
        UpperCase
        ZoomIn
        ZoomNormal
        ZoomOut
        CustomAction1
        CustomAction2
        CustomAction3
        CustomAction4
        CustomAction5
        CustomAction6
        CustomAction7
        CustomAction8
        CustomAction9
        CustomAction10
        CustomAction11
        CustomAction12
        CustomAction13
        CustomAction14
        CustomAction15
        CustomAction16
        CustomAction17
        CustomAction18
        CustomAction19
        CustomAction20
    End Enum

    Friend Class HotkeysEditor
        Inherits UITypeEditor

        Public Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As UITypeEditorEditStyle
            Return UITypeEditorEditStyle.Modal
        End Function

        Public Overrides Function EditValue(ByVal context As ITypeDescriptorContext, ByVal provider As IServiceProvider, ByVal value As Object) As Object
            If (provider IsNot Nothing) AndAlso ((CType(provider.GetService(GetType(IWindowsFormsEditorService)), IWindowsFormsEditorService)) IsNot Nothing) Then
                Dim form = New HotkeysEditorForm(HotkeysMapping.Parse(TryCast(value, String)))
                If form.ShowDialog() = DialogResult.OK Then value = form.GetHotkeys().ToString()
            End If

            Return value
        End Function
    End Class
End Namespace