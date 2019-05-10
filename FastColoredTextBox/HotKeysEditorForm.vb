Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Namespace FastColoredTextBoxNS
    Public Class HotkeysEditorForm

        Private wrappers As BindingList(Of HotkeyWrapper) = New BindingList(Of HotkeyWrapper)()

        Public Sub New(ByVal hotkeys As HotkeysMapping)
            InitializeComponent()
            BuildWrappers(hotkeys)
            dgv.DataSource = wrappers
        End Sub

        Private Function CompereKeys(ByVal key1 As Keys, ByVal key2 As Keys) As Integer
            Dim res = (CInt(key1) And &HFF).CompareTo(CInt(key2) And &HFF)
            If res = 0 Then res = key1.CompareTo(key2)
            Return res
        End Function

        Private Sub BuildWrappers(ByVal hotkeys As HotkeysMapping)
            Dim keys = New List(Of Keys)(hotkeys.Keys)
            keys.Sort(AddressOf CompereKeys)
            wrappers.Clear()

            For Each k In keys
                wrappers.Add(New HotkeyWrapper(k, hotkeys(k)))
            Next
        End Sub

        Public Function GetHotkeys() As HotkeysMapping
            Dim result = New HotkeysMapping()

            For Each w In wrappers
                result(w.ToKeyData()) = w.Action
            Next

            Return result
        End Function

        Private Sub btAdd_Click(ByVal sender As Object, ByVal e As EventArgs)
            wrappers.Add(New HotkeyWrapper(Keys.None, FCTBAction.None))
        End Sub

        Private Sub dgv_RowsAdded(ByVal sender As Object, ByVal e As DataGridViewRowsAddedEventArgs)
            Dim cell = (TryCast(dgv(0, e.RowIndex), DataGridViewComboBoxCell))

            If cell.Items.Count = 0 Then

                For Each item In New String() {"", "Ctrl", "Ctrl + Shift", "Ctrl + Alt", "Shift", "Shift + Alt", "Alt", "Ctrl + Shift + Alt"}
                    cell.Items.Add(item)
                Next
            End If

            cell = (TryCast(dgv(1, e.RowIndex), DataGridViewComboBoxCell))

            If cell.Items.Count = 0 Then

                For Each item In [Enum].GetValues(GetType(Keys))
                    cell.Items.Add(item)
                Next
            End If

            cell = (TryCast(dgv(2, e.RowIndex), DataGridViewComboBoxCell))

            If cell.Items.Count = 0 Then

                For Each item In [Enum].GetValues(GetType(FCTBAction))
                    cell.Items.Add(item)
                Next
            End If
        End Sub

        Private Sub btResore_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim h As HotkeysMapping = New HotkeysMapping()
            h.InitDefault()
            BuildWrappers(h)
        End Sub

        Private Sub btRemove_Click(ByVal sender As Object, ByVal e As EventArgs)
            For i As Integer = dgv.RowCount - 1 To 0
                If dgv.Rows(i).Selected Then dgv.Rows.RemoveAt(i)
            Next
        End Sub

        Private Sub HotkeysEditorForm_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs)
            If DialogResult = System.Windows.Forms.DialogResult.OK Then
                Dim actions = GetUnAssignedActions()

                If Not String.IsNullOrEmpty(actions) Then
                    If MessageBox.Show("Some actions are not assigned!" & vbCrLf & "Actions: " & actions & vbCrLf & "Press Yes to save and exit, press No to continue editing", "Some actions is not assigned", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = System.Windows.Forms.DialogResult.No Then e.Cancel = True
                End If
            End If
        End Sub

        Private Function GetUnAssignedActions() As String
            Dim sb As StringBuilder = New StringBuilder()
            Dim dic = New Dictionary(Of FCTBAction, FCTBAction)()

            For Each w In wrappers
                dic(w.Action) = w.Action
            Next

            For Each item In [Enum].GetValues(GetType(FCTBAction))

                If CType(item, FCTBAction) <> FCTBAction.None Then

                    If Not (CType(item, FCTBAction)).ToString().StartsWith("CustomAction") Then
                        If Not dic.ContainsKey(CType(item, FCTBAction)) Then sb.Append(item & ", ")
                    End If
                End If
            Next

            Return sb.ToString().TrimEnd(" "c, ","c)
        End Function
    End Class

    Friend Class HotkeyWrapper
        Public Sub New(ByVal keyData As Keys, ByVal action As FCTBAction)
            Dim a As KeyEventArgs = New KeyEventArgs(keyData)
            Ctrl = a.Control
            Shift = a.Shift
            Alt = a.Alt
            Key = a.KeyCode
            action = action
        End Sub

        Public Function ToKeyData() As Keys
            Dim res = Key
            If Ctrl Then res = res Or Keys.Control
            If Alt Then res = res Or Keys.Alt
            If Shift Then res = res Or Keys.Shift
            Return res
        End Function

        Private Ctrl As Boolean
        Private Shift As Boolean
        Private Alt As Boolean

        Public Property Modifiers As String
            Get
                Dim res = ""
                If Ctrl Then res += "Ctrl + "
                If Shift Then res += "Shift + "
                If Alt Then res += "Alt + "
                Return res.Trim(" "c, "+"c)
            End Get
            Set(ByVal value As String)

                If value Is Nothing Then
                    Ctrl = CSharpImpl.__Assign(Alt, CSharpImpl.__Assign(Shift, False))
                Else
                    Ctrl = value.Contains("Ctrl")
                    Shift = value.Contains("Shift")
                    Alt = value.Contains("Alt")
                End If
            End Set
        End Property

        Public Property Key As Keys
        Public Property Action As FCTBAction

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace