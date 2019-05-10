Imports System
Imports System.Windows.Forms

Namespace FastColoredTextBoxNS
    Public Class GoToForm
        Inherits Form

        Public Property SelectedLineNumber As Integer
        Public Property TotalLineCount As Integer

        Public Sub New()
            InitializeComponent()
        End Sub

        Protected Overrides Sub OnLoad(ByVal e As EventArgs)
            MyBase.OnLoad(e)
            Me.tbLineNumber.Text = Me.SelectedLineNumber.ToString()
            Me.label.Text = String.Format("Line number (1 - {0}):", Me.TotalLineCount)
        End Sub

        Protected Overrides Sub OnShown(ByVal e As EventArgs)
            MyBase.OnShown(e)
            Me.tbLineNumber.Focus()
        End Sub

        Private Sub btnOk_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim enteredLine As Integer

            If Integer.TryParse(Me.tbLineNumber.Text, enteredLine) Then
                enteredLine = Math.Min(enteredLine, Me.TotalLineCount)
                enteredLine = Math.Max(1, enteredLine)
                Me.SelectedLineNumber = enteredLine
            End If

            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub

        Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs)
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        End Sub
    End Class
End Namespace