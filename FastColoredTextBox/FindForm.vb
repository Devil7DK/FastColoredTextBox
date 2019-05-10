Imports System
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Namespace FastColoredTextBoxNS
    Partial Public Class FindForm

        Private firstSearch As Boolean = True
        Private startPlace As Place
        Private tb As FastColoredTextBox

        Public Sub New(ByVal tb As FastColoredTextBox)
            InitializeComponent()
            Me.tb = tb
        End Sub

        Private Sub btClose_Click(ByVal sender As Object, ByVal e As EventArgs)
            Close()
        End Sub

        Private Sub btFindNext_Click(ByVal sender As Object, ByVal e As EventArgs)
            FindNext(tbFind.Text)
        End Sub

        Public Overridable Sub FindNext(ByVal pattern As String)
            Try
                Dim opt As RegexOptions = If(cbMatchCase.Checked, RegexOptions.None, RegexOptions.IgnoreCase)
                If Not cbRegex.Checked Then pattern = Regex.Escape(pattern)
                If cbWholeWord.Checked Then pattern = "\b" & pattern & "\b"
                Dim range As Range = tb.Selection.Clone()
                range.Normalize()

                If firstSearch Then
                    startPlace = range.Start
                    firstSearch = False
                End If

                range.Start = range.[End]

                If range.Start >= startPlace Then
                    range.[End] = New Place(tb.GetLineLength(tb.LinesCount - 1), tb.LinesCount - 1)
                Else
                    range.[End] = startPlace
                End If

                For Each r In range.GetRangesByLines(pattern, opt)
                    tb.Selection = r
                    tb.DoSelectionVisible()
                    tb.Invalidate()
                    Return
                Next

                If range.Start >= startPlace AndAlso startPlace > Place.Empty Then
                    tb.Selection.Start = New Place(0, 0)
                    FindNext(pattern)
                    Return
                End If

                MessageBox.Show("Not found")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Private Sub tbFind_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
            If e.KeyChar = vbCr Then
                btFindNext.PerformClick()
                e.Handled = True
                Return
            End If

            If e.KeyChar = ChrW(27) Then
                Hide()
                e.Handled = True
                Return
            End If
        End Sub

        Private Sub FindForm_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs)
            If e.CloseReason = CloseReason.UserClosing Then
                e.Cancel = True
                Hide()
            End If

            Me.tb.Focus()
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
            If keyData = Keys.Escape Then
                Me.Close()
                Return True
            End If

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Protected Overrides Sub OnActivated(ByVal e As EventArgs)
            tbFind.Focus()
            ResetSerach()
        End Sub

        Private Sub ResetSerach()
            firstSearch = True
        End Sub

        Private Sub cbMatchCase_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
            ResetSerach()
        End Sub
    End Class
End Namespace
