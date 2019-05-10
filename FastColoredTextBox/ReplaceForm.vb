Imports System
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Namespace FastColoredTextBoxNS
    Partial Public Class ReplaceForm

        Private tb As FastColoredTextBox
        Private firstSearch As Boolean = True
        Private startPlace As Place

        Public Sub New(ByVal tb As FastColoredTextBox)
            InitializeComponent()
            Me.tb = tb
        End Sub

        Private Sub btClose_Click(ByVal sender As Object, ByVal e As EventArgs)
            Close()
        End Sub

        Private Sub btFindNext_Click(ByVal sender As Object, ByVal e As EventArgs)
            Try
                If Not Find(tbFind.Text) Then MessageBox.Show("Not found")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Function FindAll(ByVal pattern As String) As List(Of Range)
            Dim opt = If(cbMatchCase.Checked, RegexOptions.None, RegexOptions.IgnoreCase)
            If Not cbRegex.Checked Then pattern = Regex.Escape(pattern)
            If cbWholeWord.Checked Then pattern = "\b" & pattern & "\b"
            Dim range = If(tb.Selection.IsEmpty, tb.Range.Clone(), tb.Selection.Clone())
            Dim list = New List(Of Range)()

            For Each r In range.GetRangesByLines(pattern, opt)
                list.Add(r)
            Next

            Return list
        End Function

        Public Function Find(ByVal pattern As String) As Boolean
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
                tb.Selection.Start = r.Start
                tb.Selection.[End] = r.[End]
                tb.DoSelectionVisible()
                tb.Invalidate()
                Return True
            Next

            If range.Start >= startPlace AndAlso startPlace > Place.Empty Then
                tb.Selection.Start = New Place(0, 0)
                Return Find(pattern)
            End If

            Return False
        End Function

        Private Sub tbFind_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
            If e.KeyChar = vbCr Then btFindNext_Click(sender, Nothing)
            If e.KeyChar = ChrW(27) Then Hide()
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
            If keyData = Keys.Escape Then
                Me.Close()
                Return True
            End If

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Private Sub ReplaceForm_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs)
            If e.CloseReason = CloseReason.UserClosing Then
                e.Cancel = True
                Hide()
            End If

            Me.tb.Focus()
        End Sub

        Private Sub btReplace_Click(ByVal sender As Object, ByVal e As EventArgs)
            Try

                If tb.SelectionLength <> 0 Then
                    If Not tb.Selection.[ReadOnly] Then tb.InsertText(tbReplace.Text)
                End If

                btFindNext_Click(sender, Nothing)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Private Sub btReplaceAll_Click(ByVal sender As Object, ByVal e As EventArgs)
            Try
                tb.Selection.BeginUpdate()
                Dim ranges = FindAll(tbFind.Text)
                Dim ro = False

                For Each r In ranges

                    If r.[ReadOnly] Then
                        ro = True
                        Exit For
                    End If
                Next

                If Not ro Then

                    If ranges.Count > 0 Then
                        tb.TextSource.Manager.ExecuteCommand(New ReplaceTextCommand(tb.TextSource, ranges, tbReplace.Text))
                        tb.Selection.Start = New Place(0, 0)
                    End If
                End If

                tb.Invalidate()
                MessageBox.Show(ranges.Count & " occurrence(s) replaced")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            tb.Selection.EndUpdate()
        End Sub

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