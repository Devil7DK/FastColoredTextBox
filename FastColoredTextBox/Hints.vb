Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Text
Imports System.Windows.Forms

Namespace FastColoredTextBoxNS
    Public Class Hints
        Implements ICollection(Of Hint), IDisposable

        Private tb As FastColoredTextBox
        Private items As List(Of Hint) = New List(Of Hint)()

        Public Sub New(ByVal tb As FastColoredTextBox)
            Me.tb = tb
            tb.TextChanged += AddressOf OnTextBoxTextChanged
            tb.KeyDown += AddressOf OnTextBoxKeyDown
            tb.VisibleRangeChanged += AddressOf OnTextBoxVisibleRangeChanged
        End Sub

        Protected Overridable Sub OnTextBoxKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = System.Windows.Forms.Keys.Escape AndAlso e.Modifiers = System.Windows.Forms.Keys.None Then Clear()
        End Sub

        Protected Overridable Sub OnTextBoxTextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
            Clear()
        End Sub

        Public Sub Dispose()
            tb.TextChanged -= AddressOf OnTextBoxTextChanged
            tb.KeyDown -= AddressOf OnTextBoxKeyDown
            tb.VisibleRangeChanged -= AddressOf OnTextBoxVisibleRangeChanged
        End Sub

        Private Sub OnTextBoxVisibleRangeChanged(ByVal sender As Object, ByVal e As EventArgs)
            If items.Count = 0 Then Return
            tb.NeedRecalc(True)

            For Each item In items
                LayoutHint(item)
                item.HostPanel.Invalidate()
            Next
        End Sub

        Private Sub LayoutHint(ByVal hint As Hint)
            If hint.Inline Then

                If hint.Range.Start.iLine < tb.LineInfos.Count - 1 Then
                    hint.HostPanel.Top = tb.LineInfos(hint.Range.Start.iLine + 1).startY - hint.TopPadding - hint.HostPanel.Height - tb.VerticalScroll.Value
                Else
                    hint.HostPanel.Top = tb.TextHeight + tb.Paddings.Top - hint.HostPanel.Height - tb.VerticalScroll.Value
                End If
            Else
                If hint.Range.Start.iLine > tb.LinesCount - 1 Then Return

                If hint.Range.Start.iLine = tb.LinesCount - 1 Then
                    Dim y = tb.LineInfos(hint.Range.Start.iLine).startY - tb.VerticalScroll.Value + tb.CharHeight

                    If y + hint.HostPanel.Height + 1 > tb.ClientRectangle.Bottom Then
                        hint.HostPanel.Top = Math.Max(0, tb.LineInfos(hint.Range.Start.iLine).startY - tb.VerticalScroll.Value - hint.HostPanel.Height)
                    Else
                        hint.HostPanel.Top = y
                    End If
                Else
                    hint.HostPanel.Top = tb.LineInfos(hint.Range.Start.iLine + 1).startY - tb.VerticalScroll.Value
                    If hint.HostPanel.Bottom > tb.ClientRectangle.Bottom Then hint.HostPanel.Top = tb.LineInfos(hint.Range.Start.iLine + 1).startY - tb.CharHeight - hint.TopPadding - hint.HostPanel.Height - tb.VerticalScroll.Value
                End If
            End If

            If hint.Dock = DockStyle.Fill Then
                hint.Width = tb.ClientSize.Width - tb.LeftIndent - 2
                hint.HostPanel.Left = tb.LeftIndent
            Else
                Dim p1 = tb.PlaceToPoint(hint.Range.Start)
                Dim p2 = tb.PlaceToPoint(hint.Range.[End])
                Dim cx = (p1.X + p2.X) / 2
                Dim x = cx - hint.HostPanel.Width / 2
                hint.HostPanel.Left = Math.Max(tb.LeftIndent, x)
                If hint.HostPanel.Right > tb.ClientSize.Width Then hint.HostPanel.Left = Math.Max(tb.LeftIndent, x - (hint.HostPanel.Right - tb.ClientSize.Width))
            End If
        End Sub

        Public Iterator Function GetEnumerator() As IEnumerator(Of Hint)
            For Each item In items
                Yield item
            Next
        End Function

        Private Function GetEnumerator() As System.Collections.IEnumerator
            Return GetEnumerator()
        End Function

        Public Sub Clear()
            items.Clear()

            If tb.Controls.Count <> 0 Then
                Dim toDelete = New List(Of Control)()

                For Each item As Control In tb.Controls
                    If TypeOf item Is UnfocusablePanel Then toDelete.Add(item)
                Next

                For Each item In toDelete
                    tb.Controls.Remove(item)
                Next

                For i As Integer = 0 To tb.LineInfos.Count - 1
                    Dim li = tb.LineInfos(i)
                    li.bottomPadding = 0
                    tb.LineInfos(i) = li
                Next

                tb.NeedRecalc()
                tb.Invalidate()
                tb.[Select]()
                tb.ActiveControl = Nothing
            End If
        End Sub

        Public Sub Add(ByVal hint As Hint)
            items.Add(hint)

            If hint.Inline Then
                Dim li = tb.LineInfos(hint.Range.Start.iLine)
                hint.TopPadding = li.bottomPadding
                li.bottomPadding += hint.HostPanel.Height
                tb.LineInfos(hint.Range.Start.iLine) = li
                tb.NeedRecalc(True)
            End If

            LayoutHint(hint)
            tb.OnVisibleRangeChanged()
            hint.HostPanel.Parent = tb
            tb.[Select]()
            tb.ActiveControl = Nothing
            tb.Invalidate()
        End Sub

        Public Function Contains(ByVal item As Hint) As Boolean
            Return items.Contains(item)
        End Function

        Public Sub CopyTo(ByVal array As Hint(), ByVal arrayIndex As Integer)
            items.CopyTo(array, arrayIndex)
        End Sub

        Public ReadOnly Property Count As Integer
            Get
                Return items.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean
            Get
                Return False
            End Get
        End Property

        Public Function Remove(ByVal item As Hint) As Boolean
            Throw New NotImplementedException()
        End Function
    End Class

    Public Class Hint
        Public Property Text As String
            Get
                Return HostPanel.Text
            End Get
            Set(ByVal value As String)
                HostPanel.Text = value
            End Set
        End Property

        Public Property Range As Range

        Public Property BackColor As Color
            Get
                Return HostPanel.BackColor
            End Get
            Set(ByVal value As Color)
                HostPanel.BackColor = value
            End Set
        End Property

        Public Property BackColor2 As Color
            Get
                Return HostPanel.BackColor2
            End Get
            Set(ByVal value As Color)
                HostPanel.BackColor2 = value
            End Set
        End Property

        Public Property BorderColor As Color
            Get
                Return HostPanel.BorderColor
            End Get
            Set(ByVal value As Color)
                HostPanel.BorderColor = value
            End Set
        End Property

        Public Property ForeColor As Color
            Get
                Return HostPanel.ForeColor
            End Get
            Set(ByVal value As Color)
                HostPanel.ForeColor = value
            End Set
        End Property

        Public Property TextAlignment As StringAlignment
            Get
                Return HostPanel.TextAlignment
            End Get
            Set(ByVal value As StringAlignment)
                HostPanel.TextAlignment = value
            End Set
        End Property

        Public Property Font As Font
            Get
                Return HostPanel.Font
            End Get
            Set(ByVal value As Font)
                HostPanel.Font = value
            End Set
        End Property

        Public Event Click As EventHandler
        AddHandler(ByVal value As EventHandler)
        AddHandler() HostPanel.Click, value
            End AddHandler
        RemoveHandler(ByVal value As EventHandler)
        RemoveHandler() HostPanel.Click, value
            End RemoveHandler
        End Event

        Public Property InnerControl As Control
        Public Property Dock As DockStyle

        Public Property Width As Integer
            Get
                Return HostPanel.Width
            End Get
            Set(ByVal value As Integer)
                HostPanel.Width = value
            End Set
        End Property

        Public Property Height As Integer
            Get
                Return HostPanel.Height
            End Get
            Set(ByVal value As Integer)
                HostPanel.Height = value
            End Set
        End Property

        Public Property HostPanel As UnfocusablePanel
        Friend Property TopPadding As Integer
        Public Property Tag As Object

        Public Property Cursor As Cursor
            Get
                Return HostPanel.Cursor
            End Get
            Set(ByVal value As Cursor)
                HostPanel.Cursor = value
            End Set
        End Property

        Public Property Inline As Boolean

        Public Overridable Sub DoVisible()
            Range.tb.DoRangeVisible(Range, True)
            Range.tb.DoVisibleRectangle(HostPanel.Bounds)
            Range.tb.Invalidate()
        End Sub

        Private Sub New(ByVal range As Range, ByVal innerControl As Control, ByVal text As String, ByVal inline As Boolean, ByVal dock As Boolean)
            Me.Range = range
            Me.Inline = inline
            Me.InnerControl = innerControl
            Init()
            dock = If(dock, DockStyle.Fill, DockStyle.None)
            text = text
        End Sub

        Public Sub New(ByVal range As Range, ByVal text As String, ByVal inline As Boolean, ByVal dock As Boolean)
            Me.New(range, Nothing, text, inline, dock)
        End Sub

        Public Sub New(ByVal range As Range, ByVal text As String)
            Me.New(range, Nothing, text, True, True)
        End Sub

        Public Sub New(ByVal range As Range, ByVal innerControl As Control, ByVal inline As Boolean, ByVal dock As Boolean)
            Me.New(range, innerControl, Nothing, inline, dock)
        End Sub

        Public Sub New(ByVal range As Range, ByVal innerControl As Control)
            Me.New(range, innerControl, Nothing, True, True)
        End Sub

        Protected Overridable Sub Init()
            HostPanel = New UnfocusablePanel()
            HostPanel.Click += AddressOf OnClick
            Cursor = Cursors.[Default]
            BorderColor = Color.Silver
            BackColor2 = Color.White
            BackColor = If(InnerControl Is Nothing, Color.Silver, SystemColors.Control)
            ForeColor = Color.Black
            TextAlignment = StringAlignment.Near
            Font = If(Range.tb.Parent Is Nothing, Range.tb.Font, Range.tb.Parent.Font)

            If InnerControl IsNot Nothing Then
                HostPanel.Controls.Add(InnerControl)
                Dim size = InnerControl.GetPreferredSize(InnerControl.Size)
                HostPanel.Width = size.Width + 2
                HostPanel.Height = size.Height + 2
                InnerControl.Dock = DockStyle.Fill
                InnerControl.Visible = True
                BackColor = SystemColors.Control
            Else
                HostPanel.Height = Range.tb.CharHeight + 5
            End If
        End Sub

        Protected Overridable Sub OnClick(ByVal sender As Object, ByVal e As EventArgs)
            Range.tb.OnHintClick(Me)
        End Sub
    End Class
End Namespace