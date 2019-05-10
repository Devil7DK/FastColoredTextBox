Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions

Namespace FastColoredTextBoxNS
    <Browsable(False)>
    Public Class AutocompleteMenu
        Inherits ToolStripDropDown
        Implements IDisposable

        Private listView As AutocompleteListView
        Public host As ToolStripControlHost
        Public Property Fragment As Range
        Public Property SearchPattern As String
        Public Property MinFragmentLength As Integer
        Public Event Selecting As EventHandler(Of SelectingEventArgs)
        Public Event Selected As EventHandler(Of SelectedEventArgs)
        Public Overloads Event Opening As EventHandler(Of CancelEventArgs)

        Public Property AllowTabKey As Boolean
            Get
                Return listView.AllowTabKey
            End Get
            Set(ByVal value As Boolean)
                listView.AllowTabKey = value
            End Set
        End Property

        Public Property AppearInterval As Integer
            Get
                Return listView.AppearInterval
            End Get
            Set(ByVal value As Integer)
                listView.AppearInterval = value
            End Set
        End Property

        Public Property MaxTooltipSize As Size
            Get
                Return listView.MaxToolTipSize
            End Get
            Set(ByVal value As Size)
                listView.MaxToolTipSize = value
            End Set
        End Property

        Public Property AlwaysShowTooltip As Boolean
            Get
                Return listView.AlwaysShowTooltip
            End Get
            Set(ByVal value As Boolean)
                listView.AlwaysShowTooltip = value
            End Set
        End Property

        <DefaultValue(GetType(Color), "Orange")>
        Public Property SelectedColor As Color
            Get
                Return listView.SelectedColor
            End Get
            Set(ByVal value As Color)
                listView.SelectedColor = value
            End Set
        End Property

        <DefaultValue(GetType(Color), "Red")>
        Public Property HoveredColor As Color
            Get
                Return listView.HoveredColor
            End Get
            Set(ByVal value As Color)
                listView.HoveredColor = value
            End Set
        End Property

        Public Sub New(ByVal tb As FastColoredTextBox)
            AutoClose = False
            AutoSize = False
            Margin = Padding.Empty
            Padding = Padding.Empty
            BackColor = Color.White
            listView = New AutocompleteListView(tb)
            host = New ToolStripControlHost(listView)
            host.Margin = New Padding(2, 2, 2, 2)
            host.Padding = Padding.Empty
            host.AutoSize = False
            host.AutoToolTip = False
            CalcSize()
            MyBase.Items.Add(host)
            listView.Parent = Me
            SearchPattern = "[\w\.]"
            MinFragmentLength = 2
        End Sub

        Public Overloads Property Font As Font
            Get
                Return listView.Font
            End Get
            Set(ByVal value As Font)
                listView.Font = value
            End Set
        End Property

        Friend Overloads Sub OnOpening(ByVal args As CancelEventArgs)
            RaiseEvent Opening(Me, args)
        End Sub

        Public Overloads Sub Close()
            listView.toolTip.Hide(listView)
            MyBase.Close()
        End Sub

        Friend Sub CalcSize()
            host.Size = listView.Size
            Size = New System.Drawing.Size(listView.Size.Width + 4, listView.Size.Height + 4)
        End Sub

        Public Overridable Sub OnSelecting()
            listView.OnSelecting()
        End Sub

        Public Sub SelectNext(ByVal shift As Integer)
            listView.SelectNext(shift)
        End Sub

        Friend Sub OnSelecting(ByVal args As SelectingEventArgs)
            RaiseEvent Selecting(Me, args)
        End Sub

        Public Sub OnSelected(ByVal args As SelectedEventArgs)
            RaiseEvent Selected(Me, args)
        End Sub

        Public Overloads ReadOnly Property Items As AutocompleteListView
            Get
                Return listView
            End Get
        End Property

        Public Sub Show(ByVal forced As Boolean)
            Items.DoAutocomplete(forced)
        End Sub

        Public Overloads Property MinimumSize As Size
            Get
                Return Items.MinimumSize
            End Get
            Set(ByVal value As Size)
                Items.MinimumSize = value
            End Set
        End Property

        Public Overloads Property ImageList As ImageList
            Get
                Return Items.ImageList
            End Get
            Set(ByVal value As ImageList)
                Items.ImageList = value
            End Set
        End Property

        Public Property ToolTipDuration As Integer
            Get
                Return Items.ToolTipDuration
            End Get
            Set(ByVal value As Integer)
                Items.ToolTipDuration = value
            End Set
        End Property

        Public Property ToolTip As ToolTip
            Get
                Return Items.toolTip
            End Get
            Set(ByVal value As ToolTip)
                Items.toolTip = value
            End Set
        End Property

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            MyBase.Dispose(disposing)
            If listView IsNot Nothing AndAlso Not listView.IsDisposed Then listView.Dispose()
        End Sub
    End Class

    <System.ComponentModel.ToolboxItem(False)>
    Public Class AutocompleteListView
        Inherits UserControl
        Implements IDisposable

        Public Event FocussedItemIndexChanged As EventHandler
        Friend visibleItems As List(Of AutocompleteItem)
        Private sourceItems As IEnumerable(Of AutocompleteItem) = New List(Of AutocompleteItem)()
        Private focussedItemIndex As Integer = 0
        Private hoveredItemIndex As Integer = -1

        Private ReadOnly Property ItemHeight As Integer
            Get
                Return Font.Height + 2
            End Get
        End Property

        Private ReadOnly Property Menu As AutocompleteMenu
            Get
                Return TryCast(Parent, AutocompleteMenu)
            End Get
        End Property

        Private oldItemCount As Integer = 0
        Private tb As FastColoredTextBox
        Friend toolTip As ToolTip = New ToolTip()
        Private timer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer()
        Friend Property AllowTabKey As Boolean
        Public Property ImageList As ImageList

        Friend Property AppearInterval As Integer
            Get
                Return timer.Interval
            End Get
            Set(ByVal value As Integer)
                timer.Interval = value
            End Set
        End Property

        Friend Property ToolTipDuration As Integer
        Friend Property MaxToolTipSize As Size

        Friend Property AlwaysShowTooltip As Boolean
            Get
                Return toolTip.ShowAlways
            End Get
            Set(ByVal value As Boolean)
                toolTip.ShowAlways = value
            End Set
        End Property

        Public Property SelectedColor As Color
        Public Property HoveredColor As Color

        Public Property FocussedItemIndex As Integer
            Get
                Return FocussedItemIndex
            End Get
            Set(ByVal value As Integer)

                If focussedItemIndex <> value Then
                    focussedItemIndex = value
                    RaiseEvent FocussedItemIndexChanged(Me, EventArgs.Empty)
                End If
            End Set
        End Property

        Public Property FocussedItem As AutocompleteItem
            Get
                If focussedItemIndex >= 0 AndAlso focussedItemIndex < visibleItems.Count Then Return visibleItems(focussedItemIndex)
                Return Nothing
            End Get
            Set(ByVal value As AutocompleteItem)
                focussedItemIndex = visibleItems.IndexOf(value)
            End Set
        End Property

        Friend Sub New(ByVal tb As FastColoredTextBox)
            SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
            MyBase.Font = New Font(FontFamily.GenericSansSerif, 9)
            visibleItems = New List(Of AutocompleteItem)()
            VerticalScroll.SmallChange = ItemHeight
            MaximumSize = New Size(Size.Width, 180)
            toolTip.ShowAlways = False
            AppearInterval = 500
            AddHandler timer.Tick, New EventHandler(AddressOf timer_Tick)
            SelectedColor = Color.Orange
            HoveredColor = Color.Red
            ToolTipDuration = 3000
            toolTip.Popup += AddressOf ToolTip_Popup
            Me.tb = tb
            tb.KeyDown += New KeyEventHandler(AddressOf tb_KeyDown)
            AddHandler tb.SelectionChanged, New EventHandler(AddressOf tb_SelectionChanged)
            tb.KeyPressed += New KeyPressEventHandler(AddressOf tb_KeyPressed)
            Dim form As Form = tb.FindForm()

            If form IsNot Nothing Then
                form.LocationChanged += Function()
                                            SafetyClose()
                                        End Function

                form.ResizeBegin += Function()
                                        SafetyClose()
                                    End Function

                form.FormClosing += Function()
                                        SafetyClose()
                                    End Function

                form.LostFocus += Function()
                                      SafetyClose()
                                  End Function
            End If

            tb.LostFocus += Function(o, e)

                                If Menu IsNot Nothing AndAlso Not Menu.IsDisposed Then
                                    If Not Menu.Focused Then SafetyClose()
                                End If
                            End Function

            tb.Scroll += Function()
                             SafetyClose()
                         End Function

            Me.VisibleChanged += Function(o, e)
                                     If Me.Visible Then DoSelectedVisible()
                                 End Function
        End Sub

        Private Sub ToolTip_Popup(ByVal sender As Object, ByVal e As PopupEventArgs)
            If MaxToolTipSize.Height > 0 AndAlso MaxToolTipSize.Width > 0 Then e.ToolTipSize = MaxToolTipSize
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If toolTip IsNot Nothing Then
                toolTip.Popup -= AddressOf ToolTip_Popup
                toolTip.Dispose()
            End If

            If tb IsNot Nothing Then
                tb.KeyDown -= AddressOf tb_KeyDown
                tb.KeyPressed -= AddressOf tb_KeyPressed
                tb.SelectionChanged -= AddressOf tb_SelectionChanged
            End If

            If timer IsNot Nothing Then
                timer.[Stop]()
                timer.Tick -= AddressOf timer_Tick
                timer.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

        Private Sub SafetyClose()
            If Menu IsNot Nothing AndAlso Not Menu.IsDisposed Then Menu.Close()
        End Sub

        Private Sub tb_KeyPressed(ByVal sender As Object, ByVal e As KeyPressEventArgs)
            Dim backspaceORdel As Boolean = e.KeyChar = vbBack OrElse e.KeyChar = &HFF

            If Menu.Visible AndAlso Not backspaceORdel Then
                DoAutocomplete(False)
            Else
                ResetTimer(timer)
            End If
        End Sub

        Private Sub timer_Tick(ByVal sender As Object, ByVal e As EventArgs)
            timer.[Stop]()
            DoAutocomplete(False)
        End Sub

        Private Sub ResetTimer(ByVal timer As System.Windows.Forms.Timer)
            timer.[Stop]()
            timer.Start()
        End Sub

        Friend Sub DoAutocomplete()
            DoAutocomplete(False)
        End Sub

        Friend Sub DoAutocomplete(ByVal forced As Boolean)
            If Not Menu.Enabled Then
                Menu.Close()
                Return
            End If

            visibleItems.Clear()
            focussedItemIndex = 0
            VerticalScroll.Value = 0
            AutoScrollMinSize -= New Size(1, 0)
            AutoScrollMinSize += New Size(1, 0)
            Dim fragment As Range = tb.Selection.GetFragment(Menu.SearchPattern)
            Dim text As String = fragment.Text
            Dim point As Point = tb.PlaceToPoint(fragment.[End])
            point.Offset(2, tb.CharHeight)

            If forced OrElse (text.Length >= Menu.MinFragmentLength AndAlso tb.Selection.IsEmpty AndAlso (tb.Selection.Start > fragment.Start OrElse text.Length = 0)) Then
                Menu.Fragment = fragment
                Dim foundSelected As Boolean = False

                For Each item In sourceItems
                    item.Parent = Menu
                    Dim res As CompareResult = item.Compare(text)
                    If res <> CompareResult.Hidden Then visibleItems.Add(item)

                    If res = CompareResult.VisibleAndSelected AndAlso Not foundSelected Then
                        foundSelected = True
                        focussedItemIndex = visibleItems.Count - 1
                    End If
                Next

                If foundSelected Then
                    AdjustScroll()
                    DoSelectedVisible()
                End If
            End If

            If Count > 0 Then

                If Not Menu.Visible Then
                    Dim args As CancelEventArgs = New CancelEventArgs()
                    Menu.OnOpening(args)
                    If Not args.Cancel Then Menu.Show(tb, point)
                End If

                DoSelectedVisible()
                Invalidate()
            Else
                Menu.Close()
            End If
        End Sub

        Private Sub tb_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
            If Menu.Visible Then
                Dim needClose As Boolean = False

                If Not tb.Selection.IsEmpty Then
                    needClose = True
                ElseIf Not Menu.Fragment.Contains(tb.Selection.Start) Then

                    If tb.Selection.Start.iLine = Menu.Fragment.[End].iLine AndAlso tb.Selection.Start.iChar = Menu.Fragment.[End].iChar + 1 Then
                        Dim c As Char = tb.Selection.CharBeforeStart
                        If Not Regex.IsMatch(c.ToString(), Menu.SearchPattern) Then needClose = True
                    Else
                        needClose = True
                    End If
                End If

                If needClose Then Menu.Close()
            End If
        End Sub

        Private Sub tb_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            Dim tb = TryCast(sender, FastColoredTextBox)

            If Menu.Visible Then
                If ProcessKey(e.KeyCode, e.Modifiers) Then e.Handled = True
            End If

            If Not Menu.Visible Then

                If tb.HotkeysMapping.ContainsKey(e.KeyData) AndAlso tb.HotkeysMapping(e.KeyData) = FCTBAction.AutocompleteMenu Then
                    DoAutocomplete()
                    e.Handled = True
                Else
                    If e.KeyCode = Keys.Escape AndAlso timer.Enabled Then timer.[Stop]()
                End If
            End If
        End Sub

        Private Sub AdjustScroll()
            If oldItemCount = visibleItems.Count Then Return
            Dim needHeight As Integer = ItemHeight * visibleItems.Count + 1
            Height = Math.Min(needHeight, MaximumSize.Height)
            Menu.CalcSize()
            AutoScrollMinSize = New Size(0, needHeight)
            oldItemCount = visibleItems.Count
        End Sub

        Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
            AdjustScroll()
            Dim itemHeight = itemHeight
            Dim startI As Integer = VerticalScroll.Value / itemHeight - 1
            Dim finishI As Integer = (VerticalScroll.Value + ClientSize.Height) / itemHeight + 1
            startI = Math.Max(startI, 0)
            finishI = Math.Min(finishI, visibleItems.Count)
            Dim y As Integer = 0
            Dim leftPadding As Integer = 18

            For i As Integer = startI To finishI - 1
                y = i * itemHeight - VerticalScroll.Value
                Dim item = visibleItems(i)

                If item.BackColor <> Color.Transparent Then

                    Using brush = New SolidBrush(item.BackColor)
                        e.Graphics.FillRectangle(brush, 1, y, ClientSize.Width - 1 - 1, itemHeight - 1)
                    End Using
                End If

                If ImageList IsNot Nothing AndAlso visibleItems(i).ImageIndex >= 0 Then e.Graphics.DrawImage(ImageList.Images(item.ImageIndex), 1, y)

                If i = focussedItemIndex Then

                    Using selectedBrush = New LinearGradientBrush(New Point(0, y - 3), New Point(0, y + itemHeight), Color.Transparent, SelectedColor)

                        Using pen = New Pen(SelectedColor)
                            e.Graphics.FillRectangle(selectedBrush, leftPadding, y, ClientSize.Width - 1 - leftPadding, itemHeight - 1)
                            e.Graphics.DrawRectangle(pen, leftPadding, y, ClientSize.Width - 1 - leftPadding, itemHeight - 1)
                        End Using
                    End Using
                End If

                If i = hoveredItemIndex Then

                    Using pen = New Pen(HoveredColor)
                        e.Graphics.DrawRectangle(pen, leftPadding, y, ClientSize.Width - 1 - leftPadding, itemHeight - 1)
                    End Using
                End If

                Using brush = New SolidBrush(If(item.ForeColor <> Color.Transparent, item.ForeColor, ForeColor))
                    e.Graphics.DrawString(item.ToString(), Font, brush, leftPadding, y)
                End Using
            Next
        End Sub

        Protected Overrides Sub OnScroll(ByVal se As ScrollEventArgs)
            MyBase.OnScroll(se)
            Invalidate()
        End Sub

        Protected Overrides Sub OnMouseClick(ByVal e As MouseEventArgs)
            MyBase.OnMouseClick(e)

            If e.Button = System.Windows.Forms.MouseButtons.Left Then
                focussedItemIndex = PointToItemIndex(e.Location)
                DoSelectedVisible()
                Invalidate()
            End If
        End Sub

        Protected Overrides Sub OnMouseDoubleClick(ByVal e As MouseEventArgs)
            MyBase.OnMouseDoubleClick(e)
            focussedItemIndex = PointToItemIndex(e.Location)
            Invalidate()
            OnSelecting()
        End Sub

        Friend Overridable Sub OnSelecting()
            If focussedItemIndex < 0 OrElse focussedItemIndex >= visibleItems.Count Then Return
            tb.TextSource.Manager.BeginAutoUndoCommands()

            Try
                Dim item As AutocompleteItem = FocussedItem
                Dim args As SelectingEventArgs = New SelectingEventArgs() With {
                    .Item = item,
                    .SelectedIndex = focussedItemIndex
                }
                Menu.OnSelecting(args)

                If args.Cancel Then
                    focussedItemIndex = args.SelectedIndex
                    Invalidate()
                    Return
                End If

                If Not args.Handled Then
                    Dim fragment = Menu.Fragment
                    DoAutocomplete(item, fragment)
                End If

                Menu.Close()
                Dim args2 As SelectedEventArgs = New SelectedEventArgs() With {
                    .Item = item,
                    .Tb = Menu.Fragment.tb
                }
                item.OnSelected(Menu, args2)
                Menu.OnSelected(args2)
            Finally
                tb.TextSource.Manager.EndAutoUndoCommands()
            End Try
        End Sub

        Private Sub DoAutocomplete(ByVal item As AutocompleteItem, ByVal fragment As Range)
            Dim newText As String = item.GetTextForReplace()
            Dim tb = fragment.tb
            tb.BeginAutoUndo()
            tb.TextSource.Manager.ExecuteCommand(New SelectCommand(tb.TextSource))

            If tb.Selection.ColumnSelectionMode Then
                Dim start = tb.Selection.Start
                Dim [end] = tb.Selection.[End]
                start.iChar = fragment.Start.iChar
                [end].iChar = fragment.[End].iChar
                tb.Selection.Start = start
                tb.Selection.[End] = [end]
            Else
                tb.Selection.Start = fragment.Start
                tb.Selection.[End] = fragment.[End]
            End If

            tb.InsertText(newText)
            tb.TextSource.Manager.ExecuteCommand(New SelectCommand(tb.TextSource))
            tb.EndAutoUndo()
            tb.Focus()
        End Sub

        Private Function PointToItemIndex(ByVal p As Point) As Integer
            Return (p.Y + VerticalScroll.Value) / ItemHeight
        End Function

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
            ProcessKey(keyData, Keys.None)
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Private Function ProcessKey(ByVal keyData As Keys, ByVal keyModifiers As Keys) As Boolean
            If keyModifiers = Keys.None Then

                Select Case keyData
                    Case Keys.Down
                        SelectNext(+1)
                        Return True
                    Case Keys.PageDown
                        SelectNext(+10)
                        Return True
                    Case Keys.Up
                        SelectNext(-1)
                        Return True
                    Case Keys.PageUp
                        SelectNext(-10)
                        Return True
                    Case Keys.Enter
                        OnSelecting()
                        Return True
                    Case Keys.Tab
                        If Not AllowTabKey Then Exit Select
                        OnSelecting()
                        Return True
                    Case Keys.Escape
                        Menu.Close()
                        Return True
                End Select
            End If

            Return False
        End Function

        Public Sub SelectNext(ByVal shift As Integer)
            focussedItemIndex = Math.Max(0, Math.Min(focussedItemIndex + shift, visibleItems.Count - 1))
            DoSelectedVisible()
            Invalidate()
        End Sub

        Private Sub DoSelectedVisible()
            If FocussedItem IsNot Nothing Then SetToolTip(FocussedItem)
            Dim y = focussedItemIndex * ItemHeight - VerticalScroll.Value
            If y < 0 Then VerticalScroll.Value = focussedItemIndex * ItemHeight
            If y > ClientSize.Height - ItemHeight Then VerticalScroll.Value = Math.Min(VerticalScroll.Maximum, focussedItemIndex * ItemHeight - ClientSize.Height + ItemHeight)
            AutoScrollMinSize -= New Size(1, 0)
            AutoScrollMinSize += New Size(1, 0)
        End Sub

        Private Sub SetToolTip(ByVal autocompleteItem As AutocompleteItem)
            Dim title = autocompleteItem.ToolTipTitle
            Dim text = autocompleteItem.ToolTipText

            If String.IsNullOrEmpty(title) Then
                toolTip.ToolTipTitle = Nothing
                toolTip.SetToolTip(Me, Nothing)
                Return
            End If

            If Me.Parent IsNot Nothing Then
                Dim window As IWin32Window = If(Me.Parent, Me)
                Dim location As Point

                If (Me.PointToScreen(Me.Location).X + MaxToolTipSize.Width + 105) < Screen.FromControl(Me.Parent).WorkingArea.Right Then
                    location = New Point(Right() + 5, 0)
                Else
                    location = New Point(Left() - 105 - MaximumSize.Width, 0)
                End If

                If String.IsNullOrEmpty(text) Then
                    toolTip.ToolTipTitle = Nothing
                    toolTip.Show(title, window, location.X, location.Y, ToolTipDuration)
                Else
                    toolTip.ToolTipTitle = title
                    toolTip.Show(text, window, location.X, location.Y, ToolTipDuration)
                End If
            End If
        End Sub

        Public ReadOnly Property Count As Integer
            Get
                Return visibleItems.Count
            End Get
        End Property

        Public Sub SetAutocompleteItems(ByVal items As ICollection(Of String))
            Dim list As List(Of AutocompleteItem) = New List(Of AutocompleteItem)(items.Count)

            For Each item In items
                list.Add(New AutocompleteItem(item))
            Next

            SetAutocompleteItems(list)
        End Sub

        Public Sub SetAutocompleteItems(ByVal items As IEnumerable(Of AutocompleteItem))
            sourceItems = items
        End Sub
    End Class

    Public Class SelectingEventArgs
        Inherits EventArgs

        Public Property Item As AutocompleteItem
        Public Property Cancel As Boolean
        Public Property SelectedIndex As Integer
        Public Property Handled As Boolean
    End Class

    Public Class SelectedEventArgs
        Inherits EventArgs

        Public Property Item As AutocompleteItem
        Public Property Tb As FastColoredTextBox
    End Class
End Namespace