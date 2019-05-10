Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Drawing.Drawing2D
Imports System.Text
Imports System.Windows.Forms

Namespace FastColoredTextBoxNS
    Public Class DocumentMap
        Inherits Control

        Public TargetChanged As EventHandler
        Private target As FastColoredTextBox
        Private scale As Single = 0.3F
        Private needRepaint As Boolean = True
        Private startPlace As Place = Place.Empty
        Private scrollbarVisible As Boolean = True

        <Description("Target FastColoredTextBox")>
        Public Property Target As FastColoredTextBox
            Get
                Return Target
            End Get
            Set(ByVal value As FastColoredTextBox)
                If target IsNot Nothing Then UnSubscribe(target)
                target = value

                If value IsNot Nothing Then
                    Subscribe(target)
                End If

                OnTargetChanged()
            End Set
        End Property

        <Description("Scale")>
        <DefaultValue(0.3F)>
        Public Property Scale As Single
            Get
                Return Scale
            End Get
            Set(ByVal value As Single)
                scale = value
                needRepaint()
            End Set
        End Property

        <Description("Scrollbar visibility")>
        <DefaultValue(True)>
        Public Property ScrollbarVisible As Boolean
            Get
                Return ScrollbarVisible
            End Get
            Set(ByVal value As Boolean)
                scrollbarVisible = value
                needRepaint()
            End Set
        End Property

        Public Sub New()
            ForeColor = Color.Maroon
            SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.ResizeRedraw, True)
            Application.Idle += AddressOf Application_Idle
        End Sub

        Private Sub Application_Idle(ByVal sender As Object, ByVal e As EventArgs)
            If needRepaint Then Invalidate()
        End Sub

        Protected Overridable Sub OnTargetChanged()
            needRepaint()
            RaiseEvent TargetChanged(Me, EventArgs.Empty)
        End Sub

        Protected Overridable Sub UnSubscribe(ByVal target As FastColoredTextBox)
            target.Scroll -= New ScrollEventHandler(AddressOf Target_Scroll)
            RemoveHandler target.SelectionChangedDelayed, New EventHandler(AddressOf Target_SelectionChanged)
            RemoveHandler target.VisibleRangeChanged, New EventHandler(AddressOf Target_VisibleRangeChanged)
        End Sub

        Protected Overridable Sub Subscribe(ByVal target As FastColoredTextBox)
            target.Scroll += New ScrollEventHandler(AddressOf Target_Scroll)
            AddHandler target.SelectionChangedDelayed, New EventHandler(AddressOf Target_SelectionChanged)
            AddHandler target.VisibleRangeChanged, New EventHandler(AddressOf Target_VisibleRangeChanged)
        End Sub

        Protected Overridable Sub Target_VisibleRangeChanged(ByVal sender As Object, ByVal e As EventArgs)
            needRepaint()
        End Sub

        Protected Overridable Sub Target_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
            needRepaint()
        End Sub

        Protected Overridable Sub Target_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs)
            needRepaint()
        End Sub

        Protected Overrides Sub OnResize(ByVal e As EventArgs)
            MyBase.OnResize(e)
            needRepaint()
        End Sub

        Public Sub NeedRepaint()
            needRepaint = True
        End Sub

        Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
            If target Is Nothing Then Return
            Dim zoom = Me.scale * 100 / target.Zoom
            If zoom <= Single.Epsilon Then Return
            Dim r = target.VisibleRange

            If startPlace.iLine > r.Start.iLine Then
                startPlace.iLine = r.Start.iLine
            Else
                Dim endP = target.PlaceToPoint(r.[End])
                endP.Offset(0, -CInt((ClientSize.Height / zoom)) + target.CharHeight)
                Dim pp = target.PointToPlace(endP)
                If pp.iLine > startPlace.iLine Then startPlace.iLine = pp.iLine
            End If

            startPlace.iChar = 0
            Dim linesCount = target.Lines.Count
            Dim sp1 = CSng(r.Start.iLine) / linesCount
            Dim sp2 = CSng(r.[End].iLine) / linesCount
            e.Graphics.ScaleTransform(zoom, zoom)
            Dim size = New SizeF(ClientSize.Width / zoom, ClientSize.Height / zoom)
            target.DrawText(e.Graphics, startPlace, size.ToSize())
            Dim p0 = target.PlaceToPoint(startPlace)
            Dim p1 = target.PlaceToPoint(r.Start)
            Dim p2 = target.PlaceToPoint(r.[End])
            Dim y1 = p1.Y - p0.Y
            Dim y2 = p2.Y + target.CharHeight - p0.Y
            e.Graphics.SmoothingMode = SmoothingMode.HighQuality

            Using brush = New SolidBrush(Color.FromArgb(50, ForeColor))

                Using pen = New Pen(brush, 1 / zoom)
                    Dim rect = New Rectangle(0, y1, CInt(((ClientSize.Width - 1) / zoom)), y2 - y1)
                    e.Graphics.FillRectangle(brush, rect)
                    e.Graphics.DrawRectangle(pen, rect)
                End Using
            End Using

            If scrollbarVisible Then
                e.Graphics.ResetTransform()
                e.Graphics.SmoothingMode = SmoothingMode.None

                Using brush = New SolidBrush(Color.FromArgb(200, ForeColor))
                    Dim rect = New RectangleF(ClientSize.Width - 3, ClientSize.Height * sp1, 2, ClientSize.Height * (sp2 - sp1))
                    e.Graphics.FillRectangle(brush, rect)
                End Using
            End If

            needRepaint = False
        End Sub

        Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
            If e.Button = System.Windows.Forms.MouseButtons.Left Then Scroll(e.Location)
            MyBase.OnMouseDown(e)
        End Sub

        Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
            If e.Button = System.Windows.Forms.MouseButtons.Left Then Scroll(e.Location)
            MyBase.OnMouseMove(e)
        End Sub

        Private Sub Scroll(ByVal point As Point)
            If target Is Nothing Then Return
            Dim zoom = Me.scale * 100 / target.Zoom
            If zoom <= Single.Epsilon Then Return
            Dim p0 = target.PlaceToPoint(startPlace)
            p0 = New Point(0, p0.Y + CInt((point.Y / zoom)))
            Dim pp = target.PointToPlace(p0)
            target.DoRangeVisible(New Range(target, pp, pp), True)
            BeginInvoke(CType(AddressOf OnScroll, MethodInvoker))
        End Sub

        Private Sub OnScroll()
            Refresh()
            target.Refresh()
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                Application.Idle -= AddressOf Application_Idle
                If target IsNot Nothing Then UnSubscribe(target)
            End If

            MyBase.Dispose(disposing)
        End Sub
    End Class
End Namespace