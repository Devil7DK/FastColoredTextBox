Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Drawing.Drawing2D
Imports System.Text
Imports System.Windows.Forms

Namespace FastColoredTextBoxNS
    Partial Public Class Ruler

        Public TargetChanged As EventHandler
        <DefaultValue(GetType(Color), "ControlLight")>
        Public Property BackColor2 As Color
        <DefaultValue(GetType(Color), "DarkGray")>
        Public Property TickColor As Color
        <DefaultValue(GetType(Color), "Black")>
        Public Property CaretTickColor As Color
        Private target As FastColoredTextBox

        <Description("Target FastColoredTextBox")>
        Public Property Target As FastColoredTextBox
            Get
                Return Target
            End Get
            Set(ByVal value As FastColoredTextBox)
                If target IsNot Nothing Then UnSubscribe(target)
                target = value
                Subscribe(target)
                OnTargetChanged()
            End Set
        End Property

        Public Sub New()
            InitializeComponent()
            SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
            MinimumSize = New Size(0, 24)
            MaximumSize = New Size(Integer.MaxValue / 2, 24)
            BackColor2 = SystemColors.ControlLight
            TickColor = Color.DarkGray
            CaretTickColor = Color.Black
        End Sub

        Protected Overridable Sub OnTargetChanged()
            RaiseEvent TargetChanged(Me, EventArgs.Empty)
        End Sub

        Protected Overridable Sub UnSubscribe(ByVal target As FastColoredTextBox)
            target.Scroll -= New ScrollEventHandler(AddressOf target_Scroll)
            RemoveHandler target.SelectionChanged, New EventHandler(AddressOf target_SelectionChanged)
            RemoveHandler target.VisibleRangeChanged, New EventHandler(AddressOf target_VisibleRangeChanged)
        End Sub

        Protected Overridable Sub Subscribe(ByVal target As FastColoredTextBox)
            target.Scroll += New ScrollEventHandler(AddressOf target_Scroll)
            AddHandler target.SelectionChanged, New EventHandler(AddressOf target_SelectionChanged)
            AddHandler target.VisibleRangeChanged, New EventHandler(AddressOf target_VisibleRangeChanged)
        End Sub

        Private Sub target_VisibleRangeChanged(ByVal sender As Object, ByVal e As EventArgs)
            Invalidate()
        End Sub

        Private Sub target_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
            Invalidate()
        End Sub

        Protected Overridable Sub target_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs)
            Invalidate()
        End Sub

        Protected Overrides Sub OnResize(ByVal e As EventArgs)
            MyBase.OnResize(e)
            Invalidate()
        End Sub

        Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
            If target Is Nothing Then Return
            Dim car As Point = PointToClient(target.PointToScreen(target.PlaceToPoint(target.Selection.Start)))
            Dim fontSize As Size = TextRenderer.MeasureText("W", Font)
            Dim column As Integer = 0
            e.Graphics.FillRectangle(New LinearGradientBrush(New Rectangle(0, 0, Width, Height), BackColor, BackColor2, 270), New Rectangle(0, 0, Width, Height))
            Dim columnWidth As Single = target.CharWidth
            Dim sf = New StringFormat()
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Near
            Dim zeroPoint = target.PositionToPoint(0)
            zeroPoint = PointToClient(target.PointToScreen(zeroPoint))

            Using pen = New Pen(TickColor)

                Using textBrush = New SolidBrush(ForeColor)
                    Dim x As Single = zeroPoint.X

                    While x < Right
                        If column Mod 10 = 0 Then e.Graphics.DrawString(column.ToString(), Font, textBrush, x, 0F, sf)
                        e.Graphics.DrawLine(pen, CInt(x), fontSize.Height + (If(column Mod 5 = 0, 1, 3)), CInt(x), Height - 4)
                        x += columnWidth
                        column += 1
                    End While
                End Using
            End Using

            Using pen = New Pen(TickColor)
                e.Graphics.DrawLine(pen, New Point(car.X - 3, Height - 3), New Point(car.X + 3, Height - 3))
            End Using

            Using pen = New Pen(CaretTickColor)
                e.Graphics.DrawLine(pen, New Point(car.X - 2, fontSize.Height + 3), New Point(car.X - 2, Height - 4))
                e.Graphics.DrawLine(pen, New Point(car.X, fontSize.Height + 1), New Point(car.X, Height - 4))
                e.Graphics.DrawLine(pen, New Point(car.X + 2, fontSize.Height + 3), New Point(car.X + 2, Height - 4))
            End Using
        End Sub
    End Class
End Namespace