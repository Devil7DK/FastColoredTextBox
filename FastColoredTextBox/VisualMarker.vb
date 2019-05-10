Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Drawing
Imports System.Windows.Forms

Namespace FastColoredTextBoxNS
    Public Class VisualMarker
        Public ReadOnly rectangle As Rectangle

        Public Sub New(ByVal rectangle As Rectangle)
            Me.rectangle = rectangle
        End Sub

        Public Overridable Sub Draw(ByVal gr As Graphics, ByVal pen As Pen)
        End Sub

        Public Overridable ReadOnly Property Cursor As Cursor
            Get
                Return Cursors.Hand
            End Get
        End Property
    End Class

    Public Class CollapseFoldingMarker
        Inherits VisualMarker

        Public ReadOnly iLine As Integer

        Public Sub New(ByVal iLine As Integer, ByVal rectangle As Rectangle)
            MyBase.New(rectangle)
            Me.iLine = iLine
        End Sub

        Public Sub Draw(ByVal gr As Graphics, ByVal pen As Pen, ByVal backgroundBrush As Brush, ByVal forePen As Pen)
            gr.FillRectangle(backgroundBrush, rectangle)
            gr.DrawRectangle(pen, rectangle)
            gr.DrawLine(forePen, rectangle.Left + 2, rectangle.Top + rectangle.Height / 2, rectangle.Right - 2, rectangle.Top + rectangle.Height / 2)
        End Sub
    End Class

    Public Class ExpandFoldingMarker
        Inherits VisualMarker

        Public ReadOnly iLine As Integer

        Public Sub New(ByVal iLine As Integer, ByVal rectangle As Rectangle)
            MyBase.New(rectangle)
            Me.iLine = iLine
        End Sub

        Public Sub Draw(ByVal gr As Graphics, ByVal pen As Pen, ByVal backgroundBrush As Brush, ByVal forePen As Pen)
            gr.FillRectangle(backgroundBrush, rectangle)
            gr.DrawRectangle(pen, rectangle)
            gr.DrawLine(forePen, rectangle.Left + 2, rectangle.Top + rectangle.Height / 2, rectangle.Right - 2, rectangle.Top + rectangle.Height / 2)
            gr.DrawLine(forePen, rectangle.Left + rectangle.Width / 2, rectangle.Top + 2, rectangle.Left + rectangle.Width / 2, rectangle.Bottom - 2)
        End Sub
    End Class

    Public Class FoldedAreaMarker
        Inherits VisualMarker

        Public ReadOnly iLine As Integer

        Public Sub New(ByVal iLine As Integer, ByVal rectangle As Rectangle)
            MyBase.New(rectangle)
            Me.iLine = iLine
        End Sub

        Public Overrides Sub Draw(ByVal gr As Graphics, ByVal pen As Pen)
            gr.DrawRectangle(pen, rectangle)
        End Sub
    End Class

    Public Class StyleVisualMarker
        Inherits VisualMarker

        Public Property Style As Style

        Public Sub New(ByVal rectangle As Rectangle, ByVal style As Style)
            MyBase.New(rectangle)
            Me.Style = style
        End Sub
    End Class

    Public Class VisualMarkerEventArgs
        Inherits MouseEventArgs

        Public Property Style As Style
        Public Property Marker As StyleVisualMarker

        Public Sub New(ByVal style As Style, ByVal marker As StyleVisualMarker, ByVal args As MouseEventArgs)
            MyBase.New(args.Button, args.Clicks, args.X, args.Y, args.Delta)
            Me.Style = style
            Me.Marker = marker
        End Sub
    End Class
End Namespace