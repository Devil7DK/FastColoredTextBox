Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Text
Imports System.Windows.Forms

Namespace FastColoredTextBoxNS
    <System.ComponentModel.ToolboxItem(False)>
    Public Class UnfocusablePanel
        Inherits UserControl

        Public Property BackColor2 As Color
        Public Property BorderColor As Color
        Public Overloads Property Text As String
        Public Property TextAlignment As StringAlignment

        Public Sub New()
            SetStyle(ControlStyles.Selectable, False)
            SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
        End Sub

        Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
            Using brush = New LinearGradientBrush(ClientRectangle, BackColor2, BackColor, 90)
                e.Graphics.FillRectangle(brush, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1)
            End Using

            Using pen = New Pen(BorderColor)
                e.Graphics.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1)
            End Using

            If Not String.IsNullOrEmpty(Text) Then
                Dim sf As StringFormat = New StringFormat()
                sf.Alignment = TextAlignment
                sf.LineAlignment = StringAlignment.Center

                Using brush = New SolidBrush(ForeColor)
                    e.Graphics.DrawString(Text, Font, brush, New RectangleF(1, 1, ClientSize.Width - 2, ClientSize.Height - 2), sf)
                End Using
            End If
        End Sub
    End Class
End Namespace