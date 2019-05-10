Imports System.Drawing
Imports System
Imports System.Drawing.Drawing2D
Imports System.Collections.Generic

Namespace FastColoredTextBoxNS
    Public MustInherit Class Style
        Implements IDisposable

        Public Overridable Property IsExportable As Boolean
        Public Event VisualMarkerClick As EventHandler(Of VisualMarkerEventArgs)

        Public Sub New()
            IsExportable = True
        End Sub

        Public MustOverride Sub Draw(ByVal gr As Graphics, ByVal position As Point, ByVal range As Range)

        Public Overridable Sub OnVisualMarkerClick(ByVal tb As FastColoredTextBox, ByVal args As VisualMarkerEventArgs)
            RaiseEvent VisualMarkerClick(tb, args)
        End Sub

        Protected Overridable Sub AddVisualMarker(ByVal tb As FastColoredTextBox, ByVal marker As StyleVisualMarker)
            tb.AddVisualMarker(marker)
        End Sub

        Public Shared Function GetSizeOfRange(ByVal range As Range) As Size
            Return New Size((range.[End].iChar - range.Start.iChar) * range.tb.CharWidth, range.tb.CharHeight)
        End Function

        Public Shared Function GetRoundedRectangle(ByVal rect As Rectangle, ByVal d As Integer) As GraphicsPath
            Dim gp As GraphicsPath = New GraphicsPath()
            gp.AddArc(rect.X, rect.Y, d, d, 180, 90)
            gp.AddArc(rect.X + rect.Width - d, rect.Y, d, d, 270, 90)
            gp.AddArc(rect.X + rect.Width - d, rect.Y + rect.Height - d, d, d, 0, 90)
            gp.AddArc(rect.X, rect.Y + rect.Height - d, d, d, 90, 90)
            gp.AddLine(rect.X, rect.Y + rect.Height - d, rect.X, rect.Y + d / 2)
            Return gp
        End Function

        Public Overridable Sub Dispose()
        End Sub

        Public Overridable Function GetCSS() As String
            Return ""
        End Function

        Public Overridable Function GetRTF() As RTFStyleDescriptor
            Return New RTFStyleDescriptor()
        End Function
    End Class

    Public Class TextStyle
        Inherits Style

        Public Property ForeBrush As Brush
        Public Property BackgroundBrush As Brush
        Public Property FontStyle As FontStyle
        Public stringFormat As StringFormat

        Public Sub New(ByVal foreBrush As Brush, ByVal backgroundBrush As Brush, ByVal fontStyle As FontStyle)
            Me.ForeBrush = foreBrush
            Me.BackgroundBrush = backgroundBrush
            Me.FontStyle = fontStyle
            stringFormat = New StringFormat(StringFormatFlags.MeasureTrailingSpaces)
        End Sub

        Public Overrides Sub Draw(ByVal gr As Graphics, ByVal position As Point, ByVal range As Range)
            If BackgroundBrush IsNot Nothing Then gr.FillRectangle(BackgroundBrush, position.X, position.Y, (range.[End].iChar - range.Start.iChar) * range.tb.CharWidth, range.tb.CharHeight)

            Using f = New Font(range.tb.Font, FontStyle)
                Dim line As Line = range.tb(range.Start.iLine)
                Dim dx As Single = range.tb.CharWidth
                Dim y As Single = position.Y + range.tb.LineInterval / 2
                Dim x As Single = position.X - range.tb.CharWidth / 3
                If ForeBrush Is Nothing Then ForeBrush = New SolidBrush(range.tb.ForeColor)

                If range.tb.ImeAllowed Then

                    For i As Integer = range.Start.iChar To range.[End].iChar - 1
                        Dim size As SizeF = FastColoredTextBox.GetCharSize(f, line(i).c)
                        Dim gs = gr.Save()
                        Dim k As Single = If(size.Width > range.tb.CharWidth + 1, range.tb.CharWidth / size.Width, 1)
                        gr.TranslateTransform(x, y + (1 - k) * range.tb.CharHeight / 2)
                        gr.ScaleTransform(k, CSng(Math.Sqrt(k)))
                        gr.DrawString(line(i).c.ToString(), f, ForeBrush, 0, 0, stringFormat)
                        gr.Restore(gs)
                        x += dx
                    Next
                Else

                    For i As Integer = range.Start.iChar To range.[End].iChar - 1
                        gr.DrawString(line(i).c.ToString(), f, ForeBrush, x, y, stringFormat)
                        x += dx
                    Next
                End If
            End Using
        End Sub

        Public Overrides Function GetCSS() As String
            Dim result As String = ""

            If TypeOf BackgroundBrush Is SolidBrush Then
                Dim s = ExportToHTML.GetColorAsString((TryCast(BackgroundBrush, SolidBrush)).Color)
                If s <> "" Then result += "background-color:" & s & ";"
            End If

            If TypeOf ForeBrush Is SolidBrush Then
                Dim s = ExportToHTML.GetColorAsString((TryCast(ForeBrush, SolidBrush)).Color)
                If s <> "" Then result += "color:" & s & ";"
            End If

            If (FontStyle And FontStyle.Bold) <> 0 Then result += "font-weight:bold;"
            If (FontStyle And FontStyle.Italic) <> 0 Then result += "font-style:oblique;"
            If (FontStyle And FontStyle.Strikeout) <> 0 Then result += "text-decoration:line-through;"
            If (FontStyle And FontStyle.Underline) <> 0 Then result += "text-decoration:underline;"
            Return result
        End Function

        Public Overrides Function GetRTF() As RTFStyleDescriptor
            Dim result = New RTFStyleDescriptor()
            If TypeOf BackgroundBrush Is SolidBrush Then result.BackColor = (TryCast(BackgroundBrush, SolidBrush)).Color
            If TypeOf ForeBrush Is SolidBrush Then result.ForeColor = (TryCast(ForeBrush, SolidBrush)).Color
            If (FontStyle And FontStyle.Bold) <> 0 Then result.AdditionalTags += "\b"
            If (FontStyle And FontStyle.Italic) <> 0 Then result.AdditionalTags += "\i"
            If (FontStyle And FontStyle.Strikeout) <> 0 Then result.AdditionalTags += "\strike"
            If (FontStyle And FontStyle.Underline) <> 0 Then result.AdditionalTags += "\ul"
            Return result
        End Function
    End Class

    Public Class FoldedBlockStyle
        Inherits TextStyle

        Public Sub New(ByVal foreBrush As Brush, ByVal backgroundBrush As Brush, ByVal fontStyle As FontStyle)
            MyBase.New(foreBrush, backgroundBrush, fontStyle)
        End Sub

        Public Overrides Sub Draw(ByVal gr As Graphics, ByVal position As Point, ByVal range As Range)
            If range.[End].iChar > range.Start.iChar Then
                MyBase.Draw(gr, position, range)
                Dim firstNonSpaceSymbolX As Integer = position.X

                For i As Integer = range.Start.iChar To range.[End].iChar - 1

                    If range.tb(range.Start.iLine)(i).c <> " "c Then
                        Exit For
                    Else
                        firstNonSpaceSymbolX += range.tb.CharWidth
                    End If
                Next

                range.tb.AddVisualMarker(New FoldedAreaMarker(range.Start.iLine, New Rectangle(firstNonSpaceSymbolX, position.Y, position.X + (range.[End].iChar - range.Start.iChar) * range.tb.CharWidth - firstNonSpaceSymbolX, range.tb.CharHeight)))
            Else

                Using f As Font = New Font(range.tb.Font, FontStyle)
                    gr.DrawString("...", f, ForeBrush, range.tb.LeftIndent, position.Y - 2)
                End Using

                range.tb.AddVisualMarker(New FoldedAreaMarker(range.Start.iLine, New Rectangle(range.tb.LeftIndent + 2, position.Y, 2 * range.tb.CharHeight, range.tb.CharHeight)))
            End If
        End Sub
    End Class

    Public Class SelectionStyle
        Inherits Style

        Public Property BackgroundBrush As Brush
        Public Property ForegroundBrush As Brush

        Public Overrides Property IsExportable As Boolean
            Get
                Return False
            End Get
            Set(ByVal value As Boolean)
            End Set
        End Property

        Public Sub New(ByVal backgroundBrush As Brush, ByVal Optional foregroundBrush As Brush = Nothing)
            Me.BackgroundBrush = backgroundBrush
            Me.ForegroundBrush = foregroundBrush
        End Sub

        Public Overrides Sub Draw(ByVal gr As Graphics, ByVal position As Point, ByVal range As Range)
            If BackgroundBrush IsNot Nothing Then
                gr.SmoothingMode = SmoothingMode.None
                Dim rect = New Rectangle(position.X, position.Y, (range.[End].iChar - range.Start.iChar) * range.tb.CharWidth, range.tb.CharHeight)
                If rect.Width = 0 Then Return
                gr.FillRectangle(BackgroundBrush, rect)

                If ForegroundBrush IsNot Nothing Then
                    gr.SmoothingMode = SmoothingMode.AntiAlias
                    Dim r = New Range(range.tb, range.Start.iChar, range.Start.iLine, Math.Min(range.tb(range.[End].iLine).Count, range.[End].iChar), range.[End].iLine)

                    Using style = New TextStyle(ForegroundBrush, Nothing, FontStyle.Regular)
                        style.Draw(gr, New Point(position.X, position.Y - 1), r)
                    End Using
                End If
            End If
        End Sub
    End Class

    Public Class MarkerStyle
        Inherits Style

        Public Property BackgroundBrush As Brush

        Public Sub New(ByVal backgroundBrush As Brush)
            Me.BackgroundBrush = backgroundBrush
            IsExportable = True
        End Sub

        Public Overrides Sub Draw(ByVal gr As Graphics, ByVal position As Point, ByVal range As Range)
            If BackgroundBrush IsNot Nothing Then
                Dim rect As Rectangle = New Rectangle(position.X, position.Y, (range.[End].iChar - range.Start.iChar) * range.tb.CharWidth, range.tb.CharHeight)
                If rect.Width = 0 Then Return
                gr.FillRectangle(BackgroundBrush, rect)
            End If
        End Sub

        Public Overrides Function GetCSS() As String
            Dim result As String = ""

            If TypeOf BackgroundBrush Is SolidBrush Then
                Dim s = ExportToHTML.GetColorAsString((TryCast(BackgroundBrush, SolidBrush)).Color)
                If s <> "" Then result += "background-color:" & s & ";"
            End If

            Return result
        End Function
    End Class

    Public Class ShortcutStyle
        Inherits Style

        Public borderPen As Pen

        Public Sub New(ByVal borderPen As Pen)
            Me.borderPen = borderPen
        End Sub

        Public Overrides Sub Draw(ByVal gr As Graphics, ByVal position As Point, ByVal range As Range)
            Dim p As Point = range.tb.PlaceToPoint(range.[End])
            Dim rect As Rectangle = New Rectangle(p.X - 5, p.Y + range.tb.CharHeight - 2, 4, 3)
            gr.FillPath(Brushes.White, GetRoundedRectangle(rect, 1))
            gr.DrawPath(borderPen, GetRoundedRectangle(rect, 1))
            AddVisualMarker(range.tb, New StyleVisualMarker(New Rectangle(p.X - range.tb.CharWidth, p.Y, range.tb.CharWidth, range.tb.CharHeight), Me))
        End Sub
    End Class

    Public Class WavyLineStyle
        Inherits Style

        Private Property Pen As Pen

        Public Sub New(ByVal alpha As Integer, ByVal color As Color)
            Pen = New Pen(Color.FromArgb(alpha, color))
        End Sub

        Public Overrides Sub Draw(ByVal gr As Graphics, ByVal pos As Point, ByVal range As Range)
            Dim size = GetSizeOfRange(range)
            Dim start = New Point(pos.X, pos.Y + size.Height - 1)
            Dim [end] = New Point(pos.X + size.Width, pos.Y + size.Height - 1)
            DrawWavyLine(gr, start, [end])
        End Sub

        Private Sub DrawWavyLine(ByVal graphics As Graphics, ByVal start As Point, ByVal [end] As Point)
            If [end].X - start.X < 2 Then
                graphics.DrawLine(Pen, start, [end])
                Return
            End If

            Dim offset = -1
            Dim points = New List(Of Point)()

            For i As Integer = start.X To [end].X Step 2
                points.Add(New Point(i, start.Y + offset))
                offset = -offset
            Next

            graphics.DrawLines(Pen, points.ToArray())
        End Sub

        Public Overrides Sub Dispose()
            MyBase.Dispose()
            If Pen IsNot Nothing Then Pen.Dispose()
        End Sub
    End Class

    Public Class ReadOnlyStyle
        Inherits Style

        Public Sub New()
            IsExportable = False
        End Sub

        Public Overrides Sub Draw(ByVal gr As Graphics, ByVal position As Point, ByVal range As Range)
        End Sub
    End Class
End Namespace