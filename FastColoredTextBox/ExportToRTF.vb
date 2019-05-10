Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text

Namespace FastColoredTextBoxNS
    Public Class ExportToRTF
        Public Property IncludeLineNumbers As Boolean
        Public Property UseOriginalFont As Boolean
        Private tb As FastColoredTextBox
        Private colorTable As Dictionary(Of Color, Integer) = New Dictionary(Of Color, Integer)()

        Public Sub New()
            UseOriginalFont = True
        End Sub

        Public Function GetRtf(ByVal tb As FastColoredTextBox) As String
            Me.tb = tb
            Dim sel As Range = New Range(tb)
            sel.SelectAll()
            Return GetRtf(sel)
        End Function

        Public Function GetRtf(ByVal r As Range) As String
            Me.tb = r.tb
            Dim styles = New Dictionary(Of StyleIndex, Object)()
            Dim sb = New StringBuilder()
            Dim tempSB = New StringBuilder()
            Dim currentStyleId = StyleIndex.None
            r.Normalize()
            Dim currentLine As Integer = r.Start.iLine
            styles(currentStyleId) = Nothing
            colorTable.Clear()
            Dim lineNumberColor = GetColorTableNumber(r.tb.LineNumberColor)
            If IncludeLineNumbers Then tempSB.AppendFormat("{{\cf{1} {0}}}\tab", currentLine + 1, lineNumberColor)

            For Each p As Place In r
                Dim c As Char = r.tb(p.iLine)(p.iChar)

                If c.style <> currentStyleId Then
                    Flush(sb, tempSB, currentStyleId)
                    currentStyleId = c.style
                    styles(currentStyleId) = Nothing
                End If

                If p.iLine <> currentLine Then

                    For i As Integer = currentLine To p.iLine - 1
                        tempSB.AppendLine("\line")
                        If IncludeLineNumbers Then tempSB.AppendFormat("{{\cf{1} {0}}}\tab", i + 2, lineNumberColor)
                    Next

                    currentLine = p.iLine
                End If

                Select Case c.c
                    Case "\"c
                        tempSB.Append("\\")
                    Case "{"c
                        tempSB.Append("\{")
                    Case "}"c
                        tempSB.Append("\}")
                    Case Else
                        Dim ch = c.c
                        Dim code = CInt(ch)

                        If code < 128 Then
                            tempSB.Append(c.c)
                        Else
                            tempSB.AppendFormat("{{\u{0}}}", code)
                        End If
                End Select
            Next

            Flush(sb, tempSB, currentStyleId)
            Dim list = New SortedList(Of Integer, Color)()

            For Each pair In colorTable
                list.Add(pair.Value, pair.Key)
            Next

            tempSB.Length = 0
            tempSB.AppendFormat("{{\colortbl;")

            For Each pair In list
                tempSB.Append(GetColorAsString(pair.Value) & ";")
            Next

            tempSB.AppendLine("}")

            If UseOriginalFont Then
                sb.Insert(0, String.Format("{{\fonttbl{{\f0\fmodern {0};}}}}{{\fs{1} ", tb.Font.Name, CInt((2 * tb.Font.SizeInPoints)), tb.CharHeight))
                sb.AppendLine("}")
            End If

            sb.Insert(0, tempSB.ToString())
            sb.Insert(0, "{\rtf1\ud\deff0")
            sb.AppendLine("}")
            Return sb.ToString()
        End Function

        Private Function GetRtfDescriptor(ByVal styleIndex As StyleIndex) As RTFStyleDescriptor
            Dim styles As List(Of Style) = New List(Of Style)()
            Dim textStyle As TextStyle = Nothing
            Dim mask As Integer = 1
            Dim hasTextStyle As Boolean = False

            For i As Integer = 0 To tb.Styles.Length - 1

                If tb.Styles(i) IsNot Nothing AndAlso (CInt(styleIndex) And mask) <> 0 Then

                    If tb.Styles(i).IsExportable Then
                        Dim style = tb.Styles(i)
                        styles.Add(style)
                        Dim isTextStyle As Boolean = TypeOf style Is TextStyle

                        If isTextStyle Then

                            If Not hasTextStyle OrElse tb.AllowSeveralTextStyleDrawing Then
                                hasTextStyle = True
                                textStyle = TryCast(style, TextStyle)
                            End If
                        End If
                    End If
                End If

                ''' Cannot convert ExpressionStatementSyntax, System.ArgumentOutOfRangeException: Exception of type 'System.ArgumentOutOfRangeException' was thrown.
                ''' Parameter name: op
                ''' Actual value was LeftShiftExpression.
                '''    at ICSharpCode.CodeConverter.Util.VBUtil.GetExpressionOperatorTokenKind(SyntaxKind op)
                '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node)
                '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
                '''    at ICSharpCode.CodeConverter.VB.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
                '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.VisitBinaryExpression(BinaryExpressionSyntax node)
                '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.MakeAssignmentStatement(AssignmentExpressionSyntax node)
                '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitAssignmentExpression(AssignmentExpressionSyntax node)
                '''    at Microsoft.CodeAnalysis.CSharp.Syntax.AssignmentExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
                '''    at ICSharpCode.CodeConverter.VB.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
                '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.VisitAssignmentExpression(AssignmentExpressionSyntax node)
                '''    at Microsoft.CodeAnalysis.CSharp.Syntax.AssignmentExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                '''    at ICSharpCode.CodeConverter.VB.MethodBodyVisitor.ConvertSingleExpression(ExpressionSyntax node)
                '''    at ICSharpCode.CodeConverter.VB.MethodBodyVisitor.VisitExpressionStatement(ExpressionStatementSyntax node)
                '''    at Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionStatementSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
                '''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
                '''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)
                ''' 
                ''' Input: 
                '''                 mask = mask << 1;

                ''' 
            Next

            Dim result As RTFStyleDescriptor = Nothing

            If Not hasTextStyle Then
                result = tb.DefaultStyle.GetRTF()
            Else
                result = textStyle.GetRTF()
            End If

            Return result
        End Function

        Public Shared Function GetColorAsString(ByVal color As Color) As String
            If color = Color.Transparent Then Return ""
            Return String.Format("\red{0}\green{1}\blue{2}", color.R, color.G, color.B)
        End Function

        Private Sub Flush(ByVal sb As StringBuilder, ByVal tempSB As StringBuilder, ByVal currentStyle As StyleIndex)
            If tempSB.Length = 0 Then Return
            Dim desc = GetRtfDescriptor(currentStyle)
            Dim cf = GetColorTableNumber(desc.ForeColor)
            Dim cb = GetColorTableNumber(desc.BackColor)
            Dim tags = New StringBuilder()
            If cf >= 0 Then tags.AppendFormat("\cf{0}", cf)
            If cb >= 0 Then tags.AppendFormat("\highlight{0}", cb)
            If Not String.IsNullOrEmpty(desc.AdditionalTags) Then tags.Append(desc.AdditionalTags.Trim())

            If tags.Length > 0 Then
                sb.AppendFormat("{{{0} {1}}}", tags, tempSB.ToString())
            Else
                sb.Append(tempSB.ToString())
            End If

            tempSB.Length = 0
        End Sub

        Private Function GetColorTableNumber(ByVal color As Color) As Integer
            If color.A = 0 Then Return -1
            If Not colorTable.ContainsKey(color) Then colorTable(color) = colorTable.Count + 1
            Return colorTable(color)
        End Function
    End Class

    Public Class RTFStyleDescriptor
        Public Property ForeColor As Color
        Public Property BackColor As Color
        Public Property AdditionalTags As String
    End Class
End Namespace