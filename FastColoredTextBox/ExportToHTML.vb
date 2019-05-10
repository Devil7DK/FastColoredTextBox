Imports System.Text
Imports System.Drawing
Imports System.Collections.Generic

Namespace FastColoredTextBoxNS
    Public Class ExportToHTML
        Public LineNumbersCSS As String = "<style type=""text/css""> .lineNumber{font-family : monospace; font-size : small; font-style : normal; font-weight : normal; color : Teal; background-color : ThreedFace;} </style>"
        Public Property UseNbsp As Boolean
        Public Property UseForwardNbsp As Boolean
        Public Property UseOriginalFont As Boolean
        Public Property UseStyleTag As Boolean
        Public Property UseBr As Boolean
        Public Property IncludeLineNumbers As Boolean
        Private tb As FastColoredTextBox

        Public Sub New()
            UseNbsp = True
            UseOriginalFont = True
            UseStyleTag = True
            UseBr = True
        End Sub

        Public Function GetHtml(ByVal tb As FastColoredTextBox) As String
            Me.tb = tb
            Dim sel As Range = New Range(tb)
            sel.SelectAll()
            Return GetHtml(sel)
        End Function

        Public Function GetHtml(ByVal r As Range) As String
            Me.tb = r.tb
            Dim styles As Dictionary(Of StyleIndex, Object) = New Dictionary(Of StyleIndex, Object)()
            Dim sb As StringBuilder = New StringBuilder()
            Dim tempSB As StringBuilder = New StringBuilder()
            Dim currentStyleId As StyleIndex = StyleIndex.None
            r.Normalize()
            Dim currentLine As Integer = r.Start.iLine
            styles(currentStyleId) = Nothing
            If UseOriginalFont Then sb.AppendFormat("<font style=""font-family: {0}, monospace; font-size: {1}pt; line-height: {2}px;"">", r.tb.Font.Name, r.tb.Font.SizeInPoints, r.tb.CharHeight)
            If IncludeLineNumbers Then tempSB.AppendFormat("<span class=lineNumber>{0}</span>  ", currentLine + 1)
            Dim hasNonSpace As Boolean = False

            For Each p As Place In r
                Dim c As Char = r.tb(p.iLine)(p.iChar)

                If c.style <> currentStyleId Then
                    Flush(sb, tempSB, currentStyleId)
                    currentStyleId = c.style
                    styles(currentStyleId) = Nothing
                End If

                If p.iLine <> currentLine Then

                    For i As Integer = currentLine To p.iLine - 1
                        tempSB.Append(If(UseBr, "<br>", vbCrLf))
                        If IncludeLineNumbers Then tempSB.AppendFormat("<span class=lineNumber>{0}</span>  ", i + 2)
                    Next

                    currentLine = p.iLine
                    hasNonSpace = False
                End If

                Select Case c.c
                    Case " "c

                        If (hasNonSpace OrElse Not UseForwardNbsp) AndAlso Not UseNbsp Then
                            GoTo _Select0_CaseDefault
                        End If

                        tempSB.Append("&nbsp;")
                    Case "<"c
                        tempSB.Append("&lt;")
                    Case ">"c
                        tempSB.Append("&gt;")
                    Case "&"c
                        tempSB.Append("&amp;")
                    Case Else
_Select0_CaseDefault:
                        hasNonSpace = True
                        tempSB.Append(c.c)
                End Select
            Next

            Flush(sb, tempSB, currentStyleId)
            If UseOriginalFont Then sb.Append("</font>")

            If UseStyleTag Then
                tempSB.Length = 0
                tempSB.Append("<style type=""text/css"">")

                For Each styleId In styles.Keys
                    tempSB.AppendFormat(".fctb{0}{{ {1} }}" & vbCrLf, GetStyleName(styleId), GetCss(styleId))
                Next

                tempSB.Append("</style>")
                sb.Insert(0, tempSB.ToString())
            End If

            If IncludeLineNumbers Then sb.Insert(0, LineNumbersCSS)
            Return sb.ToString()
        End Function

        Private Function GetCss(ByVal styleIndex As StyleIndex) As String
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

            Dim result As String = ""

            If Not hasTextStyle Then
                result = tb.DefaultStyle.GetCSS()
            Else
                result = textStyle.GetCSS()
            End If

            For Each style In styles
                If Not (TypeOf style Is TextStyle) Then result += style.GetCSS()
            Next

            Return result
        End Function

        Public Shared Function GetColorAsString(ByVal color As Color) As String
            If color = Color.Transparent Then Return ""
            Return String.Format("#{0:x2}{1:x2}{2:x2}", color.R, color.G, color.B)
        End Function

        Private Function GetStyleName(ByVal styleIndex As StyleIndex) As String
            Return styleIndex.ToString().Replace(" ", "").Replace(",", "")
        End Function

        Private Sub Flush(ByVal sb As StringBuilder, ByVal tempSB As StringBuilder, ByVal currentStyle As StyleIndex)
            If tempSB.Length = 0 Then Return

            If UseStyleTag Then
                sb.AppendFormat("<font class=fctb{0}>{1}</font>", GetStyleName(currentStyle), tempSB.ToString())
            Else
                Dim css As String = GetCss(currentStyle)
                If css <> "" Then sb.AppendFormat("<font style=""{0}"">", css)
                sb.Append(tempSB.ToString())
                If css <> "" Then sb.Append("</font>")
            End If

            tempSB.Length = 0
        End Sub
    End Class
End Namespace