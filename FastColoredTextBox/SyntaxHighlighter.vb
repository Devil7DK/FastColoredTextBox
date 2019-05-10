Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml

Namespace FastColoredTextBoxNS
    Public Class SyntaxHighlighter
        Implements IDisposable

        Protected Shared ReadOnly platformType As Platform = platformType.GetOperationSystemPlatform()
        Public ReadOnly BlueBoldStyle As Style = New TextStyle(Brushes.Blue, Nothing, FontStyle.Bold)
        Public ReadOnly BlueStyle As Style = New TextStyle(Brushes.Blue, Nothing, FontStyle.Regular)
        Public ReadOnly BoldStyle As Style = New TextStyle(Nothing, Nothing, FontStyle.Bold Or FontStyle.Underline)
        Public ReadOnly BrownStyle As Style = New TextStyle(Brushes.Brown, Nothing, FontStyle.Italic)
        Public ReadOnly GrayStyle As Style = New TextStyle(Brushes.Gray, Nothing, FontStyle.Regular)
        Public ReadOnly GreenStyle As Style = New TextStyle(Brushes.Green, Nothing, FontStyle.Italic)
        Public ReadOnly MagentaStyle As Style = New TextStyle(Brushes.Magenta, Nothing, FontStyle.Regular)
        Public ReadOnly MaroonStyle As Style = New TextStyle(Brushes.Maroon, Nothing, FontStyle.Regular)
        Public ReadOnly RedStyle As Style = New TextStyle(Brushes.Red, Nothing, FontStyle.Regular)
        Public ReadOnly BlackStyle As Style = New TextStyle(Brushes.Black, Nothing, FontStyle.Regular)
        Protected ReadOnly descByXMLfileNames As Dictionary(Of String, SyntaxDescriptor) = New Dictionary(Of String, SyntaxDescriptor)()
        Protected ReadOnly resilientStyles As List(Of Style) = New List(Of Style)(5)
        Protected CSharpAttributeRegex, CSharpClassNameRegex As Regex
        Protected CSharpCommentRegex1, CSharpCommentRegex2, CSharpCommentRegex3 As Regex
        Protected CSharpKeywordRegex As Regex
        Protected CSharpNumberRegex As Regex
        Protected CSharpStringRegex As Regex
        Protected HTMLAttrRegex, HTMLAttrValRegex, HTMLCommentRegex1, HTMLCommentRegex2 As Regex
        Protected HTMLEndTagRegex As Regex
        Protected HTMLEntityRegex, HTMLTagContentRegex As Regex
        Protected HTMLTagNameRegex As Regex
        Protected HTMLTagRegex As Regex
        Protected XMLAttrRegex, XMLAttrValRegex, XMLCommentRegex1, XMLCommentRegex2 As Regex
        Protected XMLEndTagRegex As Regex
        Protected XMLEntityRegex, XMLTagContentRegex As Regex
        Protected XMLTagNameRegex As Regex
        Protected XMLTagRegex As Regex
        Protected XMLCDataRegex As Regex
        Protected XMLFoldingRegex As Regex
        Protected JScriptCommentRegex1, JScriptCommentRegex2, JScriptCommentRegex3 As Regex
        Protected JScriptKeywordRegex As Regex
        Protected JScriptNumberRegex As Regex
        Protected JScriptStringRegex As Regex
        Protected LuaCommentRegex1, LuaCommentRegex2, LuaCommentRegex3 As Regex
        Protected LuaKeywordRegex As Regex
        Protected LuaNumberRegex As Regex
        Protected LuaStringRegex As Regex
        Protected LuaFunctionsRegex As Regex
        Protected PHPCommentRegex1, PHPCommentRegex2, PHPCommentRegex3 As Regex
        Protected PHPKeywordRegex1, PHPKeywordRegex2, PHPKeywordRegex3 As Regex
        Protected PHPNumberRegex As Regex
        Protected PHPStringRegex As Regex
        Protected PHPVarRegex As Regex
        Protected SQLCommentRegex1, SQLCommentRegex2, SQLCommentRegex3, SQLCommentRegex4 As Regex
        Protected SQLFunctionsRegex As Regex
        Protected SQLKeywordsRegex As Regex
        Protected SQLNumberRegex As Regex
        Protected SQLStatementsRegex As Regex
        Protected SQLStringRegex As Regex
        Protected SQLTypesRegex As Regex
        Protected SQLVarRegex As Regex
        Protected VBClassNameRegex As Regex
        Protected VBCommentRegex As Regex
        Protected VBKeywordRegex As Regex
        Protected VBNumberRegex As Regex
        Protected VBStringRegex As Regex
        Protected currentTb As FastColoredTextBox

        Public Shared ReadOnly Property RegexCompiledOption As RegexOptions
            Get

                If platformType = Platform.X86 Then
                    Return RegexOptions.Compiled
                Else
                    Return RegexOptions.None
                End If
            End Get
        End Property

        Public Sub New(ByVal currentTb As FastColoredTextBox)
            Me.currentTb = currentTb
        End Sub

        Public Sub Dispose()
            For Each desc As SyntaxDescriptor In descByXMLfileNames.Values
                desc.Dispose()
            Next
        End Sub

        Public Overridable Sub HighlightSyntax(ByVal language As Language, ByVal range As Range)
            Select Case language
                Case Language.CSharp
                    CSharpSyntaxHighlight(range)
                Case Language.VB
                    VBSyntaxHighlight(range)
                Case Language.HTML
                    HTMLSyntaxHighlight(range)
                Case Language.XML
                    XMLSyntaxHighlight(range)
                Case Language.SQL
                    SQLSyntaxHighlight(range)
                Case Language.PHP
                    PHPSyntaxHighlight(range)
                Case Language.JS
                    JScriptSyntaxHighlight(range)
                Case Language.Lua
                    LuaSyntaxHighlight(range)
                Case Else
            End Select
        End Sub

        Public Overridable Sub HighlightSyntax(ByVal XMLdescriptionFile As String, ByVal range As Range)
            Dim desc As SyntaxDescriptor = Nothing

            If Not descByXMLfileNames.TryGetValue(XMLdescriptionFile, desc) Then
                Dim doc = New XmlDocument()
                Dim file As String = XMLdescriptionFile
                If Not file.Exists(file) Then file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileName(file))
                doc.LoadXml(file.ReadAllText(file))
                desc = ParseXmlDescription(doc)
                descByXMLfileNames(XMLdescriptionFile) = desc
            End If

            HighlightSyntax(desc, range)
        End Sub

        Public Overridable Sub AutoIndentNeeded(ByVal sender As Object, ByVal args As AutoIndentEventArgs)
            Dim tb = (TryCast(sender, FastColoredTextBox))
            Dim language As Language = tb.Language

            Select Case language
                Case Language.CSharp
                    CSharpAutoIndentNeeded(sender, args)
                Case Language.VB
                    VBAutoIndentNeeded(sender, args)
                Case Language.HTML
                    HTMLAutoIndentNeeded(sender, args)
                Case Language.XML
                    XMLAutoIndentNeeded(sender, args)
                Case Language.SQL
                    SQLAutoIndentNeeded(sender, args)
                Case Language.PHP
                    PHPAutoIndentNeeded(sender, args)
                Case Language.JS
                    CSharpAutoIndentNeeded(sender, args)
                Case Language.Lua
                    LuaAutoIndentNeeded(sender, args)
                Case Else
            End Select
        End Sub

        Protected Sub PHPAutoIndentNeeded(ByVal sender As Object, ByVal args As AutoIndentEventArgs)
            If Regex.IsMatch(args.LineText, "^[^""']*\{.*\}[^""']*$") Then Return

            If Regex.IsMatch(args.LineText, "^[^""']*\{") Then
                args.ShiftNextLines = args.TabLength
                Return
            End If

            If Regex.IsMatch(args.LineText, "}[^""']*$") Then
                args.Shift = -args.TabLength
                args.ShiftNextLines = -args.TabLength
                Return
            End If

            If Regex.IsMatch(args.PrevLineText, "^\s*(if|for|foreach|while|[\}\s]*else)\b[^{]*$") Then

                If Not Regex.IsMatch(args.PrevLineText, "(;\s*$)|(;\s*//)") Then
                    args.Shift = args.TabLength
                    Return
                End If
            End If
        End Sub

        Protected Sub SQLAutoIndentNeeded(ByVal sender As Object, ByVal args As AutoIndentEventArgs)
            Dim tb = TryCast(sender, FastColoredTextBox)
            tb.CalcAutoIndentShiftByCodeFolding(sender, args)
        End Sub

        Protected Sub HTMLAutoIndentNeeded(ByVal sender As Object, ByVal args As AutoIndentEventArgs)
            Dim tb = TryCast(sender, FastColoredTextBox)
            tb.CalcAutoIndentShiftByCodeFolding(sender, args)
        End Sub

        Protected Sub XMLAutoIndentNeeded(ByVal sender As Object, ByVal args As AutoIndentEventArgs)
            Dim tb = TryCast(sender, FastColoredTextBox)
            tb.CalcAutoIndentShiftByCodeFolding(sender, args)
        End Sub

        Protected Sub VBAutoIndentNeeded(ByVal sender As Object, ByVal args As AutoIndentEventArgs)
            If Regex.IsMatch(args.LineText, "^\s*(End|EndIf|Next|Loop)\b", RegexOptions.IgnoreCase) Then
                args.Shift = -args.TabLength
                args.ShiftNextLines = -args.TabLength
                Return
            End If

            If Regex.IsMatch(args.LineText, "\b(Class|Property|Enum|Structure|Sub|Function|Namespace|Interface|Get)\b|(Set\s*\()", RegexOptions.IgnoreCase) Then
                args.ShiftNextLines = args.TabLength
                Return
            End If

            If Regex.IsMatch(args.LineText, "\b(Then)\s*\S+", RegexOptions.IgnoreCase) Then Return

            If Regex.IsMatch(args.LineText, "^\s*(If|While|For|Do|Try|With|Using|Select)\b", RegexOptions.IgnoreCase) Then
                args.ShiftNextLines = args.TabLength
                Return
            End If

            If Regex.IsMatch(args.LineText, "^\s*(Else|ElseIf|Case|Catch|Finally)\b", RegexOptions.IgnoreCase) Then
                args.Shift = -args.TabLength
                Return
            End If

            If args.PrevLineText.TrimEnd().EndsWith("_") Then
                args.Shift = args.TabLength
                Return
            End If
        End Sub

        Protected Sub CSharpAutoIndentNeeded(ByVal sender As Object, ByVal args As AutoIndentEventArgs)
            If Regex.IsMatch(args.LineText, "^[^""']*\{.*\}[^""']*$") Then Return

            If Regex.IsMatch(args.LineText, "^[^""']*\{") Then
                args.ShiftNextLines = args.TabLength
                Return
            End If

            If Regex.IsMatch(args.LineText, "}[^""']*$") Then
                args.Shift = -args.TabLength
                args.ShiftNextLines = -args.TabLength
                Return
            End If

            If Regex.IsMatch(args.LineText, "^\s*\w+\s*:\s*($|//)") AndAlso Not Regex.IsMatch(args.LineText, "^\s*default\s*:") Then
                args.Shift = -args.TabLength
                Return
            End If

            If Regex.IsMatch(args.LineText, "^\s*(case|default)\b.*:\s*($|//)") Then
                args.Shift = -args.TabLength / 2
                Return
            End If

            If Regex.IsMatch(args.PrevLineText, "^\s*(if|for|foreach|while|[\}\s]*else)\b[^{]*$") Then

                If Not Regex.IsMatch(args.PrevLineText, "(;\s*$)|(;\s*//)") Then
                    args.Shift = args.TabLength
                    Return
                End If
            End If
        End Sub

        Public Overridable Sub AddXmlDescription(ByVal descriptionFileName As String, ByVal doc As XmlDocument)
            Dim desc As SyntaxDescriptor = ParseXmlDescription(doc)
            descByXMLfileNames(descriptionFileName) = desc
        End Sub

        Public Overridable Sub AddResilientStyle(ByVal style As Style)
            If resilientStyles.Contains(style) Then Return
            currentTb.CheckStylesBufferSize()
            resilientStyles.Add(style)
        End Sub

        Public Shared Function ParseXmlDescription(ByVal doc As XmlDocument) As SyntaxDescriptor
            Dim desc = New SyntaxDescriptor()
            Dim brackets As XmlNode = doc.SelectSingleNode("doc/brackets")

            If brackets IsNot Nothing Then

                If brackets.Attributes("left") Is Nothing OrElse brackets.Attributes("right") Is Nothing OrElse brackets.Attributes("left").Value = "" OrElse brackets.Attributes("right").Value = "" Then
                    desc.leftBracket = vbNullChar
                    desc.rightBracket = vbNullChar
                Else
                    desc.leftBracket = brackets.Attributes("left").Value(0)
                    desc.rightBracket = brackets.Attributes("right").Value(0)
                End If

                If brackets.Attributes("left2") Is Nothing OrElse brackets.Attributes("right2") Is Nothing OrElse brackets.Attributes("left2").Value = "" OrElse brackets.Attributes("right2").Value = "" Then
                    desc.leftBracket2 = vbNullChar
                    desc.rightBracket2 = vbNullChar
                Else
                    desc.leftBracket2 = brackets.Attributes("left2").Value(0)
                    desc.rightBracket2 = brackets.Attributes("right2").Value(0)
                End If

                If brackets.Attributes("strategy") Is Nothing OrElse brackets.Attributes("strategy").Value = "" Then
                    desc.bracketsHighlightStrategy = BracketsHighlightStrategy.Strategy2
                Else
                    desc.bracketsHighlightStrategy = CType([Enum].Parse(GetType(BracketsHighlightStrategy), brackets.Attributes("strategy").Value), BracketsHighlightStrategy)
                End If
            End If

            Dim styleByName = New Dictionary(Of String, Style)()

            For Each style As XmlNode In doc.SelectNodes("doc/style")
                Dim s As Style = ParseStyle(style)
                styleByName(style.Attributes("name").Value) = s
                desc.styles.Add(s)
            Next

            For Each rule As XmlNode In doc.SelectNodes("doc/rule")
                desc.rules.Add(ParseRule(rule, styleByName))
            Next

            For Each folding As XmlNode In doc.SelectNodes("doc/folding")
                desc.foldings.Add(ParseFolding(folding))
            Next

            Return desc
        End Function

        Protected Shared Function ParseFolding(ByVal foldingNode As XmlNode) As FoldingDesc
            Dim folding = New FoldingDesc()
            folding.startMarkerRegex = foldingNode.Attributes("start").Value
            folding.finishMarkerRegex = foldingNode.Attributes("finish").Value
            Dim optionsA As XmlAttribute = foldingNode.Attributes("options")
            If optionsA IsNot Nothing Then folding.options = CType([Enum].Parse(GetType(RegexOptions), optionsA.Value), RegexOptions)
            Return folding
        End Function

        Protected Shared Function ParseRule(ByVal ruleNode As XmlNode, ByVal styles As Dictionary(Of String, Style)) As RuleDesc
            Dim rule = New RuleDesc()
            rule.pattern = ruleNode.InnerText
            Dim styleA As XmlAttribute = ruleNode.Attributes("style")
            Dim optionsA As XmlAttribute = ruleNode.Attributes("options")
            If styleA Is Nothing Then Throw New Exception("Rule must contain style name.")
            If Not styles.ContainsKey(styleA.Value) Then Throw New Exception("Style '" & styleA.Value & "' is not found.")
            rule.style = styles(styleA.Value)
            If optionsA IsNot Nothing Then rule.options = CType([Enum].Parse(GetType(RegexOptions), optionsA.Value), RegexOptions)
            Return rule
        End Function

        Protected Shared Function ParseStyle(ByVal styleNode As XmlNode) As Style
            Dim typeA As XmlAttribute = styleNode.Attributes("type")
            Dim colorA As XmlAttribute = styleNode.Attributes("color")
            Dim backColorA As XmlAttribute = styleNode.Attributes("backColor")
            Dim fontStyleA As XmlAttribute = styleNode.Attributes("fontStyle")
            Dim nameA As XmlAttribute = styleNode.Attributes("name")
            Dim foreBrush As SolidBrush = Nothing
            If colorA IsNot Nothing Then foreBrush = New SolidBrush(ParseColor(colorA.Value))
            Dim backBrush As SolidBrush = Nothing
            If backColorA IsNot Nothing Then backBrush = New SolidBrush(ParseColor(backColorA.Value))
            Dim fontStyle As FontStyle = FontStyle.Regular
            If fontStyleA IsNot Nothing Then fontStyle = CType([Enum].Parse(GetType(FontStyle), fontStyleA.Value), FontStyle)
            Return New TextStyle(foreBrush, backBrush, fontStyle)
        End Function

        Protected Shared Function ParseColor(ByVal s As String) As Color
            If s.StartsWith("#") Then

                If s.Length <= 7 Then
                    Return Color.FromArgb(255, Color.FromArgb(Int32.Parse(s.Substring(1), NumberStyles.AllowHexSpecifier)))
                Else
                    Return Color.FromArgb(Int32.Parse(s.Substring(1), NumberStyles.AllowHexSpecifier))
                End If
            Else
                Return Color.FromName(s)
            End If
        End Function

        Public Sub HighlightSyntax(ByVal desc As SyntaxDescriptor, ByVal range As Range)
            range.tb.ClearStylesBuffer()

            For i As Integer = 0 To desc.styles.Count - 1
                range.tb.Styles(i) = desc.styles(i)
            Next

            Dim l As Integer = desc.styles.Count

            For i As Integer = 0 To resilientStyles.Count - 1
                range.tb.Styles(l + i) = resilientStyles(i)
            Next

            Dim oldBrackets As Char() = RememberBrackets(range.tb)
            range.tb.LeftBracket = desc.leftBracket
            range.tb.RightBracket = desc.rightBracket
            range.tb.LeftBracket2 = desc.leftBracket2
            range.tb.RightBracket2 = desc.rightBracket2
            range.ClearStyle(desc.styles.ToArray())

            For Each rule As RuleDesc In desc.rules
                range.SetStyle(rule.style, rule.Regex)
            Next

            range.ClearFoldingMarkers()

            For Each folding As FoldingDesc In desc.foldings
                range.SetFoldingMarkers(folding.startMarkerRegex, folding.finishMarkerRegex, folding.options)
            Next

            RestoreBrackets(range.tb, oldBrackets)
        End Sub

        Protected Sub RestoreBrackets(ByVal tb As FastColoredTextBox, ByVal oldBrackets As Char())
            tb.LeftBracket = oldBrackets(0)
            tb.RightBracket = oldBrackets(1)
            tb.LeftBracket2 = oldBrackets(2)
            tb.RightBracket2 = oldBrackets(3)
        End Sub

        Protected Function RememberBrackets(ByVal tb As FastColoredTextBox) As Char()
            Return {tb.LeftBracket, tb.RightBracket, tb.LeftBracket2, tb.RightBracket2}
        End Function

        Protected Sub InitCShaprRegex()
            CSharpStringRegex = New Regex("
                            # Character definitions:
                            '
                            (?> # disable backtracking
                              (?:
                                \\[^\r\n]|    # escaped meta char
                                [^'\r\n]      # any character except '
                              )*
                            )
                            '?
                            |
                            # Normal string & verbatim strings definitions:
                            (?<verbatimIdentifier>@)?         # this group matches if it is an verbatim string
                            ""
                            (?> # disable backtracking
                              (?:
                                # match and consume an escaped character including escaped double quote ("") char
                                (?(verbatimIdentifier)        # if it is a verbatim string ...
                                  """"|                         #   then: only match an escaped double quote ("") char
                                  \\.                         #   else: match an escaped sequence
                                )
                                | # OR
            
                                # match any char except double quote char ("")
                                [^""]
                              )*
                            )
                            ""
                        ", RegexOptions.ExplicitCapture Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace Or RegexCompiledOption)
            CSharpCommentRegex1 = New Regex("//.*$", RegexOptions.Multiline Or RegexCompiledOption)
            CSharpCommentRegex2 = New Regex("(/\*.*?\*/)|(/\*.*)", RegexOptions.Singleline Or RegexCompiledOption)
            CSharpCommentRegex3 = New Regex("(/\*.*?\*/)|(.*\*/)", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexCompiledOption)
            CSharpNumberRegex = New Regex("\b\d+[\.]?\d*([eE]\-?\d+)?[lLdDfF]?\b|\b0x[a-fA-F\d]+\b", RegexCompiledOption)
            CSharpAttributeRegex = New Regex("^\s*(?<range>\[.+?\])\s*$", RegexOptions.Multiline Or RegexCompiledOption)
            CSharpClassNameRegex = New Regex("\b(class|struct|enum|interface)\s+(?<range>\w+?)\b", RegexCompiledOption)
            CSharpKeywordRegex = New Regex("\b(abstract|add|alias|as|ascending|async|await|base|bool|break|by|byte|case|catch|char|checked|class|const|continue|decimal|default|delegate|descending|do|double|dynamic|else|enum|equals|event|explicit|extern|false|finally|fixed|float|for|foreach|from|get|global|goto|group|if|implicit|in|int|interface|internal|into|is|join|let|lock|long|nameof|namespace|new|null|object|on|operator|orderby|out|override|params|partial|private|protected|public|readonly|ref|remove|return|sbyte|sealed|select|set|short|sizeof|stackalloc|static|static|string|struct|switch|this|throw|true|try|typeof|uint|ulong|unchecked|unsafe|ushort|using|using|value|var|virtual|void|volatile|when|where|while|yield)\b|#region\b|#endregion\b", RegexCompiledOption)
        End Sub

        Public Sub InitStyleSchema(ByVal lang As Language)
            Select Case lang
                Case Language.CSharp
                    StringStyle = BrownStyle
                    CommentStyle = GreenStyle
                    NumberStyle = MagentaStyle
                    AttributeStyle = GreenStyle
                    ClassNameStyle = BoldStyle
                    KeywordStyle = BlueStyle
                    CommentTagStyle = GrayStyle
                Case Language.VB
                    StringStyle = BrownStyle
                    CommentStyle = GreenStyle
                    NumberStyle = MagentaStyle
                    ClassNameStyle = BoldStyle
                    KeywordStyle = BlueStyle
                Case Language.HTML
                    CommentStyle = GreenStyle
                    TagBracketStyle = BlueStyle
                    TagNameStyle = MaroonStyle
                    AttributeStyle = RedStyle
                    AttributeValueStyle = BlueStyle
                    HtmlEntityStyle = RedStyle
                Case Language.XML
                    CommentStyle = GreenStyle
                    XmlTagBracketStyle = BlueStyle
                    XmlTagNameStyle = MaroonStyle
                    XmlAttributeStyle = RedStyle
                    XmlAttributeValueStyle = BlueStyle
                    XmlEntityStyle = RedStyle
                    XmlCDataStyle = BlackStyle
                Case Language.JS
                    StringStyle = BrownStyle
                    CommentStyle = GreenStyle
                    NumberStyle = MagentaStyle
                    KeywordStyle = BlueStyle
                Case Language.Lua
                    StringStyle = BrownStyle
                    CommentStyle = GreenStyle
                    NumberStyle = MagentaStyle
                    KeywordStyle = BlueBoldStyle
                    FunctionsStyle = MaroonStyle
                Case Language.PHP
                    StringStyle = RedStyle
                    CommentStyle = GreenStyle
                    NumberStyle = RedStyle
                    VariableStyle = MaroonStyle
                    KeywordStyle = MagentaStyle
                    KeywordStyle2 = BlueStyle
                    KeywordStyle3 = GrayStyle
                Case Language.SQL
                    StringStyle = RedStyle
                    CommentStyle = GreenStyle
                    NumberStyle = MagentaStyle
                    KeywordStyle = BlueBoldStyle
                    StatementsStyle = BlueBoldStyle
                    FunctionsStyle = MaroonStyle
                    VariableStyle = MaroonStyle
                    TypesStyle = BrownStyle
            End Select
        End Sub

        Public Overridable Sub CSharpSyntaxHighlight(ByVal range As Range)
            range.tb.CommentPrefix = "//"
            range.tb.LeftBracket = "("c
            range.tb.RightBracket = ")"c
            range.tb.LeftBracket2 = "{"c
            range.tb.RightBracket2 = "}"c
            range.tb.BracketsHighlightStrategy = BracketsHighlightStrategy.Strategy2
            range.tb.AutoIndentCharsPatterns = "
^\s*[\w\.]+(\s\w+)?\s*(?<range>=)\s*(?<range>[^;=]+);
^\s*(case|default)\s*[^:]*(?<range>:)\s*(?<range>[^;]+);
"
            range.ClearStyle(StringStyle, CommentStyle, NumberStyle, AttributeStyle, ClassNameStyle, KeywordStyle)
            If CSharpStringRegex Is Nothing Then InitCShaprRegex()
            range.SetStyle(StringStyle, CSharpStringRegex)
            range.SetStyle(CommentStyle, CSharpCommentRegex1)
            range.SetStyle(CommentStyle, CSharpCommentRegex2)
            range.SetStyle(CommentStyle, CSharpCommentRegex3)
            range.SetStyle(NumberStyle, CSharpNumberRegex)
            range.SetStyle(AttributeStyle, CSharpAttributeRegex)
            range.SetStyle(ClassNameStyle, CSharpClassNameRegex)
            range.SetStyle(KeywordStyle, CSharpKeywordRegex)

            For Each r As Range In range.GetRanges("^\s*///.*$", RegexOptions.Multiline)
                r.ClearStyle(StyleIndex.All)
                If HTMLTagRegex Is Nothing Then InitHTMLRegex()
                r.SetStyle(CommentStyle)

                For Each rr As Range In r.GetRanges(HTMLTagContentRegex)
                    rr.ClearStyle(StyleIndex.All)
                    rr.SetStyle(CommentTagStyle)
                Next

                For Each rr As Range In r.GetRanges("^\s*///", RegexOptions.Multiline)
                    rr.ClearStyle(StyleIndex.All)
                    rr.SetStyle(CommentTagStyle)
                Next
            Next

            range.ClearFoldingMarkers()
            range.SetFoldingMarkers("{", "}")
            range.SetFoldingMarkers("#region\b", "#endregion\b")
            range.SetFoldingMarkers("/\*", "\*/")
        End Sub

        Protected Sub InitVBRegex()
            VBStringRegex = New Regex("""""|"".*?[^\\]""", RegexCompiledOption)
            VBCommentRegex = New Regex("'.*$", RegexOptions.Multiline Or RegexCompiledOption)
            VBNumberRegex = New Regex("\b\d+[\.]?\d*([eE]\-?\d+)?\b", RegexCompiledOption)
            VBClassNameRegex = New Regex("\b(Class|Structure|Enum|Interface)[ ]+(?<range>\w+?)\b", RegexOptions.IgnoreCase Or RegexCompiledOption)
            VBKeywordRegex = New Regex("\b(AddHandler|AddressOf|Alias|And|AndAlso|As|Boolean|ByRef|Byte|ByVal|Call|Case|Catch|CBool|CByte|CChar|CDate|CDbl|CDec|Char|CInt|Class|CLng|CObj|Const|Continue|CSByte|CShort|CSng|CStr|CType|CUInt|CULng|CUShort|Date|Decimal|Declare|Default|Delegate|Dim|DirectCast|Do|Double|Each|Else|ElseIf|End|EndIf|Enum|Erase|Error|Event|Exit|False|Finally|For|Friend|Function|Get|GetType|GetXMLNamespace|Global|GoSub|GoTo|Handles|If|Implements|Imports|In|Inherits|Integer|Interface|Is|IsNot|Let|Lib|Like|Long|Loop|Me|Mod|Module|MustInherit|MustOverride|MyBase|MyClass|Namespace|Narrowing|New|Next|Not|Nothing|NotInheritable|NotOverridable|Object|Of|On|Operator|Option|Optional|Or|OrElse|Overloads|Overridable|Overrides|ParamArray|Partial|Private|Property|Protected|Public|RaiseEvent|ReadOnly|ReDim|REM|RemoveHandler|Resume|Return|SByte|Select|Set|Shadows|Shared|Short|Single|Static|Step|Stop|String|Structure|Sub|SyncLock|Then|Throw|To|True|Try|TryCast|TypeOf|UInteger|ULong|UShort|Using|Variant|Wend|When|While|Widening|With|WithEvents|WriteOnly|Xor|Region)\b|(#Const|#Else|#ElseIf|#End|#If|#Region)\b", RegexOptions.IgnoreCase Or RegexCompiledOption)
        End Sub

        Public Overridable Sub VBSyntaxHighlight(ByVal range As Range)
            range.tb.CommentPrefix = "'"
            range.tb.LeftBracket = "("c
            range.tb.RightBracket = ")"c
            range.tb.LeftBracket2 = vbNullChar
            range.tb.RightBracket2 = vbNullChar
            range.tb.AutoIndentCharsPatterns = "
^\s*[\w\.\(\)]+\s*(?<range>=)\s*(?<range>.+)
"
            range.ClearStyle(StringStyle, CommentStyle, NumberStyle, ClassNameStyle, KeywordStyle)
            If VBStringRegex Is Nothing Then InitVBRegex()
            range.SetStyle(StringStyle, VBStringRegex)
            range.SetStyle(CommentStyle, VBCommentRegex)
            range.SetStyle(NumberStyle, VBNumberRegex)
            range.SetStyle(ClassNameStyle, VBClassNameRegex)
            range.SetStyle(KeywordStyle, VBKeywordRegex)
            range.ClearFoldingMarkers()
            range.SetFoldingMarkers("#Region\b", "#End\s+Region\b", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("\b(Class|Property|Enum|Structure|Interface)[ \t]+\S+", "\bEnd (Class|Property|Enum|Structure|Interface)\b", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("^\s*(?<range>While)[ \t]+\S+", "^\s*(?<range>End While)\b", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("\b(Sub|Function)[ \t]+[^\s']+", "\bEnd (Sub|Function)\b", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("(\r|\n|^)[ \t]*(?<range>Get|Set)[ \t]*(\r|\n|$)", "\bEnd (Get|Set)\b", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("^\s*(?<range>For|For\s+Each)\b", "^\s*(?<range>Next)\b", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("^\s*(?<range>Do)\b", "^\s*(?<range>Loop)\b", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        End Sub

        Protected Sub InitHTMLRegex()
            HTMLCommentRegex1 = New Regex("(<!--.*?-->)|(<!--.*)", RegexOptions.Singleline Or RegexCompiledOption)
            HTMLCommentRegex2 = New Regex("(<!--.*?-->)|(.*-->)", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexCompiledOption)
            HTMLTagRegex = New Regex("<|/>|</|>", RegexCompiledOption)
            HTMLTagNameRegex = New Regex("<(?<range>[!\w:]+)", RegexCompiledOption)
            HTMLEndTagRegex = New Regex("</(?<range>[\w:]+)>", RegexCompiledOption)
            HTMLTagContentRegex = New Regex("<[^>]+>", RegexCompiledOption)
            HTMLAttrRegex = New Regex("(?<range>[\w\d\-]{1,20}?)='[^']*'|(?<range>[\w\d\-]{1,20})=""[^""]*""|(?<range>[\w\d\-]{1,20})=[\w\d\-]{1,20}", RegexCompiledOption)
            HTMLAttrValRegex = New Regex("[\w\d\-]{1,20}?=(?<range>'[^']*')|[\w\d\-]{1,20}=(?<range>""[^""]*"")|[\w\d\-]{1,20}=(?<range>[\w\d\-]{1,20})", RegexCompiledOption)
            HTMLEntityRegex = New Regex("\&(amp|gt|lt|nbsp|quot|apos|copy|reg|#[0-9]{1,8}|#x[0-9a-f]{1,8});", RegexCompiledOption Or RegexOptions.IgnoreCase)
        End Sub

        Public Overridable Sub HTMLSyntaxHighlight(ByVal range As Range)
            range.tb.CommentPrefix = Nothing
            range.tb.LeftBracket = "<"c
            range.tb.RightBracket = ">"c
            range.tb.LeftBracket2 = "("c
            range.tb.RightBracket2 = ")"c
            range.tb.AutoIndentCharsPatterns = ""
            range.ClearStyle(CommentStyle, TagBracketStyle, TagNameStyle, AttributeStyle, AttributeValueStyle, HtmlEntityStyle)
            If HTMLTagRegex Is Nothing Then InitHTMLRegex()
            range.SetStyle(CommentStyle, HTMLCommentRegex1)
            range.SetStyle(CommentStyle, HTMLCommentRegex2)
            range.SetStyle(TagBracketStyle, HTMLTagRegex)
            range.SetStyle(TagNameStyle, HTMLTagNameRegex)
            range.SetStyle(TagNameStyle, HTMLEndTagRegex)
            range.SetStyle(AttributeStyle, HTMLAttrRegex)
            range.SetStyle(AttributeValueStyle, HTMLAttrValRegex)
            range.SetStyle(HtmlEntityStyle, HTMLEntityRegex)
            range.ClearFoldingMarkers()
            range.SetFoldingMarkers("<head", "</head>", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("<body", "</body>", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("<table", "</table>", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("<form", "</form>", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("<div", "</div>", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("<script", "</script>", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("<tr", "</tr>", RegexOptions.IgnoreCase)
        End Sub

        Protected Sub InitXMLRegex()
            XMLCommentRegex1 = New Regex("(<!--.*?-->)|(<!--.*)", RegexOptions.Singleline Or RegexCompiledOption)
            XMLCommentRegex2 = New Regex("(<!--.*?-->)|(.*-->)", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexCompiledOption)
            XMLTagRegex = New Regex("<\?|<|/>|</|>|\?>", RegexCompiledOption)
            XMLTagNameRegex = New Regex("<[?](?<range1>[x][m][l]{1})|<(?<range>[!\w:]+)", RegexCompiledOption)
            XMLEndTagRegex = New Regex("</(?<range>[\w:]+)>", RegexCompiledOption)
            XMLTagContentRegex = New Regex("<[^>]+>", RegexCompiledOption)
            XMLAttrRegex = New Regex("(?<range>[\w\d\-\:]+)[ ]*=[ ]*'[^']*'|(?<range>[\w\d\-\:]+)[ ]*=[ ]*""[^""]*""|(?<range>[\w\d\-\:]+)[ ]*=[ ]*[\w\d\-\:]+", RegexCompiledOption)
            XMLAttrValRegex = New Regex("[\w\d\-]+?=(?<range>'[^']*')|[\w\d\-]+[ ]*=[ ]*(?<range>""[^""]*"")|[\w\d\-]+[ ]*=[ ]*(?<range>[\w\d\-]+)", RegexCompiledOption)
            XMLEntityRegex = New Regex("\&(amp|gt|lt|nbsp|quot|apos|copy|reg|#[0-9]{1,8}|#x[0-9a-f]{1,8});", RegexCompiledOption Or RegexOptions.IgnoreCase)
            XMLCDataRegex = New Regex("<!\s*\[CDATA\s*\[(?<text>(?>[^]]+|](?!]>))*)]]>", RegexCompiledOption Or RegexOptions.IgnoreCase)
            XMLFoldingRegex = New Regex("<(?<range>/?\w+)\s[^>]*?[^/]>|<(?<range>/?\w+)\s*>", RegexOptions.Singleline Or RegexCompiledOption)
        End Sub

        Public Overridable Sub XMLSyntaxHighlight(ByVal range As Range)
            range.tb.CommentPrefix = Nothing
            range.tb.LeftBracket = "<"c
            range.tb.RightBracket = ">"c
            range.tb.LeftBracket2 = "("c
            range.tb.RightBracket2 = ")"c
            range.tb.AutoIndentCharsPatterns = ""
            range.ClearStyle(CommentStyle, XmlTagBracketStyle, XmlTagNameStyle, XmlAttributeStyle, XmlAttributeValueStyle, XmlEntityStyle, XmlCDataStyle)

            If XMLTagRegex Is Nothing Then
                InitXMLRegex()
            End If

            range.SetStyle(XmlCDataStyle, XMLCDataRegex)
            range.SetStyle(CommentStyle, XMLCommentRegex1)
            range.SetStyle(CommentStyle, XMLCommentRegex2)
            range.SetStyle(XmlTagBracketStyle, XMLTagRegex)
            range.SetStyle(XmlTagNameStyle, XMLTagNameRegex)
            range.SetStyle(XmlTagNameStyle, XMLEndTagRegex)
            range.SetStyle(XmlAttributeStyle, XMLAttrRegex)
            range.SetStyle(XmlAttributeValueStyle, XMLAttrValRegex)
            range.SetStyle(XmlEntityStyle, XMLEntityRegex)
            range.ClearFoldingMarkers()
            XmlFolding(range)
        End Sub

        Private Sub XmlFolding(ByVal range As Range)
            Dim stack = New Stack(Of XmlFoldingTag)()
            Dim id = 0
            Dim fctb = range.tb

            For Each r In range.GetRanges(XMLFoldingRegex)
                Dim tagName = r.Text
                Dim iLine = r.Start.iLine

                If tagName(0) <> "/"c Then
                    Dim tag = New XmlFoldingTag With {
                        .Name = tagName,
                        .id = Math.Min(System.Threading.Interlocked.Increment(id), id - 1),
                        .startLine = r.Start.iLine
                    }
                    stack.Push(tag)
                    If String.IsNullOrEmpty(fctb(iLine).FoldingStartMarker) Then fctb(iLine).FoldingStartMarker = tag.Marker
                Else

                    If stack.Count > 0 Then
                        Dim tag = stack.Pop()

                        If iLine = tag.startLine Then
                            If fctb(iLine).FoldingStartMarker = tag.Marker Then fctb(iLine).FoldingStartMarker = Nothing
                        Else
                            If String.IsNullOrEmpty(fctb(iLine).FoldingEndMarker) Then fctb(iLine).FoldingEndMarker = tag.Marker
                        End If
                    End If
                End If
            Next
        End Sub

        Class XmlFoldingTag
            Public Name As String
            Public id As Integer
            Public startLine As Integer

            Public ReadOnly Property Marker As String
                Get
                    Return Name & id
                End Get
            End Property
        End Class

        Protected Sub InitSQLRegex()
            SQLStringRegex = New Regex("""""|''|"".*?[^\\]""|'.*?[^\\]'", RegexCompiledOption)
            SQLNumberRegex = New Regex("\b\d+[\.]?\d*([eE]\-?\d+)?\b", RegexCompiledOption)
            SQLCommentRegex1 = New Regex("--.*$", RegexOptions.Multiline Or RegexCompiledOption)
            SQLCommentRegex2 = New Regex("(/\*.*?\*/)|(/\*.*)", RegexOptions.Singleline Or RegexCompiledOption)
            SQLCommentRegex3 = New Regex("(/\*.*?\*/)|(.*\*/)", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexCompiledOption)
            SQLCommentRegex4 = New Regex("#.*$", RegexOptions.Multiline Or RegexCompiledOption)
            SQLVarRegex = New Regex("@[a-zA-Z_\d]*\b", RegexCompiledOption)
            SQLStatementsRegex = New Regex("\b(ALTER APPLICATION ROLE|ALTER ASSEMBLY|ALTER ASYMMETRIC KEY|ALTER AUTHORIZATION|ALTER BROKER PRIORITY|ALTER CERTIFICATE|ALTER CREDENTIAL|ALTER CRYPTOGRAPHIC PROVIDER|ALTER DATABASE|ALTER DATABASE AUDIT SPECIFICATION|ALTER DATABASE ENCRYPTION KEY|ALTER ENDPOINT|ALTER EVENT SESSION|ALTER FULLTEXT CATALOG|ALTER FULLTEXT INDEX|ALTER FULLTEXT STOPLIST|ALTER FUNCTION|ALTER INDEX|ALTER LOGIN|ALTER MASTER KEY|ALTER MESSAGE TYPE|ALTER PARTITION FUNCTION|ALTER PARTITION SCHEME|ALTER PROCEDURE|ALTER QUEUE|ALTER REMOTE SERVICE BINDING|ALTER RESOURCE GOVERNOR|ALTER RESOURCE POOL|ALTER ROLE|ALTER ROUTE|ALTER SCHEMA|ALTER SERVER AUDIT|ALTER SERVER AUDIT SPECIFICATION|ALTER SERVICE|ALTER SERVICE MASTER KEY|ALTER SYMMETRIC KEY|ALTER TABLE|ALTER TRIGGER|ALTER USER|ALTER VIEW|ALTER WORKLOAD GROUP|ALTER XML SCHEMA COLLECTION|BULK INSERT|CREATE AGGREGATE|CREATE APPLICATION ROLE|CREATE ASSEMBLY|CREATE ASYMMETRIC KEY|CREATE BROKER PRIORITY|CREATE CERTIFICATE|CREATE CONTRACT|CREATE CREDENTIAL|CREATE CRYPTOGRAPHIC PROVIDER|CREATE DATABASE|CREATE DATABASE AUDIT SPECIFICATION|CREATE DATABASE ENCRYPTION KEY|CREATE DEFAULT|CREATE ENDPOINT|CREATE EVENT NOTIFICATION|CREATE EVENT SESSION|CREATE FULLTEXT CATALOG|CREATE FULLTEXT INDEX|CREATE FULLTEXT STOPLIST|CREATE FUNCTION|CREATE INDEX|CREATE LOGIN|CREATE MASTER KEY|CREATE MESSAGE TYPE|CREATE PARTITION FUNCTION|CREATE PARTITION SCHEME|CREATE PROCEDURE|CREATE QUEUE|CREATE REMOTE SERVICE BINDING|CREATE RESOURCE POOL|CREATE ROLE|CREATE ROUTE|CREATE RULE|CREATE SCHEMA|CREATE SERVER AUDIT|CREATE SERVER AUDIT SPECIFICATION|CREATE SERVICE|CREATE SPATIAL INDEX|CREATE STATISTICS|CREATE SYMMETRIC KEY|CREATE SYNONYM|CREATE TABLE|CREATE TRIGGER|CREATE TYPE|CREATE USER|CREATE VIEW|CREATE WORKLOAD GROUP|CREATE XML INDEX|CREATE XML SCHEMA COLLECTION|DELETE|DISABLE TRIGGER|DROP AGGREGATE|DROP APPLICATION ROLE|DROP ASSEMBLY|DROP ASYMMETRIC KEY|DROP BROKER PRIORITY|DROP CERTIFICATE|DROP CONTRACT|DROP CREDENTIAL|DROP CRYPTOGRAPHIC PROVIDER|DROP DATABASE|DROP DATABASE AUDIT SPECIFICATION|DROP DATABASE ENCRYPTION KEY|DROP DEFAULT|DROP ENDPOINT|DROP EVENT NOTIFICATION|DROP EVENT SESSION|DROP FULLTEXT CATALOG|DROP FULLTEXT INDEX|DROP FULLTEXT STOPLIST|DROP FUNCTION|DROP INDEX|DROP LOGIN|DROP MASTER KEY|DROP MESSAGE TYPE|DROP PARTITION FUNCTION|DROP PARTITION SCHEME|DROP PROCEDURE|DROP QUEUE|DROP REMOTE SERVICE BINDING|DROP RESOURCE POOL|DROP ROLE|DROP ROUTE|DROP RULE|DROP SCHEMA|DROP SERVER AUDIT|DROP SERVER AUDIT SPECIFICATION|DROP SERVICE|DROP SIGNATURE|DROP STATISTICS|DROP SYMMETRIC KEY|DROP SYNONYM|DROP TABLE|DROP TRIGGER|DROP TYPE|DROP USER|DROP VIEW|DROP WORKLOAD GROUP|DROP XML SCHEMA COLLECTION|ENABLE TRIGGER|EXEC|EXECUTE|REPLACE|FROM|INSERT|MERGE|OPTION|OUTPUT|SELECT|TOP|TRUNCATE TABLE|UPDATE|UPDATE STATISTICS|WHERE|WITH|INTO|IN|SET)\b", RegexOptions.IgnoreCase Or RegexCompiledOption)
            SQLKeywordsRegex = New Regex("\b(ADD|ALL|AND|ANY|AS|ASC|AUTHORIZATION|BACKUP|BEGIN|BETWEEN|BREAK|BROWSE|BY|CASCADE|CHECK|CHECKPOINT|CLOSE|CLUSTERED|COLLATE|COLUMN|COMMIT|COMPUTE|CONSTRAINT|CONTAINS|CONTINUE|CROSS|CURRENT|CURRENT_DATE|CURRENT_TIME|CURSOR|DATABASE|DBCC|DEALLOCATE|DECLARE|DEFAULT|DENY|DESC|DISK|DISTINCT|DISTRIBUTED|DOUBLE|DUMP|ELSE|END|ERRLVL|ESCAPE|EXCEPT|EXISTS|EXIT|EXTERNAL|FETCH|FILE|FILLFACTOR|FOR|FOREIGN|FREETEXT|FULL|FUNCTION|GOTO|GRANT|GROUP|HAVING|HOLDLOCK|IDENTITY|IDENTITY_INSERT|IDENTITYCOL|IF|INDEX|INNER|INTERSECT|IS|JOIN|KEY|KILL|LIKE|LINENO|LOAD|NATIONAL|NOCHECK|NONCLUSTERED|NOT|NULL|OF|OFF|OFFSETS|ON|OPEN|OR|ORDER|OUTER|OVER|PERCENT|PIVOT|PLAN|PRECISION|PRIMARY|PRINT|PROC|PROCEDURE|PUBLIC|RAISERROR|READ|READTEXT|RECONFIGURE|REFERENCES|REPLICATION|RESTORE|RESTRICT|RETURN|REVERT|REVOKE|ROLLBACK|ROWCOUNT|ROWGUIDCOL|RULE|SAVE|SCHEMA|SECURITYAUDIT|SHUTDOWN|SOME|STATISTICS|TABLE|TABLESAMPLE|TEXTSIZE|THEN|TO|TRAN|TRANSACTION|TRIGGER|TSEQUAL|UNION|UNIQUE|UNPIVOT|UPDATETEXT|USE|USER|VALUES|VARYING|VIEW|WAITFOR|WHEN|WHILE|WRITETEXT)\b", RegexOptions.IgnoreCase Or RegexCompiledOption)
            SQLFunctionsRegex = New Regex("(@@CONNECTIONS|@@CPU_BUSY|@@CURSOR_ROWS|@@DATEFIRST|@@DATEFIRST|@@DBTS|@@ERROR|@@FETCH_STATUS|@@IDENTITY|@@IDLE|@@IO_BUSY|@@LANGID|@@LANGUAGE|@@LOCK_TIMEOUT|@@MAX_CONNECTIONS|@@MAX_PRECISION|@@NESTLEVEL|@@OPTIONS|@@PACKET_ERRORS|@@PROCID|@@REMSERVER|@@ROWCOUNT|@@SERVERNAME|@@SERVICENAME|@@SPID|@@TEXTSIZE|@@TRANCOUNT|@@VERSION)\b|\b(ABS|ACOS|APP_NAME|ASCII|ASIN|ASSEMBLYPROPERTY|AsymKey_ID|ASYMKEY_ID|asymkeyproperty|ASYMKEYPROPERTY|ATAN|ATN2|AVG|CASE|CAST|CEILING|Cert_ID|Cert_ID|CertProperty|CHAR|CHARINDEX|CHECKSUM_AGG|COALESCE|COL_LENGTH|COL_NAME|COLLATIONPROPERTY|COLLATIONPROPERTY|COLUMNPROPERTY|COLUMNS_UPDATED|COLUMNS_UPDATED|CONTAINSTABLE|CONVERT|COS|COT|COUNT|COUNT_BIG|CRYPT_GEN_RANDOM|CURRENT_TIMESTAMP|CURRENT_TIMESTAMP|CURRENT_USER|CURRENT_USER|CURSOR_STATUS|DATABASE_PRINCIPAL_ID|DATABASE_PRINCIPAL_ID|DATABASEPROPERTY|DATABASEPROPERTYEX|DATALENGTH|DATALENGTH|DATEADD|DATEDIFF|DATENAME|DATEPART|DAY|DB_ID|DB_NAME|DECRYPTBYASYMKEY|DECRYPTBYCERT|DECRYPTBYKEY|DECRYPTBYKEYAUTOASYMKEY|DECRYPTBYKEYAUTOCERT|DECRYPTBYPASSPHRASE|DEGREES|DENSE_RANK|DIFFERENCE|ENCRYPTBYASYMKEY|ENCRYPTBYCERT|ENCRYPTBYKEY|ENCRYPTBYPASSPHRASE|ERROR_LINE|ERROR_MESSAGE|ERROR_NUMBER|ERROR_PROCEDURE|ERROR_SEVERITY|ERROR_STATE|EVENTDATA|EXP|FILE_ID|FILE_IDEX|FILE_NAME|FILEGROUP_ID|FILEGROUP_NAME|FILEGROUPPROPERTY|FILEPROPERTY|FLOOR|fn_helpcollations|fn_listextendedproperty|fn_servershareddrives|fn_virtualfilestats|fn_virtualfilestats|FORMATMESSAGE|FREETEXTTABLE|FULLTEXTCATALOGPROPERTY|FULLTEXTSERVICEPROPERTY|GETANSINULL|GETDATE|GETUTCDATE|GROUPING|HAS_PERMS_BY_NAME|HOST_ID|HOST_NAME|IDENT_CURRENT|IDENT_CURRENT|IDENT_INCR|IDENT_INCR|IDENT_SEED|IDENTITY\(|INDEX_COL|INDEXKEY_PROPERTY|INDEXPROPERTY|IS_MEMBER|IS_OBJECTSIGNED|IS_SRVROLEMEMBER|ISDATE|ISDATE|ISNULL|ISNUMERIC|Key_GUID|Key_GUID|Key_ID|Key_ID|KEY_NAME|KEY_NAME|LEFT|LEN|LOG|LOG10|LOWER|LTRIM|MAX|MIN|MONTH|NCHAR|NEWID|NTILE|NULLIF|OBJECT_DEFINITION|OBJECT_ID|OBJECT_NAME|OBJECT_SCHEMA_NAME|OBJECTPROPERTY|OBJECTPROPERTYEX|OPENDATASOURCE|OPENQUERY|OPENROWSET|OPENXML|ORIGINAL_LOGIN|ORIGINAL_LOGIN|PARSENAME|PATINDEX|PATINDEX|PERMISSIONS|PI|POWER|PUBLISHINGSERVERNAME|PWDCOMPARE|PWDENCRYPT|QUOTENAME|RADIANS|RAND|RANK|REPLICATE|REVERSE|RIGHT|ROUND|ROW_NUMBER|ROWCOUNT_BIG|RTRIM|SCHEMA_ID|SCHEMA_ID|SCHEMA_NAME|SCHEMA_NAME|SCOPE_IDENTITY|SERVERPROPERTY|SESSION_USER|SESSION_USER|SESSIONPROPERTY|SETUSER|SIGN|SignByAsymKey|SignByCert|SIN|SOUNDEX|SPACE|SQL_VARIANT_PROPERTY|SQRT|SQUARE|STATS_DATE|STDEV|STDEVP|STR|STUFF|SUBSTRING|SUM|SUSER_ID|SUSER_NAME|SUSER_SID|SUSER_SNAME|SWITCHOFFSET|SYMKEYPROPERTY|symkeyproperty|sys\.dm_db_index_physical_stats|sys\.fn_builtin_permissions|sys\.fn_my_permissions|SYSDATETIME|SYSDATETIMEOFFSET|SYSTEM_USER|SYSTEM_USER|SYSUTCDATETIME|TAN|TERTIARY_WEIGHTS|TEXTPTR|TODATETIMEOFFSET|TRIGGER_NESTLEVEL|TYPE_ID|TYPE_NAME|TYPEPROPERTY|UNICODE|UPDATE\(|UPPER|USER_ID|USER_NAME|USER_NAME|VAR|VARP|VerifySignedByAsymKey|VerifySignedByCert|XACT_STATE|YEAR)\b", RegexOptions.IgnoreCase Or RegexCompiledOption)
            SQLTypesRegex = New Regex("\b(BIGINT|NUMERIC|BIT|SMALLINT|DECIMAL|SMALLMONEY|INT|TINYINT|MONEY|FLOAT|REAL|DATE|DATETIMEOFFSET|DATETIME2|SMALLDATETIME|DATETIME|TIME|CHAR|VARCHAR|TEXT|NCHAR|NVARCHAR|NTEXT|BINARY|VARBINARY|IMAGE|TIMESTAMP|HIERARCHYID|TABLE|UNIQUEIDENTIFIER|SQL_VARIANT|XML)\b", RegexOptions.IgnoreCase Or RegexCompiledOption)
        End Sub

        Public Overridable Sub SQLSyntaxHighlight(ByVal range As Range)
            range.tb.CommentPrefix = "--"
            range.tb.LeftBracket = "("c
            range.tb.RightBracket = ")"c
            range.tb.LeftBracket2 = vbNullChar
            range.tb.RightBracket2 = vbNullChar
            range.tb.AutoIndentCharsPatterns = ""
            range.ClearStyle(CommentStyle, StringStyle, NumberStyle, VariableStyle, StatementsStyle, KeywordStyle, FunctionsStyle, TypesStyle)
            If SQLStringRegex Is Nothing Then InitSQLRegex()
            range.SetStyle(CommentStyle, SQLCommentRegex1)
            range.SetStyle(CommentStyle, SQLCommentRegex2)
            range.SetStyle(CommentStyle, SQLCommentRegex3)
            range.SetStyle(CommentStyle, SQLCommentRegex4)
            range.SetStyle(StringStyle, SQLStringRegex)
            range.SetStyle(NumberStyle, SQLNumberRegex)
            range.SetStyle(TypesStyle, SQLTypesRegex)
            range.SetStyle(VariableStyle, SQLVarRegex)
            range.SetStyle(StatementsStyle, SQLStatementsRegex)
            range.SetStyle(KeywordStyle, SQLKeywordsRegex)
            range.SetStyle(FunctionsStyle, SQLFunctionsRegex)
            range.ClearFoldingMarkers()
            range.SetFoldingMarkers("\bBEGIN\b", "\bEND\b", RegexOptions.IgnoreCase)
            range.SetFoldingMarkers("/\*", "\*/")
        End Sub

        Protected Sub InitPHPRegex()
            PHPStringRegex = New Regex("""""|''|"".*?[^\\]""|'.*?[^\\]'", RegexCompiledOption)
            PHPNumberRegex = New Regex("\b\d+[\.]?\d*\b", RegexCompiledOption)
            PHPCommentRegex1 = New Regex("(//|#).*$", RegexOptions.Multiline Or RegexCompiledOption)
            PHPCommentRegex2 = New Regex("(/\*.*?\*/)|(/\*.*)", RegexOptions.Singleline Or RegexCompiledOption)
            PHPCommentRegex3 = New Regex("(/\*.*?\*/)|(.*\*/)", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexCompiledOption)
            PHPVarRegex = New Regex("\$[a-zA-Z_\d]*\b", RegexCompiledOption)
            PHPKeywordRegex1 = New Regex("\b(die|echo|empty|exit|eval|include|include_once|isset|list|require|require_once|return|print|unset)\b", RegexCompiledOption)
            PHPKeywordRegex2 = New Regex("\b(abstract|and|array|as|break|case|catch|cfunction|class|clone|const|continue|declare|default|do|else|elseif|enddeclare|endfor|endforeach|endif|endswitch|endwhile|extends|final|for|foreach|function|global|goto|if|implements|instanceof|interface|namespace|new|or|private|protected|public|static|switch|throw|try|use|var|while|xor)\b", RegexCompiledOption)
            PHPKeywordRegex3 = New Regex("__CLASS__|__DIR__|__FILE__|__LINE__|__FUNCTION__|__METHOD__|__NAMESPACE__", RegexCompiledOption)
        End Sub

        Public Overridable Sub PHPSyntaxHighlight(ByVal range As Range)
            range.tb.CommentPrefix = "//"
            range.tb.LeftBracket = "("c
            range.tb.RightBracket = ")"c
            range.tb.LeftBracket2 = "{"c
            range.tb.RightBracket2 = "}"c
            range.tb.BracketsHighlightStrategy = BracketsHighlightStrategy.Strategy2
            range.ClearStyle(StringStyle, CommentStyle, NumberStyle, VariableStyle, KeywordStyle, KeywordStyle2, KeywordStyle3)
            range.tb.AutoIndentCharsPatterns = "
^\s*\$[\w\.\[\]\'\""]+\s*(?<range>=)\s*(?<range>[^;]+);
"
            If PHPStringRegex Is Nothing Then InitPHPRegex()
            range.SetStyle(StringStyle, PHPStringRegex)
            range.SetStyle(CommentStyle, PHPCommentRegex1)
            range.SetStyle(CommentStyle, PHPCommentRegex2)
            range.SetStyle(CommentStyle, PHPCommentRegex3)
            range.SetStyle(NumberStyle, PHPNumberRegex)
            range.SetStyle(VariableStyle, PHPVarRegex)
            range.SetStyle(KeywordStyle, PHPKeywordRegex1)
            range.SetStyle(KeywordStyle2, PHPKeywordRegex2)
            range.SetStyle(KeywordStyle3, PHPKeywordRegex3)
            range.ClearFoldingMarkers()
            range.SetFoldingMarkers("{", "}")
            range.SetFoldingMarkers("/\*", "\*/")
        End Sub

        Protected Sub InitJScriptRegex()
            JScriptStringRegex = New Regex("""""|''|"".*?[^\\]""|'.*?[^\\]'", RegexCompiledOption)
            JScriptCommentRegex1 = New Regex("//.*$", RegexOptions.Multiline Or RegexCompiledOption)
            JScriptCommentRegex2 = New Regex("(/\*.*?\*/)|(/\*.*)", RegexOptions.Singleline Or RegexCompiledOption)
            JScriptCommentRegex3 = New Regex("(/\*.*?\*/)|(.*\*/)", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexCompiledOption)
            JScriptNumberRegex = New Regex("\b\d+[\.]?\d*([eE]\-?\d+)?[lLdDfF]?\b|\b0x[a-fA-F\d]+\b", RegexCompiledOption)
            JScriptKeywordRegex = New Regex("\b(true|false|break|case|catch|const|continue|default|delete|do|else|export|for|function|if|in|instanceof|new|null|return|switch|this|throw|try|var|void|while|with|typeof)\b", RegexCompiledOption)
        End Sub

        Public Overridable Sub JScriptSyntaxHighlight(ByVal range As Range)
            range.tb.CommentPrefix = "//"
            range.tb.LeftBracket = "("c
            range.tb.RightBracket = ")"c
            range.tb.LeftBracket2 = "{"c
            range.tb.RightBracket2 = "}"c
            range.tb.BracketsHighlightStrategy = BracketsHighlightStrategy.Strategy2
            range.tb.AutoIndentCharsPatterns = "
^\s*[\w\.]+(\s\w+)?\s*(?<range>=)\s*(?<range>[^;]+);
"
            range.ClearStyle(StringStyle, CommentStyle, NumberStyle, KeywordStyle)
            If JScriptStringRegex Is Nothing Then InitJScriptRegex()
            range.SetStyle(StringStyle, JScriptStringRegex)
            range.SetStyle(CommentStyle, JScriptCommentRegex1)
            range.SetStyle(CommentStyle, JScriptCommentRegex2)
            range.SetStyle(CommentStyle, JScriptCommentRegex3)
            range.SetStyle(NumberStyle, JScriptNumberRegex)
            range.SetStyle(KeywordStyle, JScriptKeywordRegex)
            range.ClearFoldingMarkers()
            range.SetFoldingMarkers("{", "}")
            range.SetFoldingMarkers("/\*", "\*/")
        End Sub

        Protected Sub InitLuaRegex()
            LuaStringRegex = New Regex("""""|''|"".*?[^\\]""|'.*?[^\\]'", RegexCompiledOption)
            LuaCommentRegex1 = New Regex("--.*$", RegexOptions.Multiline Or RegexCompiledOption)
            LuaCommentRegex2 = New Regex("(--\[\[.*?\]\])|(--\[\[.*)", RegexOptions.Singleline Or RegexCompiledOption)
            LuaCommentRegex3 = New Regex("(--\[\[.*?\]\])|(.*\]\])", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexCompiledOption)
            LuaNumberRegex = New Regex("\b\d+[\.]?\d*([eE]\-?\d+)?[lLdDfF]?\b|\b0x[a-fA-F\d]+\b", RegexCompiledOption)
            LuaKeywordRegex = New Regex("\b(and|break|do|else|elseif|end|false|for|function|if|in|local|nil|not|or|repeat|return|then|true|until|while)\b", RegexCompiledOption)
            LuaFunctionsRegex = New Regex("\b(assert|collectgarbage|dofile|error|getfenv|getmetatable|ipairs|load|loadfile|loadstring|module|next|pairs|pcall|print|rawequal|rawget|rawset|require|select|setfenv|setmetatable|tonumber|tostring|type|unpack|xpcall)\b", RegexCompiledOption)
        End Sub

        Public Overridable Sub LuaSyntaxHighlight(ByVal range As Range)
            range.tb.CommentPrefix = "--"
            range.tb.LeftBracket = "("c
            range.tb.RightBracket = ")"c
            range.tb.LeftBracket2 = "{"c
            range.tb.RightBracket2 = "}"c
            range.tb.BracketsHighlightStrategy = BracketsHighlightStrategy.Strategy2
            range.tb.AutoIndentCharsPatterns = "
^\s*[\w\.]+(\s\w+)?\s*(?<range>=)\s*(?<range>.+)
"
            range.ClearStyle(StringStyle, CommentStyle, NumberStyle, KeywordStyle, FunctionsStyle)
            If LuaStringRegex Is Nothing Then InitLuaRegex()
            range.SetStyle(StringStyle, LuaStringRegex)
            range.SetStyle(CommentStyle, LuaCommentRegex1)
            range.SetStyle(CommentStyle, LuaCommentRegex2)
            range.SetStyle(CommentStyle, LuaCommentRegex3)
            range.SetStyle(NumberStyle, LuaNumberRegex)
            range.SetStyle(KeywordStyle, LuaKeywordRegex)
            range.SetStyle(FunctionsStyle, LuaFunctionsRegex)
            range.ClearFoldingMarkers()
            range.SetFoldingMarkers("{", "}")
            range.SetFoldingMarkers("--\[\[", "\]\]")
        End Sub

        Protected Sub LuaAutoIndentNeeded(ByVal sender As Object, ByVal args As AutoIndentEventArgs)
            If Regex.IsMatch(args.LineText, "^\s*(end|until)\b") Then
                args.Shift = -args.TabLength
                args.ShiftNextLines = -args.TabLength
                Return
            End If

            If Regex.IsMatch(args.LineText, "\b(then)\s*\S+") Then Return

            If Regex.IsMatch(args.LineText, "^\s*(function|do|for|while|repeat|if)\b") Then
                args.ShiftNextLines = args.TabLength
                Return
            End If

            If Regex.IsMatch(args.LineText, "^\s*(else|elseif)\b", RegexOptions.IgnoreCase) Then
                args.Shift = -args.TabLength
                Return
            End If
        End Sub

        Public Property StringStyle As Style
        Public Property CommentStyle As Style
        Public Property NumberStyle As Style
        Public Property AttributeStyle As Style
        Public Property ClassNameStyle As Style
        Public Property KeywordStyle As Style
        Public Property CommentTagStyle As Style
        Public Property AttributeValueStyle As Style
        Public Property TagBracketStyle As Style
        Public Property TagNameStyle As Style
        Public Property HtmlEntityStyle As Style
        Public Property XmlAttributeStyle As Style
        Public Property XmlAttributeValueStyle As Style
        Public Property XmlTagBracketStyle As Style
        Public Property XmlTagNameStyle As Style
        Public Property XmlEntityStyle As Style
        Public Property XmlCDataStyle As Style
        Public Property VariableStyle As Style
        Public Property KeywordStyle2 As Style
        Public Property KeywordStyle3 As Style
        Public Property StatementsStyle As Style
        Public Property FunctionsStyle As Style
        Public Property TypesStyle As Style
    End Class

    Public Enum Language
        Custom
        CSharp
        VB
        HTML
        XML
        SQL
        PHP
        JS
        Lua
    End Enum
End Namespace