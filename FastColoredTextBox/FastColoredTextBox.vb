'
'  THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
'  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
'  PURPOSE.
'
'  License: GNU Lesser General Public License (LGPLv3)
'
'  Email: pavel_torgashov@ukr.net
'
'  Copyright (C) Pavel Torgashov, 2011-2016. 
' #define debug
' -------------------------------------------------------------------------------
' By default the FastColoredTextbox supports no more 16 styles at the same time.
' This restriction saves memory.
' However, you can to compile FCTB with 32 styles supporting.
' Uncomment following definition if you need 32 styles instead of 16:
'
' #define Styles32
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Design
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Forms.Design
Imports Microsoft.Win32
Imports Timer = System.Windows.Forms.Timer

Namespace FastColoredTextBoxNS

    Partial Public Class FastColoredTextBox
        Inherits UserControl
        Implements ISupportInitialize

        Friend Const minLeftIndent As Integer = 8
        Private Const maxBracketSearchIterations As Integer = 1000
        Private Const maxLinesForFolding As Integer = 3000
        Private Const minLinesForAccuracy As Integer = 100000
        Private Const WM_IME_SETCONTEXT As Integer = &H281
        Private Const WM_HSCROLL As Integer = &H114
        Private Const WM_VSCROLL As Integer = &H115
        Private Const SB_ENDSCROLL As Integer = &H8
        Public ReadOnly LineInfos As List(Of LineInfo) = New List(Of LineInfo)()
        Private ReadOnly selection_ As Range
        Private ReadOnly timer As Timer = New Timer()
        Private ReadOnly timer2 As Timer = New Timer()
        Private ReadOnly timer3 As Timer = New Timer()
        Private ReadOnly visibleMarkers As List(Of VisualMarker) = New List(Of VisualMarker)()
        Public TextHeight As Integer
        Public AllowInsertRemoveLines As Boolean = True
        Private backBrush_ As Brush
        Private bookmarks As BaseBookmarks
        Private caretVisible As Boolean
        Private changedLineColor As Color
        Private charHeight_ As Integer
        Private currentLineColor As Color
        Private defaultCursor As Cursor
        Private delayedTextChangedRange As Range
        Private descriptionFile As String
        Private endFoldingLine As Integer = -1
        Private foldingIndicatorColor As Color
        Protected foldingPairs As Dictionary(Of Integer, Integer) = New Dictionary(Of Integer, Integer)()
        Private handledChar As Boolean
        Private highlightFoldingIndicator As Boolean
        Private hints As Hints
        Private indentBackColor As Color
        Private isChanged_ As Boolean
        Private isLineSelect As Boolean
        Private isReplaceMode As Boolean
        Private language As Language
        Private lastModifiers As Keys
        Private lastMouseCoord As Point
        Private lastNavigatedDateTime As DateTime
        Private leftBracketPosition As Range
        Private leftBracketPosition2 As Range
        Private leftPadding As Integer
        Private lineInterval As Integer
        Private lineNumberColor As Color
        Private lineNumberStartValue As UInteger
        Private lineSelectFrom As Integer
        Private lines_ As TextSource
        Private m_hImc As IntPtr
        Private maxLineLength As Integer
        Private mouseIsDrag As Boolean
        Private mouseIsDragDrop As Boolean
        Private multiline As Boolean
        Protected needRecalc_ As Boolean
        Protected needRecalcWordWrap As Boolean
        Private needRecalcWordWrapInterval As Point
        Private needRecalcFoldingLines As Boolean
        Private needRiseSelectionChangedDelayed As Boolean
        Private needRiseTextChangedDelayed As Boolean
        Private needRiseVisibleRangeChangedDelayed As Boolean
        Private paddingBackColor As Color
        Private preferredLineWidth As Integer
        Private rightBracketPosition As Range
        Private rightBracketPosition2 As Range
        Private scrollBars As Boolean
        Private selectionColor As Color
        Private serviceLinesColor As Color
        Private showFoldingLines As Boolean
        Private showLineNumbers_ As Boolean
        Private sourceTextBox As FastColoredTextBox
        Private startFoldingLine As Integer = -1
        Private updating As Integer
        Private updatingRange As Range
        Private visibleRange As Range
        Private wordWrap_ As Boolean
        Private wordWrapMode As WordWrapMode = wordWrapMode.WordWrapControlWidth
        Private reservedCountOfLineNumberChars_ As Integer = 1
        Private zoom As Integer = 100
        Private localAutoScrollMinSize As Size

        Public Sub New()
            Dim prov As TypeDescriptionProvider = TypeDescriptor.GetProvider([GetType]())
            Dim theProvider As Object = prov.[GetType]().GetField("Provider", BindingFlags.NonPublic Or BindingFlags.Instance).GetValue(prov)
            If theProvider.[GetType]() <> GetType(FCTBDescriptionProvider) Then TypeDescriptor.AddProvider(New FCTBDescriptionProvider([GetType]()), [GetType]())
            SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
            Font = New Font(FontFamily.GenericMonospace, 9.75F)
            InitTextSource(CreateTextSource())
            If lines_.Count = 0 Then lines_.InsertLine(0, lines_.CreateLine())
            selection_ = New Range(Me) With {
            .Start = New Place(0, 0)
        }
            Cursor = Cursors.IBeam
            BackColor = Color.White
            lineNumberColor = Color.Teal
            indentBackColor = Color.WhiteSmoke
            serviceLinesColor = Color.Silver
            foldingIndicatorColor = Color.Green
            currentLineColor = Color.Transparent
            changedLineColor = Color.Transparent
            highlightFoldingIndicator = True
            showLineNumbers_ = True
            TabLength = 4
            FoldedBlockStyle = New FoldedBlockStyle(Brushes.Gray, Nothing, FontStyle.Regular)
            selectionColor = Color.Blue
            BracketsStyle = New MarkerStyle(New SolidBrush(Color.FromArgb(80, Color.Lime)))
            BracketsStyle2 = New MarkerStyle(New SolidBrush(Color.FromArgb(60, Color.Red)))
            DelayedEventsInterval = 100
            DelayedTextChangedInterval = 100
            AllowSeveralTextStyleDrawing = False
            LeftBracket = vbNullChar
            RightBracket = vbNullChar
            LeftBracket2 = vbNullChar
            RightBracket2 = vbNullChar
            SyntaxHighlighter = New SyntaxHighlighter(Me)
            language = language.Custom
            preferredLineWidth = 0
            needRecalc_ = True
            lastNavigatedDateTime = DateTime.Now
            AutoIndent = True
            AutoIndentExistingLines = True
            CommentPrefix = "//"
            lineNumberStartValue = 1
            multiline = True
            scrollBars = True
            AcceptsTab = True
            AcceptsReturn = True
            caretVisible = True
            CaretColor = Color.Black
            WideCaret = False
            Paddings = New Padding(0, 0, 0, 0)
            paddingBackColor = Color.Transparent
            DisabledColor = Color.FromArgb(100, 180, 180, 180)
            needRecalcFoldingLines = True
            AllowDrop = True
            FindEndOfFoldingBlockStrategy = FindEndOfFoldingBlockStrategy.Strategy1
            VirtualSpace = False
            bookmarks = New Bookmarks(Me)
            BookmarkColor = Color.PowderBlue
            ToolTip = New ToolTip()
            timer3.Interval = 500
            hints = New Hints(Me)
            selectionHighlightingForLineBreaksEnabled = True
            textAreaBorder = TextAreaBorderType.None
            textAreaBorderColor = Color.Black
            macrosManager = New MacrosManager(Me)
            HotkeysMapping = New HotkeysMapping()
            HotkeysMapping.InitDefault()
            WordWrapAutoIndent = True
            FoldedBlocks = New Dictionary(Of Integer, Integer)()
            AutoCompleteBrackets = False
            AutoIndentCharsPatterns = "^\s*[\w\.]+(\s\w+)?\s*(?<range>=)\s*(?<range>[^;=]+);
^\s*(case|default)\s*[^:]*(?<range>:)\s*(?<range>[^;]+);"
            AutoIndentChars = True
            CaretBlinking = True
            ServiceColors = New ServiceColors()
            MyBase.AutoScroll = True
            timer.Tick += AddressOf timer_Tick
            timer2.Tick += AddressOf timer2_Tick
            timer3.Tick += AddressOf timer3_Tick
            middleClickScrollingTimer.Tick += AddressOf middleClickScrollingTimer_Tick
        End Sub

        Private autoCompleteBracketsList_ As Char() = {"("c, ")"c, "{"c, "}"c, "["c, "]"c, """"c, """"c, "'"c, "'"c}

        Public Property AutoCompleteBracketsList As Char()
            Get
                Return AutoCompleteBracketsList
            End Get
            Set(ByVal value As Char())
                autoCompleteBracketsList_ = value
            End Set
        End Property

        <DefaultValue(False)>
        <Description("AutoComplete brackets.")>
        Public Property AutoCompleteBrackets As Boolean
        <Browsable(True)>
        <Description("Colors of some service visual markers.")>
        <TypeConverter(GetType(ExpandableObjectConverter))>
        Public Property ServiceColors As ServiceColors
        <Browsable(False)>
        Public Property FoldedBlocks As Dictionary(Of Integer, Integer)
        <DefaultValue(GetType(BracketsHighlightStrategy), "Strategy1")>
        <Description("Strategy of search of brackets to highlighting.")>
        Public Property BracketsHighlightStrategy As BracketsHighlightStrategy
        <DefaultValue(True)>
        <Description("Automatically shifts secondary wordwrap lines_ on the shift amount of the first line.")>
        Public Property WordWrapAutoIndent As Boolean
        <DefaultValue(0)>
        <Description("Indent of secondary wordwrap lines_ (in chars).")>
        Public Property WordWrapIndent As Integer
        Private macrosManager As MacrosManager

        <Browsable(False)>
        Public ReadOnly Property MacrosManager As MacrosManager
            Get
                Return MacrosManager
            End Get
        End Property

        <DefaultValue(True)>
        <Description("Allows drag and drop")>
        Public Overrides Property AllowDrop As Boolean
            Get
                Return MyBase.AllowDrop
            End Get
            Set(ByVal value As Boolean)
                MyBase.AllowDrop = value
            End Set
        End Property

        <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), EditorBrowsable(EditorBrowsableState.Never)>
        Public Property Hints As Hints
            Get
                Return Hints
            End Get
            Set(ByVal value As Hints)
                hints = value
            End Set
        End Property

        <Browsable(True)>
        <DefaultValue(500)>
        <Description("Delay(ms) of ToolTip.")>
        Public Property ToolTipDelay As Integer
            Get
                Return timer3.Interval
            End Get
            Set(ByVal value As Integer)
                timer3.Interval = value
            End Set
        End Property

        <Browsable(True)>
        <Description("ToolTip component.")>
        Public Property ToolTip As ToolTip
        <Browsable(True)>
        <DefaultValue(GetType(Color), "PowderBlue")>
        <Description("Color of bookmarks.")>
        Public Property BookmarkColor As Color

        <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), EditorBrowsable(EditorBrowsableState.Never)>
        Public Property Bookmarks As BaseBookmarks
            Get
                Return Bookmarks
            End Get
            Set(ByVal value As BaseBookmarks)
                bookmarks = value
            End Set
        End Property

        <DefaultValue(False)>
        <Description("Enables virtual spaces.")>
        Public Property VirtualSpace As Boolean
        <DefaultValue(FindEndOfFoldingBlockStrategy.Strategy1)>
        <Description("Strategy of search of end of folding block.")>
        Public Property FindEndOfFoldingBlockStrategy As FindEndOfFoldingBlockStrategy
        <DefaultValue(True)>
        <Description("Indicates if tab characters are accepted as input.")>
        Public Property AcceptsTab As Boolean
        <DefaultValue(True)>
        <Description("Indicates if return characters are accepted as input.")>
        Public Property AcceptsReturn As Boolean

        <DefaultValue(True)>
        <Description("Shows or hides the caret")>
        Public Property CaretVisible As Boolean
            Get
                Return CaretVisible
            End Get
            Set(ByVal value As Boolean)
                caretVisible = value
                Invalidate()
            End Set
        End Property

        <DefaultValue(True)>
        <Description("Enables caret blinking")>
        Public Property CaretBlinking As Boolean
        <DefaultValue(False)>
        Public Property ShowCaretWhenInactive As Boolean
        Private textAreaBorderColor As Color

        <DefaultValue(GetType(Color), "Black")>
        <Description("Color of border of text area")>
        Public Property TextAreaBorderColor As Color
            Get
                Return TextAreaBorderColor
            End Get
            Set(ByVal value As Color)
                textAreaBorderColor = value
                Invalidate()
            End Set
        End Property

        Private textAreaBorder As TextAreaBorderType

        <DefaultValue(GetType(TextAreaBorderType), "None")>
        <Description("Type of border of text area")>
        Public Property TextAreaBorder As TextAreaBorderType
            Get
                Return TextAreaBorder
            End Get
            Set(ByVal value As TextAreaBorderType)
                textAreaBorder = value
                Invalidate()
            End Set
        End Property

        <DefaultValue(GetType(Color), "Transparent")>
        <Description("Background color for current line. Set to Color.Transparent to hide current line highlighting")>
        Public Property CurrentLineColor As Color
            Get
                Return CurrentLineColor
            End Get
            Set(ByVal value As Color)
                currentLineColor = value
                Invalidate()
            End Set
        End Property

        <DefaultValue(GetType(Color), "Transparent")>
        <Description("Background color for highlighting of changed lines_. Set to Color.Transparent to hide changed line highlighting")>
        Public Property ChangedLineColor As Color
            Get
                Return ChangedLineColor
            End Get
            Set(ByVal value As Color)
                changedLineColor = value
                Invalidate()
            End Set
        End Property

        Public Overrides Property ForeColor As Color
            Get
                Return MyBase.ForeColor
            End Get
            Set(ByVal value As Color)
                MyBase.ForeColor = value
                lines_.InitDefaultStyle()
                Invalidate()
            End Set
        End Property

        <Browsable(False)>
        Public Property CharHeight As Integer
            Get
                Return CharHeight
            End Get
            Set(ByVal value As Integer)
                charHeight_ = value
                NeedRecalc()
                OnCharSizeChanged()
            End Set
        End Property

        <Description("Interval between lines_ in pixels")>
        <DefaultValue(0)>
        Public Property LineInterval As Integer
            Get
                Return LineInterval
            End Get
            Set(ByVal value As Integer)
                lineInterval = value
                SetFont(Font)
                Invalidate()
            End Set
        End Property

        <Browsable(False)>
        Public Property CharWidth As Integer
        <DefaultValue(4)>
        <Description("Spaces count for tab")>
        Public Property TabLength As Integer

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property IsChanged As Boolean
            Get
                Return IsChanged
            End Get
            Set(ByVal value As Boolean)
                If Not value Then lines_.ClearIsChanged()
                isChanged_ = value
            End Set
        End Property

        <Browsable(False)>
        Public Property TextVersion As Integer
        <DefaultValue(False)>
        Public Property [ReadOnly] As Boolean

        <DefaultValue(True)>
        <Description("Shows line numbers.")>
        Public Property ShowLineNumbers As Boolean
            Get
                Return ShowLineNumbers
            End Get
            Set(ByVal value As Boolean)
                showLineNumbers_ = value
                NeedRecalc()
                Invalidate()
            End Set
        End Property

        <DefaultValue(False)>
        <Description("Shows vertical lines_ between folding start line and folding end line.")>
        Public Property ShowFoldingLines As Boolean
            Get
                Return ShowFoldingLines
            End Get
            Set(ByVal value As Boolean)
                showFoldingLines = value
                Invalidate()
            End Set
        End Property

        <Browsable(False)>
        Public ReadOnly Property TextAreaRect As Rectangle
            Get
                Dim rightPaddingStartX As Integer = LeftIndent + maxLineLength * CharWidth + Paddings.Left + 1
                rightPaddingStartX = Math.Max(ClientSize.Width - Paddings.Right, rightPaddingStartX)
                Dim bottomPaddingStartY As Integer = TextHeight + Paddings.Top
                bottomPaddingStartY = Math.Max(ClientSize.Height - Paddings.Bottom, bottomPaddingStartY)
                Dim top = Math.Max(0, Paddings.Top - 1) - VerticalScroll.Value
                Dim left = LeftIndent - HorizontalScroll.Value - 2 + Math.Max(0, Paddings.Left - 1)
                Dim rect = Rectangle.FromLTRB(left, top, rightPaddingStartX - HorizontalScroll.Value, bottomPaddingStartY - VerticalScroll.Value)
                Return rect
            End Get
        End Property

        <DefaultValue(GetType(Color), "Teal")>
        <Description("Color of line numbers.")>
        Public Property LineNumberColor As Color
            Get
                Return LineNumberColor
            End Get
            Set(ByVal value As Color)
                lineNumberColor = value
                Invalidate()
            End Set
        End Property

        <DefaultValue(GetType(UInteger), "1")>
        <Description("Start value of first line number.")>
        Public Property LineNumberStartValue As UInteger
            Get
                Return LineNumberStartValue
            End Get
            Set(ByVal value As UInteger)
                lineNumberStartValue = value
                needRecalc_ = True
                Invalidate()
            End Set
        End Property

        <DefaultValue(GetType(Color), "WhiteSmoke")>
        <Description("Background color of indent area")>
        Public Property IndentBackColor As Color
            Get
                Return IndentBackColor
            End Get
            Set(ByVal value As Color)
                indentBackColor = value
                Invalidate()
            End Set
        End Property

        <DefaultValue(GetType(Color), "Transparent")>
        <Description("Background color of padding area")>
        Public Property PaddingBackColor As Color
            Get
                Return PaddingBackColor
            End Get
            Set(ByVal value As Color)
                paddingBackColor = value
                Invalidate()
            End Set
        End Property

        <DefaultValue(GetType(Color), "100;180;180;180")>
        <Description("Color of disabled component")>
        Public Property DisabledColor As Color
        <DefaultValue(GetType(Color), "Black")>
        <Description("Color of caret.")>
        Public Property CaretColor As Color
        <DefaultValue(False)>
        <Description("Wide caret.")>
        Public Property WideCaret As Boolean

        <DefaultValue(GetType(Color), "Silver")>
        <Description("Color of service lines_ (folding lines_, borders of blocks etc.)")>
        Public Property ServiceLinesColor As Color
            Get
                Return ServiceLinesColor
            End Get
            Set(ByVal value As Color)
                serviceLinesColor = value
                Invalidate()
            End Set
        End Property

        <Browsable(True)>
        <Description("Paddings of text area.")>
        Public Property Paddings As Padding

        <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), EditorBrowsable(EditorBrowsableState.Never)>
        Public Overloads Property Padding As Padding
            Get
                Throw New NotImplementedException()
            End Get
            Set(ByVal value As Padding)
                Throw New NotImplementedException()
            End Set
        End Property

        <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), EditorBrowsable(EditorBrowsableState.Never)>
        Public Overloads Property RightToLeft As Boolean
            Get
                Throw New NotImplementedException()
            End Get
            Set(ByVal value As Boolean)
                Throw New NotImplementedException()
            End Set
        End Property

        <DefaultValue(GetType(Color), "Green")>
        <Description("Color of folding area indicator.")>
        Public Property FoldingIndicatorColor As Color
            Get
                Return FoldingIndicatorColor
            End Get
            Set(ByVal value As Color)
                foldingIndicatorColor = value
                Invalidate()
            End Set
        End Property

        <DefaultValue(True)>
        <Description("Enables folding indicator (left vertical line between folding bounds)")>
        Public Property HighlightFoldingIndicator As Boolean
            Get
                Return HighlightFoldingIndicator
            End Get
            Set(ByVal value As Boolean)
                highlightFoldingIndicator = value
                Invalidate()
            End Set
        End Property

        <Browsable(False)>
        <Description("Left distance to text beginning.")>
        Public Property LeftIndent As Integer

        <DefaultValue(0)>
        <Description("Width of left service area (in pixels)")>
        Public Property LeftPadding As Integer
            Get
                Return LeftPadding
            End Get
            Set(ByVal value As Integer)
                leftPadding = value
                Invalidate()
            End Set
        End Property

        <DefaultValue(0)>
        <Description("This property draws vertical line after defined char position. Set to 0 for disable drawing of vertical line.")>
        Public Property PreferredLineWidth As Integer
            Get
                Return PreferredLineWidth
            End Get
            Set(ByVal value As Integer)
                preferredLineWidth = value
                Invalidate()
            End Set
        End Property

        <Browsable(False)>
        Public ReadOnly Property Styles As Style()
            Get
                Return lines_.Styles
            End Get
        End Property

        <Description("Here you can change hotkeys for FastColoredTextBox.")>
        <Editor(GetType(HotkeysEditor), GetType(UITypeEditor))>
        <DefaultValue("Tab=IndentIncrease, Escape=ClearHints, PgUp=GoPageUp, PgDn=GoPageDown, End=GoEnd, Home=GoHome, Left=GoLeft, Up=GoUp, Right=GoRight, Down=GoDown, Ins=ReplaceMode, Del=DeleteCharRight, F3=FindNext, Shift+Tab=IndentDecrease, Shift+PgUp=GoPageUpWithSelection, Shift+PgDn=GoPageDownWithSelection, Shift+End=GoEndWithSelection, Shift+Home=GoHomeWithSelection, Shift+Left=GoLeftWithSelection, Shift+Up=GoUpWithSelection, Shift+Right=GoRightWithSelection, Shift+Down=GoDownWithSelection, Shift+Ins=Paste, Shift+Del=Cut, Ctrl+Back=ClearWordLeft, Ctrl+Space=AutocompleteMenu, Ctrl+End=GoLastLine, Ctrl+Home=GoFirstLine, Ctrl+Left=GoWordLeft, Ctrl+Up=ScrollUp, Ctrl+Right=GoWordRight, Ctrl+Down=ScrollDown, Ctrl+Ins=Copy, Ctrl+Del=ClearWordRight, Ctrl+0=ZoomNormal, Ctrl+A=SelectAll, Ctrl+B=BookmarkLine, Ctrl+C=Copy, Ctrl+E=MacroExecute, Ctrl+F=FindDialog, Ctrl+G=GoToDialog, Ctrl+H=ReplaceDialog, Ctrl+I=AutoIndentChars, Ctrl+M=MacroRecord, Ctrl+N=GoNextBookmark, Ctrl+R=Redo, Ctrl+U=UpperCase, Ctrl+V=Paste, Ctrl+X=Cut, Ctrl+Z=Undo, Ctrl+Add=ZoomIn, Ctrl+Subtract=ZoomOut, Ctrl+OemMinus=NavigateBackward, Ctrl+Shift+End=GoLastLineWithSelection, Ctrl+Shift+Home=GoFirstLineWithSelection, Ctrl+Shift+Left=GoWordLeftWithSelection, Ctrl+Shift+Right=GoWordRightWithSelection, Ctrl+Shift+B=UnbookmarkLine, Ctrl+Shift+C=CommentSelected, Ctrl+Shift+N=GoPrevBookmark, Ctrl+Shift+U=LowerCase, Ctrl+Shift+OemMinus=NavigateForward, Alt+Back=Undo, Alt+Up=MoveSelectedLinesUp, Alt+Down=MoveSelectedLinesDown, Alt+F=FindChar, Alt+Shift+Left=GoLeft_ColumnSelectionMode, Alt+Shift+Up=GoUp_ColumnSelectionMode, Alt+Shift+Right=GoRight_ColumnSelectionMode, Alt+Shift+Down=GoDown_ColumnSelectionMode")>
        Public Property Hotkeys As String
            Get
                Return HotkeysMapping.ToString()
            End Get
            Set(ByVal value As String)
                HotkeysMapping = HotkeysMapping.Parse(value)
            End Set
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property HotkeysMapping As HotkeysMapping

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property DefaultStyle As TextStyle
            Get
                Return lines_.DefaultStyle
            End Get
            Set(ByVal value As TextStyle)
                lines_.DefaultStyle = value
            End Set
        End Property

        <Browsable(False)>
        Public Property SelectionStyle As SelectionStyle
        <Browsable(False)>
        Public Property FoldedBlockStyle As TextStyle
        <Browsable(False)>
        Public Property BracketsStyle As MarkerStyle
        <Browsable(False)>
        Public Property BracketsStyle2 As MarkerStyle
        <DefaultValue(vbNullChar)>
        <Description("Opening bracket for brackets highlighting. Set to '\x0' for disable brackets highlighting.")>
        Public Property LeftBracket As Char
        <DefaultValue(vbNullChar)>
        <Description("Closing bracket for brackets highlighting. Set to '\x0' for disable brackets highlighting.")>
        Public Property RightBracket As Char
        <DefaultValue(vbNullChar)>
        <Description("Alternative opening bracket for brackets highlighting. Set to '\x0' for disable brackets highlighting.")>
        Public Property LeftBracket2 As Char
        <DefaultValue(vbNullChar)>
        <Description("Alternative closing bracket for brackets highlighting. Set to '\x0' for disable brackets highlighting.")>
        Public Property RightBracket2 As Char
        <DefaultValue("//")>
        <Description("Comment line prefix.")>
        Public Property CommentPrefix As String
        <DefaultValue(GetType(HighlightingRangeType), "ChangedRange")>
        <Description("This property specifies which part of the text will be highlighted as you type.")>
        Public Property HighlightingRangeType As HighlightingRangeType

        <Browsable(False)>
        Public Property IsReplaceMode As Boolean
            Get
                Return IsReplaceMode AndAlso selection_.IsEmpty AndAlso (Not selection_.ColumnSelectionMode) AndAlso selection_.Start.iChar < lines_(selection_.Start.iLine).Count
            End Get
            Set(ByVal value As Boolean)
                isReplaceMode = value
            End Set
        End Property

        <Browsable(True)>
        <DefaultValue(False)>
        <Description("Allows text rendering several styles same time.")>
        Public Property AllowSeveralTextStyleDrawing As Boolean

        <Browsable(True)>
        <DefaultValue(True)>
        <Description("Allows to record macros.")>
        Public Property AllowMacroRecording As Boolean
            Get
                Return macrosManager.AllowMacroRecordingByUser
            End Get
            Set(ByVal value As Boolean)
                macrosManager.AllowMacroRecordingByUser = value
            End Set
        End Property

        <DefaultValue(True)>
        <Description("Allows auto indent. Inserts spaces before line chars.")>
        Public Property AutoIndent As Boolean
        <DefaultValue(True)>
        <Description("Does autoindenting in existing lines_. It works only if AutoIndent is True.")>
        Public Property AutoIndentExistingLines As Boolean

        <Browsable(True)>
        <DefaultValue(100)>
        <Description("Minimal delay(ms) for delayed events (except TextChangedDelayed).")>
        Public Property DelayedEventsInterval As Integer
            Get
                Return timer.Interval
            End Get
            Set(ByVal value As Integer)
                timer.Interval = value
            End Set
        End Property

        <Browsable(True)>
        <DefaultValue(100)>
        <Description("Minimal delay(ms) for TextChangedDelayed event.")>
        Public Property DelayedTextChangedInterval As Integer
            Get
                Return timer2.Interval
            End Get
            Set(ByVal value As Integer)
                timer2.Interval = value
            End Set
        End Property

        <Browsable(True)>
        <DefaultValue(GetType(Language), "Custom")>
        <Description("Language for highlighting by built-in highlighter.")>
        Public Property Language As Language
            Get
                Return Language
            End Get
            Set(ByVal value As Language)
                language = value
                If SyntaxHighlighter IsNot Nothing Then SyntaxHighlighter.InitStyleSchema(language)
                Invalidate()
            End Set
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property SyntaxHighlighter As SyntaxHighlighter

        <Browsable(True)>
        <DefaultValue(Nothing)>
        <Editor(GetType(FileNameEditor), GetType(UITypeEditor))>
        <Description("XML file with description of syntax highlighting. This property works only with Language == Language.Custom.")>
        Public Property DescriptionFile As String
            Get
                Return DescriptionFile
            End Get
            Set(ByVal value As String)
                descriptionFile = value
                Invalidate()
            End Set
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public ReadOnly Property LeftBracketPosition As Range
            Get
                Return LeftBracketPosition
            End Get
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public ReadOnly Property RightBracketPosition As Range
            Get
                Return RightBracketPosition
            End Get
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public ReadOnly Property LeftBracketPosition2 As Range
            Get
                Return LeftBracketPosition2
            End Get
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public ReadOnly Property RightBracketPosition2 As Range
            Get
                Return RightBracketPosition2
            End Get
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public ReadOnly Property StartFoldingLine As Integer
            Get
                Return StartFoldingLine
            End Get
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public ReadOnly Property EndFoldingLine As Integer
            Get
                Return EndFoldingLine
            End Get
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property TextSource As TextSource
            Get
                Return lines_
            End Get
            Set(ByVal value As TextSource)
                InitTextSource(value)
            End Set
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public ReadOnly Property HasSourceTextBox As Boolean
            Get
                Return sourceTextBox IsNot Nothing
            End Get
        End Property

        <Browsable(True)>
        <DefaultValue(Nothing)>
        <Description("Allows to get text from other FastColoredTextBox.")>
        Public Property SourceTextBox As FastColoredTextBox
            Get
                Return SourceTextBox
            End Get
            Set(ByVal value As FastColoredTextBox)
                If value = sourceTextBox Then Return
                sourceTextBox = value

                If sourceTextBox Is Nothing Then
                    InitTextSource(CreateTextSource())
                    lines_.InsertLine(0, TextSource.CreateLine())
                    isChanged_ = False
                Else
                    InitTextSource(sourceTextBox.TextSource)
                    isChanged_ = False
                End If

                Invalidate()
            End Set
        End Property

        <Browsable(False)>
        Public ReadOnly Property VisibleRange As Range
            Get
                If VisibleRange IsNot Nothing Then Return VisibleRange
                Return GetRange(PointToPlace(New Point(LeftIndent, 0)), PointToPlace(New Point(ClientSize.Width, ClientSize.Height)))
            End Get
        End Property

        <Browsable(False)>
        Public Property Selection As Range
            Get
                Return Selection
            End Get
            Set(ByVal value As Range)
                If value = selection_ Then Return
                selection_.BeginUpdate()
                selection_.Start = value.Start
                selection_.[End] = value.[End]
                selection_.EndUpdate()
                Invalidate()
            End Set
        End Property

        <DefaultValue(GetType(Color), "White")>
        <Description("Background color.")>
        Public Overrides Property BackColor As Color
            Get
                Return MyBase.BackColor
            End Get
            Set(ByVal value As Color)
                MyBase.BackColor = value
            End Set
        End Property

        <Browsable(False)>
        Public Property BackBrush As Brush
            Get
                Return BackBrush
            End Get
            Set(ByVal value As Brush)
                backBrush_ = value
                Invalidate()
            End Set
        End Property

        <Browsable(True)>
        <DefaultValue(True)>
        <Description("Scollbars visibility.")>
        Public Property ShowScrollBars As Boolean
            Get
                Return scrollBars
            End Get
            Set(ByVal value As Boolean)
                If value = scrollBars Then Return
                scrollBars = value
                needRecalc_ = True
                Invalidate()
            End Set
        End Property

        <Browsable(True)>
        <DefaultValue(True)>
        <Description("Multiline mode.")>
        Public Property Multiline As Boolean
            Get
                Return Multiline
            End Get
            Set(ByVal value As Boolean)
                If multiline = value Then Return
                multiline = value
                needRecalc_ = True

                If multiline Then
                    MyBase.AutoScroll = True
                    ShowScrollBars = True
                Else
                    MyBase.AutoScroll = False
                    ShowScrollBars = False
                    If lines_.Count > 1 Then lines_.RemoveLine(1, lines_.Count - 1)
                    lines_.Manager.ClearHistory()
                End If

                Invalidate()
            End Set
        End Property

        <Browsable(True)>
        <DefaultValue(False)>
        <Description("WordWrap.")>
        Public Property WordWrap As Boolean
            Get
                Return WordWrap
            End Get
            Set(ByVal value As Boolean)
                If wordWrap_ = value Then Return
                wordWrap_ = value
                If wordWrap_ Then selection_.ColumnSelectionMode = False
                NeedRecalc(False, True)
                Invalidate()
            End Set
        End Property

        <Browsable(True)>
        <DefaultValue(GetType(WordWrapMode), "WordWrapControlWidth")>
        <Description("WordWrap mode.")>
        Public Property WordWrapMode As WordWrapMode
            Get
                Return WordWrapMode
            End Get
            Set(ByVal value As WordWrapMode)
                If wordWrapMode = value Then Return
                wordWrapMode = value
                NeedRecalc(False, True)
                Invalidate()
            End Set
        End Property

        Private selectionHighlightingForLineBreaksEnabled As Boolean

        <DefaultValue(True)>
        <Description("If enabled then line ends included into the selection_ will be selected too. " & "Then line ends will be shown as selected blank character.")>
        Public Property SelectionHighlightingForLineBreaksEnabled As Boolean
            Get
                Return SelectionHighlightingForLineBreaksEnabled
            End Get
            Set(ByVal value As Boolean)
                selectionHighlightingForLineBreaksEnabled = value
                Invalidate()
            End Set
        End Property

        <Browsable(False)>
        Public Property findForm As FindForm
        <Browsable(False)>
        Public Property replaceForm As ReplaceForm

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Overrides Property AutoScroll As Boolean
            Get
                Return MyBase.AutoScroll
            End Get
            Set(ByVal value As Boolean)
            End Set
        End Property

        <Browsable(False)>
        Public ReadOnly Property LinesCount As Integer
            Get
                Return lines_.Count
            End Get
        End Property

        Default Public Property Item(ByVal place As Place) As Char
            Get
                Return lines_(place.iLine)(place.iChar)
            End Get
            Set(ByVal value As Char)
                lines_(place.iLine)(place.iChar) = value
            End Set
        End Property

        Default Public ReadOnly Property Item(ByVal iLine As Integer) As Line
            Get
                Return lines_(iLine)
            End Get
        End Property

        <Browsable(True)>
        <Localizable(True)>
        <Editor("System.ComponentModel.Design.MultilineStringEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", GetType(UITypeEditor))>
        <SettingsBindable(True)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
        <Description("Text of the control.")>
        <Bindable(True)>
        Public Overrides Property Text As String
            Get
                If LinesCount = 0 Then Return ""
                Dim sel = New Range(Me)
                sel.SelectAll()
                Return sel.Text
            End Get
            Set(ByVal value As String)
                If value = Text AndAlso value <> "" Then Return
                SetAsCurrentTB()
                selection_.ColumnSelectionMode = False
                selection_.BeginUpdate()

                Try
                    selection_.SelectAll()
                    InsertText(value)
                    GoHome()
                Finally
                    selection_.EndUpdate()
                End Try
            End Set
        End Property

        Public ReadOnly Property TextLength As Integer
            Get
                If LinesCount = 0 Then Return 0
                Dim sel = New Range(Me)
                sel.SelectAll()
                Return sel.Length
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property Lines As IList(Of String)
            Get
                Return Lines.GetLines()
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property Html As String
            Get
                Dim exporter = New ExportToHTML()
                exporter.UseNbsp = False
                exporter.UseStyleTag = False
                exporter.UseBr = False
                Return "<pre>" & exporter.GetHtml(Me) & "</pre>"
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property Rtf As String
            Get
                Dim exporter = New ExportToRTF()
                Return exporter.GetRtf(Me)
            End Get
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property SelectedText As String
            Get
                Return selection_.Text
            End Get
            Set(ByVal value As String)
                InsertText(value)
            End Set
        End Property

        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property SelectionStart As Integer
            Get
                Return Math.Min(PlaceToPosition(selection_.Start), PlaceToPosition(selection_.[End]))
            End Get
            Set(ByVal value As Integer)
                selection_.Start = PositionToPlace(value)
            End Set
        End Property

        <Browsable(False)>
        <DefaultValue(0)>
        Public Property SelectionLength As Integer
            Get
                Return selection_.Length
            End Get
            Set(ByVal value As Integer)
                If value > 0 Then selection_.[End] = PositionToPlace(SelectionStart + value)
            End Set
        End Property

        <DefaultValue(GetType(Font), "Courier New, 9.75")>
        Public Overrides Property Font As Font
            Get
                Return baseFont_
            End Get
            Set(ByVal value As Font)
                originalFont = CType(value.Clone(), Font)
                SetFont(value)
            End Set
        End Property

        Private baseFont_ As Font

        <DefaultValue(GetType(Font), "Courier New, 9.75")>
        Private Property BaseFont As Font
            Get
                Return baseFont_
            End Get
            Set(ByVal value As Font)
                baseFont_ = value
            End Set
        End Property

        Private Sub SetFont(ByVal newFont As Font)
            baseFont_ = newFont
            Dim sizeM As SizeF = GetCharSize(baseFont_, "M"c)
            Dim sizeDot As SizeF = GetCharSize(baseFont_, "."c)
            If sizeM <> sizeDot Then baseFont_ = New Font("Courier New", baseFont_.SizeInPoints, FontStyle.Regular, GraphicsUnit.Point)
            Dim size As SizeF = GetCharSize(baseFont_, "M"c)
            CharWidth = CInt(Math.Round(size.Width * 1.0F)) - 1
            charHeight_ = lineInterval + CInt(Math.Round(size.Height * 1.0F)) - 1
            NeedRecalc(False, wordWrap_)
            Invalidate()
        End Sub

        Public Overloads Property AutoScrollMinSize As Size
            Set(ByVal value As Size)

                If scrollBars Then
                    If Not MyBase.AutoScroll Then MyBase.AutoScroll = True
                    Dim newSize As Size = value

                    If wordWrap_ AndAlso wordWrapMode <> FastColoredTextBoxNS.WordWrapMode.Custom Then
                        Dim maxWidth As Integer = GetMaxLineWordWrapedWidth()
                        newSize = New Size(Math.Min(newSize.Width, maxWidth), newSize.Height)
                    End If

                    MyBase.AutoScrollMinSize = newSize
                Else
                    If MyBase.AutoScroll Then MyBase.AutoScroll = False
                    MyBase.AutoScrollMinSize = New Size(0, 0)
                    VerticalScroll.Visible = False
                    HorizontalScroll.Visible = False
                    VerticalScroll.Maximum = Math.Max(0, value.Height - ClientSize.Height)
                    HorizontalScroll.Maximum = Math.Max(0, value.Width - ClientSize.Width)
                    localAutoScrollMinSize = value
                End If
            End Set
            Get

                If scrollBars Then
                    Return MyBase.AutoScrollMinSize
                Else
                    Return localAutoScrollMinSize
                End If
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property ImeAllowed As Boolean
            Get
                Return ImeMode <> ImeMode.Disable AndAlso ImeMode <> ImeMode.Off AndAlso ImeMode <> ImeMode.NoControl
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property UndoEnabled As Boolean
            Get
                Return lines_.Manager.UndoEnabled
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property RedoEnabled As Boolean
            Get
                Return lines_.Manager.RedoEnabled
            End Get
        End Property

        Private ReadOnly Property LeftIndentLine As Integer
            Get
                Return LeftIndent - minLeftIndent / 2 - 3
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property Range As Range
            Get
                Return New Range(Me, New Place(0, 0), New Place(lines_(lines_.Count - 1).Count, lines_.Count - 1))
            End Get
        End Property

        <DefaultValue(GetType(Color), "Blue")>
        <Description("Color of selected area.")>
        Public Overridable Property SelectionColor As Color
            Get
                Return SelectionColor
            End Get
            Set(ByVal value As Color)
                selectionColor = value
                If selectionColor.A = 255 Then selectionColor = Color.FromArgb(60, selectionColor)
                SelectionStyle = New SelectionStyle(New SolidBrush(selectionColor))
                Invalidate()
            End Set
        End Property

        Public Overrides Property Cursor As Cursor
            Get
                Return MyBase.Cursor
            End Get
            Set(ByVal value As Cursor)
                defaultCursor = value
                MyBase.Cursor = value
            End Set
        End Property

        <DefaultValue(1)>
        <Description("Reserved space for line number characters. If smaller than needed (e. g. line count >= 10 and " & "this value set to 1) this value will have no impact. If you want to reserve space, e. g. for line " & "numbers >= 10 or >= 100, than you can set this value to 2 or 3 or higher.")>
        Public Property ReservedCountOfLineNumberChars As Integer
            Get
                Return ReservedCountOfLineNumberChars
            End Get
            Set(ByVal value As Integer)
                reservedCountOfLineNumberChars_ = value
                NeedRecalc()
                Invalidate()
            End Set
        End Property

        <Browsable(True)>
        <Description("Occurs when mouse is moving over text and tooltip is needed.")>
        Public Event ToolTipNeeded As EventHandler(Of ToolTipNeededEventArgs)

        Public Sub ClearHints()
            If hints IsNot Nothing Then hints.Clear()
        End Sub

        Public Overridable Function AddHint(ByVal range As Range, ByVal innerControl As Control, ByVal scrollToHint As Boolean, ByVal inline As Boolean, ByVal dock As Boolean) As Hint
            Dim hint = New Hint(range, innerControl, inline, dock)
            hints.Add(hint)
            If scrollToHint Then hint.DoVisible()
            Return hint
        End Function

        Public Function AddHint(ByVal range As Range, ByVal innerControl As Control) As Hint
            Return AddHint(range, innerControl, True, True, True)
        End Function

        Public Overridable Function AddHint(ByVal range As Range, ByVal text As String, ByVal scrollToHint As Boolean, ByVal inline As Boolean, ByVal dock As Boolean) As Hint
            Dim hint = New Hint(range, text, inline, dock)
            hints.Add(hint)
            If scrollToHint Then hint.DoVisible()
            Return hint
        End Function

        Public Function AddHint(ByVal range As Range, ByVal text As String) As Hint
            Return AddHint(range, text, True, True, True)
        End Function

        Public Overridable Sub OnHintClick(ByVal hint As Hint)
            RaiseEvent HintClick(Me, New HintClickEventArgs(hint))
        End Sub

        Private Sub timer3_Tick(ByVal sender As Object, ByVal e As EventArgs)
            timer3.[Stop]()
            OnToolTip()
        End Sub

        Protected Overridable Sub OnToolTip()
            If ToolTip Is Nothing Then Return
            If ToolTipNeeded Is Nothing Then Return
            Dim place As Place = PointToPlace(lastMouseCoord)
            Dim p As Point = PlaceToPoint(place)
            If Math.Abs(p.X - lastMouseCoord.X) > CharWidth * 2 OrElse Math.Abs(p.Y - lastMouseCoord.Y) > charHeight_ * 2 Then Return
            Dim r = New Range(Me, place, place)
            Dim hoveredWord As String = r.GetFragment("[a-zA-Z]").Text
            Dim ea = New ToolTipNeededEventArgs(place, hoveredWord)
            ToolTipNeeded(Me, ea)

            If ea.ToolTipText IsNot Nothing Then
                ToolTip.ToolTipTitle = ea.ToolTipTitle
                ToolTip.ToolTipIcon = ea.ToolTipIcon
                ToolTip.Show(ea.ToolTipText, Me, New Point(lastMouseCoord.X, lastMouseCoord.Y + charHeight_))
            End If
        End Sub

        Public Overridable Sub OnVisibleRangeChanged()
            needRecalcFoldingLines = True
            needRiseVisibleRangeChangedDelayed = True
            ResetTimer(timer)
            RaiseEvent VisibleRangeChanged(Me, New EventArgs())
        End Sub

        Public Overloads Sub Invalidate()
            If InvokeRequired Then
                BeginInvoke(New MethodInvoker(AddressOf Invalidate))
            Else
                MyBase.Invalidate()
            End If
        End Sub

        Protected Overridable Sub OnCharSizeChanged()
            VerticalScroll.SmallChange = charHeight_
            VerticalScroll.LargeChange = 10 * charHeight_
            HorizontalScroll.SmallChange = CharWidth
        End Sub

        <Browsable(True)>
        <Description("It occurs if user click on the hint.")>
        Public Event HintClick As EventHandler(Of HintClickEventArgs)
        <Browsable(True)>
        <Description("It occurs after insert, delete, clear, undo and redo operations.")>
        Public Overloads Event TextChanged As EventHandler(Of TextChangedEventArgs)
        <Browsable(False)>
        Friend Event BindingTextChanged As EventHandler
        <Description("Occurs when user paste text from clipboard")>
        Public Event Pasting As EventHandler(Of TextChangingEventArgs)
        <Browsable(True)>
        <Description("It occurs before insert, delete, clear, undo and redo operations.")>
        Public Event TextChanging As EventHandler(Of TextChangingEventArgs)
        <Browsable(True)>
        <Description("It occurs after changing of selection_.")>
        Public Event SelectionChanged As EventHandler
        <Browsable(True)>
        <Description("It occurs after changing of visible range.")>
        Public Event VisibleRangeChanged As EventHandler
        <Browsable(True)>
        <Description("It occurs after insert, delete, clear, undo and redo operations. This event occurs with a delay relative to TextChanged, and fires only once.")>
        Public Event TextChangedDelayed As EventHandler(Of TextChangedEventArgs)
        <Browsable(True)>
        <Description("It occurs after changing of selection_. This event occurs with a delay relative to SelectionChanged, and fires only once.")>
        Public Event SelectionChangedDelayed As EventHandler
        <Browsable(True)>
        <Description("It occurs after changing of visible range. This event occurs with a delay relative to VisibleRangeChanged, and fires only once.")>
        Public Event VisibleRangeChangedDelayed As EventHandler
        <Browsable(True)>
        <Description("It occurs when user click on VisualMarker.")>
        Public Event VisualMarkerClick As EventHandler(Of VisualMarkerEventArgs)
        <Browsable(True)>
        <Description("It occurs when visible char is enetering (alphabetic, digit, punctuation, DEL, BACKSPACE).")>
        Public Event KeyPressing As KeyPressEventHandler
        <Browsable(True)>
        <Description("It occurs when visible char is enetered (alphabetic, digit, punctuation, DEL, BACKSPACE).")>
        Public Event KeyPressed As KeyPressEventHandler
        <Browsable(True)>
        <Description("It occurs when calculates AutoIndent for new line.")>
        Public Event AutoIndentNeeded As EventHandler(Of AutoIndentEventArgs)
        <Browsable(True)>
        <Description("It occurs when line background is painting.")>
        Public Event PaintLine As EventHandler(Of PaintLineEventArgs)
        <Browsable(True)>
        <Description("Occurs when line was inserted/added.")>
        Public Event LineInserted As EventHandler(Of LineInsertedEventArgs)
        <Browsable(True)>
        <Description("Occurs when line was removed.")>
        Public Event LineRemoved As EventHandler(Of LineRemovedEventArgs)
        <Browsable(True)>
        <Description("Occurs when current highlighted folding area is changed.")>
        Public Event FoldingHighlightChanged As EventHandler(Of EventArgs)
        <Browsable(True)>
        <Description("Occurs when undo/redo stack is changed.")>
        Public Event UndoRedoStateChanged As EventHandler(Of EventArgs)
        <Browsable(True)>
        <Description("Occurs when component was zoomed.")>
        Public Event ZoomChanged As EventHandler
        <Browsable(True)>
        <Description("Occurs when user pressed key, that specified as CustomAction.")>
        Public Event CustomAction As EventHandler(Of CustomActionEventArgs)
        <Browsable(True)>
        <Description("Occurs when scroolbars are updated.")>
        Public Event ScrollbarsUpdated As EventHandler
        <Browsable(True)>
        <Description("Occurs when custom wordwrap is needed.")>
        Public Event WordWrapNeeded As EventHandler(Of WordWrapNeededEventArgs)

        Public Function GetStylesOfChar(ByVal place As Place) As List(Of Style)
            Dim result = New List(Of Style)()

            If place.iLine < LinesCount AndAlso place.iChar < Me(place.iLine).Count Then
                Dim s = CUShort(Me(place).style)

                For i As Integer = 0 To 16 - 1
                    ''' Cannot convert IfStatementSyntax, System.ArgumentOutOfRangeException: Exception of type 'System.ArgumentOutOfRangeException' was thrown.
                    ''' Parameter name: op
                    ''' Actual value was LeftShiftExpression.
                    '''    at ICSharpCode.CodeConverter.Util.VBUtil.GetExpressionOperatorTokenKind(SyntaxKind op)
                    '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node)
                    '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                    '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
                    '''    at ICSharpCode.CodeConverter.VB.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
                    '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.VisitBinaryExpression(BinaryExpressionSyntax node)
                    '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                    '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node)
                    '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                    '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
                    '''    at ICSharpCode.CodeConverter.VB.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
                    '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.VisitBinaryExpression(BinaryExpressionSyntax node)
                    '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                    '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitParenthesizedExpression(ParenthesizedExpressionSyntax node)
                    '''    at Microsoft.CodeAnalysis.CSharp.Syntax.ParenthesizedExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                    '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
                    '''    at ICSharpCode.CodeConverter.VB.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
                    '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.VisitParenthesizedExpression(ParenthesizedExpressionSyntax node)
                    '''    at Microsoft.CodeAnalysis.CSharp.Syntax.ParenthesizedExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                    '''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node)
                    '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                    '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
                    '''    at ICSharpCode.CodeConverter.VB.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
                    '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.VisitBinaryExpression(BinaryExpressionSyntax node)
                    '''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                    '''    at ICSharpCode.CodeConverter.VB.MethodBodyVisitor.VisitIfStatement(IfStatementSyntax node)
                    '''    at Microsoft.CodeAnalysis.CSharp.Syntax.IfStatementSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
                    '''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
                    '''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
                    '''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)
                    ''' 
                    ''' Input: 
                    '''                     if ((s & ((ushort) 1) << i) != 0)
                    result.Add(Styles[i]);

''' 
            Next
            End If

            Return result
        End Function

        Protected Overridable Function CreateTextSource() As TextSource
            Return New TextSource(Me)
        End Function

        Private Sub SetAsCurrentTB()
            TextSource.CurrentTB = Me
        End Sub

        Protected Overridable Sub InitTextSource(ByVal ts As TextSource)
            If lines_ IsNot Nothing Then
                lines_.LineInserted -= AddressOf ts_LineInserted
                lines_.LineRemoved -= AddressOf ts_LineRemoved
                lines_.TextChanged -= AddressOf ts_TextChanged
                lines_.RecalcNeeded -= AddressOf ts_RecalcNeeded
                lines_.RecalcWordWrap -= AddressOf ts_RecalcWordWrap
                lines_.TextChanging -= AddressOf ts_TextChanging
                lines_.Dispose()
            End If

            LineInfos.Clear()
            ClearHints()
            If bookmarks IsNot Nothing Then bookmarks.Clear()
            lines_ = ts

            If ts IsNot Nothing Then
                ts.LineInserted += AddressOf ts_LineInserted
                ts.LineRemoved += AddressOf ts_LineRemoved
                ts.TextChanged += AddressOf ts_TextChanged
                ts.RecalcNeeded += AddressOf ts_RecalcNeeded
                ts.RecalcWordWrap += AddressOf ts_RecalcWordWrap
                ts.TextChanging += AddressOf ts_TextChanging

                While LineInfos.Count < ts.Count
                    LineInfos.Add(New LineInfo(-1))
                End While
            End If

            isChanged_ = False
            needRecalc_ = True
        End Sub

        Private Sub ts_RecalcWordWrap(ByVal sender As Object, ByVal e As TextSource.TextChangedEventArgs)
            RecalcWordWrap(e.iFromLine, e.iToLine)
        End Sub

        Private Sub ts_TextChanging(ByVal sender As Object, ByVal e As TextChangingEventArgs)
            If TextSource.CurrentTB = Me Then
                Dim text As String = e.InsertingText
                OnTextChanging(text)
                e.InsertingText = text
            End If
        End Sub

        Private Sub ts_RecalcNeeded(ByVal sender As Object, ByVal e As TextSource.TextChangedEventArgs)
            If e.iFromLine = e.iToLine AndAlso Not wordWrap_ AndAlso lines_.Count > minLinesForAccuracy Then
                RecalcScrollByOneLine(e.iFromLine)
            Else
                NeedRecalc(False, wordWrap_)
            End If
        End Sub

        Public Sub NeedRecalc()
            NeedRecalc(False)
        End Sub

        Public Sub NeedRecalc(ByVal forced As Boolean)
            NeedRecalc(forced, False)
        End Sub

        Public Sub NeedRecalc(ByVal forced As Boolean, ByVal wordWrapRecalc As Boolean)
            needRecalc_ = True

            If wordWrapRecalc Then
                needRecalcWordWrapInterval = New Point(0, LinesCount - 1)
                needRecalcWordWrap = True
            End If

            If forced Then Recalc()
        End Sub

        Private Sub ts_TextChanged(ByVal sender As Object, ByVal e As TextSource.TextChangedEventArgs)
            If e.iFromLine = e.iToLine AndAlso Not wordWrap_ Then
                RecalcScrollByOneLine(e.iFromLine)
            Else
                needRecalc_ = True
            End If

            Invalidate()
            If TextSource.CurrentTB = Me Then OnTextChanged(e.iFromLine, e.iToLine)
        End Sub

        Private Sub ts_LineRemoved(ByVal sender As Object, ByVal e As LineRemovedEventArgs)
            LineInfos.RemoveRange(e.Index, e.Count)
            OnLineRemoved(e.Index, e.Count, e.RemovedLineUniqueIds)
        End Sub

        Private Sub ts_LineInserted(ByVal sender As Object, ByVal e As LineInsertedEventArgs)
            Dim newState As VisibleState = VisibleState.Visible
            If e.Index >= 0 AndAlso e.Index < LineInfos.Count AndAlso LineInfos(e.Index).VisibleState = VisibleState.Hidden Then newState = VisibleState.Hidden
            If e.Count > 100000 Then LineInfos.Capacity = LineInfos.Count + e.Count + 1000
            Dim temp = New LineInfo(e.Count - 1) {}

            For i As Integer = 0 To e.Count - 1
                temp(i).startY = -1
                temp(i).VisibleState = newState
            Next

            LineInfos.InsertRange(e.Index, temp)
            If e.Count > 1000000 Then GC.Collect()
            OnLineInserted(e.Index, e.Count)
        End Sub

        Public Function NavigateForward() As Boolean
            Dim min As DateTime = DateTime.Now
            Dim iLine As Integer = -1

            For i As Integer = 0 To LinesCount - 1

                If lines_.IsLineLoaded(i) Then

                    If lines_(i).LastVisit > lastNavigatedDateTime AndAlso lines_(i).LastVisit < min Then
                        min = lines_(i).LastVisit
                        iLine = i
                    End If
                End If
            Next

            If iLine >= 0 Then
                Navigate(iLine)
                Return True
            Else
                Return False
            End If
        End Function

        Public Function NavigateBackward() As Boolean
            Dim max = New DateTime()
            Dim iLine As Integer = -1

            For i As Integer = 0 To LinesCount - 1

                If lines_.IsLineLoaded(i) Then

                    If lines_(i).LastVisit < lastNavigatedDateTime AndAlso lines_(i).LastVisit > max Then
                        max = lines_(i).LastVisit
                        iLine = i
                    End If
                End If
            Next

            If iLine >= 0 Then
                Navigate(iLine)
                Return True
            Else
                Return False
            End If
        End Function

        Public Sub Navigate(ByVal iLine As Integer)
            If iLine >= LinesCount Then Return
            lastNavigatedDateTime = lines_(iLine).LastVisit
            selection_.Start = New Place(0, iLine)
            DoSelectionVisible()
        End Sub

        Protected Overrides Sub OnLoad(ByVal e As EventArgs)
            MyBase.OnLoad(e)
            m_hImc = ImmGetContext(Handle)
        End Sub

        Private Sub timer2_Tick(ByVal sender As Object, ByVal e As EventArgs)
            timer2.Enabled = False

            If needRiseTextChangedDelayed Then
                needRiseTextChangedDelayed = False
                If delayedTextChangedRange Is Nothing Then Return
                delayedTextChangedRange = Range.GetIntersectionWith(delayedTextChangedRange)
                delayedTextChangedRange.Expand()
                OnTextChangedDelayed(delayedTextChangedRange)
                delayedTextChangedRange = Nothing
            End If
        End Sub

        Public Sub AddVisualMarker(ByVal marker As VisualMarker)
            visibleMarkers.Add(marker)
        End Sub

        Private Sub timer_Tick(ByVal sender As Object, ByVal e As EventArgs)
            timer.Enabled = False

            If needRiseSelectionChangedDelayed Then
                needRiseSelectionChangedDelayed = False
                OnSelectionChangedDelayed()
            End If

            If needRiseVisibleRangeChangedDelayed Then
                needRiseVisibleRangeChangedDelayed = False
                OnVisibleRangeChangedDelayed()
            End If
        End Sub

        Public Overridable Sub OnTextChangedDelayed(ByVal changedRange As Range)
            RaiseEvent TextChangedDelayed(Me, New TextChangedEventArgs(changedRange))
        End Sub

        Public Overridable Sub OnSelectionChangedDelayed()
            RecalcScrollByOneLine(selection_.Start.iLine)
            ClearBracketsPositions()
            If LeftBracket <> vbNullChar AndAlso RightBracket <> vbNullChar Then HighlightBrackets(LeftBracket, RightBracket, leftBracketPosition, rightBracketPosition)
            If LeftBracket2 <> vbNullChar AndAlso RightBracket2 <> vbNullChar Then HighlightBrackets(LeftBracket2, RightBracket2, leftBracketPosition2, rightBracketPosition2)

            If selection_.IsEmpty AndAlso selection_.Start.iLine < LinesCount Then

                If lastNavigatedDateTime <> lines_(selection_.Start.iLine).LastVisit Then
                    lines_(selection_.Start.iLine).LastVisit = DateTime.Now
                    lastNavigatedDateTime = lines_(selection_.Start.iLine).LastVisit
                End If
            End If

            RaiseEvent SelectionChangedDelayed(Me, New EventArgs())
        End Sub

        Public Overridable Sub OnVisibleRangeChangedDelayed()
            RaiseEvent VisibleRangeChangedDelayed(Me, New EventArgs())
        End Sub

        Private timersToReset As Dictionary(Of Timer, Timer) = New Dictionary(Of Timer, Timer)()

        Private Sub ResetTimer(ByVal timer As Timer)
            If InvokeRequired Then
                BeginInvoke(New MethodInvoker(Function() ResetTimer(timer)))
                Return
            End If

            timer.[Stop]()

            If IsHandleCreated Then
                timer.Start()
            Else
                timersToReset(timer) = timer
            End If
        End Sub

        Protected Overrides Sub OnHandleCreated(ByVal e As EventArgs)
            MyBase.OnHandleCreated(e)

            For Each timer In New List(Of Timer)(timersToReset.Keys)
                ResetTimer(timer)
            Next

            timersToReset.Clear()
            OnScrollbarsUpdated()
        End Sub

        Public Function AddStyle(ByVal style As Style) As Integer
            If style Is Nothing Then Return -1
            Dim i As Integer = GetStyleIndex(style)
            If i >= 0 Then Return i
            i = CheckStylesBufferSize()
            Styles(i) = style
            Return i
        End Function

        Public Function CheckStylesBufferSize() As Integer
            Dim i As Integer

            For i = Styles.Length - 1 To 0
                If Styles(i) IsNot Nothing Then Exit For
            Next

            i += 1
            If i >= Styles.Length Then Throw New Exception("Maximum count of Styles is exceeded.")
            Return i
        End Function

        Public Overridable Sub ShowFindDialog()
            ShowFindDialog(Nothing)
        End Sub

        Public Overridable Sub ShowFindDialog(ByVal findText As String)
            If findForm Is Nothing Then findForm = New FindForm(Me)

            If findText IsNot Nothing Then
                findForm.tbFind.Text = findText
            ElseIf Not selection_.IsEmpty AndAlso selection_.Start.iLine = selection_.[End].iLine Then
                findForm.tbFind.Text = selection_.Text
            End If

            findForm.tbFind.SelectAll()
            findForm.Show()
            findForm.Focus()
        End Sub

        Public Overridable Sub ShowReplaceDialog()
            ShowReplaceDialog(Nothing)
        End Sub

        Public Overridable Sub ShowReplaceDialog(ByVal findText As String)
            If [ReadOnly] Then Return
            If replaceForm Is Nothing Then replaceForm = New ReplaceForm(Me)

            If findText IsNot Nothing Then
                replaceForm.tbFind.Text = findText
            ElseIf Not selection_.IsEmpty AndAlso selection_.Start.iLine = selection_.[End].iLine Then
                replaceForm.tbFind.Text = selection_.Text
            End If

            replaceForm.tbFind.SelectAll()
            replaceForm.Show()
            replaceForm.Focus()
        End Sub

        Public Function GetLineLength(ByVal iLine As Integer) As Integer
            If iLine < 0 OrElse iLine >= lines_.Count Then Throw New ArgumentOutOfRangeException("Line index out of range")
            Return lines_(iLine).Count
        End Function

        Public Function GetLine(ByVal iLine As Integer) As Range
            If iLine < 0 OrElse iLine >= lines_.Count Then Throw New ArgumentOutOfRangeException("Line index out of range")
            Dim sel = New Range(Me)
            sel.Start = New Place(0, iLine)
            sel.[End] = New Place(lines_(iLine).Count, iLine)
            Return sel
        End Function

        Public Overridable Sub Copy()
            If selection_.IsEmpty Then selection_.Expand()

            If Not selection_.IsEmpty Then
                Dim data = New DataObject()
                OnCreateClipboardData(data)
                Dim thread = New Thread(Function() SetClipboard(data))
                thread.SetApartmentState(ApartmentState.STA)
                thread.Start()
                thread.Join()
            End If
        End Sub

        Protected Overridable Sub OnCreateClipboardData(ByVal data As DataObject)
            Dim exp = New ExportToHTML()
            exp.UseBr = False
            exp.UseNbsp = False
            exp.UseStyleTag = True
            Dim html As String = "<pre>" & exp.GetHtml(selection_.Clone()) & "</pre>"
            data.SetData(DataFormats.UnicodeText, True, selection_.Text)
            data.SetData(DataFormats.Html, PrepareHtmlForClipboard(html))
            data.SetData(DataFormats.Rtf, New ExportToRTF().GetRtf(selection_.Clone()))
        End Sub

        <DllImport("user32.dll")>
        Private Shared Function GetOpenClipboardWindow() As IntPtr
    <DllImport("user32.dll")>
        Private Shared Function CloseClipboard() As IntPtr

    Protected Sub SetClipboard(ByVal data As DataObject)
            Try
                CloseClipboard()
                Clipboard.SetDataObject(data, True, 5, 100)
            Catch __unusedExternalException1__ As ExternalException
            End Try
        End Sub

        Public Shared Function PrepareHtmlForClipboard(ByVal html As String) As MemoryStream
            Dim enc As Encoding = Encoding.UTF8
            Dim begin As String = "Version:0.9" & vbCrLf & "StartHTML:{0:000000}" & vbCrLf & "EndHTML:{1:000000}" & vbCrLf & "StartFragment:{2:000000}" & vbCrLf & "EndFragment:{3:000000}" & vbCrLf
            Dim html_begin As String = "<html>" & vbCrLf & "<head>" & vbCrLf & "<meta http-equiv=""Content-Type""" & " content=""text/html; charset=" & enc.WebName & """>" & vbCrLf & "<title>HTML clipboard</title>" & vbCrLf & "</head>" & vbCrLf & "<body>" & vbCrLf & "<!--StartFragment-->"
            Dim html_end As String = "<!--EndFragment-->" & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
            Dim begin_sample As String = String.Format(begin, 0, 0, 0, 0)
            Dim count_begin As Integer = enc.GetByteCount(begin_sample)
            Dim count_html_begin As Integer = enc.GetByteCount(html_begin)
            Dim count_html As Integer = enc.GetByteCount(html)
            Dim count_html_end As Integer = enc.GetByteCount(html_end)
            Dim html_total As String = String.Format(begin, count_begin, count_begin + count_html_begin + count_html + count_html_end, count_begin + count_html_begin, count_begin + count_html_begin + count_html) & html_begin & html & html_end
            Return New MemoryStream(enc.GetBytes(html_total))
        End Function

        Public Overridable Sub Cut()
            If Not selection_.IsEmpty Then
                Copy()
                ClearSelected()
            ElseIf LinesCount = 1 Then
                selection_.SelectAll()
                Copy()
                ClearSelected()
            Else
                Dim data = New DataObject()
                OnCreateClipboardData(data)
                Dim thread = New Thread(Function() SetClipboard(data))
                thread.SetApartmentState(ApartmentState.STA)
                thread.Start()
                thread.Join()

                If selection_.Start.iLine >= 0 AndAlso selection_.Start.iLine < LinesCount Then
                    Dim iLine As Integer = selection_.Start.iLine
                    RemoveLines(New List(Of Integer) From {
                    iLine
                })
                    selection_.Start = New Place(0, Math.Max(0, Math.Min(iLine, LinesCount - 1)))
                End If
            End If
        End Sub

        Public Overridable Sub Paste()
            Dim text As String = Nothing
            Dim thread = New Thread(Function()
                                        If Clipboard.ContainsText() Then text = Clipboard.GetText()
                                    End Function)
            thread.SetApartmentState(ApartmentState.STA)
            thread.Start()
            thread.Join()

            If Pasting IsNot Nothing Then
                Dim args = New TextChangingEventArgs With {
                .Cancel = False,
                .InsertingText = text
            }
                Pasting(Me, args)

                If args.Cancel Then
                    text = String.Empty
                Else
                    text = args.InsertingText
                End If
            End If

            If Not String.IsNullOrEmpty(text) Then InsertText(text)
        End Sub

        Public Sub SelectAll()
            selection_.SelectAll()
        End Sub

        Public Sub GoEnd()
            If lines_.Count > 0 Then
                selection_.Start = New Place(lines_(lines_.Count - 1).Count, lines_.Count - 1)
            Else
                selection_.Start = New Place(0, 0)
            End If

            DoCaretVisible()
        End Sub

        Public Sub GoHome()
            selection_.Start = New Place(0, 0)
            DoCaretVisible()
        End Sub

        Public Overridable Sub Clear()
            selection_.BeginUpdate()

            Try
                selection_.SelectAll()
                ClearSelected()
                lines_.Manager.ClearHistory()
                Invalidate()
            Finally
                selection_.EndUpdate()
            End Try
        End Sub

        Public Sub ClearStylesBuffer()
            For i As Integer = 0 To Styles.Length - 1
                Styles(i) = Nothing
            Next
        End Sub

        Public Sub ClearStyle(ByVal styleIndex As StyleIndex)
            For Each line As Line In lines_
                line.ClearStyle(styleIndex)
            Next

            For i As Integer = 0 To LineInfos.Count - 1
                SetVisibleState(i, VisibleState.Visible)
            Next

            Invalidate()
        End Sub

        Public Sub ClearUndo()
            lines_.Manager.ClearHistory()
        End Sub

        Public Overridable Sub InsertText(ByVal text As String)
            InsertText(text, True)
        End Sub

        Public Overridable Sub InsertText(ByVal text As String, ByVal jumpToCaret As Boolean)
            If text Is Nothing Then Return
            If text = vbCr Then text = vbLf
            lines_.Manager.BeginAutoUndoCommands()

            Try
                If Not selection_.IsEmpty Then lines_.Manager.ExecuteCommand(New ClearSelectedCommand(TextSource))

                If Me.TextSource.Count > 0 Then
                    If selection_.IsEmpty AndAlso selection_.Start.iChar > GetLineLength(selection_.Start.iLine) AndAlso VirtualSpace Then InsertVirtualSpaces()
                End If

                lines_.Manager.ExecuteCommand(New InsertTextCommand(TextSource, text))
                If updating <= 0 AndAlso jumpToCaret Then DoCaretVisible()
            Finally
                lines_.Manager.EndAutoUndoCommands()
            End Try

            Invalidate()
        End Sub

        Public Overridable Function InsertText(ByVal text As String, ByVal style As Style) As Range
            Return InsertText(text, style, True)
        End Function

        Public Overridable Function InsertText(ByVal text As String, ByVal style As Style, ByVal jumpToCaret As Boolean) As Range
            If text Is Nothing Then Return Nothing
            Dim last As Place = If(selection_.Start > selection_.[End], selection_.[End], selection_.Start)
            InsertText(text, jumpToCaret)
            Dim range = New Range(Me, last, selection_.Start) With {
            .ColumnSelectionMode = selection_.ColumnSelectionMode
        }
            range = range.GetIntersectionWith(range)
            range.SetStyle(style)
            Return range
        End Function

        Public Overridable Function InsertTextAndRestoreSelection(ByVal replaceRange As Range, ByVal text As String, ByVal style As Style) As Range
            If text Is Nothing Then Return Nothing
            Dim oldStart = PlaceToPosition(selection_.Start)
            Dim oldEnd = PlaceToPosition(selection_.[End])
            Dim count = replaceRange.Text.Length
            Dim pos = PlaceToPosition(replaceRange.Start)
            selection_.BeginUpdate()
            selection_ = replaceRange
            Dim range = InsertText(text, style)
            count = range.Text.Length - count
            selection_.Start = PositionToPlace(oldStart + (If(oldStart >= pos, count, 0)))
            selection_.[End] = PositionToPlace(oldEnd + (If(oldEnd >= pos, count, 0)))
            selection_.EndUpdate()
            Return range
        End Function

        Public Overridable Sub AppendText(ByVal text As String)
            AppendText(text, Nothing)
        End Sub

        Public Overridable Sub AppendText(ByVal text As String, ByVal style As Style)
            If text Is Nothing Then Return
            selection_.ColumnSelectionMode = False
            Dim oldStart As Place = selection_.Start
            Dim oldEnd As Place = selection_.[End]
            selection_.BeginUpdate()
            lines_.Manager.BeginAutoUndoCommands()

            Try

                If lines_.Count > 0 Then
                    selection_.Start = New Place(lines_(lines_.Count - 1).Count, lines_.Count - 1)
                Else
                    selection_.Start = New Place(0, 0)
                End If

                Dim last As Place = selection_.Start
                lines_.Manager.ExecuteCommand(New InsertTextCommand(TextSource, text))
                If style IsNot Nothing Then New Range(Me, last, Selection.Start).SetStyle(style)
        Finally
                lines_.Manager.EndAutoUndoCommands()
                selection_.Start = oldStart
                selection_.[End] = oldEnd
                selection_.EndUpdate()
            End Try

            Invalidate()
        End Sub

        Public Function GetStyleIndex(ByVal style As Style) As Integer
            Return Array.IndexOf(Styles, style)
        End Function

        Public Function GetStyleIndexMask(ByVal styles As Style()) As StyleIndex
            Dim mask As StyleIndex = StyleIndex.None

            For Each style As Style In styles
                Dim i As Integer = GetStyleIndex(style)
                If i >= 0 Then mask = mask Or Range.ToStyleIndex(i)
            Next

            Return mask
        End Function

        Friend Function GetOrSetStyleLayerIndex(ByVal style As Style) As Integer
            Dim i As Integer = GetStyleIndex(style)
            If i < 0 Then i = AddStyle(style)
            Return i
        End Function

        Public Shared Function GetCharSize(ByVal font As Font, ByVal c As Char) As SizeF
            Dim sz2 As Size = TextRenderer.MeasureText("<" & c.ToString() & ">", font)
            Dim sz3 As Size = TextRenderer.MeasureText("<>", font)
            Return New SizeF(sz2.Width - sz3.Width + 1, font.Height)
        End Function

        <DllImport("Imm32.dll")>
        Public Shared Function ImmGetContext(ByVal hWnd As IntPtr) As IntPtr
    <DllImport("Imm32.dll")>
        Public Shared Function ImmAssociateContext(ByVal hWnd As IntPtr, ByVal hIMC As IntPtr) As IntPtr

    Protected Overrides Sub WndProc(ByRef m As Message)
            If m.Msg = WM_HSCROLL OrElse m.Msg = WM_VSCROLL Then
                If m.WParam.ToInt32() <> SB_ENDSCROLL Then Invalidate()
            End If

            MyBase.WndProc(m)

            If ImeAllowed Then

                If m.Msg = WM_IME_SETCONTEXT AndAlso m.WParam.ToInt32() = 1 Then
                    ImmAssociateContext(Handle, m_hImc)
                End If
            End If
        End Sub

        Private tempHintsList As List(Of Control) = New List(Of Control)()

        Private Sub HideHints()
            If Not ShowScrollBars AndAlso hints.Count > 0 Then
            (TryCast(Me, Control)).SuspendLayout()

            For Each c As Control In Controls
                    tempHintsList.Add(c)
                Next

                Controls.Clear()
            End If
        End Sub

        Private Sub RestoreHints()
            If Not ShowScrollBars AndAlso hints.Count > 0 Then

                For Each c In tempHintsList
                    Controls.Add(c)
                Next

                tempHintsList.Clear()
            (TryCast(Me, Control)).ResumeLayout(False)
            If Not Focused Then Focus()
            End If
        End Sub

        Public Sub OnScroll(ByVal se As ScrollEventArgs, ByVal alignByLines As Boolean)
            HideHints()

            If se.ScrollOrientation = ScrollOrientation.VerticalScroll Then
                Dim newValue As Integer = se.NewValue
                If alignByLines Then newValue = CInt((Math.Ceiling(1.0R * newValue / charHeight_) * charHeight_))
                VerticalScroll.Value = Math.Max(VerticalScroll.Minimum, Math.Min(VerticalScroll.Maximum, newValue))
            End If

            If se.ScrollOrientation = ScrollOrientation.HorizontalScroll Then HorizontalScroll.Value = Math.Max(HorizontalScroll.Minimum, Math.Min(HorizontalScroll.Maximum, se.NewValue))
            UpdateScrollbars()
            RestoreHints()
            Invalidate()
            MyBase.OnScroll(se)
            OnVisibleRangeChanged()
        End Sub

        Protected Overrides Sub OnScroll(ByVal se As ScrollEventArgs)
            OnScroll(se, True)
        End Sub

        Protected Overridable Sub InsertChar(ByVal c As Char)
            lines_.Manager.BeginAutoUndoCommands()

            Try
                If Not selection_.IsEmpty Then lines_.Manager.ExecuteCommand(New ClearSelectedCommand(TextSource))
                If selection_.IsEmpty AndAlso selection_.Start.iChar > GetLineLength(selection_.Start.iLine) AndAlso VirtualSpace Then InsertVirtualSpaces()
                lines_.Manager.ExecuteCommand(New InsertCharCommand(TextSource, c))
            Finally
                lines_.Manager.EndAutoUndoCommands()
            End Try

            Invalidate()
        End Sub

        Private Sub InsertVirtualSpaces()
            Dim lineLength As Integer = GetLineLength(selection_.Start.iLine)
            Dim count As Integer = selection_.Start.iChar - lineLength
            selection_.BeginUpdate()

            Try
                selection_.Start = New Place(lineLength, selection_.Start.iLine)
                lines_.Manager.ExecuteCommand(New InsertTextCommand(TextSource, New String(" "c, count)))
            Finally
                selection_.EndUpdate()
            End Try
        End Sub

        Public Overridable Sub ClearSelected()
            If Not selection_.IsEmpty Then
                lines_.Manager.ExecuteCommand(New ClearSelectedCommand(TextSource))
                Invalidate()
            End If
        End Sub

        Public Sub ClearCurrentLine()
            selection_.Expand()
            lines_.Manager.ExecuteCommand(New ClearSelectedCommand(TextSource))

            If selection_.Start.iLine = 0 Then
                If Not selection_.GoRightThroughFolded() Then Return
            End If

            If selection_.Start.iLine > 0 Then lines_.Manager.ExecuteCommand(New InsertCharCommand(TextSource, vbBack))
            Invalidate()
        End Sub

        Private Sub Recalc()
            If Not needRecalc_ Then Return
            needRecalc_ = False
            LeftIndent = leftPadding
            Dim maxLineNumber As Long = LinesCount + lineNumberStartValue - 1
            Dim charsForLineNumber As Integer = 2 + (If(maxLineNumber > 0, CInt(Math.Log10(maxLineNumber)), 0))
            If Me.reservedCountOfLineNumberChars_ + 1 > charsForLineNumber Then charsForLineNumber = Me.reservedCountOfLineNumberChars_ + 1

            If Created Then
                If showLineNumbers_ Then LeftIndent += charsForLineNumber * CharWidth + minLeftIndent + 1

                If needRecalcWordWrap Then
                    RecalcWordWrap(needRecalcWordWrapInterval.X, needRecalcWordWrapInterval.Y)
                    needRecalcWordWrap = False
                End If
            Else
                needRecalc_ = True
            End If

            TextHeight = 0
            maxLineLength = RecalcMaxLineLength()
            Dim minWidth As Integer
            CalcMinAutosizeWidth(minWidth, maxLineLength)
            AutoScrollMinSize = New Size(minWidth, TextHeight + Paddings.Top + Paddings.Bottom)
            UpdateScrollbars()
        End Sub

        Private Sub CalcMinAutosizeWidth(<Out> ByRef minWidth As Integer, ByRef maxLineLength As Integer)
            minWidth = LeftIndent + (maxLineLength) * CharWidth + 2 + Paddings.Left + Paddings.Right

            If wordWrap_ Then

                Select Case wordWrapMode
                    Case wordWrapMode.WordWrapControlWidth, wordWrapMode.CharWrapControlWidth
                        maxLineLength = Math.Min(maxLineLength, (ClientSize.Width - LeftIndent - Paddings.Left - Paddings.Right) / CharWidth)
                        minWidth = 0
                    Case wordWrapMode.WordWrapPreferredWidth, wordWrapMode.CharWrapPreferredWidth
                        maxLineLength = Math.Min(maxLineLength, preferredLineWidth)
                        minWidth = LeftIndent + preferredLineWidth * CharWidth + 2 + Paddings.Left + Paddings.Right
                End Select
            End If
        End Sub

        Private Sub RecalcScrollByOneLine(ByVal iLine As Integer)
            If iLine >= lines_.Count Then Return
            Dim maxLineLength As Integer = lines_(iLine).Count
            If Me.maxLineLength < maxLineLength AndAlso Not wordWrap_ Then Me.maxLineLength = maxLineLength
            Dim minWidth As Integer
            CalcMinAutosizeWidth(minWidth, maxLineLength)
            If AutoScrollMinSize.Width < minWidth Then AutoScrollMinSize = New Size(minWidth, AutoScrollMinSize.Height)
        End Sub

        Private Function RecalcMaxLineLength() As Integer
            Dim maxLineLength As Integer = 0
            Dim lines_ As TextSource = Me.lines_
            Dim count As Integer = lines_.Count
            Dim charHeight_ As Integer = charHeight_
            Dim topIndent As Integer = Paddings.Top
            TextHeight = topIndent

            For i As Integer = 0 To count - 1
                Dim lineLength As Integer = lines_.GetLineLength(i)
                Dim lineInfo As LineInfo = LineInfos(i)
                If lineLength > maxLineLength AndAlso lineInfo.VisibleState = VisibleState.Visible Then maxLineLength = lineLength
                lineInfo.startY = TextHeight
                TextHeight += lineInfo.WordWrapStringsCount * charHeight_ + lineInfo.bottomPadding
                LineInfos(i) = lineInfo
            Next

            TextHeight -= topIndent
            Return maxLineLength
        End Function

        Private Function GetMaxLineWordWrapedWidth() As Integer
            If wordWrap_ Then

                Select Case wordWrapMode
                    Case wordWrapMode.WordWrapControlWidth, wordWrapMode.CharWrapControlWidth
                        Return ClientSize.Width
                    Case wordWrapMode.WordWrapPreferredWidth, wordWrapMode.CharWrapPreferredWidth
                        Return LeftIndent + preferredLineWidth * CharWidth + 2 + Paddings.Left + Paddings.Right
                End Select
            End If

            Return Integer.MaxValue
        End Function

        Private Sub RecalcWordWrap(ByVal fromLine As Integer, ByVal toLine As Integer)
            Dim maxCharsPerLine As Integer = 0
            Dim charWrap As Boolean = False
            toLine = Math.Min(LinesCount - 1, toLine)

            Select Case wordWrapMode
                Case wordWrapMode.WordWrapControlWidth
                    maxCharsPerLine = (ClientSize.Width - LeftIndent - Paddings.Left - Paddings.Right) / CharWidth
                Case wordWrapMode.CharWrapControlWidth
                    maxCharsPerLine = (ClientSize.Width - LeftIndent - Paddings.Left - Paddings.Right) / CharWidth
                    charWrap = True
                Case wordWrapMode.WordWrapPreferredWidth
                    maxCharsPerLine = preferredLineWidth
                Case wordWrapMode.CharWrapPreferredWidth
                    maxCharsPerLine = preferredLineWidth
                    charWrap = True
            End Select

            For iLine As Integer = fromLine To toLine

                If lines_.IsLineLoaded(iLine) Then

                    If Not wordWrap_ Then
                        LineInfos(iLine).CutOffPositions.Clear()
                    Else
                        Dim li As LineInfo = LineInfos(iLine)
                        li.wordWrapIndent = If(WordWrapAutoIndent, lines_(iLine).StartSpacesCount + WordWrapIndent, WordWrapIndent)

                        If wordWrapMode = wordWrapMode.Custom Then
                            RaiseEvent WordWrapNeeded(Me, New WordWrapNeededEventArgs(li.CutOffPositions, ImeAllowed, lines_(iLine)))
                        Else
                            CalcCutOffs(li.CutOffPositions, maxCharsPerLine, maxCharsPerLine - li.wordWrapIndent, ImeAllowed, charWrap, lines_(iLine))
                        End If

                        LineInfos(iLine) = li
                    End If
                End If
            Next

            needRecalc_ = True
        End Sub

        Public Shared Sub CalcCutOffs(ByVal cutOffPositions As List(Of Integer), ByVal maxCharsPerLine As Integer, ByVal maxCharsPerSecondaryLine As Integer, ByVal allowIME As Boolean, ByVal charWrap As Boolean, ByVal line As Line)
            If maxCharsPerSecondaryLine < 1 Then maxCharsPerSecondaryLine = 1
            If maxCharsPerLine < 1 Then maxCharsPerLine = 1
            Dim segmentLength As Integer = 0
            Dim cutOff As Integer = 0
            cutOffPositions.Clear()

            For i As Integer = 0 To line.Count - 1 - 1
                Dim c As Char = line(i).c

                If charWrap Then
                    cutOff = i + 1
                Else

                    If allowIME AndAlso IsCJKLetter(c) Then
                        cutOff = i
                    ElseIf Not Char.IsLetterOrDigit(c) AndAlso c <> "_"c AndAlso c <> "'"c AndAlso c <> " "c AndAlso ((c <> "."c AndAlso c <> ","c) OrElse Not Char.IsDigit(line(i + 1).c)) Then
                        cutOff = Math.Min(i + 1, line.Count - 1)
                    End If
                End If

                segmentLength += 1

                If segmentLength = maxCharsPerLine Then
                    If cutOff = 0 OrElse (cutOffPositions.Count > 0 AndAlso cutOff = cutOffPositions(cutOffPositions.Count - 1)) Then cutOff = i + 1
                    cutOffPositions.Add(cutOff)
                    segmentLength = 1 + i - cutOff
                    maxCharsPerLine = maxCharsPerSecondaryLine
                End If
            Next
        End Sub

        Public Shared Function IsCJKLetter(ByVal c As Char) As Boolean
            Dim code As Integer = Convert.ToInt32(c)
            Return (code >= &H3300 AndAlso code <= &H33FF) OrElse (code >= &HFE30 AndAlso code <= &HFE4F) OrElse (code >= &HF900 AndAlso code <= &HFAFF) OrElse (code >= &H2E80 AndAlso code <= &H2EFF) OrElse (code >= &H31C0 AndAlso code <= &H31EF) OrElse (code >= &H4E00 AndAlso code <= &H9FFF) OrElse (code >= &H3400 AndAlso code <= &H4DBF) OrElse (code >= &H3200 AndAlso code <= &H32FF) OrElse (code >= &H2460 AndAlso code <= &H24FF) OrElse (code >= &H3040 AndAlso code <= &H309F) OrElse (code >= &H2F00 AndAlso code <= &H2FDF) OrElse (code >= &H31A0 AndAlso code <= &H31BF) OrElse (code >= &H4DC0 AndAlso code <= &H4DFF) OrElse (code >= &H3100 AndAlso code <= &H312F) OrElse (code >= &H30A0 AndAlso code <= &H30FF) OrElse (code >= &H31F0 AndAlso code <= &H31FF) OrElse (code >= &H2FF0 AndAlso code <= &H2FFF) OrElse (code >= &H1100 AndAlso code <= &H11FF) OrElse (code >= &HA960 AndAlso code <= &HA97F) OrElse (code >= &HD7B0 AndAlso code <= &HD7FF) OrElse (code >= &H3130 AndAlso code <= &H318F) OrElse (code >= &HAC00 AndAlso code <= &HD7AF)
        End Function

        Protected Overrides Sub OnClientSizeChanged(ByVal e As EventArgs)
            MyBase.OnClientSizeChanged(e)

            If wordWrap_ Then
                NeedRecalc(False, True)
                Invalidate()
            End If

            OnVisibleRangeChanged()
            UpdateScrollbars()
        End Sub

        Friend Sub DoVisibleRectangle(ByVal rect As Rectangle)
            HideHints()
            Dim oldV As Integer = VerticalScroll.Value
            Dim v As Integer = VerticalScroll.Value
            Dim h As Integer = HorizontalScroll.Value

            If rect.Bottom > ClientRectangle.Height Then
                v += rect.Bottom - ClientRectangle.Height
            ElseIf rect.Top < 0 Then
                v += rect.Top
            End If

            If rect.Right > ClientRectangle.Width Then
                h += rect.Right - ClientRectangle.Width
            ElseIf rect.Left < LeftIndent Then
                h += rect.Left - LeftIndent
            End If

            If Not multiline Then v = 0
            v = Math.Max(VerticalScroll.Minimum, v)
            h = Math.Max(HorizontalScroll.Minimum, h)

            Try
                If VerticalScroll.Visible OrElse Not ShowScrollBars Then VerticalScroll.Value = Math.Min(v, VerticalScroll.Maximum)
                If HorizontalScroll.Visible OrElse Not ShowScrollBars Then HorizontalScroll.Value = Math.Min(h, HorizontalScroll.Maximum)
            Catch __unusedArgumentOutOfRangeException1__ As ArgumentOutOfRangeException
            End Try

            UpdateScrollbars()
            RestoreHints()
            If oldV <> VerticalScroll.Value Then OnVisibleRangeChanged()
        End Sub

        Public Sub UpdateScrollbars()
            If ShowScrollBars Then
                MyBase.AutoScrollMinSize -= New Size(1, 0)
                MyBase.AutoScrollMinSize += New Size(1, 0)
            Else
                PerformLayout()
            End If

            If IsHandleCreated Then BeginInvoke(CType(AddressOf OnScrollbarsUpdated, MethodInvoker))
        End Sub

        Protected Overridable Sub OnScrollbarsUpdated()
            RaiseEvent ScrollbarsUpdated(Me, EventArgs.Empty)
        End Sub

        Public Sub DoCaretVisible()
            Invalidate()
            Recalc()
            Dim car As Point = PlaceToPoint(selection_.Start)
            car.Offset(-CharWidth, 0)
            DoVisibleRectangle(New Rectangle(car, New Size(2 * CharWidth, 2 * charHeight_)))
        End Sub

        Public Sub ScrollLeft()
            Invalidate()
            HorizontalScroll.Value = 0
            AutoScrollMinSize -= New Size(1, 0)
            AutoScrollMinSize += New Size(1, 0)
        End Sub

        Public Sub DoSelectionVisible()
            If LineInfos(selection_.[End].iLine).VisibleState <> VisibleState.Visible Then ExpandBlock(selection_.[End].iLine)
            If LineInfos(selection_.Start.iLine).VisibleState <> VisibleState.Visible Then ExpandBlock(selection_.Start.iLine)
            Recalc()
            DoVisibleRectangle(New Rectangle(PlaceToPoint(New Place(0, selection_.[End].iLine)), New Size(2 * CharWidth, 2 * charHeight_)))
            Dim car As Point = PlaceToPoint(selection_.Start)
            Dim car2 As Point = PlaceToPoint(selection_.[End])
            car.Offset(-CharWidth, -ClientSize.Height / 2)
            DoVisibleRectangle(New Rectangle(car, New Size(Math.Abs(car2.X - car.X), ClientSize.Height)))
            Invalidate()
        End Sub

        Public Sub DoRangeVisible(ByVal range As Range)
            DoRangeVisible(range, False)
        End Sub

        Public Sub DoRangeVisible(ByVal range As Range, ByVal tryToCentre As Boolean)
            range = range.Clone()
            range.Normalize()
            range.[End] = New Place(range.[End].iChar, Math.Min(range.[End].iLine, range.Start.iLine + ClientSize.Height / charHeight_))
            If LineInfos(range.[End].iLine).VisibleState <> VisibleState.Visible Then ExpandBlock(range.[End].iLine)
            If LineInfos(range.Start.iLine).VisibleState <> VisibleState.Visible Then ExpandBlock(range.Start.iLine)
            Recalc()
            Dim h As Integer = (1 + range.[End].iLine - range.Start.iLine) * charHeight_
            Dim p As Point = PlaceToPoint(New Place(0, range.Start.iLine))

            If tryToCentre Then
                p.Offset(0, -ClientSize.Height / 2)
                h = ClientSize.Height
            End If

            DoVisibleRectangle(New Rectangle(p, New Size(2 * CharWidth, h)))
            Invalidate()
        End Sub

        Protected Overrides Sub OnKeyUp(ByVal e As KeyEventArgs)
            MyBase.OnKeyUp(e)
            If e.KeyCode = Keys.ShiftKey Then lastModifiers = lastModifiers And Not Keys.Shift
            If e.KeyCode = Keys.Alt Then lastModifiers = lastModifiers And Not Keys.Alt
            If e.KeyCode = Keys.ControlKey Then lastModifiers = lastModifiers And Not Keys.Control
        End Sub

        Private findCharMode As Boolean

        Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)
            If middleClickScrollingActivated Then Return
            MyBase.OnKeyDown(e)
            If Focused Then lastModifiers = e.Modifiers
            handledChar = False

            If e.Handled Then
                handledChar = True
                Return
            End If

            If ProcessKey(e.KeyData) Then Return
            e.Handled = True
            DoCaretVisible()
            Invalidate()
        End Sub

        Protected Overrides Function ProcessDialogKey(ByVal keyData As Keys) As Boolean
            If (keyData And Keys.Alt) > 0 Then

                If HotkeysMapping.ContainsKey(keyData) Then
                    ProcessKey(keyData)
                    Return True
                End If
            End If

            Return MyBase.ProcessDialogKey(keyData)
        End Function

        Shared scrollActions As Dictionary(Of FCTBAction, Boolean) = New Dictionary(Of FCTBAction, Boolean)() From {
        {FCTBAction.ScrollDown, True},
        {FCTBAction.ScrollUp, True},
        {FCTBAction.ZoomOut, True},
        {FCTBAction.ZoomIn, True},
        {FCTBAction.ZoomNormal, True}
    }

        Public Overridable Function ProcessKey(ByVal keyData As Keys) As Boolean
            Dim a As KeyEventArgs = New KeyEventArgs(keyData)
            If a.KeyCode = Keys.Tab AndAlso Not AcceptsTab Then Return False

            If macrosManager IsNot Nothing Then
                If Not HotkeysMapping.ContainsKey(keyData) OrElse (HotkeysMapping(keyData) <> FCTBAction.MacroExecute AndAlso HotkeysMapping(keyData) <> FCTBAction.MacroRecord) Then macrosManager.ProcessKey(keyData)
            End If

            If HotkeysMapping.ContainsKey(keyData) Then
                Dim act = HotkeysMapping(keyData)
                DoAction(act)
                If scrollActions.ContainsKey(act) Then Return True

                If keyData = Keys.Tab OrElse keyData = (Keys.Tab Or Keys.Shift) Then
                    handledChar = True
                    Return True
                End If
            Else
                If a.KeyCode = Keys.Alt Then Return True
                If (a.Modifiers And Keys.Control) <> 0 Then Return True

                If (a.Modifiers And Keys.Alt) <> 0 Then
                    If (MouseButtons And MouseButtons.Left) <> 0 Then CheckAndChangeSelectionType()
                    Return True
                End If

                If a.KeyCode = Keys.ShiftKey Then Return True
            End If

            Return False
        End Function

        Private Sub DoAction(ByVal action As FCTBAction)
            Select Case action
                Case FCTBAction.ZoomIn
                    ChangeFontSize(2)
                Case FCTBAction.ZoomOut
                    ChangeFontSize(-2)
                Case FCTBAction.ZoomNormal
                    RestoreFontSize()
                Case FCTBAction.ScrollDown
                    DoScrollVertical(1, -1)
                Case FCTBAction.ScrollUp
                    DoScrollVertical(1, 1)
                Case FCTBAction.GoToDialog
                    ShowGoToDialog()
                Case FCTBAction.FindDialog
                    ShowFindDialog()
                Case FCTBAction.FindChar
                    findCharMode = True
                Case FCTBAction.FindNext

                    If findForm Is Nothing OrElse findForm.tbFind.Text = "" Then
                        ShowFindDialog()
                    Else
                        findForm.FindNext(findForm.tbFind.Text)
                    End If

                Case FCTBAction.ReplaceDialog
                    ShowReplaceDialog()
                Case FCTBAction.Copy
                    Copy()
                Case FCTBAction.CommentSelected
                    CommentSelected()
                Case FCTBAction.Cut
                    If Not selection_.[ReadOnly] Then Cut()
                Case FCTBAction.Paste
                    If Not selection_.[ReadOnly] Then Paste()
                Case FCTBAction.SelectAll
                    selection_.SelectAll()
                Case FCTBAction.Undo
                    If Not [ReadOnly] Then Undo()
                Case FCTBAction.Redo
                    If Not [ReadOnly] Then Redo()
                Case FCTBAction.LowerCase
                    If Not selection_.[ReadOnly] Then LowerCase()
                Case FCTBAction.UpperCase
                    If Not selection_.[ReadOnly] Then UpperCase()
                Case FCTBAction.IndentDecrease

                    If Not selection_.[ReadOnly] Then
                        Dim sel = selection_.Clone()

                        If sel.Start.iLine = sel.[End].iLine Then
                            Dim line = Me(sel.Start.iLine)

                            If sel.Start.iChar = 0 AndAlso sel.[End].iChar = line.Count Then
                                selection_ = New Range(Me, line.StartSpacesCount, sel.Start.iLine, line.Count, sel.Start.iLine)
                            ElseIf sel.Start.iChar = line.Count AndAlso sel.[End].iChar = 0 Then
                                selection_ = New Range(Me, line.Count, sel.Start.iLine, line.StartSpacesCount, sel.Start.iLine)
                            End If
                        End If

                        DecreaseIndent()
                    End If

                Case FCTBAction.IndentIncrease

                    If Not selection_.[ReadOnly] Then
                        Dim sel = selection_.Clone()
                        Dim inverted = sel.Start > sel.[End]
                        sel.Normalize()
                        Dim spaces = Me(sel.Start.iLine).StartSpacesCount

                        If sel.Start.iLine <> sel.[End].iLine OrElse (sel.Start.iChar <= spaces AndAlso sel.[End].iChar = Me(sel.Start.iLine).Count) OrElse sel.[End].iChar <= spaces Then
                            IncreaseIndent()

                            If sel.Start.iLine = sel.[End].iLine AndAlso Not sel.IsEmpty Then
                                selection_ = New Range(Me, Me(sel.Start.iLine).StartSpacesCount, sel.[End].iLine, Me(sel.Start.iLine).Count, sel.[End].iLine)
                                If inverted Then selection_.Inverse()
                            End If
                        Else
                            ProcessKey(vbTab, Keys.None)
                        End If
                    End If

                Case FCTBAction.AutoIndentChars
                    If Not selection_.[ReadOnly] Then DoAutoIndentChars(selection_.Start.iLine)
                Case FCTBAction.NavigateBackward
                    NavigateBackward()
                Case FCTBAction.NavigateForward
                    NavigateForward()
                Case FCTBAction.UnbookmarkLine
                    UnbookmarkLine(selection_.Start.iLine)
                Case FCTBAction.BookmarkLine
                    BookmarkLine(selection_.Start.iLine)
                Case FCTBAction.GoNextBookmark
                    GotoNextBookmark(selection_.Start.iLine)
                Case FCTBAction.GoPrevBookmark
                    GotoPrevBookmark(selection_.Start.iLine)
                Case FCTBAction.ClearWordLeft
                    If OnKeyPressing(vbBack) Then Exit Select

                    If Not selection_.[ReadOnly] Then
                        If Not selection_.IsEmpty Then ClearSelected()
                        selection_.GoWordLeft(True)
                        If Not selection_.[ReadOnly] Then ClearSelected()
                    End If

                    OnKeyPressed(vbBack)
                Case FCTBAction.ReplaceMode
                    If Not [ReadOnly] Then isReplaceMode = Not isReplaceMode
                Case FCTBAction.DeleteCharRight

                    If Not selection_.[ReadOnly] Then
                        If OnKeyPressing(ChrW(&HFF)) Then Exit Select

                        If Not selection_.IsEmpty Then
                            ClearSelected()
                        Else
                            If Me(selection_.Start.iLine).StartSpacesCount = Me(selection_.Start.iLine).Count Then RemoveSpacesAfterCaret()

                            If Not selection_.IsReadOnlyRightChar() Then

                                If selection_.GoRightThroughFolded() Then
                                    Dim iLine As Integer = selection_.Start.iLine
                                    InsertChar(vbBack)

                                    If iLine <> selection_.Start.iLine AndAlso AutoIndent Then
                                        If selection_.Start.iChar > 0 Then RemoveSpacesAfterCaret()
                                    End If
                                End If
                            End If
                        End If

                        If AutoIndentChars Then DoAutoIndentChars(selection_.Start.iLine)
                        OnKeyPressed(ChrW(&HFF))
                    End If

                Case FCTBAction.ClearWordRight
                    If OnKeyPressing(ChrW(&HFF)) Then Exit Select

                    If Not selection_.[ReadOnly] Then
                        If Not selection_.IsEmpty Then ClearSelected()
                        selection_.GoWordRight(True)
                        If Not selection_.[ReadOnly] Then ClearSelected()
                    End If

                    OnKeyPressed(ChrW(&HFF))
                Case FCTBAction.GoWordLeft
                    selection_.GoWordLeft(False)
                Case FCTBAction.GoWordLeftWithSelection
                    selection_.GoWordLeft(True)
                Case FCTBAction.GoLeft
                    selection_.GoLeft(False)
                Case FCTBAction.GoLeftWithSelection
                    selection_.GoLeft(True)
                Case FCTBAction.GoLeft_ColumnSelectionMode
                    CheckAndChangeSelectionType()
                    If selection_.ColumnSelectionMode Then selection_.GoLeft_ColumnSelectionMode()
                    Invalidate()
                Case FCTBAction.GoWordRight
                    selection_.GoWordRight(False, True)
                Case FCTBAction.GoWordRightWithSelection
                    selection_.GoWordRight(True, True)
                Case FCTBAction.GoRight
                    selection_.GoRight(False)
                Case FCTBAction.GoRightWithSelection
                    selection_.GoRight(True)
                Case FCTBAction.GoRight_ColumnSelectionMode
                    CheckAndChangeSelectionType()
                    If selection_.ColumnSelectionMode Then selection_.GoRight_ColumnSelectionMode()
                    Invalidate()
                Case FCTBAction.GoUp
                    selection_.GoUp(False)
                    ScrollLeft()
                Case FCTBAction.GoUpWithSelection
                    selection_.GoUp(True)
                    ScrollLeft()
                Case FCTBAction.GoUp_ColumnSelectionMode
                    CheckAndChangeSelectionType()
                    If selection_.ColumnSelectionMode Then selection_.GoUp_ColumnSelectionMode()
                    Invalidate()
                Case FCTBAction.MoveSelectedLinesUp
                    If Not selection_.ColumnSelectionMode Then MoveSelectedLinesUp()
                Case FCTBAction.GoDown
                    selection_.GoDown(False)
                    ScrollLeft()
                Case FCTBAction.GoDownWithSelection
                    selection_.GoDown(True)
                    ScrollLeft()
                Case FCTBAction.GoDown_ColumnSelectionMode
                    CheckAndChangeSelectionType()
                    If selection_.ColumnSelectionMode Then selection_.GoDown_ColumnSelectionMode()
                    Invalidate()
                Case FCTBAction.MoveSelectedLinesDown
                    If Not selection_.ColumnSelectionMode Then MoveSelectedLinesDown()
                Case FCTBAction.GoPageUp
                    selection_.GoPageUp(False)
                    ScrollLeft()
                Case FCTBAction.GoPageUpWithSelection
                    selection_.GoPageUp(True)
                    ScrollLeft()
                Case FCTBAction.GoPageDown
                    selection_.GoPageDown(False)
                    ScrollLeft()
                Case FCTBAction.GoPageDownWithSelection
                    selection_.GoPageDown(True)
                    ScrollLeft()
                Case FCTBAction.GoFirstLine
                    selection_.GoFirst(False)
                Case FCTBAction.GoFirstLineWithSelection
                    selection_.GoFirst(True)
                Case FCTBAction.GoHome
                    GoHome(False)
                    ScrollLeft()
                Case FCTBAction.GoHomeWithSelection
                    GoHome(True)
                    ScrollLeft()
                Case FCTBAction.GoLastLine
                    selection_.GoLast(False)
                Case FCTBAction.GoLastLineWithSelection
                    selection_.GoLast(True)
                Case FCTBAction.GoEnd
                    selection_.GoEnd(False)
                Case FCTBAction.GoEndWithSelection
                    selection_.GoEnd(True)
                Case FCTBAction.ClearHints
                    ClearHints()
                    If macrosManager IsNot Nothing Then macrosManager.IsRecording = False
                Case FCTBAction.MacroRecord

                    If macrosManager IsNot Nothing Then
                        If macrosManager.AllowMacroRecordingByUser Then macrosManager.IsRecording = Not macrosManager.IsRecording
                        If macrosManager.IsRecording Then macrosManager.ClearMacros()
                    End If

                Case FCTBAction.MacroExecute

                    If macrosManager IsNot Nothing Then
                        macrosManager.IsRecording = False
                        macrosManager.ExecuteMacros()
                    End If

                Case FCTBAction.CustomAction1, FCTBAction.CustomAction2, FCTBAction.CustomAction3, FCTBAction.CustomAction4, FCTBAction.CustomAction5, FCTBAction.CustomAction6, FCTBAction.CustomAction7, FCTBAction.CustomAction8, FCTBAction.CustomAction9, FCTBAction.CustomAction10, FCTBAction.CustomAction11, FCTBAction.CustomAction12, FCTBAction.CustomAction13, FCTBAction.CustomAction14, FCTBAction.CustomAction15, FCTBAction.CustomAction16, FCTBAction.CustomAction17, FCTBAction.CustomAction18, FCTBAction.CustomAction19, FCTBAction.CustomAction20
                    OnCustomAction(New CustomActionEventArgs(action))
            End Select
        End Sub

        Protected Overridable Sub OnCustomAction(ByVal e As CustomActionEventArgs)
            RaiseEvent CustomAction(Me, e)
        End Sub

        Private originalFont As Font

        Private Sub RestoreFontSize()
            zoom = 100
        End Sub

        Public Function GotoNextBookmark(ByVal iLine As Integer) As Boolean
            Dim nearestBookmark As Bookmark = Nothing
            Dim minNextLineIndex As Integer = Integer.MaxValue
            Dim minBookmark As Bookmark = Nothing
            Dim minLineIndex As Integer = Integer.MaxValue

            For Each bookmark As Bookmark In bookmarks

                If bookmark.LineIndex < minLineIndex Then
                    minLineIndex = bookmark.LineIndex
                    minBookmark = bookmark
                End If

                If bookmark.LineIndex > iLine AndAlso bookmark.LineIndex < minNextLineIndex Then
                    minNextLineIndex = bookmark.LineIndex
                    nearestBookmark = bookmark
                End If
            Next

            If nearestBookmark IsNot Nothing Then
                nearestBookmark.DoVisible()
                Return True
            ElseIf minBookmark IsNot Nothing Then
                minBookmark.DoVisible()
                Return True
            End If

            Return False
        End Function

        Public Function GotoPrevBookmark(ByVal iLine As Integer) As Boolean
            Dim nearestBookmark As Bookmark = Nothing
            Dim maxPrevLineIndex As Integer = -1
            Dim maxBookmark As Bookmark = Nothing
            Dim maxLineIndex As Integer = -1

            For Each bookmark As Bookmark In bookmarks

                If bookmark.LineIndex > maxLineIndex Then
                    maxLineIndex = bookmark.LineIndex
                    maxBookmark = bookmark
                End If

                If bookmark.LineIndex < iLine AndAlso bookmark.LineIndex > maxPrevLineIndex Then
                    maxPrevLineIndex = bookmark.LineIndex
                    nearestBookmark = bookmark
                End If
            Next

            If nearestBookmark IsNot Nothing Then
                nearestBookmark.DoVisible()
                Return True
            ElseIf maxBookmark IsNot Nothing Then
                maxBookmark.DoVisible()
                Return True
            End If

            Return False
        End Function

        Public Overridable Sub BookmarkLine(ByVal iLine As Integer)
            If Not bookmarks.Contains(iLine) Then bookmarks.Add(iLine)
        End Sub

        Public Overridable Sub UnbookmarkLine(ByVal iLine As Integer)
            bookmarks.Remove(iLine)
        End Sub

        Public Overridable Sub MoveSelectedLinesDown()
            Dim prevSelection As Range = selection_.Clone()
            selection_.Expand()

            If Not selection_.[ReadOnly] Then
                Dim iLine As Integer = selection_.Start.iLine

                If selection_.[End].iLine >= LinesCount - 1 Then
                    selection_ = prevSelection
                    Return
                End If

                Dim text As String = SelectedText
                Dim temp = New List(Of Integer)()

                For i As Integer = selection_.Start.iLine To selection_.[End].iLine
                    temp.Add(i)
                Next

                RemoveLines(temp)
                selection_.Start = New Place(GetLineLength(iLine), iLine)
                SelectedText = vbLf & text
                selection_.Start = New Place(prevSelection.Start.iChar, prevSelection.Start.iLine + 1)
                selection_.[End] = New Place(prevSelection.[End].iChar, prevSelection.[End].iLine + 1)
            Else
                selection_ = prevSelection
            End If
        End Sub

        Public Overridable Sub MoveSelectedLinesUp()
            Dim prevSelection As Range = selection_.Clone()
            selection_.Expand()

            If Not selection_.[ReadOnly] Then
                Dim iLine As Integer = selection_.Start.iLine

                If iLine = 0 Then
                    selection_ = prevSelection
                    Return
                End If

                Dim text As String = SelectedText
                Dim temp = New List(Of Integer)()

                For i As Integer = selection_.Start.iLine To selection_.[End].iLine
                    temp.Add(i)
                Next

                RemoveLines(temp)
                selection_.Start = New Place(0, iLine - 1)
                SelectedText = text & vbLf
                selection_.Start = New Place(prevSelection.Start.iChar, prevSelection.Start.iLine - 1)
                selection_.[End] = New Place(prevSelection.[End].iChar, prevSelection.[End].iLine - 1)
            Else
                selection_ = prevSelection
            End If
        End Sub

        Private Sub GoHome(ByVal shift As Boolean)
            selection_.BeginUpdate()

            Try
                Dim iLine As Integer = selection_.Start.iLine
                Dim spaces As Integer = Me(iLine).StartSpacesCount

                If selection_.Start.iChar <= spaces Then
                    selection_.GoHome(shift)
                Else
                    selection_.GoHome(shift)

                    For i As Integer = 0 To spaces - 1
                        selection_.GoRight(shift)
                    Next
                End If

            Finally
                selection_.EndUpdate()
            End Try
        End Sub

        Public Overridable Sub UpperCase()
            Dim old As Range = selection_.Clone()
            SelectedText = SelectedText.ToUpper()
            selection_.Start = old.Start
            selection_.[End] = old.[End]
        End Sub

        Public Overridable Sub LowerCase()
            Dim old As Range = selection_.Clone()
            SelectedText = SelectedText.ToLower()
            selection_.Start = old.Start
            selection_.[End] = old.[End]
        End Sub

        Public Overridable Sub TitleCase()
            Dim old As Range = selection_.Clone()
            SelectedText = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(SelectedText.ToLower())
            selection_.Start = old.Start
            selection_.[End] = old.[End]
        End Sub

        Public Overridable Sub SentenceCase()
            Dim old As Range = selection_.Clone()
            Dim lowerCase = SelectedText.ToLower()
            Dim r = New Regex("(^\S)|[\.\?!:]\s+(\S)", RegexOptions.ExplicitCapture)
            SelectedText = r.Replace(lowerCase, Function(s) s.Value.ToUpper())
            selection_.Start = old.Start
            selection_.[End] = old.[End]
        End Sub

        Public Sub CommentSelected()
            CommentSelected(CommentPrefix)
        End Sub

        Public Overridable Sub CommentSelected(ByVal commentPrefix As String)
            If String.IsNullOrEmpty(commentPrefix) Then Return
            selection_.Normalize()
            Dim isCommented As Boolean = lines_(selection_.Start.iLine).Text.TrimStart().StartsWith(commentPrefix)

            If isCommented Then
                RemoveLinePrefix(commentPrefix)
            Else
                InsertLinePrefix(commentPrefix)
            End If
        End Sub

        Public Sub OnKeyPressing(ByVal args As KeyPressEventArgs)
            RaiseEvent KeyPressing(Me, args)
        End Sub

        Private Function OnKeyPressing(ByVal c As Char) As Boolean
            If findCharMode Then
                findCharMode = False
                FindChar(c)
                Return True
            End If

            Dim args = New KeyPressEventArgs(c)
            OnKeyPressing(args)
            Return args.Handled
        End Function

        Public Sub OnKeyPressed(ByVal c As Char)
            Dim args = New KeyPressEventArgs(c)
            RaiseEvent KeyPressed(Me, args)
        End Sub

        Protected Overrides Function ProcessMnemonic(ByVal charCode As Char) As Boolean
            If middleClickScrollingActivated Then Return False

            If Focused Then
                Return ProcessKey(charCode, lastModifiers) OrElse MyBase.ProcessMnemonic(charCode)
            Else
                Return False
            End If
        End Function

        Const WM_CHAR As Integer = &H102

        Protected Overrides Function ProcessKeyMessage(ByRef m As Message) As Boolean
            If m.Msg = WM_CHAR Then ProcessMnemonic(Convert.ToChar(m.WParam.ToInt32()))
            Return MyBase.ProcessKeyMessage(m)
        End Function

        Public Overridable Function ProcessKey(ByVal c As Char, ByVal modifiers As Keys) As Boolean
            If handledChar Then Return True
            If macrosManager IsNot Nothing Then macrosManager.ProcessKey(c, modifiers)

            If c = vbBack AndAlso (modifiers = Keys.None OrElse modifiers = Keys.Shift OrElse (modifiers And Keys.Alt) <> 0) Then
                If [ReadOnly] OrElse Not Enabled Then Return False
                If OnKeyPressing(c) Then Return True
                If selection_.[ReadOnly] Then Return False

                If Not selection_.IsEmpty Then
                    ClearSelected()
                ElseIf Not selection_.IsReadOnlyLeftChar() Then
                    InsertChar(vbBack)
                End If

                If AutoIndentChars Then DoAutoIndentChars(selection_.Start.iLine)
                OnKeyPressed(vbBack)
                Return True
            End If

            If Char.IsControl(c) AndAlso c <> vbCr AndAlso c <> vbTab Then Return False
            If [ReadOnly] OrElse Not Enabled Then Return False
            If modifiers <> Keys.None AndAlso modifiers <> Keys.Shift AndAlso modifiers <> (Keys.Control Or Keys.Alt) AndAlso modifiers <> (Keys.Shift Or Keys.Control Or Keys.Alt) AndAlso (modifiers <> (Keys.Alt) OrElse Char.IsLetterOrDigit(c)) Then Return False
            Dim sourceC As Char = c
            If OnKeyPressing(sourceC) Then Return True
            If selection_.[ReadOnly] Then Return False
            If c = vbCr AndAlso Not AcceptsReturn Then Return False
            If c = vbCr Then c = vbLf

            If isReplaceMode Then
                selection_.GoRight(True)
                selection_.Inverse()
            End If

            If Not selection_.[ReadOnly] Then
                If Not DoAutocompleteBrackets(c) Then InsertChar(c)
            End If

            If c = vbLf OrElse AutoIndentExistingLines Then DoAutoIndentIfNeed()
            If AutoIndentChars Then DoAutoIndentChars(selection_.Start.iLine)
            DoCaretVisible()
            Invalidate()
            OnKeyPressed(sourceC)
            Return True
        End Function

        <Description("Enables AutoIndentChars mode")>
        <DefaultValue(True)>
        Public Property AutoIndentChars As Boolean
        <Description("Regex patterns for AutoIndentChars (one regex per line)")>
        <Editor("System.ComponentModel.Design.MultilineStringEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", GetType(UITypeEditor))>
        <DefaultValue("^\s*[\w\.]+\s*(?<range>=)\s*(?<range>[^;]+);")>
        Public Property AutoIndentCharsPatterns As String

        Public Sub DoAutoIndentChars(ByVal iLine As Integer)
            Dim patterns = AutoIndentCharsPatterns.Split(New Char() {vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each pattern In patterns
                Dim m = Regex.Match(Me(iLine).Text, pattern)

                If m.Success Then
                    DoAutoIndentChars(iLine, New Regex(pattern))
                    Exit For
                End If
            Next
        End Sub

        Protected Sub DoAutoIndentChars(ByVal iLine As Integer, ByVal regex As Regex)
            Dim oldSel = selection_.Clone()
            Dim captures = New SortedDictionary(Of Integer, CaptureCollection)()
            Dim texts = New SortedDictionary(Of Integer, String)()
            Dim maxCapturesCount = 0
            Dim spaces = Me(iLine).StartSpacesCount

            For i = iLine To 0
                If spaces <> Me(i).StartSpacesCount Then Exit For
                Dim text = Me(i).Text
                Dim m = regex.Match(text)

                If m.Success Then
                    captures(i) = m.Groups("range").Captures
                    texts(i) = text
                    If captures(i).Count > maxCapturesCount Then maxCapturesCount = captures(i).Count
                Else
                    Exit For
                End If
            Next

            For i = iLine + 1 To LinesCount - 1
                If spaces <> Me(i).StartSpacesCount Then Exit For
                Dim text = Me(i).Text
                Dim m = regex.Match(text)

                If m.Success Then
                    captures(i) = m.Groups("range").Captures
                    texts(i) = text
                    If captures(i).Count > maxCapturesCount Then maxCapturesCount = captures(i).Count
                Else
                    Exit For
                End If
            Next

            Dim changed = New Dictionary(Of Integer, Boolean)()
            Dim was = False

            For iCapture As Integer = maxCapturesCount - 1 To 0
                Dim maxDist = 0

                For Each i In captures.Keys
                    Dim caps = captures(i)
                    If caps.Count <= iCapture Then Continue For
                    Dim dist = 0
                    Dim cap = caps(iCapture)
                    Dim index = cap.Index
                    Dim text = texts(i)

                    While index > 0 AndAlso text(index - 1) = " "c
                        index -= 1
                    End While

                    If iCapture = 0 Then
                        dist = index
                    Else
                        dist = index - caps(iCapture - 1).Index - 1
                    End If

                    If dist > maxDist Then maxDist = dist
                Next

                For Each i In New List(Of Integer)(texts.Keys)
                    If captures(i).Count <= iCapture Then Continue For
                    Dim dist = 0
                    Dim cap = captures(i)(iCapture)

                    If iCapture = 0 Then
                        dist = cap.Index
                    Else
                        dist = cap.Index - captures(i)(iCapture - 1).Index - 1
                    End If

                    Dim addSpaces = maxDist - dist + 1
                    If addSpaces = 0 Then Continue For
                    If oldSel.Start.iLine = i AndAlso oldSel.Start.iChar > cap.Index Then oldSel.Start = New Place(oldSel.Start.iChar + addSpaces, i)

                    If addSpaces > 0 Then
                        texts(i) = texts(i).Insert(cap.Index, New String(" "c, addSpaces))
                    Else
                        texts(i) = texts(i).Remove(cap.Index + addSpaces, -addSpaces)
                    End If

                    changed(i) = True
                    was = True
                Next
            Next

            If was Then
                selection_.BeginUpdate()
                BeginAutoUndo()
                BeginUpdate()
                TextSource.Manager.ExecuteCommand(New SelectCommand(TextSource))

                For Each i In texts.Keys

                    If changed.ContainsKey(i) Then
                        selection_ = New Range(Me, 0, i, Me(i).Count, i)
                        If Not selection_.[ReadOnly] Then InsertText(texts(i))
                    End If
                Next

                selection_ = oldSel
                EndUpdate()
                EndAutoUndo()
                selection_.EndUpdate()
            End If
        End Sub

        Private Function DoAutocompleteBrackets(ByVal c As Char) As Boolean
            If AutoCompleteBrackets Then

                If Not selection_.ColumnSelectionMode Then

                    For i As Integer = 1 To autoCompleteBracketsList_.Length - 1 Step 2

                        If c = autoCompleteBracketsList_(i) AndAlso c = selection_.CharAfterStart Then
                            selection_.GoRight()
                            Return True
                        End If
                    Next
                End If

                For i As Integer = 0 To autoCompleteBracketsList_.Length - 1 Step 2

                    If c = autoCompleteBracketsList_(i) Then
                        InsertBrackets(autoCompleteBracketsList_(i), autoCompleteBracketsList_(i + 1))
                        Return True
                    End If
                Next
            End If

            Return False
        End Function

        Private Function InsertBrackets(ByVal left As Char, ByVal right As Char) As Boolean
            If selection_.ColumnSelectionMode Then
                Dim range = selection_.Clone()
                range.Normalize()
                selection_.BeginUpdate()
                BeginAutoUndo()
                selection_ = New Range(Me, range.Start.iChar, range.Start.iLine, range.Start.iChar, range.[End].iLine) With {
                .ColumnSelectionMode = True
            }
                InsertChar(left)
                selection_ = New Range(Me, range.[End].iChar + 1, range.Start.iLine, range.[End].iChar + 1, range.[End].iLine) With {
                .ColumnSelectionMode = True
            }
                InsertChar(right)
                If range.IsEmpty Then selection_ = New Range(Me, range.[End].iChar + 1, range.Start.iLine, range.[End].iChar + 1, range.[End].iLine) With {
                .ColumnSelectionMode = True
            }
                EndAutoUndo()
                selection_.EndUpdate()
            ElseIf selection_.IsEmpty Then
                InsertText(left & "" & right)
                selection_.GoLeft()
            Else
                InsertText(left & SelectedText & right)
            End If

            Return True
        End Function

        Protected Overridable Sub FindChar(ByVal c As Char)
            If c = vbCr Then c = vbLf
            Dim r = selection_.Clone()

            While r.GoRight()

                If r.CharBeforeStart = c Then
                    selection_ = r
                    DoCaretVisible()
                    Return
                End If
            End While
        End Sub

        Public Overridable Sub DoAutoIndentIfNeed()
            If selection_.ColumnSelectionMode Then Return

            If AutoIndent Then
                DoCaretVisible()
                Dim needSpaces As Integer = CalcAutoIndent(selection_.Start.iLine)

                If Me(selection_.Start.iLine).AutoIndentSpacesNeededCount <> needSpaces Then
                    DoAutoIndent(selection_.Start.iLine)
                    Me(selection_.Start.iLine).AutoIndentSpacesNeededCount = needSpaces
                End If
            End If
        End Sub

        Private Sub RemoveSpacesAfterCaret()
            If Not selection_.IsEmpty Then Return
            Dim [end] As Place = selection_.Start

            While selection_.CharAfterStart = " "c
                selection_.GoRight(True)
            End While

            ClearSelected()
        End Sub

        Public Overridable Sub DoAutoIndent(ByVal iLine As Integer)
            If selection_.ColumnSelectionMode Then Return
            Dim oldStart As Place = selection_.Start
            Dim needSpaces As Integer = CalcAutoIndent(iLine)
            Dim spaces As Integer = lines_(iLine).StartSpacesCount
            Dim needToInsert As Integer = needSpaces - spaces
            If needToInsert < 0 Then needToInsert = -Math.Min(-needToInsert, spaces)
            If needToInsert = 0 Then Return
            selection_.Start = New Place(0, iLine)

            If needToInsert > 0 Then
                InsertText(New String(" "c, needToInsert))
            Else
                selection_.Start = New Place(0, iLine)
                selection_.[End] = New Place(-needToInsert, iLine)
                ClearSelected()
            End If

            selection_.Start = New Place(Math.Min(lines_(iLine).Count, Math.Max(0, oldStart.iChar + needToInsert)), iLine)
        End Sub

        Public Overridable Function CalcAutoIndent(ByVal iLine As Integer) As Integer
            If iLine < 0 OrElse iLine >= LinesCount Then Return 0
            Dim calculator As EventHandler(Of AutoIndentEventArgs) = AutoIndentNeeded

            If calculator Is Nothing Then

                If language <> language.Custom AndAlso SyntaxHighlighter IsNot Nothing Then
                    calculator = SyntaxHighlighter.AutoIndentNeeded
                Else
                    calculator = AddressOf CalcAutoIndentShiftByCodeFolding
                End If
            End If

            Dim needSpaces As Integer = 0
            Dim stack = New Stack(Of AutoIndentEventArgs)()
            Dim i As Integer

            For i = iLine - 1 To 0
                Dim args = New AutoIndentEventArgs(i, lines_(i).Text, If(i > 0, lines_(i - 1).Text, ""), TabLength, 0)
                calculator(Me, args)
                stack.Push(args)
                If args.Shift = 0 AndAlso args.AbsoluteIndentation = 0 AndAlso args.LineText.Trim() <> "" Then Exit For
            Next

            Dim indent As Integer = lines_(If(i >= 0, i, 0)).StartSpacesCount

            While stack.Count <> 0
                Dim arg = stack.Pop()

                If arg.AbsoluteIndentation <> 0 Then
                    indent = arg.AbsoluteIndentation + arg.ShiftNextLines
                Else
                    indent += arg.ShiftNextLines
                End If
            End While

            Dim a = New AutoIndentEventArgs(iLine, lines_(iLine).Text, If(iLine > 0, lines_(iLine - 1).Text, ""), TabLength, indent)
            calculator(Me, a)
            needSpaces = a.AbsoluteIndentation + a.Shift
            Return needSpaces
        End Function

        Friend Overridable Sub CalcAutoIndentShiftByCodeFolding(ByVal sender As Object, ByVal args As AutoIndentEventArgs)
            If String.IsNullOrEmpty(lines_(args.iLine).FoldingEndMarker) AndAlso Not String.IsNullOrEmpty(lines_(args.iLine).FoldingStartMarker) Then
                args.ShiftNextLines = TabLength
                Return
            End If

            If Not String.IsNullOrEmpty(lines_(args.iLine).FoldingEndMarker) AndAlso String.IsNullOrEmpty(lines_(args.iLine).FoldingStartMarker) Then
                args.Shift = -TabLength
                args.ShiftNextLines = -TabLength
                Return
            End If
        End Sub

        Protected Function GetMinStartSpacesCount(ByVal fromLine As Integer, ByVal toLine As Integer) As Integer
            If fromLine > toLine Then Return 0
            Dim result As Integer = Integer.MaxValue

            For i As Integer = fromLine To toLine
                Dim count As Integer = lines_(i).StartSpacesCount
                If count < result Then result = count
            Next

            Return result
        End Function

        Protected Function GetMaxStartSpacesCount(ByVal fromLine As Integer, ByVal toLine As Integer) As Integer
            If fromLine > toLine Then Return 0
            Dim result As Integer = 0

            For i As Integer = fromLine To toLine
                Dim count As Integer = lines_(i).StartSpacesCount
                If count > result Then result = count
            Next

            Return result
        End Function

        Public Overridable Sub Undo()
            lines_.Manager.Undo()
            DoCaretVisible()
            Invalidate()
        End Sub

        Public Overridable Sub Redo()
            lines_.Manager.Redo()
            DoCaretVisible()
            Invalidate()
        End Sub

        Protected Overrides Function IsInputKey(ByVal keyData As Keys) As Boolean
            If (keyData = Keys.Tab OrElse keyData = (Keys.Shift Or Keys.Tab)) AndAlso Not AcceptsTab Then Return False
            If keyData = Keys.Enter AndAlso Not AcceptsReturn Then Return False

            If (keyData And Keys.Alt) = Keys.None Then
                Dim keys As Keys = keyData And Keys.KeyCode
                If keys = Keys.[Return] Then Return True
            End If

            If (keyData And Keys.Alt) <> Keys.Alt Then

                Select Case (keyData And Keys.KeyCode)
                    Case Keys.Prior, Keys.[Next], Keys.[End], Keys.Home, Keys.Left, Keys.Right, Keys.Up, Keys.Down
                        Return True
                    Case Keys.Escape
                        Return False
                    Case Keys.Tab
                        Return (keyData And Keys.Control) = Keys.None
                End Select
            End If

            Return MyBase.IsInputKey(keyData)
        End Function

        <DllImport("User32.dll")>
        Private Shared Function CreateCaret(ByVal hWnd As IntPtr, ByVal hBitmap As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Boolean
    <DllImport("User32.dll")>
        Private Shared Function SetCaretPos(ByVal x As Integer, ByVal y As Integer) As Boolean
    <DllImport("User32.dll")>
        Private Shared Function DestroyCaret() As Boolean
    <DllImport("User32.dll")>
        Private Shared Function ShowCaret(ByVal hWnd As IntPtr) As Boolean
    <DllImport("User32.dll")>
        Private Shared Function HideCaret(ByVal hWnd As IntPtr) As Boolean

    Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
            If backBrush_ Is Nothing Then
                MyBase.OnPaintBackground(e)
            Else
                e.Graphics.FillRectangle(backBrush_, ClientRectangle)
            End If
        End Sub

        Public Sub DrawText(ByVal gr As Graphics, ByVal start As Place, ByVal size As Size)
            If needRecalc_ Then Recalc()
            If needRecalcFoldingLines Then RecalcFoldingLines()
            Dim startPoint = PlaceToPoint(start)
            Dim startY = startPoint.Y + VerticalScroll.Value
            Dim startX = startPoint.X + HorizontalScroll.Value - LeftIndent - Paddings.Left
            Dim firstChar As Integer = start.iChar
            Dim lastChar As Integer = (startX + size.Width) / CharWidth
            Dim startLine = start.iLine

            For iLine As Integer = startLine To lines_.Count - 1
                Dim line As Line = lines_(iLine)
                Dim lineInfo As LineInfo = LineInfos(iLine)
                If lineInfo.startY > startY + size.Height Then Exit For
                If lineInfo.startY + lineInfo.WordWrapStringsCount * charHeight_ < startY Then Continue For
                If lineInfo.VisibleState = VisibleState.Hidden Then Continue For
                Dim y As Integer = lineInfo.startY - startY
                gr.SmoothingMode = SmoothingMode.None

                If lineInfo.VisibleState = VisibleState.Visible Then
                    If line.BackgroundBrush IsNot Nothing Then gr.FillRectangle(line.BackgroundBrush, New Rectangle(0, y, size.Width, charHeight_ * lineInfo.WordWrapStringsCount))
                End If

                gr.SmoothingMode = SmoothingMode.AntiAlias

                For iWordWrapLine As Integer = 0 To lineInfo.WordWrapStringsCount - 1
                    y = lineInfo.startY + iWordWrapLine * charHeight_ - startY
                    Dim indent = If(iWordWrapLine = 0, 0, lineInfo.wordWrapIndent * CharWidth)
                    DrawLineChars(gr, firstChar, lastChar, iLine, iWordWrapLine, -startX + indent, y)
                Next
            Next
        End Sub

        Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
            If needRecalc_ Then Recalc()
            If needRecalcFoldingLines Then RecalcFoldingLines()
            visibleMarkers.Clear()
            e.Graphics.SmoothingMode = SmoothingMode.None
            Dim servicePen = New Pen(serviceLinesColor)
            Dim changedLineBrush As Brush = New SolidBrush(changedLineColor)
            Dim indentBrush As Brush = New SolidBrush(indentBackColor)
            Dim paddingBrush As Brush = New SolidBrush(paddingBackColor)
            Dim currentLineBrush As Brush = New SolidBrush(Color.FromArgb(If(currentLineColor.A = 255, 50, currentLineColor.A), currentLineColor))
            Dim textAreaRect = textAreaRect
            e.Graphics.FillRectangle(paddingBrush, 0, -VerticalScroll.Value, ClientSize.Width, Math.Max(0, Paddings.Top - 1))
            e.Graphics.FillRectangle(paddingBrush, 0, textAreaRect.Bottom, ClientSize.Width, ClientSize.Height)
            e.Graphics.FillRectangle(paddingBrush, textAreaRect.Right, 0, ClientSize.Width, ClientSize.Height)
            e.Graphics.FillRectangle(paddingBrush, LeftIndentLine, 0, LeftIndent - LeftIndentLine - 1, ClientSize.Height)
            If HorizontalScroll.Value <= Paddings.Left Then e.Graphics.FillRectangle(paddingBrush, LeftIndent - HorizontalScroll.Value - 2, 0, Math.Max(0, Paddings.Left - 1), ClientSize.Height)
            Dim leftTextIndent As Integer = Math.Max(LeftIndent, LeftIndent + Paddings.Left - HorizontalScroll.Value)
            Dim textWidth As Integer = textAreaRect.Width
            e.Graphics.FillRectangle(indentBrush, 0, 0, LeftIndentLine, ClientSize.Height)
            If LeftIndent > minLeftIndent Then e.Graphics.DrawLine(servicePen, LeftIndentLine, 0, LeftIndentLine, ClientSize.Height)
            If preferredLineWidth > 0 Then e.Graphics.DrawLine(servicePen, New Point(LeftIndent + Paddings.Left + preferredLineWidth * CharWidth - HorizontalScroll.Value + 1, textAreaRect.Top + 1), New Point(LeftIndent + Paddings.Left + preferredLineWidth * CharWidth - HorizontalScroll.Value + 1, textAreaRect.Bottom - 1))
            DrawTextAreaBorder(e.Graphics)
            Dim firstChar As Integer = (Math.Max(0, HorizontalScroll.Value - Paddings.Left)) / CharWidth
            Dim lastChar As Integer = (HorizontalScroll.Value + ClientSize.Width) / CharWidth
            Dim x = LeftIndent + Paddings.Left - HorizontalScroll.Value
            If x < LeftIndent Then firstChar += 1
            Dim bookmarksByLineIndex = New Dictionary(Of Integer, Bookmark)()

            For Each item As Bookmark In bookmarks
                bookmarksByLineIndex(item.LineIndex) = item
            Next

            Dim startLine As Integer = YtoLineIndex(VerticalScroll.Value)
            Dim iLine As Integer
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

            For iLine = startLine To lines_.Count - 1
                Dim line As Line = lines_(iLine)
                Dim lineInfo As LineInfo = LineInfos(iLine)
                If lineInfo.startY > VerticalScroll.Value + ClientSize.Height Then Exit For
                If lineInfo.startY + lineInfo.WordWrapStringsCount * charHeight_ < VerticalScroll.Value Then Continue For
                If lineInfo.VisibleState = VisibleState.Hidden Then Continue For
                Dim y As Integer = lineInfo.startY - VerticalScroll.Value
                e.Graphics.SmoothingMode = SmoothingMode.None

                If lineInfo.VisibleState = VisibleState.Visible Then
                    If line.BackgroundBrush IsNot Nothing Then e.Graphics.FillRectangle(line.BackgroundBrush, New Rectangle(textAreaRect.Left, y, textAreaRect.Width, charHeight_ * lineInfo.WordWrapStringsCount))
                End If

                If currentLineColor <> Color.Transparent AndAlso iLine = selection_.Start.iLine Then
                    If selection_.IsEmpty Then e.Graphics.FillRectangle(currentLineBrush, New Rectangle(textAreaRect.Left, y, textAreaRect.Width, charHeight_))
                End If

                If changedLineColor <> Color.Transparent AndAlso line.IsChanged Then e.Graphics.FillRectangle(changedLineBrush, New RectangleF(-10, y, LeftIndent - minLeftIndent - 2 + 10, charHeight_ + 1))
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                If bookmarksByLineIndex.ContainsKey(iLine) Then bookmarksByLineIndex(iLine).Paint(e.Graphics, New Rectangle(LeftIndent, y, Width, charHeight_ * lineInfo.WordWrapStringsCount))
                If lineInfo.VisibleState = VisibleState.Visible Then OnPaintLine(New PaintLineEventArgs(iLine, New Rectangle(LeftIndent, y, Width, charHeight_ * lineInfo.WordWrapStringsCount), e.Graphics, e.ClipRectangle))

                If showLineNumbers_ Then

                    Using lineNumberBrush = New SolidBrush(lineNumberColor)
                        e.Graphics.DrawString((iLine + lineNumberStartValue).ToString(), Font, lineNumberBrush, New RectangleF(-10, y, LeftIndent - minLeftIndent - 2 + 10, charHeight_ + CInt((lineInterval * 0.5F))), New StringFormat(StringFormatFlags.DirectionRightToLeft) With {
                        .LineAlignment = StringAlignment.Center
                    })
                    End Using
                End If

                If lineInfo.VisibleState = VisibleState.StartOfHiddenBlock Then visibleMarkers.Add(New ExpandFoldingMarker(iLine, New Rectangle(LeftIndentLine - 4, y + charHeight_ / 2 - 3, 8, 8)))
                If Not String.IsNullOrEmpty(line.FoldingStartMarker) AndAlso lineInfo.VisibleState = VisibleState.Visible AndAlso String.IsNullOrEmpty(line.FoldingEndMarker) Then visibleMarkers.Add(New CollapseFoldingMarker(iLine, New Rectangle(LeftIndentLine - 4, y + charHeight_ / 2 - 3, 8, 8)))
                If lineInfo.VisibleState = VisibleState.Visible AndAlso Not String.IsNullOrEmpty(line.FoldingEndMarker) AndAlso String.IsNullOrEmpty(line.FoldingStartMarker) Then e.Graphics.DrawLine(servicePen, LeftIndentLine, y + charHeight_ * lineInfo.WordWrapStringsCount - 1, LeftIndentLine + 4, y + charHeight_ * lineInfo.WordWrapStringsCount - 1)

                For iWordWrapLine As Integer = 0 To lineInfo.WordWrapStringsCount - 1
                    y = lineInfo.startY + iWordWrapLine * charHeight_ - VerticalScroll.Value
                    If y > VerticalScroll.Value + ClientSize.Height Then Exit For
                    If lineInfo.startY + iWordWrapLine * charHeight_ < VerticalScroll.Value Then Continue For
                    Dim indent = If(iWordWrapLine = 0, 0, lineInfo.wordWrapIndent * CharWidth)
                    DrawLineChars(e.Graphics, firstChar, lastChar, iLine, iWordWrapLine, x + indent, y)
                Next
            Next

            Dim endLine As Integer = iLine - 1
            If showFoldingLines Then DrawFoldingLines(e, startLine, endLine)

            If selection_.ColumnSelectionMode Then

                If TypeOf SelectionStyle.BackgroundBrush Is SolidBrush Then
                    Dim color As Color = (CType(SelectionStyle.BackgroundBrush, SolidBrush)).Color
                    Dim p1 As Point = PlaceToPoint(selection_.Start)
                    Dim p2 As Point = PlaceToPoint(selection_.[End])

                    Using pen = New Pen(color)
                        e.Graphics.DrawRectangle(pen, Rectangle.FromLTRB(Math.Min(p1.X, p2.X) - 1, Math.Min(p1.Y, p2.Y), Math.Max(p1.X, p2.X), Math.Max(p1.Y, p2.Y) + charHeight_))
                    End Using
                End If
            End If

            If BracketsStyle IsNot Nothing AndAlso leftBracketPosition IsNot Nothing AndAlso rightBracketPosition IsNot Nothing Then
                BracketsStyle.Draw(e.Graphics, PlaceToPoint(leftBracketPosition.Start), leftBracketPosition)
                BracketsStyle.Draw(e.Graphics, PlaceToPoint(rightBracketPosition.Start), rightBracketPosition)
            End If

            If BracketsStyle2 IsNot Nothing AndAlso leftBracketPosition2 IsNot Nothing AndAlso rightBracketPosition2 IsNot Nothing Then
                BracketsStyle2.Draw(e.Graphics, PlaceToPoint(leftBracketPosition2.Start), leftBracketPosition2)
                BracketsStyle2.Draw(e.Graphics, PlaceToPoint(rightBracketPosition2.Start), rightBracketPosition2)
            End If

            e.Graphics.SmoothingMode = SmoothingMode.None

            If (startFoldingLine >= 0 OrElse endFoldingLine >= 0) AndAlso selection_.Start = selection_.[End] Then

                If endFoldingLine < LineInfos.Count Then
                    Dim startFoldingY As Integer = (If(startFoldingLine >= 0, LineInfos(startFoldingLine).startY, 0)) - VerticalScroll.Value + charHeight_ / 2
                    Dim endFoldingY As Integer = (If(endFoldingLine >= 0, LineInfos(endFoldingLine).startY + (LineInfos(endFoldingLine).WordWrapStringsCount - 1) * charHeight_, TextHeight + charHeight_)) - VerticalScroll.Value + charHeight_

                    Using indicatorPen = New Pen(Color.FromArgb(100, foldingIndicatorColor), 4)
                        e.Graphics.DrawLine(indicatorPen, LeftIndent - 5, startFoldingY, LeftIndent - 5, endFoldingY)
                    End Using
                End If
            End If

            PaintHintBrackets(e.Graphics)
            DrawMarkers(e, servicePen)
            Dim car As Point = PlaceToPoint(selection_.Start)
            Dim caretHeight = charHeight_ - lineInterval
            car.Offset(0, lineInterval / 2)

            If (Focused OrElse IsDragDrop OrElse ShowCaretWhenInactive) AndAlso car.X >= LeftIndent AndAlso caretVisible Then
                Dim carWidth As Integer = If((isReplaceMode OrElse WideCaret), CharWidth, 1)

                If WideCaret Then

                    Using brush = New SolidBrush(CaretColor)
                        e.Graphics.FillRectangle(brush, car.X, car.Y, carWidth, caretHeight + 1)
                    End Using
                Else

                    Using pen = New Pen(CaretColor)
                        e.Graphics.DrawLine(pen, car.X, car.Y, car.X, car.Y + caretHeight)
                    End Using
                End If

                Dim caretRect = New Rectangle(HorizontalScroll.Value + car.X, VerticalScroll.Value + car.Y, carWidth, caretHeight + 1)

                If CaretBlinking Then

                    If prevCaretRect <> caretRect OrElse Not ShowScrollBars Then
                        CreateCaret(Handle, 0, carWidth, caretHeight + 1)
                        SetCaretPos(car.X, car.Y)
                        ShowCaret(Handle)
                    End If
                End If

                prevCaretRect = caretRect
            Else
                HideCaret(Handle)
                prevCaretRect = Rectangle.Empty
            End If

            If Not Enabled Then

                Using brush = New SolidBrush(DisabledColor)
                    e.Graphics.FillRectangle(brush, ClientRectangle)
                End Using
            End If

            If macrosManager.IsRecording Then DrawRecordingHint(e.Graphics)
            If middleClickScrollingActivated Then DrawMiddleClickScrolling(e.Graphics)
            servicePen.Dispose()
            changedLineBrush.Dispose()
            indentBrush.Dispose()
            currentLineBrush.Dispose()
            paddingBrush.Dispose()
            MyBase.OnPaint(e)
        End Sub

        Private Sub DrawMarkers(ByVal e As PaintEventArgs, ByVal servicePen As Pen)
            For Each m As VisualMarker In visibleMarkers

                If TypeOf m Is CollapseFoldingMarker Then

                    Using bk = New SolidBrush(ServiceColors.CollapseMarkerBackColor)

                        Using fore = New Pen(ServiceColors.CollapseMarkerForeColor)

                            Using border = New Pen(ServiceColors.CollapseMarkerBorderColor)
                            (TryCast(m, CollapseFoldingMarker)).Draw(e.Graphics, border, bk, fore)
                        End Using
                        End Using
                    End Using
                ElseIf TypeOf m Is ExpandFoldingMarker Then

                    Using bk = New SolidBrush(ServiceColors.ExpandMarkerBackColor)

                        Using fore = New Pen(ServiceColors.ExpandMarkerForeColor)

                            Using border = New Pen(ServiceColors.ExpandMarkerBorderColor)
                            (TryCast(m, ExpandFoldingMarker)).Draw(e.Graphics, border, bk, fore)
                        End Using
                        End Using
                    End Using
                Else
                    m.Draw(e.Graphics, servicePen)
                End If
            Next
        End Sub

        Private prevCaretRect As Rectangle

        Private Sub DrawRecordingHint(ByVal graphics As Graphics)
            Const w As Integer = 75
            Const h As Integer = 13
            Dim rect = New Rectangle(ClientRectangle.Right - w, ClientRectangle.Bottom - h, w, h)
            Dim iconRect = New Rectangle(-h / 2 + 3, -h / 2 + 3, h - 7, h - 7)
            Dim state = graphics.Save()
            graphics.SmoothingMode = SmoothingMode.HighQuality
            graphics.TranslateTransform(rect.Left + h / 2, rect.Top + h / 2)
            Dim ts = New TimeSpan(DateTime.Now.Ticks)
            graphics.RotateTransform(180 * (DateTime.Now.Millisecond / 1000.0F))

            Using pen = New Pen(Color.Red, 2)
                graphics.DrawArc(pen, iconRect, 0, 90)
                graphics.DrawArc(pen, iconRect, 180, 90)
            End Using

            graphics.DrawEllipse(Pens.Red, iconRect)
            graphics.Restore(state)

            Using font = New Font(FontFamily.GenericSansSerif, 8.0F)
                graphics.DrawString("Recording...", font, Brushes.Red, New PointF(rect.Left + h, rect.Top))
            End Using

            Dim tm As System.Threading.Timer = Nothing
            tm = New System.Threading.Timer(Sub(o)
                                                Invalidate(rect)
                                                tm.Dispose()
                                            End Sub, Nothing, 200, System.Threading.Timeout.Infinite)
        End Sub

        Private Sub DrawTextAreaBorder(ByVal graphics As Graphics)
            If textAreaBorder = TextAreaBorderType.None Then Return
            Dim rect = TextAreaRect

            If textAreaBorder = TextAreaBorderType.Shadow Then
                Const shadowSize As Integer = 4
                Dim rBottom = New Rectangle(rect.Left + shadowSize, rect.Bottom, rect.Width - shadowSize, shadowSize)
                Dim rCorner = New Rectangle(rect.Right, rect.Bottom, shadowSize, shadowSize)
                Dim rRight = New Rectangle(rect.Right, rect.Top + shadowSize, shadowSize, rect.Height - shadowSize)

                Using brush = New SolidBrush(Color.FromArgb(80, textAreaBorderColor))
                    graphics.FillRectangle(brush, rBottom)
                    graphics.FillRectangle(brush, rRight)
                    graphics.FillRectangle(brush, rCorner)
                End Using
            End If

            Using pen As Pen = New Pen(textAreaBorderColor)
                graphics.DrawRectangle(pen, rect)
            End Using
        End Sub

        Private Sub PaintHintBrackets(ByVal gr As Graphics)
            For Each hint As Hint In hints
                Dim r As Range = hint.Range.Clone()
                r.Normalize()
                Dim p1 As Point = PlaceToPoint(r.Start)
                Dim p2 As Point = PlaceToPoint(r.[End])
                If GetVisibleState(r.Start.iLine) <> VisibleState.Visible OrElse GetVisibleState(r.[End].iLine) <> VisibleState.Visible Then Continue For

                Using pen = New Pen(hint.BorderColor)
                    pen.DashStyle = DashStyle.Dash

                    If r.IsEmpty Then
                        p1.Offset(1, -1)
                        gr.DrawLines(pen, {p1, New Point(p1.X, p1.Y + charHeight_ + 2)})
                    Else
                        p1.Offset(-1, -1)
                        p2.Offset(1, -1)
                        gr.DrawLines(pen, {New Point(p1.X + CharWidth / 2, p1.Y), p1, New Point(p1.X, p1.Y + charHeight_ + 2), New Point(p1.X + CharWidth / 2, p1.Y + charHeight_ + 2)})
                        gr.DrawLines(pen, {New Point(p2.X - CharWidth / 2, p2.Y), p2, New Point(p2.X, p2.Y + charHeight_ + 2), New Point(p2.X - CharWidth / 2, p2.Y + charHeight_ + 2)})
                    End If
                End Using
            Next
        End Sub

        Protected Overridable Sub DrawFoldingLines(ByVal e As PaintEventArgs, ByVal startLine As Integer, ByVal endLine As Integer)
            e.Graphics.SmoothingMode = SmoothingMode.None

            Using pen = New Pen(Color.FromArgb(200, serviceLinesColor)) With {
            .DashStyle = DashStyle.Dot
        }

                For Each iLine In foldingPairs

                    If iLine.Key < endLine AndAlso iLine.Value > startLine Then
                        Dim line As Line = lines_(iLine.Key)
                        Dim y As Integer = LineInfos(iLine.Key).startY - VerticalScroll.Value + charHeight_
                        y += y Mod 2
                        Dim y2 As Integer

                        If iLine.Value >= LinesCount Then
                            y2 = LineInfos(LinesCount - 1).startY + charHeight_ - VerticalScroll.Value
                        ElseIf LineInfos(iLine.Value).VisibleState = VisibleState.Visible Then
                            Dim d As Integer = 0
                            Dim spaceCount As Integer = line.StartSpacesCount
                            If lines_(iLine.Value).Count <= spaceCount OrElse lines_(iLine.Value)(spaceCount).c = " "c Then d = charHeight_
                            y2 = LineInfos(iLine.Value).startY - VerticalScroll.Value + d
                        Else
                            Continue For
                        End If

                        Dim x As Integer = LeftIndent + Paddings.Left + line.StartSpacesCount * CharWidth - HorizontalScroll.Value
                        If x >= LeftIndent + Paddings.Left Then e.Graphics.DrawLine(pen, x, If(y >= 0, y, 0), x, If(y2 < ClientSize.Height, y2, ClientSize.Height))
                    End If
                Next
            End Using
        End Sub

        Private Sub DrawLineChars(ByVal gr As Graphics, ByVal firstChar As Integer, ByVal lastChar As Integer, ByVal iLine As Integer, ByVal iWordWrapLine As Integer, ByVal startX As Integer, ByVal y As Integer)
            Dim line As Line = lines_(iLine)
            Dim lineInfo As LineInfo = LineInfos(iLine)
            Dim from As Integer = lineInfo.GetWordWrapStringStartPosition(iWordWrapLine)
            Dim [to] As Integer = lineInfo.GetWordWrapStringFinishPosition(iWordWrapLine, line)
            lastChar = Math.Min([to] - from, lastChar)
            gr.SmoothingMode = SmoothingMode.AntiAlias

            If lineInfo.VisibleState = VisibleState.StartOfHiddenBlock Then
                FoldedBlockStyle.Draw(gr, New Point(startX + firstChar * CharWidth, y), New Range(Me, from + firstChar, iLine, from + lastChar + 1, iLine))
            Else
                Dim currentStyleIndex As StyleIndex = StyleIndex.None
                Dim iLastFlushedChar As Integer = firstChar - 1

                For iChar As Integer = firstChar To lastChar
                    Dim style As StyleIndex = line(from + iChar).style

                    If currentStyleIndex <> style Then
                        FlushRendering(gr, currentStyleIndex, New Point(startX + (iLastFlushedChar + 1) * CharWidth, y), New Range(Me, from + iLastFlushedChar + 1, iLine, from + iChar, iLine))
                        iLastFlushedChar = iChar - 1
                        currentStyleIndex = style
                    End If
                Next

                FlushRendering(gr, currentStyleIndex, New Point(startX + (iLastFlushedChar + 1) * CharWidth, y), New Range(Me, from + iLastFlushedChar + 1, iLine, from + lastChar + 1, iLine))
            End If

            If selectionHighlightingForLineBreaksEnabled AndAlso iWordWrapLine = lineInfo.WordWrapStringsCount - 1 Then lastChar += 1

            If Not selection_.IsEmpty AndAlso lastChar >= firstChar Then
                gr.SmoothingMode = SmoothingMode.None
                Dim textRange = New Range(Me, from + firstChar, iLine, from + lastChar + 1, iLine)
                textRange = selection_.GetIntersectionWith(textRange)

                If textRange IsNot Nothing AndAlso SelectionStyle IsNot Nothing Then
                    SelectionStyle.Draw(gr, New Point(startX + (textRange.Start.iChar - from) * CharWidth, 1 + y), textRange)
                End If
            End If
        End Sub

        Private Sub FlushRendering(ByVal gr As Graphics, ByVal styleIndex As StyleIndex, ByVal pos As Point, ByVal range As Range)
            If range.[End] > range.Start Then
                Dim mask As Integer = 1
                Dim hasTextStyle As Boolean = False

                For i As Integer = 0 To Styles.Length - 1

                    If Styles(i) IsNot Nothing AndAlso (CInt(styleIndex) And mask) <> 0 Then
                        Dim style As Style = Styles(i)
                        Dim isTextStyle As Boolean = TypeOf style Is TextStyle
                        If Not hasTextStyle OrElse Not isTextStyle OrElse AllowSeveralTextStyleDrawing Then style.Draw(gr, pos, range)
                        hasTextStyle = hasTextStyle Or isTextStyle
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
                    '''                     mask = mask << 1;

                    ''' 
                Next

                If Not hasTextStyle Then DefaultStyle.Draw(gr, pos, range)
            End If
        End Sub

        Protected Overrides Sub OnEnter(ByVal e As EventArgs)
            MyBase.OnEnter(e)
            mouseIsDrag = False
            mouseIsDragDrop = False
            draggedRange = Nothing
        End Sub

        Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)
            MyBase.OnMouseUp(e)
            isLineSelect = False

            If e.Button = System.Windows.Forms.MouseButtons.Left Then
                If mouseIsDragDrop Then OnMouseClickText(e)
            End If
        End Sub

        Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
            MyBase.OnMouseDown(e)

            If middleClickScrollingActivated Then
                DeactivateMiddleClickScrollingMode()
                mouseIsDrag = False
                If e.Button = System.Windows.Forms.MouseButtons.Middle Then RestoreScrollsAfterMiddleClickScrollingMode()
                Return
            End If

            macrosManager.IsRecording = False
            [Select]()
            ActiveControl = Nothing

            If e.Button = MouseButtons.Left Then
                Dim marker As VisualMarker = FindVisualMarkerForPoint(e.Location)

                If marker IsNot Nothing Then
                    mouseIsDrag = False
                    mouseIsDragDrop = False
                    draggedRange = Nothing
                    OnMarkerClick(e, marker)
                    Return
                End If

                mouseIsDrag = True
                mouseIsDragDrop = False
                draggedRange = Nothing
                isLineSelect = (e.Location.X < LeftIndentLine)

                If Not isLineSelect Then
                    Dim p = PointToPlace(e.Location)

                    If e.Clicks = 2 Then
                        mouseIsDrag = False
                        mouseIsDragDrop = False
                        draggedRange = Nothing
                        SelectWord(p)
                        Return
                    End If

                    If selection_.IsEmpty OrElse Not selection_.Contains(p) OrElse Me(p.iLine).Count <= p.iChar OrElse [ReadOnly] Then
                        OnMouseClickText(e)
                    Else
                        mouseIsDragDrop = True
                        mouseIsDrag = False
                    End If
                Else
                    CheckAndChangeSelectionType()
                    selection_.BeginUpdate()
                    Dim iLine As Integer = PointToPlaceSimple(e.Location).iLine
                    lineSelectFrom = iLine
                    selection_.Start = New Place(0, iLine)
                    selection_.[End] = New Place(GetLineLength(iLine), iLine)
                    selection_.EndUpdate()
                    Invalidate()
                End If
            ElseIf e.Button = MouseButtons.Middle Then
                ActivateMiddleClickScrollingMode(e)
            End If
        End Sub

        Private Sub OnMouseClickText(ByVal e As MouseEventArgs)
            Dim oldEnd As Place = selection_.[End]
            selection_.BeginUpdate()

            If selection_.ColumnSelectionMode Then
                selection_.Start = PointToPlaceSimple(e.Location)
                selection_.ColumnSelectionMode = True
            Else

                If VirtualSpace Then
                    selection_.Start = PointToPlaceSimple(e.Location)
                Else
                    selection_.Start = PointToPlace(e.Location)
                End If
            End If

            If (lastModifiers And Keys.Shift) <> 0 Then selection_.[End] = oldEnd
            CheckAndChangeSelectionType()
            selection_.EndUpdate()
            Invalidate()
            Return
        End Sub

        Protected Overridable Sub CheckAndChangeSelectionType()
            If (ModifierKeys And Keys.Alt) <> 0 AndAlso Not wordWrap_ Then
                selection_.ColumnSelectionMode = True
            Else
                selection_.ColumnSelectionMode = False
            End If
        End Sub

        Protected Overrides Sub OnMouseWheel(ByVal e As MouseEventArgs)
            Invalidate()

            If lastModifiers = Keys.Control Then
                ChangeFontSize(2 * Math.Sign(e.Delta))
            (CType(e, HandledMouseEventArgs)).Handled = True
        ElseIf VerticalScroll.Visible OrElse Not ShowScrollBars Then
                Dim mouseWheelScrollLinesSetting As Integer = GetControlPanelWheelScrollLinesValue()
                DoScrollVertical(mouseWheelScrollLinesSetting, e.Delta)
            (CType(e, HandledMouseEventArgs)).Handled = True
        End If

            DeactivateMiddleClickScrollingMode()
        End Sub

        Private Sub DoScrollVertical(ByVal countLines As Integer, ByVal direction As Integer)
            If VerticalScroll.Visible OrElse Not ShowScrollBars Then
                Dim numberOfVisibleLines As Integer = ClientSize.Height / charHeight_
                Dim offset As Integer

                If (countLines = -1) OrElse (countLines > numberOfVisibleLines) Then
                    offset = charHeight_ * numberOfVisibleLines
                Else
                    offset = charHeight_ * countLines
                End If

                Dim newScrollPos = VerticalScroll.Value - Math.Sign(direction) * offset
                Dim ea = New ScrollEventArgs(If(direction > 0, ScrollEventType.SmallDecrement, ScrollEventType.SmallIncrement), VerticalScroll.Value, newScrollPos, ScrollOrientation.VerticalScroll)
                OnScroll(ea)
            End If
        End Sub

        Private Shared Function GetControlPanelWheelScrollLinesValue() As Integer
            Try

                Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop", False)
                    Return Convert.ToInt32(key.GetValue("WheelScrollLines"))
                End Using

            Catch
                Return 1
            End Try
        End Function

        Public Sub ChangeFontSize(ByVal [step] As Integer)
            Dim points = Font.SizeInPoints

            Using gr = Graphics.FromHwnd(Handle)
                Dim dpi = gr.DpiY
                Dim newPoints = points + [step] * 72.0F / dpi
                If newPoints < 1.0F Then Return
                Dim k = newPoints / originalFont.SizeInPoints
                zoom = CInt(Math.Round(100 * k))
            End Using
        End Sub

        <Browsable(False)>
        Public Property Zoom As Integer
            Get
                Return Zoom
            End Get
            Set(ByVal value As Integer)
                zoom = value
                DoZoom(zoom / 100.0F)
                OnZoomChanged()
            End Set
        End Property

        Protected Overridable Sub OnZoomChanged()
            RaiseEvent ZoomChanged(Me, EventArgs.Empty)
        End Sub

        Private Sub DoZoom(ByVal koeff As Single)
            Dim iLine = YtoLineIndex(VerticalScroll.Value)
            Dim points = originalFont.SizeInPoints
            ''' Cannot convert ExpressionStatementSyntax, System.ArgumentOutOfRangeException: Exception of type 'System.ArgumentOutOfRangeException' was thrown.
            ''' Parameter name: op
            ''' Actual value was MultiplyAssignmentStatement.
            '''    at ICSharpCode.CodeConverter.Util.VBUtil.GetExpressionOperatorTokenKind(SyntaxKind op)
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
            '''             points *= koeff;

            ''' 
            If points < 1.0F OrElse points > 300.0F Then Return
            Dim oldFont = Font
            SetFont(New Font(Font.FontFamily, points, Font.Style, GraphicsUnit.Point))
            oldFont.Dispose()
            NeedRecalc(True)
            If iLine < LinesCount Then VerticalScroll.Value = Math.Min(VerticalScroll.Maximum, LineInfos(iLine).startY - Paddings.Top)
            UpdateScrollbars()
            Invalidate()
            OnVisibleRangeChanged()
        End Sub

        Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
            MyBase.OnMouseLeave(e)
            CancelToolTip()
        End Sub

        Protected draggedRange As Range

        Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
            MyBase.OnMouseMove(e)
            If middleClickScrollingActivated Then Return

            If lastMouseCoord <> e.Location Then
                CancelToolTip()
                timer3.Start()
            End If

            lastMouseCoord = e.Location

            If e.Button = MouseButtons.Left AndAlso mouseIsDragDrop Then
                draggedRange = selection_.Clone()
                DoDragDrop(SelectedText, DragDropEffects.Copy)
                draggedRange = Nothing
                Return
            End If

            If e.Button = MouseButtons.Left AndAlso mouseIsDrag Then
                Dim place As Place

                If selection_.ColumnSelectionMode OrElse VirtualSpace Then
                    place = PointToPlaceSimple(e.Location)
                Else
                    place = PointToPlace(e.Location)
                End If

                If isLineSelect Then
                    selection_.BeginUpdate()
                    Dim iLine As Integer = place.iLine

                    If iLine < lineSelectFrom Then
                        selection_.Start = New Place(0, iLine)
                        selection_.[End] = New Place(GetLineLength(lineSelectFrom), lineSelectFrom)
                    Else
                        selection_.Start = New Place(GetLineLength(iLine), iLine)
                        selection_.[End] = New Place(0, lineSelectFrom)
                    End If

                    selection_.EndUpdate()
                    DoCaretVisible()
                    HorizontalScroll.Value = 0
                    UpdateScrollbars()
                    Invalidate()
                ElseIf place <> selection_.Start Then
                    Dim oldEnd As Place = selection_.[End]
                    selection_.BeginUpdate()

                    If selection_.ColumnSelectionMode Then
                        selection_.Start = place
                        selection_.ColumnSelectionMode = True
                    Else
                        selection_.Start = place
                    End If

                    selection_.[End] = oldEnd
                    selection_.EndUpdate()
                    DoCaretVisible()
                    Invalidate()
                    Return
                End If
            End If

            Dim marker As VisualMarker = FindVisualMarkerForPoint(e.Location)

            If marker IsNot Nothing Then
                MyBase.Cursor = marker.Cursor
            Else

                If e.Location.X < LeftIndentLine OrElse isLineSelect Then
                    MyBase.Cursor = Cursors.Arrow
                Else
                    MyBase.Cursor = defaultCursor
                End If
            End If
        End Sub

        Private Sub CancelToolTip()
            timer3.[Stop]()

            If ToolTip IsNot Nothing AndAlso Not String.IsNullOrEmpty(ToolTip.GetToolTip(Me)) Then
                ToolTip.Hide(Me)
                ToolTip.SetToolTip(Me, Nothing)
            End If
        End Sub

        Protected Overrides Sub OnMouseDoubleClick(ByVal e As MouseEventArgs)
            MyBase.OnMouseDoubleClick(e)
            Dim m = FindVisualMarkerForPoint(e.Location)
            If m IsNot Nothing Then OnMarkerDoubleClick(m)
        End Sub

        Private Sub SelectWord(ByVal p As Place)
            Dim fromX As Integer = p.iChar
            Dim toX As Integer = p.iChar

            For i As Integer = p.iChar To lines_(p.iLine).Count - 1
                Dim c As Char = lines_(p.iLine)(i).c

                If Char.IsLetterOrDigit(c) OrElse c = "_"c Then
                    toX = i + 1
                Else
                    Exit For
                End If
            Next

            For i As Integer = p.iChar - 1 To 0
                Dim c As Char = lines_(p.iLine)(i).c

                If Char.IsLetterOrDigit(c) OrElse c = "_"c Then
                    fromX = i
                Else
                    Exit For
                End If
            Next

            selection_ = New Range(Me, toX, p.iLine, fromX, p.iLine)
        End Sub

        Public Function YtoLineIndex(ByVal y As Integer) As Integer
            Dim i As Integer = LineInfos.BinarySearch(New LineInfo(-10), New LineYComparer(y))
            i = If(i < 0, -i - 2, i)
            If i < 0 Then Return 0
            If i > lines_.Count - 1 Then Return lines_.Count - 1
            Return i
        End Function

        Public Function PointToPlace(ByVal point As Point) As Place
            point.Offset(HorizontalScroll.Value, VerticalScroll.Value)
            point.Offset(-LeftIndent - Paddings.Left, 0)
            Dim iLine As Integer = YtoLineIndex(point.Y)
            If iLine < 0 Then Return Place.Empty
            Dim y As Integer = 0

            While iLine < lines_.Count
                y = LineInfos(iLine).startY + LineInfos(iLine).WordWrapStringsCount * charHeight_
                If y > point.Y AndAlso LineInfos(iLine).VisibleState = VisibleState.Visible Then Exit For
                iLine += 1
            End While

            If iLine >= lines_.Count Then iLine = lines_.Count - 1
            If LineInfos(iLine).VisibleState <> VisibleState.Visible Then iLine = FindPrevVisibleLine(iLine)
            Dim iWordWrapLine As Integer = LineInfos(iLine).WordWrapStringsCount

            If y > point.Y Then
                Dim approximatelyLines As Integer = (y - point.Y - charHeight_) / charHeight_
                y -= approximatelyLines * charHeight_
                iWordWrapLine -= approximatelyLines
            End If

            Do
                iWordWrapLine -= 1
                y -= charHeight_
            Loop While y > point.Y

            If iWordWrapLine < 0 Then iWordWrapLine = 0
            Dim start As Integer = LineInfos(iLine).GetWordWrapStringStartPosition(iWordWrapLine)
            Dim finish As Integer = LineInfos(iLine).GetWordWrapStringFinishPosition(iWordWrapLine, lines_(iLine))
            Dim x = CInt(Math.Round(CSng(point.X) / CharWidth))
            If iWordWrapLine > 0 Then x -= LineInfos(iLine).wordWrapIndent
            x = If(x < 0, start, start + x)
            If x > finish Then x = finish + 1
            If x > lines_(iLine).Count Then x = lines_(iLine).Count
            Return New Place(x, iLine)
        End Function

        Private Function PointToPlaceSimple(ByVal point As Point) As Place
            point.Offset(HorizontalScroll.Value, VerticalScroll.Value)
            point.Offset(-LeftIndent - Paddings.Left, 0)
            Dim iLine As Integer = YtoLineIndex(point.Y)
            Dim x = CInt(Math.Round(CSng(point.X) / CharWidth))
            If x < 0 Then x = 0
            Return New Place(x, iLine)
        End Function

        Public Function PointToPosition(ByVal point As Point) As Integer
            Return PlaceToPosition(PointToPlace(point))
        End Function

        Public Overridable Sub OnTextChanging(ByRef text As String)
            ClearBracketsPositions()

            If TextChanging IsNot Nothing Then
                Dim args = New TextChangingEventArgs With {
                .InsertingText = text
            }
                TextChanging(Me, args)
                text = args.InsertingText
                If args.Cancel Then text = String.Empty
            End If
        End Sub

        Public Overridable Sub OnTextChanging()
            Dim temp As String = Nothing
            OnTextChanging(temp)
        End Sub

        Public Overridable Sub OnTextChanged()
            Dim r = New Range(Me)
            r.SelectAll()
            OnTextChanged(New TextChangedEventArgs(r))
        End Sub

        Public Overridable Sub OnTextChanged(ByVal fromLine As Integer, ByVal toLine As Integer)
            Dim r = New Range(Me)
            r.Start = New Place(0, Math.Min(fromLine, toLine))
            r.[End] = New Place(lines_(Math.Max(fromLine, toLine)).Count, Math.Max(fromLine, toLine))
            OnTextChanged(New TextChangedEventArgs(r))
        End Sub

        Public Overridable Sub OnTextChanged(ByVal r As Range)
            OnTextChanged(New TextChangedEventArgs(r))
        End Sub

        Public Sub BeginUpdate()
            If updating = 0 Then updatingRange = Nothing
            updating += 1
        End Sub

        Public Sub EndUpdate()
            updating -= 1

            If updating = 0 AndAlso updatingRange IsNot Nothing Then
                updatingRange.Expand()
                OnTextChanged(updatingRange)
            End If
        End Sub

        Protected Overridable Sub OnTextChanged(ByVal args As TextChangedEventArgs)
            args.ChangedRange.Normalize()

            If updating > 0 Then

                If updatingRange Is Nothing Then
                    updatingRange = args.ChangedRange.Clone()
                Else
                    If updatingRange.Start.iLine > args.ChangedRange.Start.iLine Then updatingRange.Start = New Place(0, args.ChangedRange.Start.iLine)
                    If updatingRange.[End].iLine < args.ChangedRange.[End].iLine Then updatingRange.[End] = New Place(lines_(args.ChangedRange.[End].iLine).Count, args.ChangedRange.[End].iLine)
                    updatingRange = updatingRange.GetIntersectionWith(Range)
                End If

                Return
            End If

            CancelToolTip()
            ClearHints()
            isChanged_ = True
            TextVersion += 1
            MarkLinesAsChanged(args.ChangedRange)
            ClearFoldingState(args.ChangedRange)
            If wordWrap_ Then RecalcWordWrap(args.ChangedRange.Start.iLine, args.ChangedRange.[End].iLine)
            MyBase.OnTextChanged(args)

            If delayedTextChangedRange Is Nothing Then
                delayedTextChangedRange = args.ChangedRange.Clone()
            Else
                delayedTextChangedRange = delayedTextChangedRange.GetUnionWith(args.ChangedRange)
            End If

            needRiseTextChangedDelayed = True
            ResetTimer(timer2)
            OnSyntaxHighlight(args)
            RaiseEvent TextChanged(Me, args)
            RaiseEvent BindingTextChanged(Me, EventArgs.Empty)
            MyBase.OnTextChanged(EventArgs.Empty)
            OnVisibleRangeChanged()
        End Sub

        Private Sub ClearFoldingState(ByVal range As Range)
            For iLine As Integer = range.Start.iLine To range.[End].iLine
                If iLine >= 0 AndAlso iLine < lines_.Count Then FoldedBlocks.Remove(Me(iLine).UniqueId)
            Next
        End Sub

        Private Sub MarkLinesAsChanged(ByVal range As Range)
            For iLine As Integer = range.Start.iLine To range.[End].iLine
                If iLine >= 0 AndAlso iLine < lines_.Count Then lines_(iLine).IsChanged = True
            Next
        End Sub

        Public Overridable Sub OnSelectionChanged()
            If highlightFoldingIndicator Then HighlightFoldings()
            needRiseSelectionChangedDelayed = True
            ResetTimer(timer)
            RaiseEvent SelectionChanged(Me, New EventArgs())
        End Sub

        Private Sub HighlightFoldings()
            If LinesCount = 0 Then Return
            Dim prevStartFoldingLine As Integer = startFoldingLine
            Dim prevEndFoldingLine As Integer = endFoldingLine
            startFoldingLine = -1
            endFoldingLine = -1
            Dim counter As Integer = 0

            For i As Integer = selection_.Start.iLine To Math.Max(selection_.Start.iLine - maxLinesForFolding, 0)
                Dim hasStartMarker As Boolean = lines_.LineHasFoldingStartMarker(i)
                Dim hasEndMarker As Boolean = lines_.LineHasFoldingEndMarker(i)
                If hasEndMarker AndAlso hasStartMarker Then Continue For

                If hasStartMarker Then
                    counter -= 1

                    If counter = -1 Then
                        startFoldingLine = i
                        Exit For
                    End If
                End If

                If hasEndMarker AndAlso i <> selection_.Start.iLine Then counter += 1
            Next

            If startFoldingLine >= 0 Then
                endFoldingLine = FindEndOfFoldingBlock(startFoldingLine, maxLinesForFolding)
                If endFoldingLine = startFoldingLine Then endFoldingLine = -1
            End If

            If startFoldingLine <> prevStartFoldingLine OrElse endFoldingLine <> prevEndFoldingLine Then OnFoldingHighlightChanged()
        End Sub

        Protected Overridable Sub OnFoldingHighlightChanged()
            RaiseEvent FoldingHighlightChanged(Me, EventArgs.Empty)
        End Sub

        Protected Overrides Sub OnGotFocus(ByVal e As EventArgs)
            SetAsCurrentTB()
            MyBase.OnGotFocus(e)
            Invalidate()
        End Sub

        Protected Overrides Sub OnLostFocus(ByVal e As EventArgs)
            lastModifiers = Keys.None
            DeactivateMiddleClickScrollingMode()
            MyBase.OnLostFocus(e)
            Invalidate()
        End Sub

        Public Function PlaceToPosition(ByVal point As Place) As Integer
            If point.iLine < 0 OrElse point.iLine >= lines_.Count OrElse point.iChar >= lines_(point.iLine).Count + Environment.NewLine.Length Then Return -1
            Dim result As Integer = 0

            For i As Integer = 0 To point.iLine - 1
                result += lines_(i).Count + Environment.NewLine.Length
            Next

            result += point.iChar
            Return result
        End Function

        Public Function PositionToPlace(ByVal pos As Integer) As Place
            If pos < 0 Then Return New Place(0, 0)

            For i As Integer = 0 To lines_.Count - 1
                Dim lineLength As Integer = lines_(i).Count + Environment.NewLine.Length
                If pos < lines_(i).Count Then Return New Place(pos, i)
                If pos < lineLength Then Return New Place(lines_(i).Count, i)
                pos -= lineLength
            Next

            If lines_.Count > 0 Then
                Return New Place(lines_(lines_.Count - 1).Count, lines_.Count - 1)
            Else
                Return New Place(0, 0)
            End If
        End Function

        Public Function PositionToPoint(ByVal pos As Integer) As Point
            Return PlaceToPoint(PositionToPlace(pos))
        End Function

        Public Function PlaceToPoint(ByVal place As Place) As Point
            If place.iLine >= LineInfos.Count Then Return New Point()
            Dim y As Integer = LineInfos(place.iLine).startY
            Dim iWordWrapIndex As Integer = LineInfos(place.iLine).GetWordWrapStringIndex(place.iChar)
            y += iWordWrapIndex * charHeight_
            Dim x As Integer = (place.iChar - LineInfos(place.iLine).GetWordWrapStringStartPosition(iWordWrapIndex)) * CharWidth
            If iWordWrapIndex > 0 Then x += LineInfos(place.iLine).wordWrapIndent * CharWidth
            y = y - VerticalScroll.Value
            x = LeftIndent + Paddings.Left + x - HorizontalScroll.Value
            Return New Point(x, y)
        End Function

        Public Function GetRange(ByVal fromPos As Integer, ByVal toPos As Integer) As Range
            Dim sel = New Range(Me)
            sel.Start = PositionToPlace(fromPos)
            sel.[End] = PositionToPlace(toPos)
            Return sel
        End Function

        Public Function GetRange(ByVal fromPlace As Place, ByVal toPlace As Place) As Range
            Return New Range(Me, fromPlace, toPlace)
        End Function

        Public Iterator Function GetRanges(ByVal regexPattern As String) As IEnumerable(Of Range)
            Dim range = New Range(Me)
            range.SelectAll()

            For Each r As Range In range.GetRanges(regexPattern, RegexOptions.None)
                Yield r
            Next
        End Function

        Public Iterator Function GetRanges(ByVal regexPattern As String, ByVal options As RegexOptions) As IEnumerable(Of Range)
            Dim range = New Range(Me)
            range.SelectAll()

            For Each r As Range In range.GetRanges(regexPattern, options)
                Yield r
            Next
        End Function

        Public Function GetLineText(ByVal iLine As Integer) As String
            If iLine < 0 OrElse iLine >= lines_.Count Then Throw New ArgumentOutOfRangeException("Line index out of range")
            Dim sb = New StringBuilder(lines_(iLine).Count)

            For Each c As Char In lines_(iLine)
                sb.Append(c.c)
            Next

            Return sb.ToString()
        End Function

        Public Overridable Sub ExpandFoldedBlock(ByVal iLine As Integer)
            If iLine < 0 OrElse iLine >= lines_.Count Then Throw New ArgumentOutOfRangeException("Line index out of range")
            Dim [end] As Integer = iLine

            While [end] < LinesCount - 1
                If LineInfos([end] + 1).VisibleState <> VisibleState.Hidden Then Exit For
                [end] += 1
            End While

            ExpandBlock(iLine, [end])
            FoldedBlocks.Remove(Me(iLine).UniqueId)
            AdjustFolding()
        End Sub

        Public Overridable Sub AdjustFolding()
            For iLine As Integer = 0 To LinesCount - 1

                If LineInfos(iLine).VisibleState = VisibleState.Visible Then
                    If FoldedBlocks.ContainsKey(Me(iLine).UniqueId) Then CollapseFoldingBlock(iLine)
                End If
            Next
        End Sub

        Public Overridable Sub ExpandBlock(ByVal fromLine As Integer, ByVal toLine As Integer)
            Dim from As Integer = Math.Min(fromLine, toLine)
            Dim [to] As Integer = Math.Max(fromLine, toLine)

            For i As Integer = from To [to]
                SetVisibleState(i, VisibleState.Visible)
            Next

            needRecalc_ = True
            Invalidate()
            OnVisibleRangeChanged()
        End Sub

        Public Sub ExpandBlock(ByVal iLine As Integer)
            If LineInfos(iLine).VisibleState = VisibleState.Visible Then Return

            For i As Integer = iLine To LinesCount - 1

                If LineInfos(i).VisibleState = VisibleState.Visible Then
                    Exit For
                Else
                    SetVisibleState(i, VisibleState.Visible)
                    needRecalc_ = True
                End If
            Next

            For i As Integer = iLine - 1 To 0

                If LineInfos(i).VisibleState = VisibleState.Visible Then
                    Exit For
                Else
                    SetVisibleState(i, VisibleState.Visible)
                    needRecalc_ = True
                End If
            Next

            Invalidate()
            OnVisibleRangeChanged()
        End Sub

        Public Overridable Sub CollapseAllFoldingBlocks()
            For i As Integer = 0 To LinesCount - 1

                If lines_.LineHasFoldingStartMarker(i) Then
                    Dim iFinish As Integer = FindEndOfFoldingBlock(i)

                    If iFinish >= 0 Then
                        CollapseBlock(i, iFinish)
                        i = iFinish
                    End If
                End If
            Next

            OnVisibleRangeChanged()
            UpdateScrollbars()
        End Sub

        Public Overridable Sub ExpandAllFoldingBlocks()
            For i As Integer = 0 To LinesCount - 1
                SetVisibleState(i, VisibleState.Visible)
            Next

            FoldedBlocks.Clear()
            OnVisibleRangeChanged()
            Invalidate()
            UpdateScrollbars()
        End Sub

        Public Overridable Sub CollapseFoldingBlock(ByVal iLine As Integer)
            If iLine < 0 OrElse iLine >= lines_.Count Then Throw New ArgumentOutOfRangeException("Line index out of range")
            If String.IsNullOrEmpty(lines_(iLine).FoldingStartMarker) Then Throw New ArgumentOutOfRangeException("This line is not folding start line")
            Dim i As Integer = FindEndOfFoldingBlock(iLine)

            If i >= 0 Then
                CollapseBlock(iLine, i)
                Dim id = Me(iLine).UniqueId
                FoldedBlocks(id) = id
            End If
        End Sub

        Private Function FindEndOfFoldingBlock(ByVal iStartLine As Integer) As Integer
            Return FindEndOfFoldingBlock(iStartLine, Integer.MaxValue)
        End Function

        Protected Overridable Function FindEndOfFoldingBlock(ByVal iStartLine As Integer, ByVal maxLines As Integer) As Integer
            Dim i As Integer
            Dim marker As String = lines_(iStartLine).FoldingStartMarker
            Dim stack = New Stack(Of String)()

            Select Case FindEndOfFoldingBlockStrategy
                Case FindEndOfFoldingBlockStrategy.Strategy1

                    For i = iStartLine To LinesCount - 1
                        If lines_.LineHasFoldingStartMarker(i) Then stack.Push(lines_(i).FoldingStartMarker)

                        If lines_.LineHasFoldingEndMarker(i) Then
                            Dim m As String = lines_(i).FoldingEndMarker

                            While stack.Count > 0 AndAlso stack.Pop() <> m
                            End While

                            If stack.Count = 0 Then Return i
                        End If

                        maxLines -= 1
                        If maxLines < 0 Then Return i
                    Next

                Case FindEndOfFoldingBlockStrategy.Strategy2

                    For i = iStartLine To LinesCount - 1

                        If lines_.LineHasFoldingEndMarker(i) Then
                            Dim m As String = lines_(i).FoldingEndMarker

                            While stack.Count > 0 AndAlso stack.Pop() <> m
                            End While

                            If stack.Count = 0 Then Return i
                        End If

                        If lines_.LineHasFoldingStartMarker(i) Then stack.Push(lines_(i).FoldingStartMarker)
                        maxLines -= 1
                        If maxLines < 0 Then Return i
                    Next
            End Select

            Return LinesCount - 1
        End Function

        Public Function GetLineFoldingStartMarker(ByVal iLine As Integer) As String
            If lines_.LineHasFoldingStartMarker(iLine) Then Return lines_(iLine).FoldingStartMarker
            Return Nothing
        End Function

        Public Function GetLineFoldingEndMarker(ByVal iLine As Integer) As String
            If lines_.LineHasFoldingEndMarker(iLine) Then Return lines_(iLine).FoldingEndMarker
            Return Nothing
        End Function

        Protected Overridable Sub RecalcFoldingLines()
            If Not needRecalcFoldingLines Then Return
            needRecalcFoldingLines = False
            If Not showFoldingLines Then Return
            foldingPairs.Clear()
            Dim range As Range = visibleRange
            Dim startLine As Integer = Math.Max(range.Start.iLine - maxLinesForFolding, 0)
            Dim endLine As Integer = Math.Min(range.[End].iLine + maxLinesForFolding, Math.Max(range.[End].iLine, LinesCount - 1))
            Dim stack = New Stack(Of Integer)()

            For i As Integer = startLine To endLine
                Dim hasStartMarker As Boolean = lines_.LineHasFoldingStartMarker(i)
                Dim hasEndMarker As Boolean = lines_.LineHasFoldingEndMarker(i)
                If hasEndMarker AndAlso hasStartMarker Then Continue For

                If hasStartMarker Then
                    stack.Push(i)
                End If

                If hasEndMarker Then
                    Dim m As String = lines_(i).FoldingEndMarker

                    While stack.Count > 0
                        Dim iStartLine As Integer = stack.Pop()
                        foldingPairs(iStartLine) = i
                        If m = lines_(iStartLine).FoldingStartMarker Then Exit While
                    End While
                End If
            Next

            While stack.Count > 0
                foldingPairs(stack.Pop()) = endLine + 1
            End While
        End Sub

        Public Overridable Sub CollapseBlock(ByVal fromLine As Integer, ByVal toLine As Integer)
            Dim from As Integer = Math.Min(fromLine, toLine)
            Dim [to] As Integer = Math.Max(fromLine, toLine)
            If from = [to] Then Return

            While from <= [to]

                If GetLineText(from).Trim().Length > 0 Then

                    For i As Integer = from + 1 To [to]
                        SetVisibleState(i, VisibleState.Hidden)
                    Next

                    SetVisibleState(from, VisibleState.StartOfHiddenBlock)
                    Invalidate()
                    Exit For
                End If

                from += 1
            End While

            from = Math.Min(fromLine, toLine)
            [to] = Math.Max(fromLine, toLine)
            Dim newLine As Integer = FindNextVisibleLine([to])
            If newLine = [to] Then newLine = FindPrevVisibleLine(from)
            selection_.Start = New Place(0, newLine)
            needRecalc_ = True
            Invalidate()
            OnVisibleRangeChanged()
        End Sub

        Friend Function FindNextVisibleLine(ByVal iLine As Integer) As Integer
            If iLine >= lines_.Count - 1 Then Return iLine
            Dim old As Integer = iLine

            Do
                iLine += 1
            Loop While iLine < lines_.Count - 1 AndAlso LineInfos(iLine).VisibleState <> VisibleState.Visible

            If LineInfos(iLine).VisibleState <> VisibleState.Visible Then
                Return old
            Else
                Return iLine
            End If
        End Function

        Friend Function FindPrevVisibleLine(ByVal iLine As Integer) As Integer
            If iLine <= 0 Then Return iLine
            Dim old As Integer = iLine

            Do
                iLine -= 1
            Loop While iLine > 0 AndAlso LineInfos(iLine).VisibleState <> VisibleState.Visible

            If LineInfos(iLine).VisibleState <> VisibleState.Visible Then
                Return old
            Else
                Return iLine
            End If
        End Function

        Private Function FindVisualMarkerForPoint(ByVal p As Point) As VisualMarker
            For Each m As VisualMarker In visibleMarkers
                If m.rectangle.Contains(p) Then Return m
            Next

            Return Nothing
        End Function

        Public Overridable Sub IncreaseIndent()
            If selection_.Start = selection_.[End] Then

                If Not selection_.[ReadOnly] Then
                    selection_.Start = New Place(Me(selection_.Start.iLine).StartSpacesCount, selection_.Start.iLine)
                    Dim spaces As Integer = TabLength - (selection_.Start.iChar Mod TabLength)

                    If isReplaceMode Then

                        For i As Integer = 0 To spaces - 1
                            selection_.GoRight(True)
                        Next

                        selection_.Inverse()
                    End If

                    InsertText(New String(" "c, spaces))
                End If

                Return
            End If

            Dim carretAtEnd As Boolean = (selection_.Start > selection_.[End]) AndAlso Not selection_.ColumnSelectionMode
            Dim startChar As Integer = 0
            If selection_.ColumnSelectionMode Then startChar = Math.Min(selection_.[End].iChar, selection_.Start.iChar)
            BeginUpdate()
            selection_.BeginUpdate()
            lines_.Manager.BeginAutoUndoCommands()
            Dim old = selection_.Clone()
            lines_.Manager.ExecuteCommand(New SelectCommand(TextSource))
            selection_.Normalize()
            Dim currentSelection As Range = Me.selection_.Clone()
            Dim from As Integer = selection_.Start.iLine
            Dim [to] As Integer = selection_.[End].iLine

            If Not selection_.ColumnSelectionMode Then
                If selection_.[End].iChar = 0 Then [to] -= 1
            End If

            For i As Integer = from To [to]
                If lines_(i).Count = 0 Then Continue For
                selection_.Start = New Place(startChar, i)
                lines_.Manager.ExecuteCommand(New InsertTextCommand(TextSource, New String(" "c, TabLength)))
            Next

            If selection_.ColumnSelectionMode = False Then
                Dim newSelectionStartCharacterIndex As Integer = currentSelection.Start.iChar + Me.TabLength
                Dim newSelectionEndCharacterIndex As Integer = currentSelection.[End].iChar + (If(currentSelection.[End].iLine = [to], Me.TabLength, 0))
                Me.selection_.Start = New Place(newSelectionStartCharacterIndex, currentSelection.Start.iLine)
                Me.selection_.[End] = New Place(newSelectionEndCharacterIndex, currentSelection.[End].iLine)
            Else
                selection_ = old
            End If

            lines_.Manager.EndAutoUndoCommands()
            If carretAtEnd Then selection_.Inverse()
            needRecalc_ = True
            selection_.EndUpdate()
            EndUpdate()
            Invalidate()
        End Sub

        Public Overridable Sub DecreaseIndent()
            If selection_.Start.iLine = selection_.[End].iLine Then
                DecreaseIndentOfSingleLine()
                Return
            End If

            Dim startCharIndex As Integer = 0
            If selection_.ColumnSelectionMode Then startCharIndex = Math.Min(selection_.[End].iChar, selection_.Start.iChar)
            BeginUpdate()
            selection_.BeginUpdate()
            lines_.Manager.BeginAutoUndoCommands()
            Dim old = selection_.Clone()
            lines_.Manager.ExecuteCommand(New SelectCommand(TextSource))
            Dim currentSelection As Range = Me.selection_.Clone()
            selection_.Normalize()
            Dim from As Integer = selection_.Start.iLine
            Dim [to] As Integer = selection_.[End].iLine

            If Not selection_.ColumnSelectionMode Then
                If selection_.[End].iChar = 0 Then [to] -= 1
            End If

            Dim numberOfDeletedWhitespacesOfFirstLine As Integer = 0
            Dim numberOfDeletetWhitespacesOfLastLine As Integer = 0

            For i As Integer = from To [to]
                If startCharIndex > lines_(i).Count Then Continue For
                Dim endIndex As Integer = Math.Min(Me.lines_(i).Count, startCharIndex + Me.TabLength)
                Dim wasteText As String = Me.lines_(i).Text.Substring(startCharIndex, endIndex - startCharIndex)
                endIndex = Math.Min(endIndex, startCharIndex + wasteText.Length - wasteText.TrimStart().Length)
                Me.selection_ = New Range(Me, New Place(startCharIndex, i), New Place(endIndex, i))
                Dim numberOfWhitespacesToRemove As Integer = endIndex - startCharIndex

                If i = currentSelection.Start.iLine Then
                    numberOfDeletedWhitespacesOfFirstLine = numberOfWhitespacesToRemove
                End If

                If i = currentSelection.[End].iLine Then
                    numberOfDeletetWhitespacesOfLastLine = numberOfWhitespacesToRemove
                End If

                If Not selection_.IsEmpty Then Me.ClearSelected()
            Next

            If selection_.ColumnSelectionMode = False Then
                Dim newSelectionStartCharacterIndex As Integer = Math.Max(0, currentSelection.Start.iChar - numberOfDeletedWhitespacesOfFirstLine)
                Dim newSelectionEndCharacterIndex As Integer = Math.Max(0, currentSelection.[End].iChar - numberOfDeletetWhitespacesOfLastLine)
                Me.selection_.Start = New Place(newSelectionStartCharacterIndex, currentSelection.Start.iLine)
                Me.selection_.[End] = New Place(newSelectionEndCharacterIndex, currentSelection.[End].iLine)
            Else
                selection_ = old
            End If

            lines_.Manager.EndAutoUndoCommands()
            needRecalc_ = True
            selection_.EndUpdate()
            EndUpdate()
            Invalidate()
        End Sub

        Protected Overridable Sub DecreaseIndentOfSingleLine()
            If Me.selection_.Start.iLine <> Me.selection_.[End].iLine Then Return
            Dim currentSelection As Range = Me.selection_.Clone()
            Dim currentLineIndex As Integer = Me.selection_.Start.iLine
            Dim currentLeftSelectionStartIndex As Integer = Math.Min(Me.selection_.Start.iChar, Me.selection_.[End].iChar)
            Dim lineText As String = Me.lines_(currentLineIndex).Text
            Dim whitespacesLeftOfSelectionStartMatch As Match = New Regex("\s*", RegexOptions.RightToLeft).Match(lineText, currentLeftSelectionStartIndex)
            Dim leftOffset As Integer = whitespacesLeftOfSelectionStartMatch.Index
            Dim countOfWhitespaces As Integer = whitespacesLeftOfSelectionStartMatch.Length
            Dim numberOfCharactersToRemove As Integer = 0

            If countOfWhitespaces > 0 Then
                Dim remainder As Integer = If((Me.TabLength > 0), currentLeftSelectionStartIndex Mod Me.TabLength, 0)
                numberOfCharactersToRemove = If((remainder <> 0), Math.Min(remainder, countOfWhitespaces), Math.Min(Me.TabLength, countOfWhitespaces))
            End If

            If numberOfCharactersToRemove > 0 Then
                Me.BeginUpdate()
                Me.selection_.BeginUpdate()
                lines_.Manager.BeginAutoUndoCommands()
                lines_.Manager.ExecuteCommand(New SelectCommand(TextSource))
                Me.selection_.Start = New Place(leftOffset, currentLineIndex)
                Me.selection_.[End] = New Place(leftOffset + numberOfCharactersToRemove, currentLineIndex)
                ClearSelected()
                Dim newSelectionStartCharacterIndex As Integer = currentSelection.Start.iChar - numberOfCharactersToRemove
                Dim newSelectionEndCharacterIndex As Integer = currentSelection.[End].iChar - numberOfCharactersToRemove
                Me.selection_.Start = New Place(newSelectionStartCharacterIndex, currentLineIndex)
                Me.selection_.[End] = New Place(newSelectionEndCharacterIndex, currentLineIndex)
                lines_.Manager.ExecuteCommand(New SelectCommand(TextSource))
                lines_.Manager.EndAutoUndoCommands()
                Me.selection_.EndUpdate()
                Me.EndUpdate()
            End If

            Invalidate()
        End Sub

        Public Overridable Sub DoAutoIndent()
            If selection_.ColumnSelectionMode Then Return
            Dim r As Range = selection_.Clone()
            r.Normalize()
            BeginUpdate()
            selection_.BeginUpdate()
            lines_.Manager.BeginAutoUndoCommands()

            For i As Integer = r.Start.iLine To r.[End].iLine
                DoAutoIndent(i)
            Next

            lines_.Manager.EndAutoUndoCommands()
            selection_.Start = r.Start
            selection_.[End] = r.[End]
            selection_.Expand()
            selection_.EndUpdate()
            EndUpdate()
        End Sub

        Public Overridable Sub InsertLinePrefix(ByVal prefix As String)
            Dim old As Range = selection_.Clone()
            Dim from As Integer = Math.Min(selection_.Start.iLine, selection_.[End].iLine)
            Dim [to] As Integer = Math.Max(selection_.Start.iLine, selection_.[End].iLine)
            BeginUpdate()
            selection_.BeginUpdate()
            lines_.Manager.BeginAutoUndoCommands()
            lines_.Manager.ExecuteCommand(New SelectCommand(TextSource))
            Dim spaces As Integer = GetMinStartSpacesCount(from, [to])

            For i As Integer = from To [to]
                selection_.Start = New Place(spaces, i)
                lines_.Manager.ExecuteCommand(New InsertTextCommand(TextSource, prefix))
            Next

            selection_.Start = New Place(0, from)
            selection_.[End] = New Place(lines_([to]).Count, [to])
            needRecalc_ = True
            lines_.Manager.EndAutoUndoCommands()
            selection_.EndUpdate()
            EndUpdate()
            Invalidate()
        End Sub

        Public Overridable Sub RemoveLinePrefix(ByVal prefix As String)
            Dim old As Range = selection_.Clone()
            Dim from As Integer = Math.Min(selection_.Start.iLine, selection_.[End].iLine)
            Dim [to] As Integer = Math.Max(selection_.Start.iLine, selection_.[End].iLine)
            BeginUpdate()
            selection_.BeginUpdate()
            lines_.Manager.BeginAutoUndoCommands()
            lines_.Manager.ExecuteCommand(New SelectCommand(TextSource))

            For i As Integer = from To [to]
                Dim text As String = lines_(i).Text
                Dim trimmedText As String = text.TrimStart()

                If trimmedText.StartsWith(prefix) Then
                    Dim spaces As Integer = text.Length - trimmedText.Length
                    selection_.Start = New Place(spaces, i)
                    selection_.[End] = New Place(spaces + prefix.Length, i)
                    ClearSelected()
                End If
            Next

            selection_.Start = New Place(0, from)
            selection_.[End] = New Place(lines_([to]).Count, [to])
            needRecalc_ = True
            lines_.Manager.EndAutoUndoCommands()
            selection_.EndUpdate()
            EndUpdate()
        End Sub

        Public Sub BeginAutoUndo()
            lines_.Manager.BeginAutoUndoCommands()
        End Sub

        Public Sub EndAutoUndo()
            lines_.Manager.EndAutoUndoCommands()
        End Sub

        Public Overridable Sub OnVisualMarkerClick(ByVal args As MouseEventArgs, ByVal marker As StyleVisualMarker)
            RaiseEvent VisualMarkerClick(Me, New VisualMarkerEventArgs(marker.Style, marker, args))
            marker.Style.OnVisualMarkerClick(Me, New VisualMarkerEventArgs(marker.Style, marker, args))
        End Sub

        Protected Overridable Sub OnMarkerClick(ByVal args As MouseEventArgs, ByVal marker As VisualMarker)
            If TypeOf marker Is StyleVisualMarker Then
                OnVisualMarkerClick(args, TryCast(marker, StyleVisualMarker))
                Return
            End If

            If TypeOf marker Is CollapseFoldingMarker Then
                CollapseFoldingBlock((TryCast(marker, CollapseFoldingMarker)).iLine)
                Return
            End If

            If TypeOf marker Is ExpandFoldingMarker Then
                ExpandFoldedBlock((TryCast(marker, ExpandFoldingMarker)).iLine)
                Return
            End If

            If TypeOf marker Is FoldedAreaMarker Then
                Dim iStart As Integer = (TryCast(marker, FoldedAreaMarker)).iLine
                Dim iEnd As Integer = FindEndOfFoldingBlock(iStart)
                If iEnd < 0 Then Return
                selection_.BeginUpdate()
                selection_.Start = New Place(0, iStart)
                selection_.[End] = New Place(lines_(iEnd).Count, iEnd)
                selection_.EndUpdate()
                Invalidate()
                Return
            End If
        End Sub

        Protected Overridable Sub OnMarkerDoubleClick(ByVal marker As VisualMarker)
            If TypeOf marker Is FoldedAreaMarker Then
                ExpandFoldedBlock((TryCast(marker, FoldedAreaMarker)).iLine)
                Invalidate()
                Return
            End If
        End Sub

        Private Sub ClearBracketsPositions()
            leftBracketPosition = Nothing
            rightBracketPosition = Nothing
            leftBracketPosition2 = Nothing
            rightBracketPosition2 = Nothing
        End Sub

        Private Sub HighlightBrackets(ByVal LeftBracket As Char, ByVal RightBracket As Char, ByRef leftBracketPosition As Range, ByRef rightBracketPosition As Range)
            Select Case BracketsHighlightStrategy
                Case BracketsHighlightStrategy.Strategy1
                    HighlightBrackets1(LeftBracket, RightBracket, leftBracketPosition, rightBracketPosition)
                Case BracketsHighlightStrategy.Strategy2
                    HighlightBrackets2(LeftBracket, RightBracket, leftBracketPosition, rightBracketPosition)
            End Select
        End Sub

        Private Sub HighlightBrackets1(ByVal LeftBracket As Char, ByVal RightBracket As Char, ByRef leftBracketPosition As Range, ByRef rightBracketPosition As Range)
            If Not selection_.IsEmpty Then Return
            If LinesCount = 0 Then Return
            Dim oldLeftBracketPosition As Range = leftBracketPosition
            Dim oldRightBracketPosition As Range = rightBracketPosition
            Dim range = GetBracketsRange(selection_.Start, LeftBracket, RightBracket, True)

            If range IsNot Nothing Then
                leftBracketPosition = New Range(Me, range.Start, New Place(range.Start.iChar + 1, range.Start.iLine))
                rightBracketPosition = New Range(Me, New Place(range.[End].iChar - 1, range.[End].iLine), range.[End])
            End If

            If oldLeftBracketPosition <> leftBracketPosition OrElse oldRightBracketPosition <> rightBracketPosition Then Invalidate()
        End Sub

        Public Function GetBracketsRange(ByVal placeInsideBrackets As Place, ByVal leftBracket As Char, ByVal rightBracket As Char, ByVal includeBrackets As Boolean) As Range
            Dim startRange = New Range(Me, placeInsideBrackets, placeInsideBrackets)
            Dim range = startRange.Clone()
            Dim leftBracketPosition As Range = Nothing
            Dim rightBracketPosition As Range = Nothing
            Dim counter As Integer = 0
            Dim maxIterations As Integer = maxBracketSearchIterations

            While range.GoLeftThroughFolded()
                If range.CharAfterStart = leftBracket Then counter += 1
                If range.CharAfterStart = rightBracket Then counter -= 1

                If counter = 1 Then
                    range.Start = New Place(range.Start.iChar + (If(Not includeBrackets, 1, 0)), range.Start.iLine)
                    leftBracketPosition = range
                    Exit While
                End If

                maxIterations -= 1
                If maxIterations <= 0 Then Exit While
            End While

            range = startRange.Clone()
            counter = 0
            maxIterations = maxBracketSearchIterations

            Do
                If range.CharAfterStart = leftBracket Then counter += 1
                If range.CharAfterStart = rightBracket Then counter -= 1

                If counter = -1 Then
                    range.[End] = New Place(range.Start.iChar + (If(includeBrackets, 1, 0)), range.Start.iLine)
                    rightBracketPosition = range
                    Exit Do
                End If

                maxIterations -= 1
                If maxIterations <= 0 Then Exit Do
            Loop While range.GoRightThroughFolded()

            If leftBracketPosition IsNot Nothing AndAlso rightBracketPosition IsNot Nothing Then
                Return New Range(Me, leftBracketPosition.Start, rightBracketPosition.[End])
            Else
                Return Nothing
            End If
        End Function

        Private Sub HighlightBrackets2(ByVal LeftBracket As Char, ByVal RightBracket As Char, ByRef leftBracketPosition As Range, ByRef rightBracketPosition As Range)
            If Not selection_.IsEmpty Then Return
            If LinesCount = 0 Then Return
            Dim oldLeftBracketPosition As Range = leftBracketPosition
            Dim oldRightBracketPosition As Range = rightBracketPosition
            Dim range As Range = selection_.Clone()
            Dim found As Boolean = False
            Dim counter As Integer = 0
            Dim maxIterations As Integer = maxBracketSearchIterations

            If range.CharBeforeStart = RightBracket Then
                rightBracketPosition = New Range(Me, range.Start.iChar - 1, range.Start.iLine, range.Start.iChar, range.Start.iLine)

                While range.GoLeftThroughFolded()
                    If range.CharAfterStart = LeftBracket Then counter += 1
                    If range.CharAfterStart = RightBracket Then counter -= 1

                    If counter = 0 Then
                        range.[End] = New Place(range.Start.iChar + 1, range.Start.iLine)
                        leftBracketPosition = range
                        found = True
                        Exit While
                    End If

                    maxIterations -= 1
                    If maxIterations <= 0 Then Exit While
                End While
            End If

            range = selection_.Clone()
            counter = 0
            maxIterations = maxBracketSearchIterations

            If Not found Then

                If range.CharAfterStart = LeftBracket Then
                    leftBracketPosition = New Range(Me, range.Start.iChar, range.Start.iLine, range.Start.iChar + 1, range.Start.iLine)

                    Do
                        If range.CharAfterStart = LeftBracket Then counter += 1
                        If range.CharAfterStart = RightBracket Then counter -= 1

                        If counter = 0 Then
                            range.[End] = New Place(range.Start.iChar + 1, range.Start.iLine)
                            rightBracketPosition = range
                            found = True
                            Exit Do
                        End If

                        maxIterations -= 1
                        If maxIterations <= 0 Then Exit Do
                    Loop While range.GoRightThroughFolded()
                End If
            End If

            If oldLeftBracketPosition <> leftBracketPosition OrElse oldRightBracketPosition <> rightBracketPosition Then Invalidate()
        End Sub

        Public Function SelectNext(ByVal regexPattern As String, ByVal Optional backward As Boolean = False, ByVal Optional options As RegexOptions = RegexOptions.None) As Boolean
            Dim sel = selection_.Clone()
            sel.Normalize()
            Dim range1 = If(backward, New Range(Me, Range.Start, sel.Start), New Range(Me, sel.[End], Range.[End]))
            Dim res As Range = Nothing

            For Each r In range1.GetRanges(regexPattern, options)
                res = r
                If Not backward Then Exit For
            Next

            If res Is Nothing Then Return False
            selection_ = res
            Invalidate()
            Return True
        End Function

        Public Overridable Sub OnSyntaxHighlight(ByVal args As TextChangedEventArgs)
            Dim range As Range

            Select Case HighlightingRangeType
                Case HighlightingRangeType.VisibleRange
                    range = visibleRange.GetUnionWith(args.ChangedRange)
                Case HighlightingRangeType.AllTextRange
                    range = range
                Case Else
                    range = args.ChangedRange
            End Select

            If SyntaxHighlighter IsNot Nothing Then

                If language = language.Custom AndAlso Not String.IsNullOrEmpty(descriptionFile) Then
                    SyntaxHighlighter.HighlightSyntax(descriptionFile, range)
                Else
                    SyntaxHighlighter.HighlightSyntax(language, range)
                End If
            End If
        End Sub

        Private Sub InitializeComponent()
            SuspendLayout()
            Name = "FastColoredTextBox"
            ResumeLayout(False)
        End Sub

        Public Overridable Sub Print(ByVal range As Range, ByVal settings As PrintDialogSettings)
            Dim exporter = New ExportToHTML()
            exporter.UseBr = True
            exporter.UseForwardNbsp = True
            exporter.UseNbsp = True
            exporter.UseStyleTag = False
            exporter.IncludeLineNumbers = settings.IncludeLineNumbers
            If range Is Nothing Then range = range
            If range.Text = String.Empty Then Return
            visibleRange = range

            Try
                RaiseEvent VisibleRangeChanged(Me, New EventArgs())
                RaiseEvent VisibleRangeChangedDelayed(Me, New EventArgs())
            Finally
                visibleRange = Nothing
            End Try

            Dim HTML As String = exporter.GetHtml(range)
            HTML = "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8""><head><title>" & PrepareHtmlText(settings.Title) & "</title></head>" & HTML & "<br>" & SelectHTMLRangeScript()
            Dim tempFile As String = Path.GetTempPath() & "fctb.html"
            File.WriteAllText(tempFile, HTML)
            SetPageSetupSettings(settings)
            Dim wb = New WebBrowser()
            wb.Tag = settings
            wb.Visible = False
            wb.Location = New Point(-1000, -1000)
            wb.Parent = Me
            wb.StatusTextChanged += AddressOf wb_StatusTextChanged
            wb.Navigate(tempFile)
        End Sub

        Protected Overridable Function PrepareHtmlText(ByVal s As String) As String
            Return s.Replace("<", "&lt;").Replace(">", "&gt;").Replace("&", "&amp;")
        End Function

        Private Sub wb_StatusTextChanged(ByVal sender As Object, ByVal e As EventArgs)
            Dim wb = TryCast(sender, WebBrowser)

            If wb.StatusText.Contains("#print") Then
                Dim settings = TryCast(wb.Tag, PrintDialogSettings)

                Try

                    If settings.ShowPrintPreviewDialog Then
                        wb.ShowPrintPreviewDialog()
                    Else
                        If settings.ShowPageSetupDialog Then wb.ShowPageSetupDialog()

                        If settings.ShowPrintDialog Then
                            wb.ShowPrintDialog()
                        Else
                            wb.Print()
                        End If
                    End If

                Finally
                    wb.Parent = Nothing
                    wb.Dispose()
                End Try
            End If
        End Sub

        Public Sub Print(ByVal settings As PrintDialogSettings)
            Print(Range, settings)
        End Sub

        Public Sub Print()
            Print(Range, New PrintDialogSettings With {
            .ShowPageSetupDialog = False,
            .ShowPrintDialog = False,
            .ShowPrintPreviewDialog = False
        })
        End Sub

        Private Function SelectHTMLRangeScript() As String
            Dim sel As Range = selection_.Clone()
            sel.Normalize()
            Dim start As Integer = PlaceToPosition(sel.Start) - sel.Start.iLine
            Dim len As Integer = sel.Text.Length - (sel.[End].iLine - sel.Start.iLine)
            Return String.Format("<script type=""text/javascript"">
try{{
    var sel = document.selection_;
    var rng = sel.createRange();
    rng.moveStart(""character"", {0});
    rng.moveEnd(""character"", {1});
    rng.select();
}}catch(ex){{}}
window.status = ""#print"";
</script>", start, len)
        End Function

        Private Shared Sub SetPageSetupSettings(ByVal settings As PrintDialogSettings)
            Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Internet Explorer\PageSetup", True)

            If key IsNot Nothing Then
                key.SetValue("footer", settings.Footer)
                key.SetValue("header", settings.Header)
            End If
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            MyBase.Dispose(disposing)

            If disposing Then
                If SyntaxHighlighter IsNot Nothing Then SyntaxHighlighter.Dispose()
                timer.Dispose()
                timer2.Dispose()
                middleClickScrollingTimer.Dispose()
                If findForm IsNot Nothing Then findForm.Dispose()
                If replaceForm IsNot Nothing Then replaceForm.Dispose()
                If TextSource IsNot Nothing Then TextSource.Dispose()
                If ToolTip IsNot Nothing Then ToolTip.Dispose()
            End If
        End Sub

        Protected Overridable Sub OnPaintLine(ByVal e As PaintLineEventArgs)
            RaiseEvent PaintLine(Me, e)
        End Sub

        Friend Sub OnLineInserted(ByVal index As Integer)
            OnLineInserted(index, 1)
        End Sub

        Friend Sub OnLineInserted(ByVal index As Integer, ByVal count As Integer)
            RaiseEvent LineInserted(Me, New LineInsertedEventArgs(index, count))
        End Sub

        Friend Sub OnLineRemoved(ByVal index As Integer, ByVal count As Integer, ByVal removedLineIds As List(Of Integer))
            If count > 0 Then
                RaiseEvent LineRemoved(Me, New LineRemovedEventArgs(index, count, removedLineIds))
            End If
        End Sub

        Public Sub OpenFile(ByVal fileName As String, ByVal enc As Encoding)
            Dim ts = CreateTextSource()

            Try
                InitTextSource(ts)
                Text = File.ReadAllText(fileName, enc)
                ClearUndo()
                isChanged_ = False
                OnVisibleRangeChanged()
            Catch
                InitTextSource(CreateTextSource())
                lines_.InsertLine(0, TextSource.CreateLine())
                isChanged_ = False
                Throw
            End Try

            selection_.Start = Place.Empty
            DoSelectionVisible()
        End Sub

        Public Sub OpenFile(ByVal fileName As String)
            Try
                Dim enc = EncodingDetector.DetectTextFileEncoding(fileName)

                If enc IsNot Nothing Then
                    OpenFile(fileName, enc)
                Else
                    OpenFile(fileName, Encoding.[Default])
                End If

            Catch
                InitTextSource(CreateTextSource())
                lines_.InsertLine(0, TextSource.CreateLine())
                isChanged_ = False
                Throw
            End Try
        End Sub

        Public Sub OpenBindingFile(ByVal fileName As String, ByVal enc As Encoding)
            Dim fts = New FileTextSource(Me)

            Try
                InitTextSource(fts)
                fts.OpenFile(fileName, enc)
                isChanged_ = False
                OnVisibleRangeChanged()
            Catch
                fts.CloseFile()
                InitTextSource(CreateTextSource())
                lines_.InsertLine(0, TextSource.CreateLine())
                isChanged_ = False
                Throw
            End Try

            Invalidate()
        End Sub

        Public Sub CloseBindingFile()
            If TypeOf lines_ Is FileTextSource Then
                Dim fts = TryCast(lines_, FileTextSource)
                fts.CloseFile()
                InitTextSource(CreateTextSource())
                lines_.InsertLine(0, TextSource.CreateLine())
                isChanged_ = False
                Invalidate()
            End If
        End Sub

        Public Sub SaveToFile(ByVal fileName As String, ByVal enc As Encoding)
            lines_.SaveToFile(fileName, enc)
            isChanged_ = False
            OnVisibleRangeChanged()
            UpdateScrollbars()
        End Sub

        Public Sub SetVisibleState(ByVal iLine As Integer, ByVal state As VisibleState)
            Dim li As LineInfo = LineInfos(iLine)
            li.VisibleState = state
            LineInfos(iLine) = li
            needRecalc_ = True
        End Sub

        Public Function GetVisibleState(ByVal iLine As Integer) As VisibleState
            Return LineInfos(iLine).VisibleState
        End Function

        Public Sub ShowGoToDialog()
            Dim form = New GoToForm()
            form.TotalLineCount = LinesCount
            form.SelectedLineNumber = selection_.Start.iLine + 1

            If form.ShowDialog() = DialogResult.OK Then
                Dim line As Integer = Math.Min(LinesCount - 1, Math.Max(0, form.SelectedLineNumber - 1))
                selection_ = New Range(Me, 0, line, 0, line)
                DoSelectionVisible()
            End If
        End Sub

        Public Sub OnUndoRedoStateChanged()
            RaiseEvent UndoRedoStateChanged(Me, EventArgs.Empty)
        End Sub

        Public Function FindLines(ByVal searchPattern As String, ByVal options As RegexOptions) As List(Of Integer)
            Dim iLines = New List(Of Integer)()

            For Each r As Range In Range.GetRangesByLines(searchPattern, options)
                iLines.Add(r.Start.iLine)
            Next

            Return iLines
        End Function

        Public Sub RemoveLines(ByVal iLines As List(Of Integer))
            TextSource.Manager.ExecuteCommand(New RemoveLinesCommand(TextSource, iLines))
            If iLines.Count > 0 Then isChanged_ = True
            If LinesCount = 0 Then Text = ""
            NeedRecalc()
            Invalidate()
        End Sub

        Private Sub BeginInit() Implements ISupportInitialize.BeginInit
        End Sub

        Private Sub EndInit() Implements ISupportInitialize.EndInit
            OnTextChanged()
            selection_.Start = Place.Empty
            DoCaretVisible()
            isChanged_ = False
            ClearUndo()
        End Sub

        Private Property IsDragDrop As Boolean

        Protected Overrides Sub OnDragEnter(ByVal e As DragEventArgs)
            If e.Data.GetDataPresent(DataFormats.Text) AndAlso AllowDrop Then
                e.Effect = DragDropEffects.Copy
                IsDragDrop = True
            End If

            MyBase.OnDragEnter(e)
        End Sub

        Protected Overrides Sub OnDragDrop(ByVal e As DragEventArgs)
            If [ReadOnly] OrElse Not AllowDrop Then
                IsDragDrop = False
                Return
            End If

            If e.Data.GetDataPresent(DataFormats.Text) Then
                If ParentForm IsNot Nothing Then ParentForm.Activate()
                Focus()
                Dim p As Point = PointToClient(New Point(e.X, e.Y))
                Dim text = e.Data.GetData(DataFormats.Text).ToString()
                Dim place = PointToPlace(p)
                DoDragDrop(place, text)
                IsDragDrop = False
            End If

            MyBase.OnDragDrop(e)
        End Sub

        Private Sub DoDragDrop_old(ByVal place As Place, ByVal text As String)
            Dim insertRange As Range = New Range(Me, place, place)
            If insertRange.[ReadOnly] Then Return
            If (draggedRange IsNot Nothing) AndAlso (draggedRange.Contains(place) = True) Then Return
            Dim copyMode As Boolean = (draggedRange Is Nothing) OrElse (draggedRange.[ReadOnly]) OrElse ((ModifierKeys And Keys.Control) <> Keys.None)

            If draggedRange Is Nothing Then
                selection_.BeginUpdate()
                selection_.Start = place
                InsertText(text)
                selection_ = New Range(Me, place, selection_.Start)
                selection_.EndUpdate()
                Return
            End If

            Dim caretPositionAfterInserting As Place
            BeginAutoUndo()
            selection_.BeginUpdate()
            selection_ = draggedRange
            lines_.Manager.ExecuteCommand(New SelectCommand(lines_))

            If draggedRange.ColumnSelectionMode Then
                draggedRange.Normalize()
                insertRange = New Range(Me, place, New Place(place.iChar, place.iLine + draggedRange.[End].iLine - draggedRange.Start.iLine)) With {
                .ColumnSelectionMode = True
            }

                For i As Integer = LinesCount To insertRange.[End].iLine
                    selection_.GoLast(False)
                    InsertChar(vbLf)
                Next
            End If

            If Not insertRange.[ReadOnly] Then

                If place < draggedRange.Start Then

                    If copyMode = False Then
                        selection_ = draggedRange
                        ClearSelected()
                    End If

                    selection_ = insertRange
                    selection_.ColumnSelectionMode = insertRange.ColumnSelectionMode
                    InsertText(text)
                    caretPositionAfterInserting = selection_.Start
                Else
                    selection_ = insertRange
                    selection_.ColumnSelectionMode = insertRange.ColumnSelectionMode
                    InsertText(text)
                    caretPositionAfterInserting = selection_.Start
                    Dim lineLength = Me(caretPositionAfterInserting.iLine).Count

                    If copyMode = False Then
                        selection_ = draggedRange
                        ClearSelected()
                    End If

                    Dim shift = lineLength - Me(caretPositionAfterInserting.iLine).Count
                    caretPositionAfterInserting.iChar = caretPositionAfterInserting.iChar - shift
                    place.iChar = place.iChar - shift
                End If

                If Not draggedRange.ColumnSelectionMode Then
                    selection_ = New Range(Me, place, caretPositionAfterInserting)
                Else
                    draggedRange.Normalize()
                    selection_ = New Range(Me, place, New Place(place.iChar + draggedRange.[End].iChar - draggedRange.Start.iChar, place.iLine + draggedRange.[End].iLine - draggedRange.Start.iLine)) With {
                    .ColumnSelectionMode = True
                }
                End If
            End If

            selection_.EndUpdate()
            EndAutoUndo()
            draggedRange = Nothing
        End Sub

        Protected Overridable Sub DoDragDrop(ByVal place As Place, ByVal text As String)
            Dim insertRange As Range = New Range(Me, place, place)
            If insertRange.[ReadOnly] Then Return
            If (draggedRange IsNot Nothing) AndAlso (draggedRange.Contains(place) = True) Then Return
            Dim copyMode As Boolean = (draggedRange Is Nothing) OrElse (draggedRange.[ReadOnly]) OrElse ((ModifierKeys And Keys.Control) <> Keys.None)

            If draggedRange Is Nothing Then
                selection_.BeginUpdate()
                selection_.Start = place
                InsertText(text)
                selection_ = New Range(Me, place, selection_.Start)
                selection_.EndUpdate()
            Else

                If Not draggedRange.Contains(place) Then
                    BeginAutoUndo()
                    selection_ = draggedRange
                    lines_.Manager.ExecuteCommand(New SelectCommand(lines_))

                    If draggedRange.ColumnSelectionMode Then
                        draggedRange.Normalize()
                        insertRange = New Range(Me, place, New Place(place.iChar, place.iLine + draggedRange.[End].iLine - draggedRange.Start.iLine)) With {
                        .ColumnSelectionMode = True
                    }

                        For i As Integer = LinesCount To insertRange.[End].iLine
                            selection_.GoLast(False)
                            InsertChar(vbLf)
                        Next
                    End If

                    If Not insertRange.[ReadOnly] Then

                        If place < draggedRange.Start Then

                            If copyMode = False Then
                                selection_ = draggedRange
                                ClearSelected()
                            End If

                            selection_ = insertRange
                            selection_.ColumnSelectionMode = insertRange.ColumnSelectionMode
                            InsertText(text)
                        Else
                            selection_ = insertRange
                            selection_.ColumnSelectionMode = insertRange.ColumnSelectionMode
                            InsertText(text)

                            If copyMode = False Then
                                selection_ = draggedRange
                                ClearSelected()
                            End If
                        End If
                    End If

                    Dim startPosition As Place = place
                    Dim endPosition As Place = selection_.Start
                    Dim dR As Range = If((draggedRange.[End] > draggedRange.Start), Me.GetRange(draggedRange.Start, draggedRange.[End]), Me.GetRange(draggedRange.[End], draggedRange.Start))
                    Dim tP As Place = place
                    Dim tS_S_Line As Integer
                    Dim tS_S_Char As Integer
                    Dim tS_E_Line As Integer
                    Dim tS_E_Char As Integer

                    If (place > draggedRange.Start) AndAlso (copyMode = False) Then

                        If draggedRange.ColumnSelectionMode = False Then

                            If dR.Start.iLine <> dR.[End].iLine Then
                                tS_S_Char = If((dR.[End].iLine <> tP.iLine), tP.iChar, dR.Start.iChar + (tP.iChar - dR.[End].iChar))
                                tS_E_Char = dR.[End].iChar
                            Else

                                If dR.[End].iLine = tP.iLine Then
                                    tS_S_Char = tP.iChar - dR.Text.Length
                                    tS_E_Char = tP.iChar
                                Else
                                    tS_S_Char = tP.iChar
                                    tS_E_Char = tP.iChar + dR.Text.Length
                                End If
                            End If

                            If dR.[End].iLine <> tP.iLine Then
                                tS_S_Line = tP.iLine - (dR.[End].iLine - dR.Start.iLine)
                                tS_E_Line = tP.iLine
                            Else
                                tS_S_Line = dR.Start.iLine
                                tS_E_Line = dR.[End].iLine
                            End If

                            startPosition = New Place(tS_S_Char, tS_S_Line)
                            endPosition = New Place(tS_E_Char, tS_E_Line)
                        End If
                    End If

                    If Not draggedRange.ColumnSelectionMode Then
                        selection_ = New Range(Me, startPosition, endPosition)
                    Else

                        If (copyMode = False) AndAlso (place.iLine >= dR.Start.iLine) AndAlso (place.iLine <= dR.[End].iLine) AndAlso (place.iChar >= dR.[End].iChar) Then
                            tS_S_Char = tP.iChar - (dR.[End].iChar - dR.Start.iChar)
                            tS_E_Char = tP.iChar
                        Else
                            tS_S_Char = tP.iChar
                            tS_E_Char = tP.iChar + (dR.[End].iChar - dR.Start.iChar)
                        End If

                        tS_S_Line = tP.iLine
                        tS_E_Line = tP.iLine + (dR.[End].iLine - dR.Start.iLine)
                        startPosition = New Place(tS_S_Char, tS_S_Line)
                        endPosition = New Place(tS_E_Char, tS_E_Line)
                        selection_ = New Range(Me, startPosition, endPosition) With {
                        .ColumnSelectionMode = True
                    }
                    End If

                    EndAutoUndo()
                End If

                Me.selection_.Inverse()
                OnSelectionChanged()
            End If

            draggedRange = Nothing
        End Sub

        Protected Overrides Sub OnDragOver(ByVal e As DragEventArgs)
            If e.Data.GetDataPresent(DataFormats.Text) Then
                Dim p As Point = PointToClient(New Point(e.X, e.Y))
                selection_.Start = PointToPlace(p)
                If p.Y < 6 AndAlso VerticalScroll.Visible AndAlso VerticalScroll.Value > 0 Then VerticalScroll.Value = Math.Max(0, VerticalScroll.Value - charHeight_)
                DoCaretVisible()
                Invalidate()
            End If

            MyBase.OnDragOver(e)
        End Sub

        Protected Overrides Sub OnDragLeave(ByVal e As EventArgs)
            IsDragDrop = False
            MyBase.OnDragLeave(e)
        End Sub

        Private middleClickScrollingActivated As Boolean
        Private middleClickScrollingOriginPoint As Point
        Private middleClickScrollingOriginScroll As Point
        Private ReadOnly middleClickScrollingTimer As Timer = New Timer()
        Private middleClickScollDirection As ScrollDirection = ScrollDirection.None

        Private Sub ActivateMiddleClickScrollingMode(ByVal e As MouseEventArgs)
            If Not middleClickScrollingActivated Then

                If (Not HorizontalScroll.Visible) AndAlso (Not VerticalScroll.Visible) Then
                    If ShowScrollBars Then Return
                End If

                middleClickScrollingActivated = True
                middleClickScrollingOriginPoint = e.Location
                middleClickScrollingOriginScroll = New Point(HorizontalScroll.Value, VerticalScroll.Value)
                middleClickScrollingTimer.Interval = 50
                middleClickScrollingTimer.Enabled = True
                Capture = True
                Refresh()
                SendMessage(Handle, WM_SETREDRAW, 0, 0)
            End If
        End Sub

        Private Sub DeactivateMiddleClickScrollingMode()
            If middleClickScrollingActivated Then
                middleClickScrollingActivated = False
                middleClickScrollingTimer.Enabled = False
                Capture = False
                MyBase.Cursor = defaultCursor
                SendMessage(Handle, WM_SETREDRAW, 1, 0)
                Invalidate()
            End If
        End Sub

        Private Sub RestoreScrollsAfterMiddleClickScrollingMode()
            Dim xea = New ScrollEventArgs(ScrollEventType.ThumbPosition, HorizontalScroll.Value, middleClickScrollingOriginScroll.X, ScrollOrientation.HorizontalScroll)
            OnScroll(xea)
            Dim yea = New ScrollEventArgs(ScrollEventType.ThumbPosition, VerticalScroll.Value, middleClickScrollingOriginScroll.Y, ScrollOrientation.VerticalScroll)
            OnScroll(yea)
        End Sub

        <DllImport("user32.dll")>
        Private Shared Function SendMessage(ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
            Private Const WM_SETREDRAW As Integer = &HB

    Private Sub middleClickScrollingTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
            If IsDisposed Then Return
            If Not middleClickScrollingActivated Then Return
            Dim currentMouseLocation As Point = PointToClient(Cursor.Position)
            Capture = True
            Dim distanceX As Integer = Me.middleClickScrollingOriginPoint.X - currentMouseLocation.X
            Dim distanceY As Integer = Me.middleClickScrollingOriginPoint.Y - currentMouseLocation.Y
            If Not VerticalScroll.Visible AndAlso ShowScrollBars Then distanceY = 0
            If Not HorizontalScroll.Visible AndAlso ShowScrollBars Then distanceX = 0
            Dim angleInDegree As Double = 180 - Math.Atan2(distanceY, distanceX) * 180 / Math.PI
            Dim distance As Double = Math.Sqrt(Math.Pow(distanceX, 2) + Math.Pow(distanceY, 2))

            If distance > 10 Then

                If angleInDegree >= 325 OrElse angleInDegree <= 35 Then
                    Me.middleClickScollDirection = ScrollDirection.Right
                ElseIf angleInDegree <= 55 Then
                    Me.middleClickScollDirection = ScrollDirection.Right Or ScrollDirection.Up
                ElseIf angleInDegree <= 125 Then
                    Me.middleClickScollDirection = ScrollDirection.Up
                ElseIf angleInDegree <= 145 Then
                    Me.middleClickScollDirection = ScrollDirection.Up Or ScrollDirection.Left
                ElseIf angleInDegree <= 215 Then
                    Me.middleClickScollDirection = ScrollDirection.Left
                ElseIf angleInDegree <= 235 Then
                    Me.middleClickScollDirection = ScrollDirection.Left Or ScrollDirection.Down
                ElseIf angleInDegree <= 305 Then
                    Me.middleClickScollDirection = ScrollDirection.Down
                Else
                    Me.middleClickScollDirection = ScrollDirection.Down Or ScrollDirection.Right
                End If
            Else
                Me.middleClickScollDirection = ScrollDirection.None
            End If

            Select Case Me.middleClickScollDirection
                Case ScrollDirection.Right
                    MyBase.Cursor = Cursors.PanEast
                Case ScrollDirection.Right Or ScrollDirection.Up
                    MyBase.Cursor = Cursors.PanNE
                Case ScrollDirection.Up
                    MyBase.Cursor = Cursors.PanNorth
                Case ScrollDirection.Up Or ScrollDirection.Left
                    MyBase.Cursor = Cursors.PanNW
                Case ScrollDirection.Left
                    MyBase.Cursor = Cursors.PanWest
                Case ScrollDirection.Left Or ScrollDirection.Down
                    MyBase.Cursor = Cursors.PanSW
                Case ScrollDirection.Down
                    MyBase.Cursor = Cursors.PanSouth
                Case ScrollDirection.Down Or ScrollDirection.Right
                    MyBase.Cursor = Cursors.PanSE
                Case Else
                    MyBase.Cursor = defaultCursor
                    Return
            End Select

            Dim xScrollOffset = CInt((-distanceX / 5.0))
            Dim yScrollOffset = CInt((-distanceY / 5.0))
            Dim xea = New ScrollEventArgs(If(xScrollOffset < 0, ScrollEventType.SmallIncrement, ScrollEventType.SmallDecrement), HorizontalScroll.Value, HorizontalScroll.Value + xScrollOffset, ScrollOrientation.HorizontalScroll)
            Dim yea = New ScrollEventArgs(If(yScrollOffset < 0, ScrollEventType.SmallDecrement, ScrollEventType.SmallIncrement), VerticalScroll.Value, VerticalScroll.Value + yScrollOffset, ScrollOrientation.VerticalScroll)
            If (middleClickScollDirection And (ScrollDirection.Down Or ScrollDirection.Up)) > 0 Then OnScroll(yea, False)
            If (middleClickScollDirection And (ScrollDirection.Right Or ScrollDirection.Left)) > 0 Then OnScroll(xea)
            SendMessage(Handle, WM_SETREDRAW, 1, 0)
            Refresh()
            SendMessage(Handle, WM_SETREDRAW, 0, 0)
        End Sub

        Private Sub DrawMiddleClickScrolling(ByVal gr As Graphics)
            Dim ableToScrollVertically As Boolean = Me.VerticalScroll.Visible OrElse Not ShowScrollBars
            Dim ableToScrollHorizontally As Boolean = Me.HorizontalScroll.Visible OrElse Not ShowScrollBars
            Dim inverseColor As Color = Color.FromArgb(100, CByte(Not Me.BackColor.R), CByte(Not Me.BackColor.G), CByte(Not Me.BackColor.B))

            Using inverseColorBrush As SolidBrush = New SolidBrush(inverseColor)
                Dim p = middleClickScrollingOriginPoint
                Dim state = gr.Save()
                gr.SmoothingMode = SmoothingMode.HighQuality
                gr.TranslateTransform(p.X, p.Y)
                gr.FillEllipse(inverseColorBrush, -2, -2, 4, 4)
                If ableToScrollVertically Then DrawTriangle(gr, inverseColorBrush)
                gr.RotateTransform(90)
                If ableToScrollHorizontally Then DrawTriangle(gr, inverseColorBrush)
                gr.RotateTransform(90)
                If ableToScrollVertically Then DrawTriangle(gr, inverseColorBrush)
                gr.RotateTransform(90)
                If ableToScrollHorizontally Then DrawTriangle(gr, inverseColorBrush)
                gr.Restore(state)
            End Using
        End Sub

        Private Sub DrawTriangle(ByVal g As Graphics, ByVal brush As Brush)
            Const size As Integer = 5
            Dim points = New Point() {New Point(size, 2 * size), New Point(0, 3 * size), New Point(-size, 2 * size)}
            g.FillPolygon(brush, points)
        End Sub

        Private Class LineYComparer
            Inherits IComparer(Of LineInfo)

            Private ReadOnly Y As Integer

            Public Sub New(ByVal Y As Integer)
                Me.Y = Y
            End Sub

            Public Function Compare(ByVal x As LineInfo, ByVal y As LineInfo) As Integer
                If x.startY = -10 Then
                    Return -y.startY.CompareTo(y)
                Else
                    Return x.startY.CompareTo(y)
                End If
            End Function
        End Class
    End Class

    Public Class PaintLineEventArgs
        Inherits PaintEventArgs

        Public Sub New(ByVal iLine As Integer, ByVal rect As Rectangle, ByVal gr As Graphics, ByVal clipRect As Rectangle)
            MyBase.New(gr, clipRect)
            LineIndex = iLine
            LineRect = rect
        End Sub

        Public Property LineIndex As Integer
            Get
            End Get
            Set
            End Set
        End Property

        Public Property LineRect As Rectangle
            Get
            End Get
            Set
            End Set
        End Property
    End Class

    Public Class LineInsertedEventArgs
        Inherits EventArgs

        Public Sub New(ByVal index As Integer, ByVal count As Integer)
            MyBase.New
            index = index
            count = count
        End Sub

        ''' <summary>
        ''' Inserted line index
        ''' </summary>
        Public Property Index As Integer
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Count of inserted lines_
        ''' </summary>
        Public Property Count As Integer
            Get
            End Get
            Set
            End Set
        End Property
    End Class

    Public Class LineRemovedEventArgs
        Inherits EventArgs

        Public Sub New(ByVal index As Integer, ByVal count As Integer, ByVal removedLineIds As List(Of Integer))
            MyBase.New
            index = index
            count = count
            RemovedLineUniqueIds = removedLineIds
        End Sub

        ''' <summary>
        ''' Removed line index
        ''' </summary>
        Public Property Index As Integer
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Count of removed lines_
        ''' </summary>
        Public Property Count As Integer
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' UniqueIds of removed lines_
        ''' </summary>
        Public Property RemovedLineUniqueIds As List(Of Integer)
            Get
            End Get
            Set
            End Set
        End Property
    End Class

    ''' <summary>
    ''' TextChanged event argument
    ''' </summary>
    Public Class TextChangedEventArgs
        Inherits EventArgs

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New(ByVal changedRange As Range)
            MyBase.New
            changedRange = changedRange
        End Sub

        ''' <summary>
        ''' This range contains changed area of text
        ''' </summary>
        Public Property ChangedRange As Range
            Get
            End Get
            Set
            End Set
        End Property
    End Class

    Public Class TextChangingEventArgs
        Inherits EventArgs

        Public Property InsertingText As String
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Set to true if you want to cancel text inserting
        ''' </summary>
        Public Property Cancel As Boolean
            Get
            End Get
            Set
            End Set
        End Property
    End Class

    Public Class WordWrapNeededEventArgs
        Inherits EventArgs

        Public Property CutOffPositions As List(Of Integer)
            Get
            End Get
            Set
            End Set
        End Property

        Public Property ImeAllowed As Boolean
            Get
            End Get
            Set
            End Set
        End Property

        Public Property Line As Line
            Get
            End Get
            Set
            End Set
        End Property

        Public Sub New(ByVal cutOffPositions As List(Of Integer), ByVal imeAllowed As Boolean, ByVal line As Line)
            MyBase.New
            Me.CutOffPositions = cutOffPositions
            Me.ImeAllowed = imeAllowed
            Me.Line = line
        End Sub
    End Class

    Public Enum WordWrapMode

        WordWrapControlWidth

        WordWrapPreferredWidth

        CharWrapControlWidth

        CharWrapPreferredWidth

        Custom
    End Enum

    Public Class PrintDialogSettings

        Public Sub New()
            MyBase.New
            ShowPrintPreviewDialog = True
            Title = ""
            Footer = ""
            Header = ""
            PrinterSettings = New System.Drawing.Printing.PrinterSettings
        End Sub

        Public Property ShowPageSetupDialog As Boolean
            Get
            End Get
            Set
            End Set
        End Property

        Public Property ShowPrintDialog As Boolean
            Get
            End Get
            Set
            End Set
        End Property

        Public Property ShowPrintPreviewDialog As Boolean
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Title of page. If you want to print Title on the page, insert code &w in Footer or Header.
        ''' </summary>
        Public Property Title As String
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Footer of page.
        ''' Here you can use special codes: &w (Window title), &D, &d (Date), &t(), &4 (Time), &p (Current page number), &P (Total number of pages),  && (A single ampersand), &b (Right justify text, Center text. If &b occurs once, then anything after the &b is right justified. If &b occurs twice, then anything between the two &b is centered, and anything after the second &b is right justified).
        ''' More detailed see <see cref="http://msdn.microsoft.com/en-us/library/aa969429(v=vs.85).aspx">here</see>
        ''' </summary>
        Public Property Footer As String
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Header of page
        ''' Here you can use special codes: &w (Window title), &D, &d (Date), &t(), &4 (Time), &p (Current page number), &P (Total number of pages),  && (A single ampersand), &b (Right justify text, Center text. If &b occurs once, then anything after the &b is right justified. If &b occurs twice, then anything between the two &b is centered, and anything after the second &b is right justified).
        ''' More detailed see <see cref="http://msdn.microsoft.com/en-us/library/aa969429(v=vs.85).aspx">here</see>
        ''' </summary>
        Public Property Header As String
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Prints line numbers
        ''' </summary>
        Public Property IncludeLineNumbers As Boolean
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Printer settings
        ''' Helpful when printing without a printdialog or to handle print settings programatically
        ''' </summary>
        Public Property PrinterSettings As System.Drawing.Printing.PrinterSettings
            Get
            End Get
            Set
            End Set
        End Property
    End Class

    Public Class AutoIndentEventArgs
        Inherits EventArgs

        Public Sub New(ByVal iLine As Integer, ByVal lineText As String, ByVal prevLineText As String, ByVal tabLength As Integer, ByVal currentIndentation As Integer)
            MyBase.New
            Me.iLine = iLine
            lineText = lineText
            prevLineText = prevLineText
            tabLength = tabLength
            AbsoluteIndentation = currentIndentation
        End Sub

        Public Property iLine As Integer
            Get
            End Get
            Set
            End Set
        End Property

        Public Property TabLength As Integer
            Get
            End Get
            Set
            End Set
        End Property

        Public Property LineText As String
            Get
            End Get
            Set
            End Set
        End Property

        Public Property PrevLineText As String
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Additional spaces count for this line, relative to previous line
        ''' </summary>
        Public Property Shift As Integer
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Additional spaces count for next line, relative to previous line
        ''' </summary>
        Public Property ShiftNextLines As Integer
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Absolute indentation of current line. You can change this property if you want to set absolute indentation.
        ''' </summary>
        Public Property AbsoluteIndentation As Integer
            Get
            End Get
            Set
            End Set
        End Property
    End Class

    ''' <summary>
    ''' Type of highlighting
    ''' </summary>
    Public Enum HighlightingRangeType

        ChangedRange

        VisibleRange

        AllTextRange
    End Enum

    ''' <summary>
    ''' Strategy of search of end of folding block
    ''' </summary>
    Public Enum FindEndOfFoldingBlockStrategy

        Strategy1

        Strategy2
    End Enum

    ''' <summary>
    ''' Strategy of search of brackets to highlighting
    ''' </summary>
    Public Enum BracketsHighlightStrategy

        Strategy1

        Strategy2
    End Enum

    ''' <summary>
    ''' ToolTipNeeded event args
    ''' </summary>
    Public Class ToolTipNeededEventArgs
        Inherits EventArgs

        Public Sub New(ByVal place As Place, ByVal hoveredWord As String)
            MyBase.New
            hoveredWord = hoveredWord
            place = place
        End Sub

        Public Property Place As Place
            Get
            End Get
            Set
            End Set
        End Property

        Public Property HoveredWord As String
            Get
            End Get
            Set
            End Set
        End Property

        Public Property ToolTipTitle As String
            Get
            End Get
            Set
            End Set
        End Property

        Public Property ToolTipText As String
            Get
            End Get
            Set
            End Set
        End Property

        Public Property ToolTipIcon As ToolTipIcon
            Get
            End Get
            Set
            End Set
        End Property
    End Class

    ''' <summary>
    ''' HintClick event args
    ''' </summary>
    Public Class HintClickEventArgs
        Inherits EventArgs

        Public Sub New(ByVal hint As Hint)
            MyBase.New
            hint = hint
        End Sub

        Public Property Hint As Hint
            Get
            End Get
            Set
            End Set
        End Property
    End Class

    ''' <summary>
    ''' CustomAction event args
    ''' </summary>
    Public Class CustomActionEventArgs
        Inherits EventArgs

        Public Property Action As FCTBAction
            Get
            End Get
            Set
            End Set
        End Property

        Public Sub New(ByVal action As FCTBAction)
            MyBase.New
            Me.Action = action
        End Sub
    End Class

    Public Enum TextAreaBorderType

        None

        Single

        Shadow
    End Enum

    <Flags()>
    Public Enum ScrollDirection As System.UInt16

        None = 0

        Left = 1

        Right = 2

        Up = 4

        Down = 8
    End Enum

    <Serializable()>
    Public Class ServiceColors

        Public Property CollapseMarkerForeColor As Color
            Get
            End Get
            Set
            End Set
        End Property

        Public Property CollapseMarkerBackColor As Color
            Get
            End Get
            Set
            End Set
        End Property

        Public Property CollapseMarkerBorderColor As Color
            Get
            End Get
            Set
            End Set
        End Property

        Public Property ExpandMarkerForeColor As Color
            Get
            End Get
            Set
            End Set
        End Property

        Public Property ExpandMarkerBackColor As Color
            Get
            End Get
            Set
            End Set
        End Property

        Public Property ExpandMarkerBorderColor As Color
            Get
            End Get
            Set
            End Set
        End Property

        Public Sub New()
            MyBase.New
            Me.CollapseMarkerForeColor = Color.Silver
            Me.CollapseMarkerBackColor = Color.White
            Me.CollapseMarkerBorderColor = Color.Silver
            Me.ExpandMarkerForeColor = Color.Red
            Me.ExpandMarkerBackColor = Color.White
            Me.ExpandMarkerBorderColor = Color.Silver
        End Sub
    End Class

    ''' <summary>
    ''' Style index mask (32 styles)
    ''' </summary>
    <Flags()>
    Public Enum StyleIndex As UInteger

        None = 0

        Style0 = 1

        Style1 = 2

        Style2 = 4

        Style3 = 8

        Style4 = 16

        Style5 = 32

        Style6 = 64

        Style7 = 128

        Style8 = 256

        Style9 = 512

        Style10 = 1024

        Style11 = 2048

        Style12 = 4096

        Style13 = 8192

        Style14 = 16384

        Style15 = 32768

        Style16 = 65536

        Style17 = 131072

        Style18 = 262144

        Style19 = 524288

        Style20 = 1048576

        Style21 = 2097152

        Style22 = 4194304

        Style23 = 8388608

        Style24 = 16777216

        Style25 = 33554432

        Style26 = 67108864

        Style27 = 134217728

        Style28 = 268435456

        Style29 = 536870912

        Style30 = 1073741824

        Style31 = 2147483648

        All = 4294967295
    End Enum

    ''' <summary>
    ''' Style index mask (16 styles)
    ''' </summary>
    <Flags()>
    Public Enum StyleIndex As System.UInt16

        None = 0

        Style0 = 1

        Style1 = 2

        Style2 = 4

        Style3 = 8

        Style4 = 16

        Style5 = 32

        Style6 = 64

        Style7 = 128

        Style8 = 256

        Style9 = 512

        Style10 = 1024

        Style11 = 2048

        Style12 = 4096

        Style13 = 8192

        Style14 = 16384

        Style15 = 32768

        All = 65535
    End Enum
End Namespace