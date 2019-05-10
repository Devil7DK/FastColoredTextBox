Imports System.Collections.Generic
Imports System
Imports System.Text
Imports System.Drawing

Namespace FastColoredTextBoxNS

    ''' <summary>
    ''' Line of text
    ''' </summary>
    Public Class Line
        Implements IList(Of Char)

        Protected chars As List(Of Char)

        Public Property FoldingStartMarker As String
            Get
            End Get
            Set
            End Set
        End Property

        Public Property FoldingEndMarker As String
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Text of line was changed
        ''' </summary>
        Public Property IsChanged As Boolean
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Time of last visit of caret in this line
        ''' </summary>
        ''' <remarks>This property can be used for forward/backward navigating</remarks>
        Public Property LastVisit As DateTime
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Background brush.
        ''' </summary>
        Public Property BackgroundBrush As Brush
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Unique ID
        ''' </summary>
        Public Property UniqueId As Integer
            Get
            End Get
            Set
            End Set
        End Property

        ''' <summary>
        ''' Count of needed start spaces for AutoIndent
        ''' </summary>
        Public Property AutoIndentSpacesNeededCount As Integer
            Get
            End Get
            Set
            End Set
        End Property

        Friend Sub New(ByVal uid As Integer)
            MyBase.New
            Me.UniqueId = uid
            Me.chars = New List(Of Char)
        End Sub

        ''' <summary>
        ''' Clears style of chars, delete folding markers
        ''' </summary>
        Public Sub ClearStyle(ByVal styleIndex As StyleIndex)
            Me.FoldingStartMarker = Nothing
            Me.FoldingEndMarker = Nothing
            Dim i As Integer = 0
            Do While (i < Count)
                Dim c As Char = Me(i)
                styleIndex
                Me(i) = c
                i = (i + 1)
            Loop

        End Sub

        ''' <summary>
        ''' Text of the line
        ''' </summary>
        Public Overridable ReadOnly Property Text As String
            Get
                Dim sb As StringBuilder = New StringBuilder(Count)
                For Each c As Char In Me
                    sb.Append(c.c)
                Next
                Return sb.ToString
            End Get
        End Property

        ''' <summary>
        ''' Clears folding markers
        ''' </summary>
        Public Sub ClearFoldingMarkers()
            Me.FoldingStartMarker = Nothing
            Me.FoldingEndMarker = Nothing
        End Sub

        ''' <summary>
        ''' Count of start spaces
        ''' </summary>
        Public ReadOnly Property StartSpacesCount As Integer
            Get
                Dim spacesCount As Integer = 0
                Dim i As Integer = 0
                Do While (i < Count)
                    If (Me(i).c = Microsoft.VisualBasic.ChrW(32)) Then
                        spacesCount = (spacesCount + 1)
                    Else
                        Exit For
                    End If

                    i = (i + 1)
                Loop

                Return spacesCount
            End Get
        End Property

        Public Function IndexOf(ByVal item As Char) As Integer
            Return Me.chars.IndexOf(item)
        End Function

        Public Sub Insert(ByVal index As Integer, ByVal item As Char)
            Me.chars.Insert(index, item)
        End Sub

        Public Sub RemoveAt(ByVal index As Integer)
            Me.chars.RemoveAt(index)
        End Sub

        Default Public Property Item(ByVal index As Integer) As Char
            Get
                Return Me.chars(index)
            End Get
            Set
                Me.chars(index) = Value
            End Set
        End Property

        Public Sub Add(ByVal item As Char)
            Me.chars.Add(item)
        End Sub

        Public Sub Clear()
            Me.chars.Clear()
        End Sub

        Public Function Contains(ByVal item As Char) As Boolean
            Return Me.chars.Contains(item)
        End Function

        Public Sub CopyTo(ByVal array() As Char, ByVal arrayIndex As Integer)
            Me.chars.CopyTo(array, arrayIndex)
        End Sub

        ''' <summary>
        ''' Chars count
        ''' </summary>
        Public ReadOnly Property Count As Integer
            Get
                Return Me.chars.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean
            Get
                Return False
            End Get
        End Property

        Public Function Remove(ByVal item As Char) As Boolean
            Return Me.chars.Remove(item)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of Char)
            Return Me.chars.GetEnumerator
        End Function

        Function System_Collections_IEnumerable_GetEnumerator_(() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator.(
            Return CType(Me.chars.GetEnumerator, System.Collections.IEnumerator)
        End Function

        Public Overridable Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
            If (index >= Me.Count) Then
                Return
            End If

            Me.chars.RemoveRange(index, Math.Min((Me.Count - index), count))
        End Sub

        Public Overridable Sub TrimExcess()
            Me.chars.TrimExcess()
        End Sub

        Public Overridable Sub AddRange(ByVal collection As IEnumerable(Of Char))
            Me.chars.AddRange(collection)
        End Sub
    End Class

    Public Structure LineInfo

        Private cutOffPositions As List(Of Integer)

        'Y coordinate of line on screen
        Friend startY As Integer

        Friend bottomPadding As Integer

        'indent of secondary wordwrap strings (in chars)
        Friend wordWrapIndent As Integer

        ''' <summary>
        ''' Visible state
        ''' </summary>
        Public VisibleState As VisibleState

        Public Sub New(ByVal startY As Integer)
            MyBase.New
            Me.cutOffPositions = Nothing
            Me.VisibleState = Me.VisibleState.Visible
            Me.startY = Me.startY
            Me.bottomPadding = 0
            Me.wordWrapIndent = 0
        End Sub

        ''' <summary>
        ''' Positions for wordwrap cutoffs
        ''' </summary>
        Public ReadOnly Property CutOffPositions As List(Of Integer)
            Get
                If (Me.cutOffPositions Is Nothing) Then
                    Me.cutOffPositions = New List(Of Integer)
                End If

                Return Me.cutOffPositions
            End Get
        End Property

        ''' <summary>
        ''' Count of wordwrap string count for this line
        ''' </summary>
        Public ReadOnly Property WordWrapStringsCount As Integer
            Get
                Select Case (Me.VisibleState)
                    Case Me.VisibleState.Visible
                        If (Me.cutOffPositions Is Nothing) Then
                            Return 1
                        Else
                            Return (Me.cutOffPositions.Count + 1)
                        End If

                    Case Me.VisibleState.Hidden
                        Return 0
                    Case Me.VisibleState.StartOfHiddenBlock
                        Return 1
                End Select

                Return 0
            End Get
        End Property

        Friend Function GetWordWrapStringStartPosition(ByVal iWordWrapLine As Integer) As Integer
            Return 0
            'TODO: Warning!!!, inline IF is not supported ?
            (iWordWrapLine = 0)
            Me.cutOffPositions((iWordWrapLine - 1))
        End Function

        Friend Function GetWordWrapStringFinishPosition(ByVal iWordWrapLine As Integer, ByVal line As Line) As Integer
            If (Me.WordWrapStringsCount <= 0) Then
                Return 0
            End If

            Return (line.Count - 1)
            'TODO: Warning!!!, inline IF is not supported ?
            (iWordWrapLine  _
                        = (Me.WordWrapStringsCount - 1))
            (Me.CutOffPositions(iWordWrapLine) - 1)
        End Function

        ''' <summary>
        ''' Gets index of wordwrap string for given char position
        ''' </summary>
        Public Function GetWordWrapStringIndex(ByVal iChar As Integer) As Integer
            If ((Me.cutOffPositions Is Nothing) _
                        OrElse (Me.cutOffPositions.Count = 0)) Then
                Return 0
            End If

            Dim i As Integer = 0
            Do While (i < Me.cutOffPositions.Count)
                If (Me.cutOffPositions(i) > iChar) Then
                    Return i
                End If

                i = (i + 1)
            Loop

            Return Me.cutOffPositions.Count
        End Function
    End Structure

    Public Enum VisibleState As Byte

        Visible

        StartOfHiddenBlock

        Hidden
    End Enum

    Public Enum IndentMarker

        None

        Increased

        Decreased
    End Enum
End Namespace