Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.IO

Namespace FastColoredTextBoxNS
    Public Class FileTextSource
        Inherits TextSource
        Implements IDisposable

        Private sourceFileLinePositions As List(Of Integer) = New List(Of Integer)()
        Private fs As FileStream
        Private fileEncoding As Encoding
        Private timer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer()
        Public Event LineNeeded As EventHandler(Of LineNeededEventArgs)
        Public Event LinePushed As EventHandler(Of LinePushedEventArgs)

        Public Sub New(ByVal currentTB As FastColoredTextBox)
            MyBase.New(currentTB)
            timer.Interval = 10000
            AddHandler timer.Tick, New EventHandler(AddressOf timer_Tick)
            timer.Enabled = True
            SaveEOL = Environment.NewLine
        End Sub

        Private Sub timer_Tick(ByVal sender As Object, ByVal e As EventArgs)
            timer.Enabled = False

            Try
                UnloadUnusedLines()
            Finally
                timer.Enabled = True
            End Try
        End Sub

        Private Sub UnloadUnusedLines()
            Const margin As Integer = 2000
            Dim iStartVisibleLine = CurrentTB.VisibleRange.Start.iLine
            Dim iFinishVisibleLine = CurrentTB.VisibleRange.[End].iLine
            Dim count As Integer = 0

            For i As Integer = 0 To count - 1

                If MyBase.lines(i) IsNot Nothing AndAlso Not MyBase.lines(i).IsChanged AndAlso Math.Abs(i - iFinishVisibleLine) > margin Then
                    MyBase.lines(i) = Nothing
                    count += 1
                End If
            Next
        End Sub

        Public Sub OpenFile(ByVal fileName As String, ByVal enc As Encoding)
            Clear()
            If fs IsNot Nothing Then fs.Dispose()
            SaveEOL = Environment.NewLine
            fs = New FileStream(fileName, FileMode.Open)
            Dim length = fs.Length
            enc = DefineEncoding(enc, fs)
            Dim shift As Integer = DefineShift(enc)
            sourceFileLinePositions.Add(CInt(fs.Position))
            MyBase.lines.Add(Nothing)
            sourceFileLinePositions.Capacity = CInt((length / 7 + 1000))
            Dim prev As Integer = 0
            Dim prevPos As Integer = 0
            Dim br As BinaryReader = New BinaryReader(fs, enc)

            While fs.Position < length
                prevPos = CInt(fs.Position)
                Dim b = br.ReadChar()

                If b = 10 Then
                    sourceFileLinePositions.Add(CInt(fs.Position))
                    MyBase.lines.Add(Nothing)
                ElseIf prev = 13 Then
                    sourceFileLinePositions.Add(CInt(prevPos))
                    MyBase.lines.Add(Nothing)
                    SaveEOL = vbCr
                End If

                prev = b
            End While

            If prev = 13 Then
                sourceFileLinePositions.Add(CInt(prevPos))
                MyBase.lines.Add(Nothing)
            End If

            If length > 2000000 Then GC.Collect()
            Dim temp As Line() = New Line(99) {}
            Dim c = MyBase.lines.Count
            MyBase.lines.AddRange(temp)
            MyBase.lines.TrimExcess()
            MyBase.lines.RemoveRange(c, temp.Length)
            Dim temp2 As Integer() = New Integer(99) {}
            c = MyBase.lines.Count
            sourceFileLinePositions.AddRange(temp2)
            sourceFileLinePositions.TrimExcess()
            sourceFileLinePositions.RemoveRange(c, temp.Length)
            fileEncoding = enc
            OnLineInserted(0, Count)
            Dim linesCount = Math.Min(lines.Count, CurrentTB.ClientRectangle.Height / CurrentTB.CharHeight)

            For i As Integer = 0 To linesCount - 1
                LoadLineFromSourceFile(i)
            Next

            NeedRecalc(New TextChangedEventArgs(0, linesCount - 1))
            If CurrentTB.WordWrap Then OnRecalcWordWrap(New TextChangedEventArgs(0, linesCount - 1))
        End Sub

        Private Function DefineShift(ByVal enc As Encoding) As Integer
            If enc.IsSingleByte Then Return 0
            If enc.HeaderName = "unicodeFFFE" Then Return 0
            If enc.HeaderName = "utf-16" Then Return 1
            If enc.HeaderName = "utf-32BE" Then Return 0
            If enc.HeaderName = "utf-32" Then Return 3
            Return 0
        End Function

        Private Shared Function DefineEncoding(ByVal enc As Encoding, ByVal fs As FileStream) As Encoding
            Dim bytesPerSignature As Integer = 0
            Dim signature As Byte() = New Byte(3) {}
            Dim c As Integer = fs.Read(signature, 0, 4)

            If signature(0) = &HFF AndAlso signature(1) = &HFE AndAlso signature(2) = &H0 AndAlso signature(3) = &H0 AndAlso c >= 4 Then
                enc = Encoding.UTF32
                bytesPerSignature = 4
            ElseIf signature(0) = &H0 AndAlso signature(1) = &H0 AndAlso signature(2) = &HFE AndAlso signature(3) = &HFF Then
                enc = New UTF32Encoding(True, True)
                bytesPerSignature = 4
            ElseIf signature(0) = &HEF AndAlso signature(1) = &HBB AndAlso signature(2) = &HBF Then
                enc = Encoding.UTF8
                bytesPerSignature = 3
            ElseIf signature(0) = &HFE AndAlso signature(1) = &HFF Then
                enc = Encoding.BigEndianUnicode
                bytesPerSignature = 2
            ElseIf signature(0) = &HFF AndAlso signature(1) = &HFE Then
                enc = Encoding.Unicode
                bytesPerSignature = 2
            End If

            fs.Seek(bytesPerSignature, SeekOrigin.Begin)
            Return enc
        End Function

        Public Sub CloseFile()
            If fs IsNot Nothing Then

                Try
                    fs.Dispose()
                Catch
                End Try
            End If

            fs = Nothing
        End Sub

        Public Property SaveEOL As String

        Public Overrides Sub SaveToFile(ByVal fileName As String, ByVal enc As Encoding)
            Dim newLinePos = New List(Of Integer)(Count)
            Dim dir = Path.GetDirectoryName(fileName)
            Dim tempFileName = Path.Combine(dir, Path.GetFileNameWithoutExtension(fileName) & ".tmp")
            Dim sr As StreamReader = New StreamReader(fs, fileEncoding)

            Using tempFs As FileStream = New FileStream(tempFileName, FileMode.Create)

                Using sw As StreamWriter = New StreamWriter(tempFs, enc)
                    sw.Flush()

                    For i As Integer = 0 To Count - 1
                        newLinePos.Add(CInt(tempFs.Length))
                        Dim sourceLine = ReadLine(sr, i)
                        Dim line As String
                        Dim lineIsChanged As Boolean = lines(i) IsNot Nothing AndAlso lines(i).IsChanged

                        If lineIsChanged Then
                            line = lines(i).Text
                        Else
                            line = sourceLine
                        End If

                        If LinePushed IsNot Nothing Then
                            Dim args = New LinePushedEventArgs(sourceLine, i, If(lineIsChanged, line, Nothing))
                            LinePushed(Me, args)
                            If args.SavedText IsNot Nothing Then line = args.SavedText
                        End If

                        sw.Write(line)
                        If i < Count - 1 Then sw.Write(SaveEOL)
                        sw.Flush()
                    Next
                End Using
            End Using

            For i As Integer = 0 To Count - 1
                lines(i) = Nothing
            Next

            sr.Dispose()
            fs.Dispose()
            If File.Exists(fileName) Then File.Delete(fileName)
            File.Move(tempFileName, fileName)
            sourceFileLinePositions = newLinePos
            fs = New FileStream(fileName, FileMode.Open)
            Me.fileEncoding = enc
        End Sub

        Private Function ReadLine(ByVal sr As StreamReader, ByVal i As Integer) As String
            Dim line As String
            Dim filePos = sourceFileLinePositions(i)
            If filePos < 0 Then Return ""
            fs.Seek(filePos, SeekOrigin.Begin)
            sr.DiscardBufferedData()
            line = sr.ReadLine()
            Return line
        End Function

        Public Overrides Sub ClearIsChanged()
            For Each line In lines
                If line IsNot Nothing Then line.IsChanged = False
            Next
        End Sub

        Default Public Overrides Property Item(ByVal i As Integer) As Line
            Get

                If MyBase.lines(i) IsNot Nothing Then
                    Return lines(i)
                Else
                    LoadLineFromSourceFile(i)
                End If

                Return lines(i)
            End Get
            Set(ByVal value As Line)
                Throw New NotImplementedException()
            End Set
        End Property

        Private Sub LoadLineFromSourceFile(ByVal i As Integer)
            Dim line = CreateLine()
            fs.Seek(sourceFileLinePositions(i), SeekOrigin.Begin)
            Dim sr As StreamReader = New StreamReader(fs, fileEncoding)
            Dim s = sr.ReadLine()
            If s Is Nothing Then s = ""

            If LineNeeded IsNot Nothing Then
                Dim args = New LineNeededEventArgs(s, i)
                LineNeeded(Me, args)
                s = args.DisplayedLineText
                If s Is Nothing Then Return
            End If

            For Each c In s
                line.Add(New Char(c))
            Next

            MyBase.lines(i) = line
            If CurrentTB.WordWrap Then OnRecalcWordWrap(New TextChangedEventArgs(i, i))
        End Sub

        Public Overrides Sub InsertLine(ByVal index As Integer, ByVal line As Line)
            sourceFileLinePositions.Insert(index, -1)
            MyBase.InsertLine(index, line)
        End Sub

        Public Overrides Sub RemoveLine(ByVal index As Integer, ByVal count As Integer)
            sourceFileLinePositions.RemoveRange(index, count)
            MyBase.RemoveLine(index, count)
        End Sub

        Public Overrides Sub Clear()
            MyBase.Clear()
        End Sub

        Public Overrides Function GetLineLength(ByVal i As Integer) As Integer
            If MyBase.lines(i) Is Nothing Then
                Return 0
            Else
                Return MyBase.lines(i).Count
            End If
        End Function

        Public Overrides Function LineHasFoldingStartMarker(ByVal iLine As Integer) As Boolean
            If lines(iLine) Is Nothing Then
                Return False
            Else
                Return Not String.IsNullOrEmpty(lines(iLine).FoldingStartMarker)
            End If
        End Function

        Public Overrides Function LineHasFoldingEndMarker(ByVal iLine As Integer) As Boolean
            If lines(iLine) Is Nothing Then
                Return False
            Else
                Return Not String.IsNullOrEmpty(lines(iLine).FoldingEndMarker)
            End If
        End Function

        Public Overrides Sub Dispose()
            If fs IsNot Nothing Then fs.Dispose()
            timer.Dispose()
        End Sub

        Friend Sub UnloadLine(ByVal iLine As Integer)
            If lines(iLine) IsNot Nothing AndAlso Not lines(iLine).IsChanged Then lines(iLine) = Nothing
        End Sub
    End Class

    Public Class LineNeededEventArgs
        Inherits EventArgs

        Public Property SourceLineText As String
        Public Property DisplayedLineIndex As Integer
        Public Property DisplayedLineText As String

        Public Sub New(ByVal sourceLineText As String, ByVal displayedLineIndex As Integer)
            Me.SourceLineText = sourceLineText
            Me.DisplayedLineIndex = displayedLineIndex
            Me.DisplayedLineText = sourceLineText
        End Sub
    End Class

    Public Class LinePushedEventArgs
        Inherits EventArgs

        Public Property SourceLineText As String
        Public Property DisplayedLineIndex As Integer
        Public Property DisplayedLineText As String
        Public Property SavedText As String

        Public Sub New(ByVal sourceLineText As String, ByVal displayedLineIndex As Integer, ByVal displayedLineText As String)
            Me.SourceLineText = sourceLineText
            Me.DisplayedLineIndex = displayedLineIndex
            Me.DisplayedLineText = displayedLineText
            Me.SavedText = displayedLineText
        End Sub
    End Class

    Class CharReader
        Inherits TextReader

        Public Overrides Function Read() As Integer
            Return MyBase.Read()
        End Function
    End Class
End Namespace