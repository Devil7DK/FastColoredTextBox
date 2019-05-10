Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Namespace FastColoredTextBoxNS
    Module EncodingDetector
        Const _defaultHeuristicSampleSize As Long = &H10000

        Function DetectTextFileEncoding(ByVal InputFilename As String) As Encoding
            Using textfileStream As FileStream = File.OpenRead(InputFilename)
                Return DetectTextFileEncoding(textfileStream, _defaultHeuristicSampleSize)
            End Using
        End Function

        Function DetectTextFileEncoding(ByVal InputFileStream As FileStream, ByVal HeuristicSampleSize As Long) As Encoding
            Dim uselessBool As Boolean = False
            Return DetectTextFileEncoding(InputFileStream, _defaultHeuristicSampleSize, uselessBool)
        End Function

        Function DetectTextFileEncoding(ByVal InputFileStream As FileStream, ByVal HeuristicSampleSize As Long, <Out> ByRef HasBOM As Boolean) As Encoding
            Dim encodingFound As Encoding = Nothing
            Dim originalPos As Long = InputFileStream.Position
            InputFileStream.Position = 0
            Dim bomBytes As Byte() = New Byte(If(InputFileStream.Length > 4, 4, InputFileStream.Length) - 1) {}
            InputFileStream.Read(bomBytes, 0, bomBytes.Length)
            encodingFound = DetectBOMBytes(bomBytes)

            If encodingFound IsNot Nothing Then
                InputFileStream.Position = originalPos
                HasBOM = True
                Return encodingFound
            End If

            Dim sampleBytes As Byte() = New Byte(If(HeuristicSampleSize > InputFileStream.Length, InputFileStream.Length, HeuristicSampleSize) - 1) {}
            Array.Copy(bomBytes, sampleBytes, bomBytes.Length)
            If InputFileStream.Length > bomBytes.Length Then InputFileStream.Read(sampleBytes, bomBytes.Length, sampleBytes.Length - bomBytes.Length)
            InputFileStream.Position = originalPos
            encodingFound = DetectUnicodeInByteSampleByHeuristics(sampleBytes)
            HasBOM = False
            Return encodingFound
        End Function

        Function DetectBOMBytes(ByVal BOMBytes As Byte()) As Encoding
            If BOMBytes.Length < 2 Then Return Nothing
            If BOMBytes(0) = &HFF AndAlso BOMBytes(1) = &HFE AndAlso (BOMBytes.Length < 4 OrElse BOMBytes(2) <> 0 OrElse BOMBytes(3) <> 0) Then Return Encoding.Unicode
            If BOMBytes(0) = &HFE AndAlso BOMBytes(1) = &HFF Then Return Encoding.BigEndianUnicode
            If BOMBytes.Length < 3 Then Return Nothing
            If BOMBytes(0) = &HEF AndAlso BOMBytes(1) = &HBB AndAlso BOMBytes(2) = &HBF Then Return Encoding.UTF8
            If BOMBytes(0) = &H2B AndAlso BOMBytes(1) = &H2F AndAlso BOMBytes(2) = &H76 Then Return Encoding.UTF7
            If BOMBytes.Length < 4 Then Return Nothing
            If BOMBytes(0) = &HFF AndAlso BOMBytes(1) = &HFE AndAlso BOMBytes(2) = 0 AndAlso BOMBytes(3) = 0 Then Return Encoding.UTF32
            If BOMBytes(0) = 0 AndAlso BOMBytes(1) = 0 AndAlso BOMBytes(2) = &HFE AndAlso BOMBytes(3) = &HFF Then Return Encoding.GetEncoding(12001)
            Return Nothing
        End Function

        Function DetectUnicodeInByteSampleByHeuristics(ByVal SampleBytes As Byte()) As Encoding
            Dim oddBinaryNullsInSample As Long = 0
            Dim evenBinaryNullsInSample As Long = 0
            Dim suspiciousUTF8SequenceCount As Long = 0
            Dim suspiciousUTF8BytesTotal As Long = 0
            Dim likelyUSASCIIBytesInSample As Long = 0
            Dim currentPos As Long = 0
            Dim skipUTF8Bytes As Integer = 0

            While currentPos < SampleBytes.Length

                If SampleBytes(currentPos) = 0 Then

                    If currentPos Mod 2 = 0 Then
                        evenBinaryNullsInSample += 1
                    Else
                        oddBinaryNullsInSample += 1
                    End If
                End If

                If IsCommonUSASCIIByte(SampleBytes(currentPos)) Then likelyUSASCIIBytesInSample += 1

                If skipUTF8Bytes = 0 Then
                    Dim lengthFound As Integer = DetectSuspiciousUTF8SequenceLength(SampleBytes, currentPos)

                    If lengthFound > 0 Then
                        suspiciousUTF8SequenceCount += 1
                        suspiciousUTF8BytesTotal += lengthFound
                        skipUTF8Bytes = lengthFound - 1
                    End If
                Else
                    skipUTF8Bytes -= 1
                End If

                currentPos += 1
            End While

            If ((evenBinaryNullsInSample * 2.0) / SampleBytes.Length) < 0.2 AndAlso ((oddBinaryNullsInSample * 2.0) / SampleBytes.Length) > 0.6 Then Return Encoding.Unicode
            If ((oddBinaryNullsInSample * 2.0) / SampleBytes.Length) < 0.2 AndAlso ((evenBinaryNullsInSample * 2.0) / SampleBytes.Length) > 0.6 Then Return Encoding.BigEndianUnicode
            Dim potentiallyMangledString As String = Encoding.ASCII.GetString(SampleBytes)
            Dim UTF8Validator As Regex = New Regex("\A(" & "[\x09\x0A\x0D\x20-\x7E]" & "|[\xC2-\xDF][\x80-\xBF]" & "|\xE0[\xA0-\xBF][\x80-\xBF]" & "|[\xE1-\xEC\xEE\xEF][\x80-\xBF]{2}" & "|\xED[\x80-\x9F][\x80-\xBF]" & "|\xF0[\x90-\xBF][\x80-\xBF]{2}" & "|[\xF1-\xF3][\x80-\xBF]{3}" & "|\xF4[\x80-\x8F][\x80-\xBF]{2}" & ")*\z")

            If UTF8Validator.IsMatch(potentiallyMangledString) Then
                If (suspiciousUTF8SequenceCount * 500000.0 / SampleBytes.Length >= 1) AndAlso (SampleBytes.Length - suspiciousUTF8BytesTotal = 0 OrElse likelyUSASCIIBytesInSample * 1.0 / (SampleBytes.Length - suspiciousUTF8BytesTotal) >= 0.8) Then Return Encoding.UTF8
            End If

            Return Nothing
        End Function

        Private Function IsCommonUSASCIIByte(ByVal testByte As Byte) As Boolean
            If testByte = &HA OrElse testByte = &HD OrElse testByte = &H9 OrElse (testByte >= &H20 AndAlso testByte <= &H2F) OrElse (testByte >= &H30 AndAlso testByte <= &H39) OrElse (testByte >= &H3A AndAlso testByte <= &H40) OrElse (testByte >= &H41 AndAlso testByte <= &H5A) OrElse (testByte >= &H5B AndAlso testByte <= &H60) OrElse (testByte >= &H61 AndAlso testByte <= &H7A) OrElse (testByte >= &H7B AndAlso testByte <= &H7E) Then
                Return True
            Else
                Return False
            End If
        End Function

        Private Function DetectSuspiciousUTF8SequenceLength(ByVal SampleBytes As Byte(), ByVal currentPos As Long) As Integer
            Dim lengthFound As Integer = 0

            If SampleBytes.Length >= currentPos + 1 AndAlso SampleBytes(currentPos) = &HC2 Then

                If SampleBytes(currentPos + 1) = &H81 OrElse SampleBytes(currentPos + 1) = &H8D OrElse SampleBytes(currentPos + 1) = &H8F Then
                    lengthFound = 2
                ElseIf SampleBytes(currentPos + 1) = &H90 OrElse SampleBytes(currentPos + 1) = &H9D Then
                    lengthFound = 2
                ElseIf SampleBytes(currentPos + 1) >= &HA0 AndAlso SampleBytes(currentPos + 1) <= &HBF Then
                    lengthFound = 2
                End If
            ElseIf SampleBytes.Length >= currentPos + 1 AndAlso SampleBytes(currentPos) = &HC3 Then
                If SampleBytes(currentPos + 1) >= &H80 AndAlso SampleBytes(currentPos + 1) <= &HBF Then lengthFound = 2
            ElseIf SampleBytes.Length >= currentPos + 1 AndAlso SampleBytes(currentPos) = &HC5 Then

                If SampleBytes(currentPos + 1) = &H92 OrElse SampleBytes(currentPos + 1) = &H93 Then
                    lengthFound = 2
                ElseIf SampleBytes(currentPos + 1) = &HA0 OrElse SampleBytes(currentPos + 1) = &HA1 Then
                    lengthFound = 2
                ElseIf SampleBytes(currentPos + 1) = &HB8 OrElse SampleBytes(currentPos + 1) = &HBD OrElse SampleBytes(currentPos + 1) = &HBE Then
                    lengthFound = 2
                End If
            ElseIf SampleBytes.Length >= currentPos + 1 AndAlso SampleBytes(currentPos) = &HC6 Then
                If SampleBytes(currentPos + 1) = &H92 Then lengthFound = 2
            ElseIf SampleBytes.Length >= currentPos + 1 AndAlso SampleBytes(currentPos) = &HCB Then
                If SampleBytes(currentPos + 1) = &H86 OrElse SampleBytes(currentPos + 1) = &H9C Then lengthFound = 2
            ElseIf SampleBytes.Length >= currentPos + 2 AndAlso SampleBytes(currentPos) = &HE2 Then

                If SampleBytes(currentPos + 1) = &H80 Then
                    If SampleBytes(currentPos + 2) = &H93 OrElse SampleBytes(currentPos + 2) = &H94 Then lengthFound = 3
                    If SampleBytes(currentPos + 2) = &H98 OrElse SampleBytes(currentPos + 2) = &H99 OrElse SampleBytes(currentPos + 2) = &H9A Then lengthFound = 3
                    If SampleBytes(currentPos + 2) = &H9C OrElse SampleBytes(currentPos + 2) = &H9D OrElse SampleBytes(currentPos + 2) = &H9E Then lengthFound = 3
                    If SampleBytes(currentPos + 2) = &HA0 OrElse SampleBytes(currentPos + 2) = &HA1 OrElse SampleBytes(currentPos + 2) = &HA2 Then lengthFound = 3
                    If SampleBytes(currentPos + 2) = &HA6 Then lengthFound = 3
                    If SampleBytes(currentPos + 2) = &HB0 Then lengthFound = 3
                    If SampleBytes(currentPos + 2) = &HB9 OrElse SampleBytes(currentPos + 2) = &HBA Then lengthFound = 3
                ElseIf SampleBytes(currentPos + 1) = &H82 AndAlso SampleBytes(currentPos + 2) = &HAC Then
                    lengthFound = 3
                ElseIf SampleBytes(currentPos + 1) = &H84 AndAlso SampleBytes(currentPos + 2) = &HA2 Then
                    lengthFound = 3
                End If
            End If

            Return lengthFound
        End Function
    End Module
End Namespace