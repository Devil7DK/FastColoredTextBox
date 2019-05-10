Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports System.Xml

Namespace FastColoredTextBoxNS
    Public Class MacrosManager
        Private ReadOnly macro As List(Of Object) = New List(Of Object)()

        Friend Sub New(ByVal ctrl As FastColoredTextBox)
            UnderlayingControl = ctrl
            AllowMacroRecordingByUser = True
        End Sub

        Public Property AllowMacroRecordingByUser As Boolean
        Private isRecording As Boolean

        Public Property IsRecording As Boolean
            Get
                Return IsRecording
            End Get
            Set(ByVal value As Boolean)
                isRecording = value
                UnderlayingControl.Invalidate()
            End Set
        End Property

        Public Property UnderlayingControl As FastColoredTextBox

        Public Sub ExecuteMacros()
            isRecording = False
            UnderlayingControl.BeginUpdate()
            UnderlayingControl.Selection.BeginUpdate()
            UnderlayingControl.BeginAutoUndo()

            For Each item In macro

                If TypeOf item Is Keys Then
                    UnderlayingControl.ProcessKey(CType(item, Keys))
                End If

                If TypeOf item Is KeyValuePair(Of Char, Keys) Then
                    Dim p = CType(item, KeyValuePair(Of Char, Keys))
                    UnderlayingControl.ProcessKey(p.Key, p.Value)
                End If
            Next

            UnderlayingControl.EndAutoUndo()
            UnderlayingControl.Selection.EndUpdate()
            UnderlayingControl.EndUpdate()
        End Sub

        Public Sub AddCharToMacros(ByVal c As Char, ByVal modifiers As Keys)
            macro.Add(New KeyValuePair(Of Char, Keys)(c, modifiers))
        End Sub

        Public Sub AddKeyToMacros(ByVal keyData As Keys)
            macro.Add(keyData)
        End Sub

        Public Sub ClearMacros()
            macro.Clear()
        End Sub

        Friend Sub ProcessKey(ByVal keyData As Keys)
            If isRecording Then AddKeyToMacros(keyData)
        End Sub

        Friend Sub ProcessKey(ByVal c As Char, ByVal modifiers As Keys)
            If isRecording Then AddCharToMacros(c, modifiers)
        End Sub

        Public ReadOnly Property MacroIsEmpty As Boolean
            Get
                Return macro.Count = 0
            End Get
        End Property

        Public Property Macros As String
            Get
                Dim cult = Thread.CurrentThread.CurrentUICulture
                Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture
                Dim kc = New KeysConverter()
                Dim sb As StringBuilder = New StringBuilder()
                sb.AppendLine("<macros>")

                For Each item In macro

                    If TypeOf item Is Keys Then
                        sb.AppendFormat("<item key='{0}' />" & vbCrLf, kc.ConvertToString(CType(item, Keys)))
                    ElseIf TypeOf item Is KeyValuePair(Of Char, Keys) Then
                        Dim p = CType(item, KeyValuePair(Of Char, Keys))
                        sb.AppendFormat("<item char='{0}' key='{1}' />" & vbCrLf, CInt(p.Key), kc.ConvertToString(p.Value))
                    End If
                Next

                sb.AppendLine("</macros>")
                Thread.CurrentThread.CurrentUICulture = cult
                Return sb.ToString()
            End Get
            Set(ByVal value As String)
                isRecording = False
                ClearMacros()
                If String.IsNullOrEmpty(value) Then Return
                Dim doc = New XmlDocument()
                doc.LoadXml(value)
                Dim list = doc.SelectNodes("./macros/item")
                Dim cult = Thread.CurrentThread.CurrentUICulture
                Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture
                Dim kc = New KeysConverter()

                If list IsNot Nothing Then

                    For Each node As XmlElement In list
                        Dim ca = node.GetAttributeNode("char")
                        Dim ka = node.GetAttributeNode("key")

                        If ca IsNot Nothing Then

                            If ka IsNot Nothing Then
                                AddCharToMacros(ChrW(Integer.Parse(ca.Value)), CType(kc.ConvertFromString(ka.Value), Keys))
                            Else
                                AddCharToMacros(ChrW(Integer.Parse(ca.Value)), Keys.None)
                            End If
                        ElseIf ka IsNot Nothing Then
                            AddKeyToMacros(CType(kc.ConvertFromString(ka.Value), Keys))
                        End If
                    Next
                End If

                Thread.CurrentThread.CurrentUICulture = cult
            End Set
        End Property
    End Class
End Namespace