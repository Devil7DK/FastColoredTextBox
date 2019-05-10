Imports System.Collections.Generic
Imports System

Namespace FastColoredTextBoxNS
    Public Class CommandManager
        Public Shared MaxHistoryLength As Integer = 200
        Private history As LimitedStack(Of UndoableCommand)
        Private redoStack As Stack(Of UndoableCommand) = New Stack(Of UndoableCommand)()
        Public Property TextSource As TextSource
        Public Property UndoRedoStackIsEnabled As Boolean
        Public Event RedoCompleted As EventHandler

        Public Sub New(ByVal ts As TextSource)
            history = New LimitedStack(Of UndoableCommand)(MaxHistoryLength)
            TextSource = ts
            UndoRedoStackIsEnabled = True
        End Sub

        Public Overridable Sub ExecuteCommand(ByVal cmd As Command)
            If disabledCommands > 0 Then Return

            If cmd.ts.CurrentTB.Selection.ColumnSelectionMode Then
                If TypeOf cmd Is UndoableCommand Then cmd = New MultiRangeCommand(CType(cmd, UndoableCommand))
            End If

            If TypeOf cmd Is UndoableCommand Then
                (TryCast(cmd, UndoableCommand)).autoUndo = autoUndoCommands > 0
                history.Push(TryCast(cmd, UndoableCommand))
            End If

            Try
                cmd.Execute()
            Catch __unusedArgumentOutOfRangeException1__ As ArgumentOutOfRangeException
                If TypeOf cmd Is UndoableCommand Then history.Pop()
            End Try

            If Not UndoRedoStackIsEnabled Then ClearHistory()
            redoStack.Clear()
            TextSource.CurrentTB.OnUndoRedoStateChanged()
        End Sub

        Public Sub Undo()
            If history.Count > 0 Then
                Dim cmd = history.Pop()
                BeginDisableCommands()

                Try
                    cmd.Undo()
                Finally
                    EndDisableCommands()
                End Try

                redoStack.Push(cmd)
            End If

            If history.Count > 0 Then
                If history.Peek().autoUndo Then Undo()
            End If

            TextSource.CurrentTB.OnUndoRedoStateChanged()
        End Sub

        Protected disabledCommands As Integer = 0

        Private Sub EndDisableCommands()
            disabledCommands -= 1
        End Sub

        Private Sub BeginDisableCommands()
            disabledCommands += 1
        End Sub

        Private autoUndoCommands As Integer = 0

        Public Sub EndAutoUndoCommands()
            autoUndoCommands -= 1

            If autoUndoCommands = 0 Then
                If history.Count > 0 Then history.Peek().autoUndo = False
            End If
        End Sub

        Public Sub BeginAutoUndoCommands()
            autoUndoCommands += 1
        End Sub

        Friend Sub ClearHistory()
            history.Clear()
            redoStack.Clear()
            TextSource.CurrentTB.OnUndoRedoStateChanged()
        End Sub

        Friend Sub Redo()
            If redoStack.Count = 0 Then Return
            Dim cmd As UndoableCommand
            BeginDisableCommands()

            Try
                cmd = redoStack.Pop()
                If TextSource.CurrentTB.Selection.ColumnSelectionMode Then TextSource.CurrentTB.Selection.ColumnSelectionMode = False
                TextSource.CurrentTB.Selection.Start = cmd.sel.Start
                TextSource.CurrentTB.Selection.[End] = cmd.sel.[End]
                cmd.Execute()
                history.Push(cmd)
            Finally
                EndDisableCommands()
            End Try

            RedoCompleted(Me, EventArgs.Empty)
            If cmd.autoUndo Then Redo()
            TextSource.CurrentTB.OnUndoRedoStateChanged()
        End Sub

        Public ReadOnly Property UndoEnabled As Boolean
            Get
                Return history.Count > 0
            End Get
        End Property

        Public ReadOnly Property RedoEnabled As Boolean
            Get
                Return redoStack.Count > 0
            End Get
        End Property
    End Class

    Public MustInherit Class Command
        Public ts As TextSource
        Public MustOverride Sub Execute()
    End Class

    Friend Class RangeInfo
        Public Property Start As Place
        Public Property [End] As Place

        Public Sub New(ByVal r As Range)
            Start = r.Start
            [End] = r.[End]
        End Sub

        Friend ReadOnly Property FromX As Integer
            Get
                If [End].iLine < Start.iLine Then Return [End].iChar
                If [End].iLine > Start.iLine Then Return Start.iChar
                Return Math.Min([End].iChar, Start.iChar)
            End Get
        End Property
    End Class

    Public MustInherit Class UndoableCommand
        Inherits Command

        Friend sel As RangeInfo
        Friend lastSel As RangeInfo
        Friend autoUndo As Boolean

        Public Sub New(ByVal ts As TextSource)
            Me.ts = ts
            sel = New RangeInfo(ts.CurrentTB.Selection)
        End Sub

        Public Overridable Sub Undo()
            OnTextChanged(True)
        End Sub

        Public Overrides Sub Execute()
            lastSel = New RangeInfo(ts.CurrentTB.Selection)
            OnTextChanged(False)
        End Sub

        Protected Overridable Sub OnTextChanged(ByVal invert As Boolean)
            Dim b As Boolean = sel.Start.iLine < lastSel.Start.iLine

            If invert Then

                If b Then
                    ts.OnTextChanged(sel.Start.iLine, sel.Start.iLine)
                Else
                    ts.OnTextChanged(sel.Start.iLine, lastSel.Start.iLine)
                End If
            Else

                If b Then
                    ts.OnTextChanged(sel.Start.iLine, lastSel.Start.iLine)
                Else
                    ts.OnTextChanged(lastSel.Start.iLine, lastSel.Start.iLine)
                End If
            End If
        End Sub

        Public MustOverride Function Clone() As UndoableCommand
    End Class
End Namespace