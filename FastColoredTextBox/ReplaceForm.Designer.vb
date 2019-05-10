Namespace FastColoredTextBoxNS
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class ReplaceForm
        Inherits System.Windows.Forms.Form

        Private components As System.ComponentModel.IContainer = Nothing

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

        Private Sub InitializeComponent()
            Me.btClose = New System.Windows.Forms.Button()
            Me.btFindNext = New System.Windows.Forms.Button()
            Me.tbFind = New System.Windows.Forms.TextBox()
            Me.cbRegex = New System.Windows.Forms.CheckBox()
            Me.cbMatchCase = New System.Windows.Forms.CheckBox()
            Me.label1 = New System.Windows.Forms.Label()
            Me.cbWholeWord = New System.Windows.Forms.CheckBox()
            Me.btReplace = New System.Windows.Forms.Button()
            Me.btReplaceAll = New System.Windows.Forms.Button()
            Me.label2 = New System.Windows.Forms.Label()
            Me.tbReplace = New System.Windows.Forms.TextBox()
            Me.SuspendLayout()
            Me.btClose.Location = New System.Drawing.Point(273, 153)
            Me.btClose.Name = "btClose"
            Me.btClose.Size = New System.Drawing.Size(75, 23)
            Me.btClose.TabIndex = 8
            Me.btClose.Text = "Close"
            Me.btClose.UseVisualStyleBackColor = True
            AddHandler Me.btClose.Click, New System.EventHandler(Me.btClose_Click)
            Me.btFindNext.Location = New System.Drawing.Point(111, 124)
            Me.btFindNext.Name = "btFindNext"
            Me.btFindNext.Size = New System.Drawing.Size(75, 23)
            Me.btFindNext.TabIndex = 5
            Me.btFindNext.Text = "Find next"
            Me.btFindNext.UseVisualStyleBackColor = True
            AddHandler Me.btFindNext.Click, New System.EventHandler(Me.btFindNext_Click)
            Me.tbFind.Location = New System.Drawing.Point(62, 12)
            Me.tbFind.Name = "tbFind"
            Me.tbFind.Size = New System.Drawing.Size(286, 20)
            Me.tbFind.TabIndex = 0
            AddHandler Me.tbFind.TextChanged, New System.EventHandler(Me.cbMatchCase_CheckedChanged)
            Me.tbFind.KeyPress += New System.Windows.Forms.KeyPressEventHandler(Me.tbFind_KeyPress)
            Me.cbRegex.AutoSize = True
            Me.cbRegex.Location = New System.Drawing.Point(273, 38)
            Me.cbRegex.Name = "cbRegex"
            Me.cbRegex.Size = New System.Drawing.Size(57, 17)
            Me.cbRegex.TabIndex = 3
            Me.cbRegex.Text = "Regex"
            Me.cbRegex.UseVisualStyleBackColor = True
            AddHandler Me.cbRegex.CheckedChanged, New System.EventHandler(Me.cbMatchCase_CheckedChanged)
            Me.cbMatchCase.AutoSize = True
            Me.cbMatchCase.Location = New System.Drawing.Point(66, 38)
            Me.cbMatchCase.Name = "cbMatchCase"
            Me.cbMatchCase.Size = New System.Drawing.Size(82, 17)
            Me.cbMatchCase.TabIndex = 1
            Me.cbMatchCase.Text = "Match case"
            Me.cbMatchCase.UseVisualStyleBackColor = True
            AddHandler Me.cbMatchCase.CheckedChanged, New System.EventHandler(Me.cbMatchCase_CheckedChanged)
            Me.label1.AutoSize = True
            Me.label1.Location = New System.Drawing.Point(23, 14)
            Me.label1.Name = "label1"
            Me.label1.Size = New System.Drawing.Size(33, 13)
            Me.label1.TabIndex = 5
            Me.label1.Text = "Find: "
            Me.cbWholeWord.AutoSize = True
            Me.cbWholeWord.Location = New System.Drawing.Point(154, 38)
            Me.cbWholeWord.Name = "cbWholeWord"
            Me.cbWholeWord.Size = New System.Drawing.Size(113, 17)
            Me.cbWholeWord.TabIndex = 2
            Me.cbWholeWord.Text = "Match whole word"
            Me.cbWholeWord.UseVisualStyleBackColor = True
            AddHandler Me.cbWholeWord.CheckedChanged, New System.EventHandler(Me.cbMatchCase_CheckedChanged)
            Me.btReplace.Location = New System.Drawing.Point(192, 124)
            Me.btReplace.Name = "btReplace"
            Me.btReplace.Size = New System.Drawing.Size(75, 23)
            Me.btReplace.TabIndex = 6
            Me.btReplace.Text = "Replace"
            Me.btReplace.UseVisualStyleBackColor = True
            AddHandler Me.btReplace.Click, New System.EventHandler(Me.btReplace_Click)
            Me.btReplaceAll.Location = New System.Drawing.Point(273, 124)
            Me.btReplaceAll.Name = "btReplaceAll"
            Me.btReplaceAll.Size = New System.Drawing.Size(75, 23)
            Me.btReplaceAll.TabIndex = 7
            Me.btReplaceAll.Text = "Replace all"
            Me.btReplaceAll.UseVisualStyleBackColor = True
            AddHandler Me.btReplaceAll.Click, New System.EventHandler(Me.btReplaceAll_Click)
            Me.label2.AutoSize = True
            Me.label2.Location = New System.Drawing.Point(6, 81)
            Me.label2.Name = "label2"
            Me.label2.Size = New System.Drawing.Size(50, 13)
            Me.label2.TabIndex = 9
            Me.label2.Text = "Replace:"
            Me.tbReplace.Location = New System.Drawing.Point(62, 78)
            Me.tbReplace.Name = "tbReplace"
            Me.tbReplace.Size = New System.Drawing.Size(286, 20)
            Me.tbReplace.TabIndex = 0
            AddHandler Me.tbReplace.TextChanged, New System.EventHandler(Me.cbMatchCase_CheckedChanged)
            Me.tbReplace.KeyPress += New System.Windows.Forms.KeyPressEventHandler(Me.tbFind_KeyPress)
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(360, 191)
            Me.Controls.Add(Me.tbFind)
            Me.Controls.Add(Me.label2)
            Me.Controls.Add(Me.tbReplace)
            Me.Controls.Add(Me.btReplaceAll)
            Me.Controls.Add(Me.btReplace)
            Me.Controls.Add(Me.cbWholeWord)
            Me.Controls.Add(Me.label1)
            Me.Controls.Add(Me.cbMatchCase)
            Me.Controls.Add(Me.cbRegex)
            Me.Controls.Add(Me.btFindNext)
            Me.Controls.Add(Me.btClose)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
            Me.Name = "ReplaceForm"
            Me.ShowIcon = False
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Find and replace"
            Me.TopMost = True
            Me.FormClosing += New System.Windows.Forms.FormClosingEventHandler(Me.ReplaceForm_FormClosing)
            Me.ResumeLayout(False)
            Me.PerformLayout()
        End Sub

        Private btClose As System.Windows.Forms.Button
        Private btFindNext As System.Windows.Forms.Button
        Private cbRegex As System.Windows.Forms.CheckBox
        Private cbMatchCase As System.Windows.Forms.CheckBox
        Private label1 As System.Windows.Forms.Label
        Private cbWholeWord As System.Windows.Forms.CheckBox
        Private btReplace As System.Windows.Forms.Button
        Private btReplaceAll As System.Windows.Forms.Button
        Private label2 As System.Windows.Forms.Label
        Public tbFind As System.Windows.Forms.TextBox
        Public tbReplace As System.Windows.Forms.TextBox
    End Class
End Namespace