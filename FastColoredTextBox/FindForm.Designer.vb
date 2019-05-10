Namespace FastColoredTextBoxNS
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class FindForm
        Inherits Windows.Forms.Form

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
            Me.SuspendLayout()
            Me.btClose.Location = New System.Drawing.Point(273, 73)
            Me.btClose.Name = "btClose"
            Me.btClose.Size = New System.Drawing.Size(75, 23)
            Me.btClose.TabIndex = 5
            Me.btClose.Text = "Close"
            Me.btClose.UseVisualStyleBackColor = True
            AddHandler Me.btClose.Click, New System.EventHandler(Me.btClose_Click)
            Me.btFindNext.Location = New System.Drawing.Point(192, 73)
            Me.btFindNext.Name = "btFindNext"
            Me.btFindNext.Size = New System.Drawing.Size(75, 23)
            Me.btFindNext.TabIndex = 4
            Me.btFindNext.Text = "Find next"
            Me.btFindNext.UseVisualStyleBackColor = True
            AddHandler Me.btFindNext.Click, New System.EventHandler(Me.btFindNext_Click)
            Me.tbFind.Location = New System.Drawing.Point(42, 12)
            Me.tbFind.Name = "tbFind"
            Me.tbFind.Size = New System.Drawing.Size(306, 20)
            Me.tbFind.TabIndex = 0
            AddHandler Me.tbFind.TextChanged, New System.EventHandler(Me.cbMatchCase_CheckedChanged)
            Me.tbFind.KeyPress += New System.Windows.Forms.KeyPressEventHandler(Me.tbFind_KeyPress)
            Me.cbRegex.AutoSize = True
            Me.cbRegex.Location = New System.Drawing.Point(249, 38)
            Me.cbRegex.Name = "cbRegex"
            Me.cbRegex.Size = New System.Drawing.Size(57, 17)
            Me.cbRegex.TabIndex = 3
            Me.cbRegex.Text = "Regex"
            Me.cbRegex.UseVisualStyleBackColor = True
            AddHandler Me.cbRegex.CheckedChanged, New System.EventHandler(Me.cbMatchCase_CheckedChanged)
            Me.cbMatchCase.AutoSize = True
            Me.cbMatchCase.Location = New System.Drawing.Point(42, 38)
            Me.cbMatchCase.Name = "cbMatchCase"
            Me.cbMatchCase.Size = New System.Drawing.Size(82, 17)
            Me.cbMatchCase.TabIndex = 1
            Me.cbMatchCase.Text = "Match case"
            Me.cbMatchCase.UseVisualStyleBackColor = True
            AddHandler Me.cbMatchCase.CheckedChanged, New System.EventHandler(Me.cbMatchCase_CheckedChanged)
            Me.label1.AutoSize = True
            Me.label1.Location = New System.Drawing.Point(6, 15)
            Me.label1.Name = "label1"
            Me.label1.Size = New System.Drawing.Size(33, 13)
            Me.label1.TabIndex = 5
            Me.label1.Text = "Find: "
            Me.cbWholeWord.AutoSize = True
            Me.cbWholeWord.Location = New System.Drawing.Point(130, 38)
            Me.cbWholeWord.Name = "cbWholeWord"
            Me.cbWholeWord.Size = New System.Drawing.Size(113, 17)
            Me.cbWholeWord.TabIndex = 2
            Me.cbWholeWord.Text = "Match whole word"
            Me.cbWholeWord.UseVisualStyleBackColor = True
            AddHandler Me.cbWholeWord.CheckedChanged, New System.EventHandler(Me.cbMatchCase_CheckedChanged)
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(360, 108)
            Me.Controls.Add(Me.cbWholeWord)
            Me.Controls.Add(Me.label1)
            Me.Controls.Add(Me.cbMatchCase)
            Me.Controls.Add(Me.cbRegex)
            Me.Controls.Add(Me.tbFind)
            Me.Controls.Add(Me.btFindNext)
            Me.Controls.Add(Me.btClose)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
            Me.Name = "FindForm"
            Me.ShowIcon = False
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Find"
            Me.TopMost = True
            Me.FormClosing += New System.Windows.Forms.FormClosingEventHandler(Me.FindForm_FormClosing)
            Me.ResumeLayout(False)
            Me.PerformLayout()
        End Sub

        Private btClose As System.Windows.Forms.Button
        Private btFindNext As System.Windows.Forms.Button
        Private cbRegex As System.Windows.Forms.CheckBox
        Private cbMatchCase As System.Windows.Forms.CheckBox
        Private label1 As System.Windows.Forms.Label
        Private cbWholeWord As System.Windows.Forms.CheckBox
        Public tbFind As System.Windows.Forms.TextBox
    End Class
End Namespace