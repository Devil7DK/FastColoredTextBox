Namespace FastColoredTextBoxNS
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Public Class GoToForm
        Inherits System.Windows.Forms.Form

        Private components As System.ComponentModel.IContainer = Nothing

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

        Private Sub InitializeComponent()
            Me.label = New System.Windows.Forms.Label()
            Me.tbLineNumber = New System.Windows.Forms.TextBox()
            Me.btnOk = New System.Windows.Forms.Button()
            Me.btnCancel = New System.Windows.Forms.Button()
            Me.SuspendLayout()
            Me.label.AutoSize = True
            Me.label.Location = New System.Drawing.Point(12, 9)
            Me.label.Name = "label"
            Me.label.Size = New System.Drawing.Size(96, 13)
            Me.label.TabIndex = 0
            Me.label.Text = "Line Number (1/1):"
            Me.tbLineNumber.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles))
            Me.tbLineNumber.Location = New System.Drawing.Point(12, 29)
            Me.tbLineNumber.Name = "tbLineNumber"
            Me.tbLineNumber.Size = New System.Drawing.Size(296, 20)
            Me.tbLineNumber.TabIndex = 1
            Me.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.btnOk.Location = New System.Drawing.Point(152, 71)
            Me.btnOk.Name = "btnOk"
            Me.btnOk.Size = New System.Drawing.Size(75, 23)
            Me.btnOk.TabIndex = 2
            Me.btnOk.Text = "OK"
            Me.btnOk.UseVisualStyleBackColor = True
            AddHandler Me.btnOk.Click, New System.EventHandler(Me.btnOk_Click)
            Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnCancel.Location = New System.Drawing.Point(233, 71)
            Me.btnCancel.Name = "btnCancel"
            Me.btnCancel.Size = New System.Drawing.Size(75, 23)
            Me.btnCancel.TabIndex = 3
            Me.btnCancel.Text = "Cancel"
            Me.btnCancel.UseVisualStyleBackColor = True
            AddHandler Me.btnCancel.Click, New System.EventHandler(Me.btnCancel_Click)
            Me.AcceptButton = Me.btnOk
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.btnCancel
            Me.ClientSize = New System.Drawing.Size(320, 106)
            Me.Controls.Add(Me.btnCancel)
            Me.Controls.Add(Me.btnOk)
            Me.Controls.Add(Me.tbLineNumber)
            Me.Controls.Add(Me.label)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "GoToForm"
            Me.ShowIcon = False
            Me.ShowInTaskbar = False
            Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "Go To Line"
            Me.TopMost = True
            Me.ResumeLayout(False)
            Me.PerformLayout()
        End Sub

        Private label As System.Windows.Forms.Label
        Private tbLineNumber As System.Windows.Forms.TextBox
        Private btnOk As System.Windows.Forms.Button
        Private btnCancel As System.Windows.Forms.Button
    End Class
End Namespace