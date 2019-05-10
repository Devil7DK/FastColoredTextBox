Namespace FastColoredTextBoxNS
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class HotkeysEditorForm
        Inherits System.Windows.Forms.Form

        Private components As System.ComponentModel.IContainer = Nothing

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

        Private Sub InitializeComponent()
            Dim dataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Me.dgv = New System.Windows.Forms.DataGridView()
            Me.cbModifiers = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.cbKey = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.cbAction = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.btAdd = New System.Windows.Forms.Button()
            Me.btRemove = New System.Windows.Forms.Button()
            Me.btCancel = New System.Windows.Forms.Button()
            Me.btOk = New System.Windows.Forms.Button()
            Me.label1 = New System.Windows.Forms.Label()
            Me.btResore = New System.Windows.Forms.Button()
            (CType((Me.dgv), System.ComponentModel.ISupportInitialize)).BeginInit()
            Me.SuspendLayout()
            Me.dgv.AllowUserToAddRows = False
            Me.dgv.AllowUserToDeleteRows = False
            Me.dgv.AllowUserToResizeColumns = False
            Me.dgv.AllowUserToResizeRows = False
            Me.dgv.Anchor = (CType(((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles))
            Me.dgv.BackgroundColor = System.Drawing.SystemColors.Control
            Me.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.cbModifiers, Me.cbKey, Me.cbAction})
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
            dataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte((204))))
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.LightSteelBlue
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
            Me.dgv.DefaultCellStyle = dataGridViewCellStyle1
            Me.dgv.GridColor = System.Drawing.SystemColors.Control
            Me.dgv.Location = New System.Drawing.Point(12, 28)
            Me.dgv.Name = "dgv"
            Me.dgv.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
            Me.dgv.RowHeadersVisible = False
            Me.dgv.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
            Me.dgv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            Me.dgv.Size = New System.Drawing.Size(525, 278)
            Me.dgv.TabIndex = 0
            Me.dgv.RowsAdded += New System.Windows.Forms.DataGridViewRowsAddedEventHandler(Me.dgv_RowsAdded)
            Me.cbModifiers.DataPropertyName = "Modifiers"
            Me.cbModifiers.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
            Me.cbModifiers.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.cbModifiers.HeaderText = "Modifiers"
            Me.cbModifiers.Name = "cbModifiers"
            Me.cbKey.DataPropertyName = "Key"
            Me.cbKey.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
            Me.cbKey.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.cbKey.HeaderText = "Key"
            Me.cbKey.Name = "cbKey"
            Me.cbKey.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
            Me.cbKey.Width = 120
            Me.cbAction.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            Me.cbAction.DataPropertyName = "Action"
            Me.cbAction.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
            Me.cbAction.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.cbAction.HeaderText = "Action"
            Me.cbAction.Name = "cbAction"
            Me.btAdd.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)), System.Windows.Forms.AnchorStyles))
            Me.btAdd.Location = New System.Drawing.Point(13, 322)
            Me.btAdd.Name = "btAdd"
            Me.btAdd.Size = New System.Drawing.Size(75, 23)
            Me.btAdd.TabIndex = 1
            Me.btAdd.Text = "Add"
            Me.btAdd.UseVisualStyleBackColor = True
            AddHandler Me.btAdd.Click, New System.EventHandler(Me.btAdd_Click)
            Me.btRemove.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)), System.Windows.Forms.AnchorStyles))
            Me.btRemove.Location = New System.Drawing.Point(103, 322)
            Me.btRemove.Name = "btRemove"
            Me.btRemove.Size = New System.Drawing.Size(75, 23)
            Me.btRemove.TabIndex = 2
            Me.btRemove.Text = "Remove"
            Me.btRemove.UseVisualStyleBackColor = True
            AddHandler Me.btRemove.Click, New System.EventHandler(Me.btRemove_Click)
            Me.btCancel.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles))
            Me.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btCancel.Location = New System.Drawing.Point(460, 322)
            Me.btCancel.Name = "btCancel"
            Me.btCancel.Size = New System.Drawing.Size(75, 23)
            Me.btCancel.TabIndex = 4
            Me.btCancel.Text = "Cancel"
            Me.btCancel.UseVisualStyleBackColor = True
            Me.btOk.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles))
            Me.btOk.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.btOk.Location = New System.Drawing.Point(379, 322)
            Me.btOk.Name = "btOk"
            Me.btOk.Size = New System.Drawing.Size(75, 23)
            Me.btOk.TabIndex = 3
            Me.btOk.Text = "OK"
            Me.btOk.UseVisualStyleBackColor = True
            Me.label1.AutoSize = True
            Me.label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte((204))))
            Me.label1.Location = New System.Drawing.Point(12, 9)
            Me.label1.Name = "label1"
            Me.label1.Size = New System.Drawing.Size(114, 16)
            Me.label1.TabIndex = 5
            Me.label1.Text = "Hotkeys mapping"
            Me.btResore.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)), System.Windows.Forms.AnchorStyles))
            Me.btResore.Location = New System.Drawing.Point(194, 322)
            Me.btResore.Name = "btResore"
            Me.btResore.Size = New System.Drawing.Size(105, 23)
            Me.btResore.TabIndex = 6
            Me.btResore.Text = "Restore default"
            Me.btResore.UseVisualStyleBackColor = True
            AddHandler Me.btResore.Click, New System.EventHandler(Me.btResore_Click)
            Me.AcceptButton = Me.btOk
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.btCancel
            Me.ClientSize = New System.Drawing.Size(549, 357)
            Me.Controls.Add(Me.btResore)
            Me.Controls.Add(Me.label1)
            Me.Controls.Add(Me.btCancel)
            Me.Controls.Add(Me.btOk)
            Me.Controls.Add(Me.btRemove)
            Me.Controls.Add(Me.btAdd)
            Me.Controls.Add(Me.dgv)
            Me.MaximumSize = New System.Drawing.Size(565, 700)
            Me.MinimumSize = New System.Drawing.Size(565, 395)
            Me.Name = "HotkeysEditorForm"
            Me.ShowIcon = False
            Me.Text = "Hotkeys Editor"
            Me.FormClosing += New System.Windows.Forms.FormClosingEventHandler(Me.HotkeysEditorForm_FormClosing)
            (CType((Me.dgv), System.ComponentModel.ISupportInitialize)).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()
        End Sub

        Private dgv As System.Windows.Forms.DataGridView
        Private btAdd As System.Windows.Forms.Button
        Private btRemove As System.Windows.Forms.Button
        Private btCancel As System.Windows.Forms.Button
        Private btOk As System.Windows.Forms.Button
        Private label1 As System.Windows.Forms.Label
        Private btResore As System.Windows.Forms.Button
        Private cbModifiers As System.Windows.Forms.DataGridViewComboBoxColumn
        Private cbKey As System.Windows.Forms.DataGridViewComboBoxColumn
        Private cbAction As System.Windows.Forms.DataGridViewComboBoxColumn
    End Class
End Namespace