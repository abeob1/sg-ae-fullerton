<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSendSMS
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblType = New System.Windows.Forms.Label()
        Me.lblPaymentType = New System.Windows.Forms.Label()
        Me.txtBatchNo = New System.Windows.Forms.TextBox()
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.cmbPaymentType = New System.Windows.Forms.ComboBox()
        Me.tbpnButton = New System.Windows.Forms.TableLayoutPanel()
        Me.btnSendSMS = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lblBatchNo = New System.Windows.Forms.Label()
        Me.rtxtMessage = New System.Windows.Forms.RichTextBox()
        Me.Panel4.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.tbpnButton.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(15, 410)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel2.Location = New System.Drawing.Point(542, 0)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(15, 410)
        Me.Panel2.TabIndex = 1
        '
        'Panel3
        '
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(15, 0)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(527, 12)
        Me.Panel3.TabIndex = 2
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(15, 12)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(527, 398)
        Me.Panel4.TabIndex = 3
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 5
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.62887!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 29.69072!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 42.68041!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.lblType, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.lblPaymentType, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtBatchNo, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbType, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbPaymentType, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.tbpnButton, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.lblBatchNo, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.rtxtMessage, 1, 7)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 15
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(527, 398)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblType.Location = New System.Drawing.Point(14, 41)
        Me.lblType.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(126, 31)
        Me.lblType.TabIndex = 1
        Me.lblType.Text = "Type"
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPaymentType
        '
        Me.lblPaymentType.AutoSize = True
        Me.lblPaymentType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPaymentType.Location = New System.Drawing.Point(13, 72)
        Me.lblPaymentType.Name = "lblPaymentType"
        Me.lblPaymentType.Size = New System.Drawing.Size(128, 31)
        Me.lblPaymentType.TabIndex = 6
        Me.lblPaymentType.Text = "Payment Type"
        Me.lblPaymentType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtBatchNo
        '
        Me.txtBatchNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBatchNo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtBatchNo.Location = New System.Drawing.Point(148, 14)
        Me.txtBatchNo.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBatchNo.Name = "txtBatchNo"
        Me.txtBatchNo.Size = New System.Drawing.Size(136, 23)
        Me.txtBatchNo.TabIndex = 3
        '
        'cmbType
        '
        Me.cmbType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cmbType.Enabled = False
        Me.cmbType.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Items.AddRange(New Object() {"Panel", "Non-Panel"})
        Me.cmbType.Location = New System.Drawing.Point(148, 45)
        Me.cmbType.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(136, 24)
        Me.cmbType.TabIndex = 4
        '
        'cmbPaymentType
        '
        Me.cmbPaymentType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cmbPaymentType.FormattingEnabled = True
        Me.cmbPaymentType.Items.AddRange(New Object() {"Cheque", "GIRO"})
        Me.cmbPaymentType.Location = New System.Drawing.Point(147, 75)
        Me.cmbPaymentType.Name = "cmbPaymentType"
        Me.cmbPaymentType.Size = New System.Drawing.Size(138, 24)
        Me.cmbPaymentType.TabIndex = 7
        '
        'tbpnButton
        '
        Me.tbpnButton.ColumnCount = 4
        Me.TableLayoutPanel1.SetColumnSpan(Me.tbpnButton, 2)
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.tbpnButton.Controls.Add(Me.btnSendSMS, 0, 0)
        Me.tbpnButton.Controls.Add(Me.btnCancel, 1, 0)
        Me.tbpnButton.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbpnButton.Location = New System.Drawing.Point(14, 117)
        Me.tbpnButton.Margin = New System.Windows.Forms.Padding(4)
        Me.tbpnButton.Name = "tbpnButton"
        Me.tbpnButton.RowCount = 1
        Me.tbpnButton.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tbpnButton.Size = New System.Drawing.Size(270, 37)
        Me.tbpnButton.TabIndex = 2
        '
        'btnSendSMS
        '
        Me.btnSendSMS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnSendSMS.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSendSMS.Location = New System.Drawing.Point(4, 4)
        Me.btnSendSMS.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSendSMS.Name = "btnSendSMS"
        Me.btnSendSMS.Size = New System.Drawing.Size(103, 29)
        Me.btnSendSMS.TabIndex = 0
        Me.btnSendSMS.Text = "Send"
        Me.btnSendSMS.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(115, 4)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(103, 29)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lblBatchNo
        '
        Me.lblBatchNo.AutoSize = True
        Me.lblBatchNo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblBatchNo.Location = New System.Drawing.Point(13, 10)
        Me.lblBatchNo.Name = "lblBatchNo"
        Me.lblBatchNo.Size = New System.Drawing.Size(128, 31)
        Me.lblBatchNo.TabIndex = 8
        Me.lblBatchNo.Text = "Batch No."
        Me.lblBatchNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'rtxtMessage
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.rtxtMessage, 3)
        Me.rtxtMessage.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtxtMessage.Location = New System.Drawing.Point(14, 172)
        Me.rtxtMessage.Margin = New System.Windows.Forms.Padding(4)
        Me.rtxtMessage.Name = "rtxtMessage"
        Me.TableLayoutPanel1.SetRowSpan(Me.rtxtMessage, 7)
        Me.rtxtMessage.Size = New System.Drawing.Size(478, 209)
        Me.rtxtMessage.TabIndex = 5
        Me.rtxtMessage.Text = ""
        '
        'frmSendSMS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(557, 410)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSendSMS"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SMS Notification"
        Me.Panel4.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.tbpnButton.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents tbpnButton As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnSendSMS As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents txtBatchNo As System.Windows.Forms.TextBox
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents rtxtMessage As System.Windows.Forms.RichTextBox
    Friend WithEvents lblPaymentType As System.Windows.Forms.Label
    Friend WithEvents cmbPaymentType As System.Windows.Forms.ComboBox
    Friend WithEvents lblBatchNo As System.Windows.Forms.Label

End Class
