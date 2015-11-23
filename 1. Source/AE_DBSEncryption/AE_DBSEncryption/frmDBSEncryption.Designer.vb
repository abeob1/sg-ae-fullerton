<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDBSEncryption
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
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblSelectProcessFolder = New System.Windows.Forms.Label()
        Me.lblSelectFileFolder = New System.Windows.Forms.Label()
        Me.btnSelectProcessFolder = New System.Windows.Forms.Button()
        Me.btnSelectFileFolder = New System.Windows.Forms.Button()
        Me.txtSelectFileFolder = New System.Windows.Forms.TextBox()
        Me.txtSelectProcessFolder = New System.Windows.Forms.TextBox()
        Me.tbpnButton = New System.Windows.Forms.TableLayoutPanel()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnEncrypt = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.rtxtMessage = New System.Windows.Forms.RichTextBox()
        Me.Panel5.SuspendLayout()
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
        Me.Panel1.Size = New System.Drawing.Size(15, 456)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel2.Location = New System.Drawing.Point(800, 0)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(15, 456)
        Me.Panel2.TabIndex = 1
        '
        'Panel3
        '
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(15, 0)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(785, 12)
        Me.Panel3.TabIndex = 2
        '
        'Panel4
        '
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel4.Location = New System.Drawing.Point(15, 444)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(785, 12)
        Me.Panel4.TabIndex = 3
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel5.Location = New System.Drawing.Point(15, 12)
        Me.Panel5.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(785, 444)
        Me.Panel5.TabIndex = 3
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 5
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 45.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.lblSelectProcessFolder, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.lblSelectFileFolder, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.btnSelectProcessFolder, 3, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.btnSelectFileFolder, 3, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtSelectFileFolder, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtSelectProcessFolder, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.tbpnButton, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.rtxtMessage, 1, 6)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 16
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 12.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 12.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 12.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.090908!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 12.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(785, 444)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'lblSelectProcessFolder
        '
        Me.lblSelectProcessFolder.AutoSize = True
        Me.lblSelectProcessFolder.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblSelectProcessFolder.Location = New System.Drawing.Point(24, 43)
        Me.lblSelectProcessFolder.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSelectProcessFolder.Name = "lblSelectProcessFolder"
        Me.lblSelectProcessFolder.Size = New System.Drawing.Size(196, 31)
        Me.lblSelectProcessFolder.TabIndex = 1
        Me.lblSelectProcessFolder.Text = "Select Process Folder        "
        Me.lblSelectProcessFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSelectFileFolder
        '
        Me.lblSelectFileFolder.AutoSize = True
        Me.lblSelectFileFolder.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblSelectFileFolder.Location = New System.Drawing.Point(24, 12)
        Me.lblSelectFileFolder.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSelectFileFolder.Name = "lblSelectFileFolder"
        Me.lblSelectFileFolder.Size = New System.Drawing.Size(196, 31)
        Me.lblSelectFileFolder.TabIndex = 0
        Me.lblSelectFileFolder.Text = "Select File Folder"
        Me.lblSelectFileFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSelectProcessFolder
        '
        Me.btnSelectProcessFolder.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnSelectProcessFolder.Location = New System.Drawing.Point(724, 47)
        Me.btnSelectProcessFolder.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSelectProcessFolder.Name = "btnSelectProcessFolder"
        Me.btnSelectProcessFolder.Size = New System.Drawing.Size(37, 23)
        Me.btnSelectProcessFolder.TabIndex = 3
        Me.btnSelectProcessFolder.Text = "..."
        Me.btnSelectProcessFolder.UseVisualStyleBackColor = True
        '
        'btnSelectFileFolder
        '
        Me.btnSelectFileFolder.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnSelectFileFolder.Location = New System.Drawing.Point(724, 16)
        Me.btnSelectFileFolder.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSelectFileFolder.Name = "btnSelectFileFolder"
        Me.btnSelectFileFolder.Size = New System.Drawing.Size(37, 23)
        Me.btnSelectFileFolder.TabIndex = 2
        Me.btnSelectFileFolder.Text = "..."
        Me.btnSelectFileFolder.UseVisualStyleBackColor = True
        '
        'txtSelectFileFolder
        '
        Me.txtSelectFileFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSelectFileFolder.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSelectFileFolder.Location = New System.Drawing.Point(228, 16)
        Me.txtSelectFileFolder.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSelectFileFolder.Name = "txtSelectFileFolder"
        Me.txtSelectFileFolder.Size = New System.Drawing.Size(488, 23)
        Me.txtSelectFileFolder.TabIndex = 4
        '
        'txtSelectProcessFolder
        '
        Me.txtSelectProcessFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSelectProcessFolder.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSelectProcessFolder.Location = New System.Drawing.Point(228, 47)
        Me.txtSelectProcessFolder.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSelectProcessFolder.Name = "txtSelectProcessFolder"
        Me.txtSelectProcessFolder.Size = New System.Drawing.Size(488, 23)
        Me.txtSelectProcessFolder.TabIndex = 5
        '
        'tbpnButton
        '
        Me.tbpnButton.ColumnCount = 7
        Me.TableLayoutPanel1.SetColumnSpan(Me.tbpnButton, 2)
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.tbpnButton.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.tbpnButton.Controls.Add(Me.btnClear, 0, 0)
        Me.tbpnButton.Controls.Add(Me.btnEncrypt, 0, 0)
        Me.tbpnButton.Controls.Add(Me.btnCancel, 1, 0)
        Me.tbpnButton.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbpnButton.Location = New System.Drawing.Point(26, 91)
        Me.tbpnButton.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.tbpnButton.Name = "tbpnButton"
        Me.tbpnButton.RowCount = 1
        Me.tbpnButton.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tbpnButton.Size = New System.Drawing.Size(688, 35)
        Me.tbpnButton.TabIndex = 6
        '
        'btnClear
        '
        Me.btnClear.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClear.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnClear.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Location = New System.Drawing.Point(106, 5)
        Me.btnClear.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(88, 25)
        Me.btnClear.TabIndex = 2
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnEncrypt
        '
        Me.btnEncrypt.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnEncrypt.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEncrypt.Location = New System.Drawing.Point(6, 5)
        Me.btnEncrypt.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.btnEncrypt.Name = "btnEncrypt"
        Me.btnEncrypt.Size = New System.Drawing.Size(88, 25)
        Me.btnEncrypt.TabIndex = 0
        Me.btnEncrypt.Text = "Encrypt"
        Me.btnEncrypt.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(206, 5)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 25)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'rtxtMessage
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.rtxtMessage, 3)
        Me.rtxtMessage.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtxtMessage.Location = New System.Drawing.Point(24, 147)
        Me.rtxtMessage.Margin = New System.Windows.Forms.Padding(4)
        Me.rtxtMessage.Name = "rtxtMessage"
        Me.TableLayoutPanel1.SetRowSpan(Me.rtxtMessage, 9)
        Me.rtxtMessage.Size = New System.Drawing.Size(737, 271)
        Me.rtxtMessage.TabIndex = 7
        Me.rtxtMessage.Text = ""
        '
        'frmDBSEncryption
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(815, 456)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "frmDBSEncryption"
        Me.ShowIcon = False
        Me.Text = "DBS File Encryption"
        Me.Panel5.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.tbpnButton.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents lblSelectFileFolder As System.Windows.Forms.Label
    Friend WithEvents lblSelectProcessFolder As System.Windows.Forms.Label
    Friend WithEvents btnSelectFileFolder As System.Windows.Forms.Button
    Friend WithEvents btnSelectProcessFolder As System.Windows.Forms.Button
    Friend WithEvents txtSelectFileFolder As System.Windows.Forms.TextBox
    Friend WithEvents txtSelectProcessFolder As System.Windows.Forms.TextBox
    Friend WithEvents tbpnButton As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnEncrypt As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents rtxtMessage As System.Windows.Forms.RichTextBox

End Class
