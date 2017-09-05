<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEadvice
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEadvice))
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnTaxInvoice = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.lbltime = New System.Windows.Forms.Label()
        Me.lbldate = New System.Windows.Forms.Label()
        Me.btnPaymentadvice = New System.Windows.Forms.Button()
        Me.btnSOA = New System.Windows.Forms.Button()
        Me.DateTimer = New System.Windows.Forms.Timer(Me.components)
        Me.btnCreditnote = New System.Windows.Forms.Button()
        Me.StatusStrip.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel, Me.ToolStripStatusLabel1})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 670)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(959, 22)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(36, 17)
        Me.ToolStripStatusLabel.Text = "Msg : "
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.btnCreditnote)
        Me.Panel3.Controls.Add(Me.btnTaxInvoice)
        Me.Panel3.Controls.Add(Me.btnExit)
        Me.Panel3.Controls.Add(Me.lbltime)
        Me.Panel3.Controls.Add(Me.lbldate)
        Me.Panel3.Controls.Add(Me.btnPaymentadvice)
        Me.Panel3.Controls.Add(Me.btnSOA)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(194, 670)
        Me.Panel3.TabIndex = 11
        '
        'btnTaxInvoice
        '
        Me.btnTaxInvoice.BackColor = System.Drawing.SystemColors.Control
        Me.btnTaxInvoice.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnTaxInvoice.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gray
        Me.btnTaxInvoice.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTaxInvoice.Location = New System.Drawing.Point(3, 382)
        Me.btnTaxInvoice.Name = "btnTaxInvoice"
        Me.btnTaxInvoice.Size = New System.Drawing.Size(184, 40)
        Me.btnTaxInvoice.TabIndex = 17
        Me.btnTaxInvoice.Text = "&Tax Invoice"
        Me.btnTaxInvoice.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.SystemColors.Control
        Me.btnExit.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnExit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gray
        Me.btnExit.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Location = New System.Drawing.Point(3, 476)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(184, 40)
        Me.btnExit.TabIndex = 16
        Me.btnExit.Text = "&Exit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'lbltime
        '
        Me.lbltime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbltime.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lbltime.ForeColor = System.Drawing.SystemColors.Desktop
        Me.lbltime.Location = New System.Drawing.Point(3, 634)
        Me.lbltime.Name = "lbltime"
        Me.lbltime.Size = New System.Drawing.Size(184, 26)
        Me.lbltime.TabIndex = 15
        Me.lbltime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbldate
        '
        Me.lbldate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbldate.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lbldate.ForeColor = System.Drawing.SystemColors.Desktop
        Me.lbldate.Location = New System.Drawing.Point(3, 602)
        Me.lbldate.Name = "lbldate"
        Me.lbldate.Size = New System.Drawing.Size(184, 26)
        Me.lbldate.TabIndex = 14
        Me.lbldate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnPaymentadvice
        '
        Me.btnPaymentadvice.BackColor = System.Drawing.SystemColors.Control
        Me.btnPaymentadvice.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPaymentadvice.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gray
        Me.btnPaymentadvice.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPaymentadvice.Location = New System.Drawing.Point(3, 335)
        Me.btnPaymentadvice.Name = "btnPaymentadvice"
        Me.btnPaymentadvice.Size = New System.Drawing.Size(184, 40)
        Me.btnPaymentadvice.TabIndex = 5
        Me.btnPaymentadvice.Text = "&Payment Advice"
        Me.btnPaymentadvice.UseVisualStyleBackColor = False
        '
        'btnSOA
        '
        Me.btnSOA.BackColor = System.Drawing.SystemColors.Control
        Me.btnSOA.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSOA.FlatAppearance.BorderColor = System.Drawing.Color.Blue
        Me.btnSOA.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gray
        Me.btnSOA.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSOA.Location = New System.Drawing.Point(3, 288)
        Me.btnSOA.Name = "btnSOA"
        Me.btnSOA.Size = New System.Drawing.Size(184, 40)
        Me.btnSOA.TabIndex = 4
        Me.btnSOA.Text = "&Statement of Account"
        Me.btnSOA.UseVisualStyleBackColor = False
        '
        'DateTimer
        '
        Me.DateTimer.Enabled = True
        '
        'btnCreditnote
        '
        Me.btnCreditnote.BackColor = System.Drawing.SystemColors.Control
        Me.btnCreditnote.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCreditnote.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gray
        Me.btnCreditnote.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreditnote.Location = New System.Drawing.Point(3, 429)
        Me.btnCreditnote.Name = "btnCreditnote"
        Me.btnCreditnote.Size = New System.Drawing.Size(184, 40)
        Me.btnCreditnote.TabIndex = 18
        Me.btnCreditnote.Text = "&Credit Note"
        Me.btnCreditnote.UseVisualStyleBackColor = False
        '
        'frmEadvice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(959, 692)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.StatusStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Name = "frmEadvice"
        Me.Text = "FHG - E Advice"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnPaymentadvice As System.Windows.Forms.Button
    Friend WithEvents btnSOA As System.Windows.Forms.Button
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lbldate As System.Windows.Forms.Label
    Friend WithEvents lbltime As System.Windows.Forms.Label
    Friend WithEvents DateTimer As System.Windows.Forms.Timer
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnTaxInvoice As System.Windows.Forms.Button
    Friend WithEvents btnCreditnote As System.Windows.Forms.Button

End Class
