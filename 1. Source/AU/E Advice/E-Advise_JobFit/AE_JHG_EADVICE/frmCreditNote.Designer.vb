<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCreditNote
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCreditNote))
        Me.DocType = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtStatusMsg = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnPDFGen = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.SCompany = New System.Windows.Forms.ComboBox()
        Me.DocNo = New System.Windows.Forms.Panel()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Bt_CFL1 = New System.Windows.Forms.Button()
        Me.DateTimePicker5 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker4 = New System.Windows.Forms.DateTimePicker()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtBpCode = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDocNumTo = New System.Windows.Forms.TextBox()
        Me.txtDocNumFrom = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmbSenderEmail = New System.Windows.Forms.ComboBox()
        Me.gvtaxinvoice = New System.Windows.Forms.DataGridView()
        Me.Choose = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.clmnInvoiceno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ClmnDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clmnemail = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clmnsend = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clmnstatus = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clmdocentry = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lblMsg = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.DocNo.SuspendLayout()
        CType(Me.gvtaxinvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DocType
        '
        Me.DocType.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DocType.FormattingEnabled = True
        Me.DocType.Items.AddRange(New Object() {"Injury Management" & Global.Microsoft.VisualBasic.ChrW(9), "PreEmployment"})
        Me.DocType.Location = New System.Drawing.Point(153, 50)
        Me.DocType.Margin = New System.Windows.Forms.Padding(4)
        Me.DocType.Name = "DocType"
        Me.DocType.Size = New System.Drawing.Size(312, 24)
        Me.DocType.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(18, 54)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Type"
        '
        'txtStatusMsg
        '
        Me.txtStatusMsg.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatusMsg.Location = New System.Drawing.Point(20, 301)
        Me.txtStatusMsg.Margin = New System.Windows.Forms.Padding(4)
        Me.txtStatusMsg.Multiline = True
        Me.txtStatusMsg.Name = "txtStatusMsg"
        Me.txtStatusMsg.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStatusMsg.Size = New System.Drawing.Size(843, 285)
        Me.txtStatusMsg.TabIndex = 99
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(456, 119)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(99, 33)
        Me.Button2.TabIndex = 77
        Me.Button2.Text = "&Clear"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(349, 123)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(99, 33)
        Me.Button3.TabIndex = 88
        Me.Button3.Text = "&Close"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnPDFGen
        '
        Me.btnPDFGen.Enabled = False
        Me.btnPDFGen.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPDFGen.Location = New System.Drawing.Point(184, 123)
        Me.btnPDFGen.Margin = New System.Windows.Forms.Padding(4)
        Me.btnPDFGen.Name = "btnPDFGen"
        Me.btnPDFGen.Size = New System.Drawing.Size(157, 33)
        Me.btnPDFGen.TabIndex = 66
        Me.btnPDFGen.Text = "&Send Email"
        Me.btnPDFGen.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(15, 24)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(101, 13)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Select Company"
        '
        'SCompany
        '
        Me.SCompany.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.SCompany.FormattingEnabled = True
        Me.SCompany.Location = New System.Drawing.Point(153, 18)
        Me.SCompany.Margin = New System.Windows.Forms.Padding(4)
        Me.SCompany.Name = "SCompany"
        Me.SCompany.Size = New System.Drawing.Size(312, 24)
        Me.SCompany.TabIndex = 1
        '
        'DocNo
        '
        Me.DocNo.Controls.Add(Me.CheckBox1)
        Me.DocNo.Controls.Add(Me.Label14)
        Me.DocNo.Controls.Add(Me.DateTimePicker2)
        Me.DocNo.Controls.Add(Me.Label13)
        Me.DocNo.Controls.Add(Me.Bt_CFL1)
        Me.DocNo.Controls.Add(Me.DateTimePicker5)
        Me.DocNo.Controls.Add(Me.DateTimePicker4)
        Me.DocNo.Controls.Add(Me.Label8)
        Me.DocNo.Controls.Add(Me.Label9)
        Me.DocNo.Controls.Add(Me.Label6)
        Me.DocNo.Controls.Add(Me.Label5)
        Me.DocNo.Controls.Add(Me.Label4)
        Me.DocNo.Controls.Add(Me.txtBpCode)
        Me.DocNo.Controls.Add(Me.Label3)
        Me.DocNo.Controls.Add(Me.Label2)
        Me.DocNo.Controls.Add(Me.txtDocNumTo)
        Me.DocNo.Controls.Add(Me.txtDocNumFrom)
        Me.DocNo.Location = New System.Drawing.Point(474, 18)
        Me.DocNo.Margin = New System.Windows.Forms.Padding(4)
        Me.DocNo.Name = "DocNo"
        Me.DocNo.Size = New System.Drawing.Size(637, 113)
        Me.DocNo.TabIndex = 23
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(291, 81)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(70, 13)
        Me.Label14.TabIndex = 101
        Me.Label14.Text = "Value Date"
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Checked = False
        Me.DateTimePicker2.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker2.Location = New System.Drawing.Point(371, 74)
        Me.DateTimePicker2.Margin = New System.Windows.Forms.Padding(4)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.ShowCheckBox = True
        Me.DateTimePicker2.Size = New System.Drawing.Size(152, 23)
        Me.DateTimePicker2.TabIndex = 101
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label13.Location = New System.Drawing.Point(446, 52)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(82, 13)
        Me.Label13.TabIndex = 47
        Me.Label13.Text = "or Bp Group)"
        '
        'Bt_CFL1
        '
        Me.Bt_CFL1.Location = New System.Drawing.Point(258, 73)
        Me.Bt_CFL1.Margin = New System.Windows.Forms.Padding(4)
        Me.Bt_CFL1.Name = "Bt_CFL1"
        Me.Bt_CFL1.Size = New System.Drawing.Size(31, 28)
        Me.Bt_CFL1.TabIndex = 44
        Me.Bt_CFL1.Text = "..."
        Me.Bt_CFL1.UseVisualStyleBackColor = True
        '
        'DateTimePicker5
        '
        Me.DateTimePicker5.Checked = False
        Me.DateTimePicker5.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.DateTimePicker5.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker5.Location = New System.Drawing.Point(289, 39)
        Me.DateTimePicker5.Margin = New System.Windows.Forms.Padding(4)
        Me.DateTimePicker5.Name = "DateTimePicker5"
        Me.DateTimePicker5.ShowCheckBox = True
        Me.DateTimePicker5.Size = New System.Drawing.Size(136, 23)
        Me.DateTimePicker5.TabIndex = 43
        '
        'DateTimePicker4
        '
        Me.DateTimePicker4.Checked = False
        Me.DateTimePicker4.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.DateTimePicker4.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker4.Location = New System.Drawing.Point(105, 39)
        Me.DateTimePicker4.Margin = New System.Windows.Forms.Padding(4)
        Me.DateTimePicker4.Name = "DateTimePicker4"
        Me.DateTimePicker4.ShowCheckBox = True
        Me.DateTimePicker4.Size = New System.Drawing.Size(139, 23)
        Me.DateTimePicker4.TabIndex = 42
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(255, 47)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(21, 13)
        Me.Label8.TabIndex = 32
        Me.Label8.Text = "To"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(7, 47)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(93, 13)
        Me.Label9.TabIndex = 31
        Me.Label9.Text = "Doc Date From"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label6.Location = New System.Drawing.Point(430, 36)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(202, 13)
        Me.Label6.TabIndex = 28
        Me.Label6.Text = "Either Doc No. or Doc Date Range"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label5.Location = New System.Drawing.Point(434, 16)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(138, 13)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "(Any one is mandatory"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(7, 82)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 13)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "BP Group"
        '
        'txtBpCode
        '
        Me.txtBpCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBpCode.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBpCode.Location = New System.Drawing.Point(116, 76)
        Me.txtBpCode.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBpCode.Name = "txtBpCode"
        Me.txtBpCode.Size = New System.Drawing.Size(139, 23)
        Me.txtBpCode.TabIndex = 25
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(255, 9)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(21, 13)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "To"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(7, 10)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 13)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Doc No. From"
        '
        'txtDocNumTo
        '
        Me.txtDocNumTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDocNumTo.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDocNumTo.Location = New System.Drawing.Point(289, 4)
        Me.txtDocNumTo.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDocNumTo.Name = "txtDocNumTo"
        Me.txtDocNumTo.Size = New System.Drawing.Size(137, 23)
        Me.txtDocNumTo.TabIndex = 7
        '
        'txtDocNumFrom
        '
        Me.txtDocNumFrom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDocNumFrom.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDocNumFrom.Location = New System.Drawing.Point(105, 4)
        Me.txtDocNumFrom.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDocNumFrom.Name = "txtDocNumFrom"
        Me.txtDocNumFrom.Size = New System.Drawing.Size(139, 23)
        Me.txtDocNumFrom.TabIndex = 6
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(19, 123)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(157, 33)
        Me.Button1.TabIndex = 100
        Me.Button1.Text = "&Show"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(16, 86)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(83, 13)
        Me.Label10.TabIndex = 102
        Me.Label10.Text = "Sender Email"
        '
        'cmbSenderEmail
        '
        Me.cmbSenderEmail.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSenderEmail.FormattingEnabled = True
        Me.cmbSenderEmail.Location = New System.Drawing.Point(154, 82)
        Me.cmbSenderEmail.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbSenderEmail.Name = "cmbSenderEmail"
        Me.cmbSenderEmail.Size = New System.Drawing.Size(312, 24)
        Me.cmbSenderEmail.TabIndex = 101
        '
        'gvtaxinvoice
        '
        Me.gvtaxinvoice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gvtaxinvoice.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Choose, Me.clmnInvoiceno, Me.ClmnDate, Me.clmnemail, Me.clmnsend, Me.clmnstatus, Me.clmdocentry})
        Me.gvtaxinvoice.Location = New System.Drawing.Point(12, 168)
        Me.gvtaxinvoice.Name = "gvtaxinvoice"
        Me.gvtaxinvoice.RowHeadersVisible = False
        Me.gvtaxinvoice.Size = New System.Drawing.Size(1092, 449)
        Me.gvtaxinvoice.TabIndex = 104
        '
        'Choose
        '
        Me.Choose.HeaderText = "Choose"
        Me.Choose.Name = "Choose"
        Me.Choose.Width = 65
        '
        'clmnInvoiceno
        '
        Me.clmnInvoiceno.HeaderText = "Credit Note No."
        Me.clmnInvoiceno.Name = "clmnInvoiceno"
        Me.clmnInvoiceno.ReadOnly = True
        Me.clmnInvoiceno.Width = 160
        '
        'ClmnDate
        '
        DataGridViewCellStyle1.NullValue = Nothing
        Me.ClmnDate.DefaultCellStyle = DataGridViewCellStyle1
        Me.ClmnDate.HeaderText = "Document Date"
        Me.ClmnDate.Name = "ClmnDate"
        Me.ClmnDate.ReadOnly = True
        Me.ClmnDate.Width = 160
        '
        'clmnemail
        '
        Me.clmnemail.HeaderText = "Email "
        Me.clmnemail.Name = "clmnemail"
        Me.clmnemail.ReadOnly = True
        Me.clmnemail.Width = 280
        '
        'clmnsend
        '
        Me.clmnsend.HeaderText = "Send (Y/N)"
        Me.clmnsend.Name = "clmnsend"
        Me.clmnsend.ReadOnly = True
        '
        'clmnstatus
        '
        Me.clmnstatus.HeaderText = "Status"
        Me.clmnstatus.Name = "clmnstatus"
        Me.clmnstatus.ReadOnly = True
        Me.clmnstatus.Width = 320
        '
        'clmdocentry
        '
        Me.clmdocentry.HeaderText = "DocEntry"
        Me.clmdocentry.Name = "clmdocentry"
        Me.clmdocentry.ReadOnly = True
        Me.clmdocentry.Visible = False
        '
        'lblMsg
        '
        Me.lblMsg.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMsg.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblMsg.Location = New System.Drawing.Point(535, 137)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(576, 26)
        Me.lblMsg.TabIndex = 105
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(70, 80)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(42, 20)
        Me.CheckBox1.TabIndex = 102
        Me.CheckBox1.Text = "All"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'frmCreditNote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1111, 622)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.gvtaxinvoice)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.cmbSenderEmail)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DocNo)
        Me.Controls.Add(Me.SCompany)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnPDFGen)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.txtStatusMsg)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DocType)
        Me.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "frmCreditNote"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AU - Credit Note"
        Me.DocNo.ResumeLayout(False)
        Me.DocNo.PerformLayout()
        CType(Me.gvtaxinvoice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DocType As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtStatusMsg As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents btnPDFGen As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents SCompany As System.Windows.Forms.ComboBox
    Friend WithEvents DocNo As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtBpCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDocNumTo As System.Windows.Forms.TextBox
    Friend WithEvents txtDocNumFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker5 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker4 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Bt_CFL1 As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cmbSenderEmail As System.Windows.Forms.ComboBox
    Friend WithEvents gvtaxinvoice As System.Windows.Forms.DataGridView
    Friend WithEvents lblMsg As System.Windows.Forms.Label
    Friend WithEvents Choose As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents clmnInvoiceno As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ClmnDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clmnemail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clmnsend As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clmnstatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clmdocentry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
End Class
