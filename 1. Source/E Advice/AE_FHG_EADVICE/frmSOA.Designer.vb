<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSOA
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSOA))
        Me.DocType = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnShow = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnPDFGen = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.SCompany = New System.Windows.Forms.ComboBox()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CboxBPlist = New System.Windows.Forms.ComboBox()
        Me.Dgv_BPList = New System.Windows.Forms.DataGridView()
        Me.Choose = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.CardCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CardName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Email = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SendEmail = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Status = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtGroup = New System.Windows.Forms.TextBox()
        Me.btnGroup = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.lblMsg = New System.Windows.Forms.Label()
        Me.txtGroupNo = New System.Windows.Forms.TextBox()
        Me.lblMsg1 = New System.Windows.Forms.Label()
        CType(Me.Dgv_BPList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DocType
        '
        Me.DocType.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DocType.FormattingEnabled = True
        Me.DocType.Items.AddRange(New Object() {"SOA"})
        Me.DocType.Location = New System.Drawing.Point(153, 59)
        Me.DocType.Margin = New System.Windows.Forms.Padding(4)
        Me.DocType.Name = "DocType"
        Me.DocType.Size = New System.Drawing.Size(312, 24)
        Me.DocType.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 69)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Document Type"
        '
        'btnShow
        '
        Me.btnShow.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnShow.Location = New System.Drawing.Point(301, 129)
        Me.btnShow.Margin = New System.Windows.Forms.Padding(4)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(99, 30)
        Me.btnShow.TabIndex = 77
        Me.btnShow.Text = "&Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(191, 577)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(99, 33)
        Me.Button3.TabIndex = 88
        Me.Button3.Text = "&Close"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnPDFGen
        '
        Me.btnPDFGen.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPDFGen.Location = New System.Drawing.Point(23, 575)
        Me.btnPDFGen.Margin = New System.Windows.Forms.Padding(4)
        Me.btnPDFGen.Name = "btnPDFGen"
        Me.btnPDFGen.Size = New System.Drawing.Size(157, 33)
        Me.btnPDFGen.TabIndex = 66
        Me.btnPDFGen.Text = "&Send Email"
        Me.btnPDFGen.UseVisualStyleBackColor = True
        Me.btnPDFGen.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(16, 32)
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
        Me.SCompany.Location = New System.Drawing.Point(153, 21)
        Me.SCompany.Margin = New System.Windows.Forms.Padding(4)
        Me.SCompany.Name = "SCompany"
        Me.SCompany.Size = New System.Drawing.Size(312, 24)
        Me.SCompany.TabIndex = 1
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Checked = False
        Me.DateTimePicker1.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(153, 130)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(4)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.ShowCheckBox = True
        Me.DateTimePicker1.Size = New System.Drawing.Size(139, 23)
        Me.DateTimePicker1.TabIndex = 5
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(16, 139)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(34, 13)
        Me.Label10.TabIndex = 38
        Me.Label10.Text = "Date"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(16, 104)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 13)
        Me.Label12.TabIndex = 36
        Me.Label12.Text = "BP Group"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(23, 577)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(157, 33)
        Me.Button1.TabIndex = 100
        Me.Button1.Text = "&Send Email"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CboxBPlist
        '
        Me.CboxBPlist.FormattingEnabled = True
        Me.CboxBPlist.Location = New System.Drawing.Point(669, 47)
        Me.CboxBPlist.Margin = New System.Windows.Forms.Padding(4)
        Me.CboxBPlist.Name = "CboxBPlist"
        Me.CboxBPlist.Size = New System.Drawing.Size(312, 24)
        Me.CboxBPlist.TabIndex = 101
        Me.CboxBPlist.Visible = False
        '
        'Dgv_BPList
        '
        Me.Dgv_BPList.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Dgv_BPList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Dgv_BPList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Choose, Me.CardCode, Me.CardName, Me.Email, Me.SendEmail, Me.Status})
        Me.Dgv_BPList.Location = New System.Drawing.Point(20, 167)
        Me.Dgv_BPList.Margin = New System.Windows.Forms.Padding(4)
        Me.Dgv_BPList.Name = "Dgv_BPList"
        Me.Dgv_BPList.Size = New System.Drawing.Size(1050, 400)
        Me.Dgv_BPList.TabIndex = 102
        '
        'Choose
        '
        Me.Choose.Frozen = True
        Me.Choose.HeaderText = "Choose"
        Me.Choose.Name = "Choose"
        Me.Choose.Width = 60
        '
        'CardCode
        '
        Me.CardCode.Frozen = True
        Me.CardCode.HeaderText = "CardCode"
        Me.CardCode.Name = "CardCode"
        Me.CardCode.ReadOnly = True
        Me.CardCode.Width = 120
        '
        'CardName
        '
        Me.CardName.Frozen = True
        Me.CardName.HeaderText = "CardName"
        Me.CardName.Name = "CardName"
        Me.CardName.ReadOnly = True
        Me.CardName.Width = 250
        '
        'Email
        '
        Me.Email.Frozen = True
        Me.Email.HeaderText = "Email"
        Me.Email.Name = "Email"
        Me.Email.ReadOnly = True
        Me.Email.Width = 200
        '
        'SendEmail
        '
        Me.SendEmail.Frozen = True
        Me.SendEmail.HeaderText = "Send Email(Y/N)"
        Me.SendEmail.Name = "SendEmail"
        Me.SendEmail.ReadOnly = True
        '
        'Status
        '
        Me.Status.Frozen = True
        Me.Status.HeaderText = "Status"
        Me.Status.Name = "Status"
        Me.Status.ReadOnly = True
        Me.Status.Width = 400
        '
        'txtGroup
        '
        Me.txtGroup.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.txtGroup.Location = New System.Drawing.Point(354, 98)
        Me.txtGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.txtGroup.Name = "txtGroup"
        Me.txtGroup.Size = New System.Drawing.Size(132, 23)
        Me.txtGroup.TabIndex = 103
        Me.txtGroup.Visible = False
        '
        'btnGroup
        '
        Me.btnGroup.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGroup.Location = New System.Drawing.Point(302, 96)
        Me.btnGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.btnGroup.Name = "btnGroup"
        Me.btnGroup.Size = New System.Drawing.Size(35, 26)
        Me.btnGroup.TabIndex = 104
        Me.btnGroup.Text = "..."
        Me.btnGroup.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnGroup.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(96, 102)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(50, 20)
        Me.CheckBox1.TabIndex = 105
        Me.CheckBox1.Text = "ALL"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'lblMsg
        '
        Me.lblMsg.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMsg.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblMsg.Location = New System.Drawing.Point(502, 130)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(568, 25)
        Me.lblMsg.TabIndex = 106
        '
        'txtGroupNo
        '
        Me.txtGroupNo.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.txtGroupNo.Location = New System.Drawing.Point(153, 96)
        Me.txtGroupNo.Margin = New System.Windows.Forms.Padding(4)
        Me.txtGroupNo.Name = "txtGroupNo"
        Me.txtGroupNo.Size = New System.Drawing.Size(139, 23)
        Me.txtGroupNo.TabIndex = 107
        '
        'lblMsg1
        '
        Me.lblMsg1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMsg1.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblMsg1.Location = New System.Drawing.Point(494, 105)
        Me.lblMsg1.Name = "lblMsg1"
        Me.lblMsg1.Size = New System.Drawing.Size(428, 25)
        Me.lblMsg1.TabIndex = 108
        '
        'frmSOA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1087, 613)
        Me.Controls.Add(Me.lblMsg1)
        Me.Controls.Add(Me.txtGroupNo)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.btnGroup)
        Me.Controls.Add(Me.txtGroup)
        Me.Controls.Add(Me.Dgv_BPList)
        Me.Controls.Add(Me.CboxBPlist)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.SCompany)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.btnPDFGen)
        Me.Controls.Add(Me.btnShow)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DocType)
        Me.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "frmSOA"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FHG - E Advice"
        CType(Me.Dgv_BPList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DocType As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents btnPDFGen As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents SCompany As System.Windows.Forms.ComboBox
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CboxBPlist As System.Windows.Forms.ComboBox
    Friend WithEvents Dgv_BPList As System.Windows.Forms.DataGridView
    Friend WithEvents txtGroup As System.Windows.Forms.TextBox
    Friend WithEvents btnGroup As System.Windows.Forms.Button
    Friend WithEvents Choose As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CardCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CardName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Email As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SendEmail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Status As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents lblMsg As System.Windows.Forms.Label
    Friend WithEvents txtGroupNo As System.Windows.Forms.TextBox
    Friend WithEvents lblMsg1 As System.Windows.Forms.Label
End Class
