<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUFFGen
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUFFGen))
        Me.txtStatusMsg = New System.Windows.Forms.TextBox()
        Me.TabCon = New System.Windows.Forms.TabControl()
        Me.tabGenDBS = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmbCompany = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.datpick = New System.Windows.Forms.DateTimePicker()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.lstBatch = New System.Windows.Forms.ListBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.txtBatchno = New System.Windows.Forms.TextBox()
        Me.lblBatchNo = New System.Windows.Forms.Label()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.lblFileName = New System.Windows.Forms.Label()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtpath = New System.Windows.Forms.TextBox()
        Me.lblPathfolder = New System.Windows.Forms.Label()
        Me.tabUpdateDBS = New System.Windows.Forms.TabPage()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnFilebrowse = New System.Windows.Forms.Button()
        Me.txtFilepath = New System.Windows.Forms.TextBox()
        Me.lblFilepath = New System.Windows.Forms.Label()
        Me.btnUpload = New System.Windows.Forms.Button()
        Me.TabCon.SuspendLayout()
        Me.tabGenDBS.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.tabUpdateDBS.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtStatusMsg
        '
        Me.txtStatusMsg.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatusMsg.Location = New System.Drawing.Point(5, 252)
        Me.txtStatusMsg.Multiline = True
        Me.txtStatusMsg.Name = "txtStatusMsg"
        Me.txtStatusMsg.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStatusMsg.Size = New System.Drawing.Size(594, 247)
        Me.txtStatusMsg.TabIndex = 99
        '
        'TabCon
        '
        Me.TabCon.Controls.Add(Me.tabGenDBS)
        Me.TabCon.Controls.Add(Me.tabUpdateDBS)
        Me.TabCon.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabCon.Location = New System.Drawing.Point(4, 12)
        Me.TabCon.Name = "TabCon"
        Me.TabCon.SelectedIndex = 0
        Me.TabCon.Size = New System.Drawing.Size(594, 234)
        Me.TabCon.TabIndex = 101
        '
        'tabGenDBS
        '
        Me.tabGenDBS.Controls.Add(Me.Panel1)
        Me.tabGenDBS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.tabGenDBS.Location = New System.Drawing.Point(4, 22)
        Me.tabGenDBS.Name = "tabGenDBS"
        Me.tabGenDBS.Padding = New System.Windows.Forms.Padding(3)
        Me.tabGenDBS.Size = New System.Drawing.Size(586, 208)
        Me.tabGenDBS.TabIndex = 0
        Me.tabGenDBS.Text = "Generate DBS File"
        Me.tabGenDBS.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmbCompany)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.datpick)
        Me.Panel1.Controls.Add(Me.btnView)
        Me.Panel1.Controls.Add(Me.btnGenerate)
        Me.Panel1.Controls.Add(Me.btnClear)
        Me.Panel1.Controls.Add(Me.lstBatch)
        Me.Panel1.Controls.Add(Me.Button4)
        Me.Panel1.Controls.Add(Me.txtBatchno)
        Me.Panel1.Controls.Add(Me.lblBatchNo)
        Me.Panel1.Controls.Add(Me.lblDate)
        Me.Panel1.Controls.Add(Me.txtFileName)
        Me.Panel1.Controls.Add(Me.lblFileName)
        Me.Panel1.Controls.Add(Me.btnBrowse)
        Me.Panel1.Controls.Add(Me.txtpath)
        Me.Panel1.Controls.Add(Me.lblPathfolder)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(580, 202)
        Me.Panel1.TabIndex = 0
        '
        'cmbCompany
        '
        Me.cmbCompany.FormattingEnabled = True
        Me.cmbCompany.Items.AddRange(New Object() {"--Select--", "AON", "AIA"})
        Me.cmbCompany.Location = New System.Drawing.Point(119, 121)
        Me.cmbCompany.Name = "cmbCompany"
        Me.cmbCompany.Size = New System.Drawing.Size(121, 21)
        Me.cmbCompany.TabIndex = 17
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 127)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 13)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Template Type"
        '
        'datpick
        '
        Me.datpick.Checked = False
        Me.datpick.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.datpick.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.datpick.Location = New System.Drawing.Point(120, 64)
        Me.datpick.Name = "datpick"
        Me.datpick.ShowCheckBox = True
        Me.datpick.Size = New System.Drawing.Size(124, 23)
        Me.datpick.TabIndex = 15
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(108, 162)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 13
        Me.btnView.Text = "View File"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnGenerate
        '
        Me.btnGenerate.Location = New System.Drawing.Point(7, 162)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(94, 23)
        Me.btnGenerate.TabIndex = 12
        Me.btnGenerate.Text = "Generate File"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Location = New System.Drawing.Point(500, 109)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(19, 23)
        Me.btnClear.TabIndex = 11
        Me.btnClear.Text = "X"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'lstBatch
        '
        Me.lstBatch.FormattingEnabled = True
        Me.lstBatch.Location = New System.Drawing.Point(373, 38)
        Me.lstBatch.Name = "lstBatch"
        Me.lstBatch.Size = New System.Drawing.Size(120, 95)
        Me.lstBatch.TabIndex = 10
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(333, 93)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(34, 21)
        Me.Button4.TabIndex = 9
        Me.Button4.Text = ">>"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'txtBatchno
        '
        Me.txtBatchno.Location = New System.Drawing.Point(119, 93)
        Me.txtBatchno.Name = "txtBatchno"
        Me.txtBatchno.Size = New System.Drawing.Size(208, 21)
        Me.txtBatchno.TabIndex = 8
        '
        'lblBatchNo
        '
        Me.lblBatchNo.AutoSize = True
        Me.lblBatchNo.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBatchNo.Location = New System.Drawing.Point(3, 101)
        Me.lblBatchNo.Name = "lblBatchNo"
        Me.lblBatchNo.Size = New System.Drawing.Size(62, 13)
        Me.lblBatchNo.TabIndex = 7
        Me.lblBatchNo.Text = "Batch No."
        '
        'lblDate
        '
        Me.lblDate.AutoSize = True
        Me.lblDate.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate.Location = New System.Drawing.Point(4, 74)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(34, 13)
        Me.lblDate.TabIndex = 5
        Me.lblDate.Text = "Date"
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(120, 37)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(208, 21)
        Me.txtFileName.TabIndex = 4
        '
        'lblFileName
        '
        Me.lblFileName.AutoSize = True
        Me.lblFileName.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFileName.Location = New System.Drawing.Point(4, 45)
        Me.lblFileName.Name = "lblFileName"
        Me.lblFileName.Size = New System.Drawing.Size(63, 13)
        Me.lblFileName.TabIndex = 3
        Me.lblFileName.Text = "File Name"
        '
        'btnBrowse
        '
        Me.btnBrowse.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.Location = New System.Drawing.Point(544, 11)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(30, 21)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "..."
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtpath
        '
        Me.txtpath.Location = New System.Drawing.Point(120, 11)
        Me.txtpath.Name = "txtpath"
        Me.txtpath.Size = New System.Drawing.Size(418, 21)
        Me.txtpath.TabIndex = 1
        '
        'lblPathfolder
        '
        Me.lblPathfolder.AutoSize = True
        Me.lblPathfolder.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPathfolder.Location = New System.Drawing.Point(4, 19)
        Me.lblPathfolder.Name = "lblPathfolder"
        Me.lblPathfolder.Size = New System.Drawing.Size(110, 13)
        Me.lblPathfolder.TabIndex = 0
        Me.lblPathfolder.Text = "Select Folder Path"
        '
        'tabUpdateDBS
        '
        Me.tabUpdateDBS.Controls.Add(Me.btnUpload)
        Me.tabUpdateDBS.Controls.Add(Me.btnFilebrowse)
        Me.tabUpdateDBS.Controls.Add(Me.txtFilepath)
        Me.tabUpdateDBS.Controls.Add(Me.lblFilepath)
        Me.tabUpdateDBS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabUpdateDBS.Location = New System.Drawing.Point(4, 22)
        Me.tabUpdateDBS.Name = "tabUpdateDBS"
        Me.tabUpdateDBS.Padding = New System.Windows.Forms.Padding(3)
        Me.tabUpdateDBS.Size = New System.Drawing.Size(586, 208)
        Me.tabUpdateDBS.TabIndex = 1
        Me.tabUpdateDBS.Text = "Update DBS File"
        Me.tabUpdateDBS.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(3, 505)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(74, 27)
        Me.Button2.TabIndex = 102
        Me.Button2.Text = "&Clear"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(85, 505)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(74, 27)
        Me.Button3.TabIndex = 103
        Me.Button3.Text = "&Close"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnFilebrowse
        '
        Me.btnFilebrowse.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFilebrowse.Location = New System.Drawing.Point(546, 27)
        Me.btnFilebrowse.Name = "btnFilebrowse"
        Me.btnFilebrowse.Size = New System.Drawing.Size(30, 21)
        Me.btnFilebrowse.TabIndex = 5
        Me.btnFilebrowse.Text = "..."
        Me.btnFilebrowse.UseVisualStyleBackColor = True
        '
        'txtFilepath
        '
        Me.txtFilepath.Font = New System.Drawing.Font("Verdana", 8.25!)
        Me.txtFilepath.Location = New System.Drawing.Point(122, 27)
        Me.txtFilepath.Name = "txtFilepath"
        Me.txtFilepath.Size = New System.Drawing.Size(418, 21)
        Me.txtFilepath.TabIndex = 4
        '
        'lblFilepath
        '
        Me.lblFilepath.AutoSize = True
        Me.lblFilepath.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFilepath.Location = New System.Drawing.Point(6, 35)
        Me.lblFilepath.Name = "lblFilepath"
        Me.lblFilepath.Size = New System.Drawing.Size(94, 13)
        Me.lblFilepath.TabIndex = 3
        Me.lblFilepath.Text = "Select File Path"
        '
        'btnUpload
        '
        Me.btnUpload.Font = New System.Drawing.Font("Verdana", 8.25!)
        Me.btnUpload.Location = New System.Drawing.Point(9, 168)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(142, 23)
        Me.btnUpload.TabIndex = 6
        Me.btnUpload.Text = "Update Cheque No."
        Me.btnUpload.UseVisualStyleBackColor = True
        '
        'frmUFFGen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(599, 534)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.TabCon)
        Me.Controls.Add(Me.txtStatusMsg)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUFFGen"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FHG - UFF Generation"
        Me.TabCon.ResumeLayout(False)
        Me.tabGenDBS.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.tabUpdateDBS.ResumeLayout(False)
        Me.tabUpdateDBS.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtStatusMsg As System.Windows.Forms.TextBox
    Friend WithEvents TabCon As System.Windows.Forms.TabControl
    Friend WithEvents tabGenDBS As System.Windows.Forms.TabPage
    Friend WithEvents tabUpdateDBS As System.Windows.Forms.TabPage
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents lstBatch As System.Windows.Forms.ListBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents txtBatchno As System.Windows.Forms.TextBox
    Friend WithEvents lblBatchNo As System.Windows.Forms.Label
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents lblFileName As System.Windows.Forms.Label
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txtpath As System.Windows.Forms.TextBox
    Friend WithEvents lblPathfolder As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents datpick As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbCompany As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnFilebrowse As System.Windows.Forms.Button
    Friend WithEvents txtFilepath As System.Windows.Forms.TextBox
    Friend WithEvents lblFilepath As System.Windows.Forms.Label
    Friend WithEvents btnUpload As System.Windows.Forms.Button
End Class
