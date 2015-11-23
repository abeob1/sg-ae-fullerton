
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Text


Public Class frmUFFGen


    Public Sub WriteToStatusScreen(ByVal Clear As Boolean, ByVal msg As String)
        If Clear Then
            txtStatusMsg.Text = ""
        End If
        txtStatusMsg.HideSelection = True
        txtStatusMsg.Text &= msg & vbCrLf
        txtStatusMsg.SelectAll()
        txtStatusMsg.ScrollToCaret()
        txtStatusMsg.Refresh()
    End Sub

    Private Sub frmPDFGen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim sErrdesc As String = String.Empty
        If GetSystemIntializeInfo(p_oCompDef, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
        Me.txtFileName.Text = Now.Date.ToString("yyyyMMdd") & "_" & DateTime.Now.ToString("HH.mm.ss") & p_oCompDef.sOrgID
        Me.cmbCompany.SelectedIndex = 0
    End Sub

    Private Sub btnGenerate_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerate.Click
        Dim sErrdesc As String = String.Empty
        Dim sBatch As String = String.Empty
        Dim sFuncName As String = String.Empty
        sFuncName = "btnGenerate_Click()"
        txtStatusMsg.Clear()
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            WriteToStatusScreen(False, "Please Wait ....")

            If String.IsNullOrEmpty(Me.txtpath.Text) Then
                WriteToStatusScreen(False, "Validation Msg: Folder path cannot be blank ")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Msg: Folder path cannot be blank ", sFuncName)
                WriteToStatusScreen(False, "Completed with Validation Error ....")
                MsgBox("Validation Msg: Folder path cannot be blank", MsgBoxStyle.OkOnly, "DBS UFF File Generation")
                Exit Sub
            End If

            If Me.datpick.Checked = False Then
                WriteToStatusScreen(False, "Validation Msg: Date cannot be blank ")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Msg: Date cannot be blank  ", sFuncName)
                WriteToStatusScreen(False, "Completed with Validation Error ....")
                MsgBox("Validation Msg: Date cannot be blank", MsgBoxStyle.OkOnly, "DBS UFF File Generation")
                Exit Sub
            End If

            If Me.lstBatch.Items.Count = 0 Then
                WriteToStatusScreen(False, "Validation Msg: Batch No. cannot be blank ")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Msg: Batch No. cannot be blank ", sFuncName)
                WriteToStatusScreen(False, "Completed with Validation Error ....")
                MsgBox("Validation Msg: Batch No. cannot be blank", MsgBoxStyle.OkOnly, "DBS UFF File Generation")
                Exit Sub
            End If

            If Me.cmbCompany.SelectedItem = "--Select--" Then
                WriteToStatusScreen(False, "Validation Msg: Template Type cannot be blank ")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Msg: Template Type cannot be blank ", sFuncName)
                WriteToStatusScreen(False, "Completed with Validation Error ....")
                MsgBox("Validation Msg: Template Type cannot be blank ", MsgBoxStyle.OkOnly, "DBS UFF File Generation")
                Exit Sub
            End If

            If Me.cmbCompany.SelectedItem = "AON" Then
                sTemplateType = "S"
            Else
                sTemplateType = "M"
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Template Type " & sTemplateType, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function UFF_Generation()", sFuncName)
            If UFF_Generation(Me.datpick.Value, Me.txtpath.Text, Me.txtFileName.Text, sErrdesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrdesc)
                ''MsgBox(sErrdesc)

            End If

            WriteToStatusScreen(False, "Encrypting the file...")
            Dim oFromDirInfo = New DirectoryInfo(txtpath.Text.Trim())
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling MoveAllItemsTo()", sFuncName)
            If MoveAllItemsTo(oFromDirInfo, p_oCompDef.p_sPaymentDir, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Encryption()", sFuncName)
            If Encryption(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
            WriteToStatusScreen(False, "Encryption completed successfully.")

            WriteToStatusScreen(False, "------------------------- Process Completed ------------------------")

        Catch ex As Exception
            MsgBox("Error : " & ex.Message, MsgBoxStyle.OkOnly, "DBD UFF File Generation")
            WriteToStatusScreen(False, "------------------------- Process Completed ------------------------")
        End Try
       
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.Dispose()
    End Sub

    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click
        Try
            Dim Myfloderbrowser As New FolderBrowserDialog

            Myfloderbrowser.Description = "Select the folder to store UFF File"
            Myfloderbrowser.ShowNewFolderButton = True
            Myfloderbrowser.RootFolder = System.Environment.SpecialFolder.MyComputer
            If Myfloderbrowser.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.txtpath.Text = Myfloderbrowser.SelectedPath & "\"
            End If
        Catch ex As Exception
           
        End Try
    End Sub

    Private Sub btnView_Click(sender As System.Object, e As System.EventArgs) Handles btnView.Click
        Process.Start(Me.txtpath.Text & Me.txtFileName.Text & ".txt")
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        txtStatusMsg.Clear()
        txtBatchno.Clear()
        lstBatch.Items.Clear()
        Me.cmbCompany.SelectedIndex = 0

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        If Not String.IsNullOrEmpty(txtBatchno.Text) Then
            lstBatch.Items.Add(txtBatchno.Text)
            txtBatchno.Clear()
        End If
       
    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        Try
            If Not String.IsNullOrEmpty(Me.lstBatch.SelectedItem.ToString) Then
                lstBatch.Items.Remove(Me.lstBatch.SelectedItem.ToString)
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub btnFilebrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnFilebrowse.Click
        Try
            Dim MyFilebrowser As New OpenFileDialog

            MyFilebrowser.Multiselect = False
            MyFilebrowser.RestoreDirectory = True
            MyFilebrowser.Title = "Select File to upload in SAP B1"
            If MyFilebrowser.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.txtFilepath.Text = MyFilebrowser.FileName
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class
