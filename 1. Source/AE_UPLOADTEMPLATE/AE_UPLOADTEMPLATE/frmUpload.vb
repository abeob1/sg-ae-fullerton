Imports System.IO

Public Class frmUpload

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Me.txtFileName.Text = String.Empty Then
            MsgBox("Please Select File", MsgBoxStyle.Information, "Non-FHN3 Batch Upload")
            Exit Sub
        End If
        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        WriteToStatusScreen(True, "Please wait File Upload in progress ...")
        Me.Button1.Enabled = Not Me.Button1.Enabled

        Call btnUpload_Click(Me, New System.EventArgs)

        Me.Button1.Enabled = Not Me.Button1.Enabled

        System.Windows.Forms.Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.txtStatusMsg.Clear()

    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click

        OpenFD.InitialDirectory = "C:\"
        OpenFD.Filter = "Excel Files(*.xlsx)|*.xlsx|All Files(*.*)|*.*"
        OpenFD.ShowDialog()
        Me.txtFileName.Text = OpenFD.FileName

    End Sub

    Private Sub UploadFile()

        Dim sFuncName As String = "UploadFile()"
        Dim sErrdesc As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            p_oDtReceiptLog = CreateDataTable("Customer Code", "Inv. Number", "Payment Method", "Error Description")

            WriteToStatusScreen(False, "========================== S T A R T =======================================")

            If Me.DocType.Text = "BUPA" Then
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_BUPA()", sFuncName)
                If UploadDocument_BUPA(Me.txtFileName.Text, sErrdesc) <> RTN_SUCCESS Then
                    WriteToStatusScreen(False, "========================== COMPLETED WITH ERROR =======================================")
                    'error condition.
                End If
            End If

            If Me.DocType.Text = "AETNA" Then
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_AETNA()", sFuncName)
                If UploadDocument_AETNA(Me.txtFileName.Text, sErrdesc) <> RTN_SUCCESS Then
                    WriteToStatusScreen(False, "========================== COMPLETED WITH ERROR =======================================")
                    'error condition.
                End If
            End If

            If Me.DocType.Text = "CUSTOMER BILLING" Then
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_CUSTOMERBILLING()", sFuncName)
                If UploadDocument_CUSTOMERBILLING(Me.txtFileName.Text, sErrdesc) <> RTN_SUCCESS Then
                    WriteToStatusScreen(False, "========================== COMPLETED WITH ERROR =======================================")
                    'error condition.
                End If
            End If

            If Me.DocType.Text = "PROVIDER BILLING" Then
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_PROVIDERBILLING()", sFuncName)
                If UploadDocument_PROVIDERBILLING(Me.txtFileName.Text, sErrdesc) <> RTN_SUCCESS Then
                    WriteToStatusScreen(False, "========================== COMPLETED WITH ERROR =======================================")
                    'error condition.
                End If
            End If

            'RECEIPT UPLOAD
            If Me.DocType.Text = "RECEIPT UPLOAD" Then
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_ReceiptUpload()", sFuncName)
                If UploadDocument_ReceiptUpload(Me.txtFileName.Text, sErrdesc) <> RTN_SUCCESS Then
                    WriteToStatusScreen(False, "========================== COMPLETED WITH ERROR =======================================")
                    'error condition.
                End If
            End If

            'Write the errors in notepad file and open that file

            If p_oDtReceiptLog.Rows.Count > 0 Then
                Write_TextFile(p_oDtReceiptLog, "ReceiptLog.txt", sErrdesc)
            End If
            p_oDtReceiptLog.Rows.Clear()

            WriteToStatusScreen(False, "========================== COMPLETED =======================================")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed successfully.", sFuncName)
        Catch ex As Exception
            sErrdesc = ex.Message
            Call WriteToLogFile(sErrdesc, sFuncName)
            Throw New ArgumentException(sErrdesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        End Try
    End Sub

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Dim sErrdesc As String = String.Empty
        If GetSystemIntializeInfo(p_oCompDef, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
        UploadFile()
    End Sub

    Private Sub ReadAPFile(ByVal sFileName As String, _
                                ByVal sSheet As String, _
                                ByRef bIsError As Boolean, _
                                ByRef dv As DataView, _
                                ByRef sErrdesc As String)

        Dim iHeaderRow As Integer
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "ReadAPFile"
        Dim sBatchNo As String = String.Empty

        iHeaderRow = 0

        dv = GetDataViewFromExcel(sFileName, sSheet, sErrdesc)


        If IsNothing(dv) Then
            Exit Sub
        End If

        If dv(iHeaderRow)(0).ToString.Trim <> "Vendor Code" Then
            sErrdesc = "Invalid Excel file Format - [Vendor Code] not found at Column 1"
            WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(1).ToString.Trim <> "Vendor Name" Then
            sErrdesc = "Invalid Excel file Format - [Vendor Name] not found at Column 2"
            WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(2).ToString.Trim <> "Posting Date" Then
            sErrdesc = "Invalid Excel file Format - [Posting Date] not found at Column 3"
            WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(3).ToString.Trim <> "Due Date" Then
            sErrdesc = "Invalid Excel file Format - ([Due Date] not found at Column 4"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString.Trim <> "Document Date" Then
            sErrdesc = "Invalid Excel file Format - ([Document Date] not found at Column 5"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(5).ToString.Trim <> "Vendor Ref No" Then
            sErrdesc = "Invalid Excel file Format - ([Vendor Ref No] not found at Column 6"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(6).ToString.Trim <> "LineNumber" Then
            sErrdesc = "Invalid Excel file Format - ([LineNumber] not found at Column 7"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(7).ToString.Trim <> "Description" Then
            sErrdesc = "Invalid Excel file Format - [Description] not found at Column 8"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString.Trim <> "GL" Then
            sErrdesc = "Invalid Excel file Format - [GL] not found at Column 9"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(9).ToString.Trim <> "Distibution Rule" Then
            sErrdesc = "Invalid Excel file Format - [Distibution Rule] not found at Column 10"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString.Trim <> "Tax Code" Then
            sErrdesc = "Invalid Excel file Format - [Tax Code] not found at Column 11"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(11).ToString.Trim <> "BP Currency" Then
            sErrdesc = "Invalid Excel file Format - [BP Currency] not found at Column 12"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(12).ToString.Trim <> "Total (LC)" Then
            sErrdesc = "Invalid Excel file Format - [Total (LC)] not found at Column 13"
            WriteToLogFile(False, sErrdesc)
            WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

    End Sub

    Private Sub frmUpload_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub


End Class
