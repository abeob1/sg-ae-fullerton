Module modCustomerBilling

    Private sGLAcct As String = String.Empty

    Public Function UploadDocument_CUSTOMERBILLING(ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_CUSTOMERBILLING()"
        Dim myfile As New System.IO.FileInfo(sFileName)
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim odv As DataView = Nothing
        Dim iCnt As Integer = 0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim j As Integer = Microsoft.VisualBasic.InStrRev(frmUpload.txtFileName.Text, "\")
            Dim sFName1 As String = Microsoft.VisualBasic.Right(frmUpload.txtFileName.Text, Len(frmUpload.txtFileName.Text) - j).Trim
            Dim sFile As String = Replace(sFName1, ".xlsx", "")

            frmUpload.WriteToStatusScreen(False, "Validating Excel file format.....")

            bIsError = False
            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling Read_CustomerBillingFile()", sFuncName)
            Read_CustomerBillingFile(sFileName, "Sheet1", bIsError, odv, sErrDesc)

            If bIsError = True Then
                frmUpload.WriteToStatusScreen(False, " ERROR :::  " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Customer Billing Excel Worksheet", sFuncName)
                WriteToLogFile("Invalid Excel Worksheet " & sFileName, sFuncName)
                MsgBox("Invalid Excel Worksheet", MsgBoxStyle.Critical, "NON-FHN3 Upload")
                GoTo ExitFunc
            End If

            If IsFileNameExists("AR", sFile) = True Then
                frmUpload.WriteToStatusScreen(False, " ERROR :::  " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Duplicate File Name", sFuncName)
                WriteToLogFile("Duplicate Import..This file already imported. Please check. " & sFile, sFuncName)
                MsgBox("Duplicate Import..This file already imported. Please check. ", MsgBoxStyle.Critical, "Upload Template")
                GoTo ExitFunc
            End If

            frmUpload.WriteToStatusScreen(False, "Successfully Validated Excel file format.")

            frmUpload.WriteToStatusScreen(False, "Connecting to SAP Database...")

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ConnectToCompany()", sFuncName)
            If ConnectToCompany(p_oCompany, sErrDesc) <> RTN_SUCCESS Then
                frmUpload.WriteToStatusScreen(False, "ERROR:: Unable to connect to SAP." & sErrDesc)
                Exit Function
            End If

            frmUpload.WriteToStatusScreen(False, "Successfully connected to SAP Database...")

            If p_oCompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction..", sFuncName)
                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                frmUpload.WriteToStatusScreen(False, "Uploading A/R Invoices for Customer Billing.")
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_ARInvoice_CUSTOMERBILLING()", sFuncName)
                If UploadDocument_ARInvoice_CUSTOMERBILLING(odv, sErrDesc) <> RTN_SUCCESS Then
                    frmUpload.WriteToStatusScreen(False, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                    'error condition.
                End If
                frmUpload.WriteToStatusScreen(False, "Successfully Uploaded A/R Invoices..")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction()", sFuncName)
                If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Moving " & frmUpload.txtFileName.Text & " to " & p_oCompDef.sSuccessDir, sFuncName)
                Dim UploadedFileName As String = Mid(frmUpload.txtFileName.Text, 1, frmUpload.txtFileName.Text.Length - 5) & "_" & Now.ToString("yyyyMMddhhmmss") & ".txt"
                frmUpload.WriteToStatusScreen(False, "Moving file to Success Folder")

                Dim k As Integer = Microsoft.VisualBasic.InStrRev(UploadedFileName, "\")
                Dim sFName As String = Microsoft.VisualBasic.Right(UploadedFileName, Len(UploadedFileName) - k).Trim

                System.IO.File.Move(frmUpload.txtFileName.Text, p_oCompDef.sSuccessDir & "\" & Replace(sFName, ".txt", ".xlsx"))

                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Disconnecting from SAP Databases", sFuncName)
                p_oCompany.Disconnect()

                frmUpload.WriteToStatusScreen(False, "Disconnected from SAP Database.")

                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Disconnected from SAP Databases", sFuncName)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Unable to connect to SAP.", sFuncName)
                Throw New ArgumentException("Unable to connect to SAP.")
            End If
ExitFunc:
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            UploadDocument_CUSTOMERBILLING = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_CUSTOMERBILLING = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Private Sub Read_CustomerBillingFile(ByVal sFileName As String, _
                               ByVal sSheet As String, _
                               ByRef bIsError As Boolean, _
                               ByRef dv As DataView, _
                               ByRef sErrdesc As String)

        Dim iHeaderRow As Integer
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "Read_CustomerBillingFile"
        Dim sBatchNo As String = String.Empty

        iHeaderRow = 6

        dv = GetDataViewFromExcel(sFileName, sSheet, sErrdesc)


        If IsNothing(dv) Then
            bIsError = True
            Exit Sub
        End If

        If dv(iHeaderRow)(0).ToString.Trim <> "Patient Name" Then
            sErrdesc = "Invalid Excel file Format - [Patient Name] not found at Column 1"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(1).ToString.Trim <> "Patient ID" Then
            sErrdesc = "Invalid Excel file Format - [Patient ID] not found at Column 2"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(2).ToString.Trim <> "Company Name" Then
            sErrdesc = "Invalid Excel file Format - [Company Name] not found at Column 3"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(3).ToString.Trim <> "Hospitalisation Admisison" Then
            sErrdesc = "Invalid Excel file Format - [Hospitalisation Admisison] not found at Column 4"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString.Trim <> "Provider" Then
            sErrdesc = "Invalid Excel file Format - [Provider] not found at Column 5"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(5).ToString.Trim <> "Doctor's Name" Then
            sErrdesc = "Invalid Excel file Format - [Doctor's Name] not found at Column 6"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(6).ToString.Trim <> "Admission Date" Then
            sErrdesc = "Invalid Excel file Format - [Admission Date] not found at Column 7"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(7).ToString.Trim <> "Discharge Date" Then
            sErrdesc = "Invalid Excel file Format - [Discharge Date] not found at Column 8"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString.Trim <> "Policy Number" Then
            sErrdesc = "Invalid Excel file Format - [Policy Number] not found at Column 9"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(9).ToString.Trim <> "Procedure Name" Then
            sErrdesc = "Invalid Excel file Format - [Procedure Name] not found at Column 10"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString.Trim <> "Pre Hospitalisation Consult" Then
            sErrdesc = "Invalid Excel file Format - [Pre Hospitalisation Consult] not found at Column 11"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(11).ToString.Trim <> "Pre Hospitalisation Medication" Then
            sErrdesc = "Invalid Excel file Format - [Pre Hospitalisation Medication] not found at Column 12"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(12).ToString.Trim <> "Pre Hospitalisation Procedure" Then
            sErrdesc = "Invalid Excel file Format - [Pre Hospitalisation Procedure] not found at Column 13"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(13).ToString.Trim <> "Pre Hospitalisation (Sub-Total)" Then
            sErrdesc = "Invalid Excel file Format - [Pre Hospitalisation (Sub-Total)] not found at Column 14"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(14).ToString.Trim <> "Post Hospitalisation Consult" Then
            sErrdesc = "Invalid Excel file Format - [Post Hospitalisation Consult] not found at Column 15"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(15).ToString.Trim <> "Post Hospitalisation Medication" Then
            sErrdesc = "Invalid Excel file Format - [Post Hospitalisation Medication] not found at Column 16"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(16).ToString.Trim <> "Post Hospitalisation Procedure" Then
            sErrdesc = "Invalid Excel file Format - [Post Hospitalisation Procedure] not found at Column 17"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(17).ToString.Trim <> "Post Hospitalization (Sub-Total)" Then
            sErrdesc = "Invalid Excel file Format - [Post Hospitalization (Sub-Total)] not found at Column 18"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(18).ToString.Trim <> "Daily Room and Board" Then
            sErrdesc = "Invalid Excel file Format - [Daily Room and Board] not found at Column 19"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(19).ToString.Trim <> "Hospital Facilities Fees" Then
            sErrdesc = "Invalid Excel file Format - [Hospital Facilities Fees] not found at Column 20"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If



        If dv(iHeaderRow)(20).ToString.Trim <> "In-Hospital Doctor's Consultation" Then
            sErrdesc = "Invalid Excel file Format - [In-Hospital Doctor's Consultation] not found at Column 21"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(21).ToString.Trim <> "Doctor's Professional Fee" Then
            sErrdesc = "Invalid Excel file Format - [Doctor's Professional Fee] not found at Column 22"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(22).ToString.Trim <> "Anaesthetist Fee" Then
            sErrdesc = "Invalid Excel file Format - [ Anaesthetist Fee] not found at Column 23"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(23).ToString.Trim <> "Medications or Other Equipments" Then
            sErrdesc = "Invalid Excel file Format - [Medications or Other Equipments] not found at Column 24"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(24).ToString.Trim <> "Hospitalization (Sub-Total)" Then
            sErrdesc = "Invalid Excel file Format - [Hospitalization (Sub-Total)] not found at Column 25"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(25).ToString.Trim <> "Total" Then
            sErrdesc = "Invalid Excel file Format - [Total] not found at Column 26"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(26).ToString.Trim <> "GST" Then
            sErrdesc = "Invalid Excel file Format - [GST] not found at Column 27"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(27).ToString.Trim <> "Grand Total" Then
            sErrdesc = "Invalid Excel file Format - [Grand Total] not found at Column 28"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(28).ToString.Trim <> "Co-insurance for Ward Upgrade" Then
            sErrdesc = "Invalid Excel file Format - [Co-insurance for Ward Upgrade] not found at Column 29"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(29).ToString.Trim <> "Co-insurance for Supplementary Major Medical (SMM)" Then
            sErrdesc = "Invalid Excel file Format - [Co-insurance for Supplementary Major Medical (SMM)] not found at Column 30"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(30).ToString.Trim <> "Balance Payable" Then
            sErrdesc = "Invalid Excel file Format - [Balance Payable] not found at Column 31"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If


    End Sub

    Public Function UploadDocument_ARInvoice_CUSTOMERBILLING(ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_ARInvoice_CUSTOMERBILLING()"
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim iCnt As Integer = 0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            For Each row As DataRow In odv.Table.Rows
                iCnt += 1
                If iCnt >= 8 Then

                    If row.Item(1).ToString = String.Empty Then Exit For

                    frmUpload.WriteToStatusScreen(False, "Creating Invoice for Line ::" & iCnt - 9)
                    If AddARInvoice_CUSTOMERBILLING(odv.Table, row, sErrDesc) <> RTN_SUCCESS Then
                        frmUpload.WriteToStatusScreen(False, "ERROR :: Rollback Transaction..")
                        frmUpload.WriteToStatusScreen(False, "ERROR :: Failed to create Invoice for Line ::" & iCnt - 9)
                        frmUpload.WriteToStatusScreen(False, "ERROR :: " & sErrDesc)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File was not successfully uploaded" & frmUpload.txtFileName.Text, sFuncName)
                        If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    End If
                    frmUpload.WriteToStatusScreen(False, "Successfully created Invoice for Line ::" & iCnt - 9)
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            UploadDocument_ARInvoice_CUSTOMERBILLING = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_ARInvoice_CUSTOMERBILLING = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Private Function AddARInvoice_CUSTOMERBILLING(ByVal oDt As DataTable, _
                                      ByVal Row As DataRow, _
                                      ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim bLine As Boolean = False
        Try
            sFuncName = "AddARInvoice_CUSTOMERBILLING"


            Dim k As Integer = Microsoft.VisualBasic.InStrRev(frmUpload.txtFileName.Text, "\")
            Dim sFName As String = Microsoft.VisualBasic.Right(frmUpload.txtFileName.Text, Len(frmUpload.txtFileName.Text) - k).Trim
            Dim sFileName As String = Replace(sFName, ".xlsx", "")


            sCostCenter = oDt.Rows(3).Item("F2").ToString

            If sCostCenter = String.Empty Then
                sErrDesc = "Cost Center is blank "
                frmUpload.WriteToStatusScreen(False, "ERROR :" & sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            sGLAcct = oDt.Rows(4).Item("F2").ToString

            oDoc.CardCode = oDt.Rows(1).Item("F2").ToString
            oDoc.DocDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.NumAtCard = Row.Item(1).ToString & " " & Row.Item(5).ToString
            oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName


            'Line 1 Pre Hospitalisation Consult
            If Not IsDBNull(Row.Item(10)) Then
                If Convert.ToDouble(Row.Item(10)) <> 0 Then
                    bLine = True
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(10)), "Pre Hospitalisation Consult")
                End If
            End If

            'Line 2 Pre Hospitalisation Medication
            If Not IsDBNull(Row.Item(11)) Then
                If Convert.ToDouble(Row.Item(11)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(11)), "Pre Hospitalisation Medication")
                End If
            End If


            'Line 3 Pre Hospitalisation Procedure
            If Not IsDBNull(Row.Item(12)) Then
                If Convert.ToDouble(Row.Item(12)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(12)), "Pre Hospitalisation Procedure")
                End If
            End If

            'Line 4 Post Hospitalisation Consult
            If Not IsDBNull(Row.Item(14)) Then

                If CDbl(Row.Item(14)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(14)), "Post Hospitalisation Consult")
                End If
            End If

            'Line 5 Post Hospitalisation Medication
            If Not IsDBNull(Row.Item(15)) Then

                If CDbl(Row.Item(15)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(15)), "Post Hospitalisation Medication")
                End If
            End If

            'Line 6 Post Hospitalisation Procedure
            If Not IsDBNull(Row.Item(16)) Then

                If CDbl(Row.Item(16)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(16)), "Post Hospitalisation Procedure")
                End If
            End If

            'Line 7 Daily Room and Board
            If Not IsDBNull(Row.Item(18)) Then

                If CDbl(Row.Item(18)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(18)), "Daily Room and Board")
                End If
            End If

            'Line 8 Hospital Miscellaneous
            If Not IsDBNull(Row.Item(19)) Then

                If CDbl(Row.Item(19)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(19)), "Hospital Facilities Fees")
                End If
            End If


            'Line 9 In-Hospital Doctor's Consultation
            If Not IsDBNull(Row.Item(20)) Then

                If CDbl(Row.Item(20)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(20)), "In-Hospital Doctor's Consultation")
                End If
            End If


            'Line 10 Doctor's Professional Fee
            If Not IsDBNull(Row.Item(21)) Then

                If CDbl(Row.Item(21)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(21)), "Doctor's Professional Fee")
                End If
            End If

            'Line 11 Anaesthetist Fee
            If Not IsDBNull(Row.Item(22)) Then

                If CDbl(Row.Item(22)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(22)), "Anaesthetist Fee")
                End If
            End If


            'Line 12 Medications or Other Equipments
            If Not IsDBNull(Row.Item(23)) Then

                If CDbl(Row.Item(23)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(23)), "Medications or Other Equipments")
                End If
            End If


            'Line 13 Co-insurance for Ward Upgrade
            If Not IsDBNull(Row.Item(28)) Then

                If CDbl(Row.Item(28)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, -1 * CDbl(Row.Item(28)), "Co-insurance for Ward Upgrade", "ZO")
                End If
            End If


            'Line 14 Co-insurance for Supplementary Major Medical (SMM)
            If Not IsDBNull(Row.Item(29)) Then

                If CDbl(Row.Item(29)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, -1 * CDbl(Row.Item(29)), "Co-insurance for Supplementary Major Medical (SMM)", "ZO")
                End If
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR Invoice.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                p_oCompany.GetLastError(lErrCode, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR Invoice failed.", sFuncName)
                frmUpload.WriteToStatusScreen(False, "Adding AR Invoice Failed")
                frmUpload.WriteToStatusScreen(False, sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice_CUSTOMERBILLING = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_CUSTOMERBILLING = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Sub AddDocLine(ByVal oDoc As SAPbobsCOM.Documents, ByVal Row As DataRow, _
                        ByVal sCostCenter As String, _
                        ByVal dLineTotal As Double, _
                        ByVal sDesc As String, _
                        Optional ByVal sTax As String = "")

        If sTax = "ZO" Then oDoc.Lines.VatGroup = "ZO"
        oDoc.Lines.CostingCode = sCostCenter
        oDoc.Lines.LineTotal = dLineTotal
        oDoc.Lines.ItemDescription = sDesc
        oDoc.Lines.UserFields.Fields.Item("U_AI_CompanyName").Value = Row.Item(2).ToString
        'oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderAddress").Value = Row.Item(3).ToString
        oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = Row.Item(0).ToString
        oDoc.Lines.UserFields.Fields.Item("U_AI_PolicyNo").Value = Row.Item(8).ToString
        oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderName").Value = Row.Item(4)
        oDoc.Lines.UserFields.Fields.Item("U_AI_Admitdate").Value = Row.Item(6)
        'Patient ID
        oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = Row.Item(1).ToString
        'Discharge Date
        oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = Row.Item(7)
        'Procedure Name
        oDoc.Lines.UserFields.Fields.Item("U_AI_ProcedureName").Value = Row.Item(9).ToString

        oDoc.Lines.UserFields.Fields.Item("U_AI_DoctorName").Value = Row(5).ToString   'DoctorName


        oDoc.Lines.AccountCode = sGLAcct
    End Sub

End Module
