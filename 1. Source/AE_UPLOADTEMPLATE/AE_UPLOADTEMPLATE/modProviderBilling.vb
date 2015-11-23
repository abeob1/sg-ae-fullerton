Module modProviderBilling

    Private sGLAcct As String

    Private Sub Read_ProviderBillingFile(ByVal sFileName As String, _
                               ByVal sSheet As String, _
                               ByRef bIsError As Boolean, _
                               ByRef dv As DataView, _
                               ByRef sErrdesc As String)

        Dim iHeaderRow As Integer
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "Read_ProviderBillingFile"
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

        If dv(iHeaderRow)(1).ToString.Trim <> "Patient NRIC" Then
            sErrdesc = "Invalid Excel file Format - [Patient NRIC] not found at Column 2"
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

        If dv(iHeaderRow)(3).ToString.Trim <> "Customer Code" Then
            sErrdesc = "Invalid Excel file Format - [Customer Code] not found at Column 4"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString.Trim <> "Hospitalisation Admisison" Then
            sErrdesc = "Invalid Excel file Format - [Hospitalisation Admisison] not found at Column 5"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(5).ToString.Trim <> "Provider" Then
            sErrdesc = "Invalid Excel file Format - [Provider] not found at Column 6"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(6).ToString.Trim <> "Provider Address" Then
            sErrdesc = "Invalid Excel file Format - [Provider Address] not found at Column 7"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(7).ToString.Trim <> "Admission Date" Then
            sErrdesc = "Invalid Excel file Format - [Admission Date] not found at Column 8"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString.Trim <> "Discharge Date" Then
            sErrdesc = "Invalid Excel file Format - [Discharge Date] not found at Column 9"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(9).ToString.Trim <> "Policy Number" Then
            sErrdesc = "Invalid Excel file Format - [Policy Number] not found at Column 10"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString.Trim <> "Procedure Name" Then
            sErrdesc = "Invalid Excel file Format - [Procedure Name] not found at Column 11"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(11).ToString.Trim <> "Doctor's Name" Then
            sErrdesc = "Invalid Excel file Format - [Doctor's Name] not found at Column 12"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(12).ToString.Trim <> "Doctor's Professional Fee" Then
            sErrdesc = "Invalid Excel file Format - [Doctor's Professional Fee] not found at Column 13"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(13).ToString.Trim <> "Doctor's Inpatient Consultation" Then
            sErrdesc = "Invalid Excel file Format - [Doctor's Inpatient Consultation] not found at Column 14"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(14).ToString.Trim <> "Medications" Then
            sErrdesc = "Invalid Excel file Format - [Medications] not found at Column 15"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(15).ToString.Trim <> "Other Services" Then
            sErrdesc = "Invalid Excel file Format - [Other Services] not found at Column 16"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(16).ToString.Trim <> "Management Service Fee" Then
            sErrdesc = "Invalid Excel file Format - [Management Service Fee] not found at Column 17"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(17).ToString.Trim <> "GST" Then
            sErrdesc = "Invalid Excel file Format - [GST] not found at Column 18"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(18).ToString.Trim <> "Grand Total" Then
            sErrdesc = "Invalid Excel file Format - [Grand Total] not found at Column 19"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        
    End Sub

    Public Function UploadDocument_PROVIDERBILLING(ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_PROVIDERBILLING()"
        Dim myfile As New System.IO.FileInfo(sFileName)
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim odv As DataView = Nothing
        Dim iCnt As Integer = 0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            frmUpload.WriteToStatusScreen(False, "Validating Excel file format.....")

            bIsError = False
            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling Read_ProviderBillingFile()", sFuncName)
            Read_ProviderBillingFile(sFileName, "Sheet1", bIsError, odv, sErrDesc)

            If bIsError = True Then
                frmUpload.WriteToStatusScreen(False, " ERROR :::  " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Billl To Client Excel Worksheet", sFuncName)
                WriteToLogFile("Invalid Excel Worksheet " & sFileName, sFuncName)
                MsgBox("Invalid Excel Worksheet", MsgBoxStyle.Critical, "NON-FHN3 Upload")
                GoTo ExitFunc
            End If


            Dim j As Integer = Microsoft.VisualBasic.InStrRev(frmUpload.txtFileName.Text, "\")
            Dim sFName1 As String = Microsoft.VisualBasic.Right(frmUpload.txtFileName.Text, Len(frmUpload.txtFileName.Text) - j).Trim
            Dim sName As String = Replace(sFName1, ".xlsx", "")


            If IsFileNameExists("AR", sName) = True Then
                frmUpload.WriteToStatusScreen(False, " ERROR :::  " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Duplicate File Name", sFuncName)
                WriteToLogFile("Duplicate Import..This file already imported. Please check. " & sName, sFuncName)
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

                frmUpload.WriteToStatusScreen(False, "Uploading A/R Invoices..")
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_ARInvoice_PROVIDERBILLING()", sFuncName)
                If UploadDocument_ARInvoice_PROVIDERBILLING(odv, sErrDesc) <> RTN_SUCCESS Then
                    frmUpload.WriteToStatusScreen(False, "========================== COMPLETED WITH ERROR =======================================")
                    Throw New ArgumentException(sErrDesc)
                    'error condition..
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
            UploadDocument_PROVIDERBILLING = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_PROVIDERBILLING = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Private Function AddARInvoice_PROVIDERBILLING(ByVal oDt As DataTable, _
                                        ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim bLine As Boolean = False
        Dim dGrandTotal As Double = 0
        Dim dConsultCost As Double = 0
        Dim dDrugCost As Double = 0
        Dim dInhouseServCost As Double = 0
        Dim dExtServCost As Double = 0
        Dim dTax As Double = 0
        Dim dSubTotal As Double = 0
        Dim dUnClaimAmt As Double = 0
        Dim dCoPayment As Double = 0

        Try

            sFuncName = "AddARInvoice_PROVIDERBILLING"


            Dim k As Integer = Microsoft.VisualBasic.InStrRev(frmUpload.txtFileName.Text, "\")
            Dim sFName As String = Microsoft.VisualBasic.Right(frmUpload.txtFileName.Text, Len(frmUpload.txtFileName.Text) - k).Trim
            Dim sFileName As String = Replace(sFName, ".xlsx", "")


            'If IsFileNameExists("AR", sFileName) = True Then
            '    frmUpload.WriteToStatusScreen(False, " ERROR :::  " & sErrDesc)
            '    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Duplicate File Name", sFuncName)
            '    WriteToLogFile("Duplicate Import..This file already imported. Please check. " & sFileName, sFuncName)
            '    MsgBox("Duplicate Import..This file already imported. Please check. ", MsgBoxStyle.Critical, "Upload Template")
            '    GoTo ExitFunc
            'End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            oDoc.CardCode = oDt.Rows(1).Item("F2").ToString
            oDoc.DocDate = CDate(oDt.Rows(2).Item("F2").ToString)
            sGLAcct = oDt.Rows(4).Item("F2").ToString

            oDoc.NumAtCard = Left(oDt.Rows(7).Item("F3").ToString & " " & oDt.Rows(7).Item("F4").ToString & " " & oDt.Rows(2).Item("F2").ToString, 100)

            oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

            Dim iCnt As Integer = 0
            'l
            For Each Row As DataRow In oDt.Rows

                iCnt += 1
                If iCnt >= 8 Then
                    If iCnt > 8 Then
                        oDoc.Lines.Add()
                    End If

                    oDoc.Lines.AccountCode = sGLAcct
                    oDoc.Lines.ItemDescription = Row(10).ToString

                    If CDbl(Row(17)) > 0 Then
                        oDoc.Lines.VatGroup = "SO"
                    Else
                        oDoc.Lines.VatGroup = "ZO"
                    End If

                    'oDoc.Lines.PriceAfterVAT = CDbl(Row(18))

                    oDoc.Lines.LineTotal = CDbl(Row(16))

                    oDoc.Lines.UserFields.Fields.Item("U_AI_CompanyName").Value = Row(2).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderAddress").Value = Row(6).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = Row(0).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_PolicyNo").Value = Row(9).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderName").Value = Row(5).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_Admitdate").Value = CDate(Row(7).ToString)
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = Row(1).ToString  ' Member ID
                    oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = CDate(Row(8).ToString)  ' Discharge date
                    oDoc.Lines.UserFields.Fields.Item("U_AI_DoctorName").Value = Row(11).ToString   'DoctorName
                    oDoc.Lines.UserFields.Fields.Item("U_AI_DocProfFee").Value = CDbl(Row(12))  ' Doctor's Professional Fee
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = CDbl(Row(13))  ' Doctor's Inpatient Consultation
                    oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = CDbl(Row(14))  ' Medications
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = CDbl(Row(15))  'Other services
                    oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = CDbl(Row(16))  'Management Service Fee
                    oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = CDbl(Row(17))  'GST

                    'Cost Center
                    oDoc.Lines.CostingCode = oDt.Rows(3).Item("F2").ToString
                End If
            Next

            frmUpload.WriteToStatusScreen(False, "Creating A/R Invoice..")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR Invoice.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                p_oCompany.GetLastError(lErrCode, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR Invoice failed.", sFuncName)
                frmUpload.WriteToStatusScreen(False, "Adding AR Invoice Failed")
                frmUpload.WriteToStatusScreen(False, sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If
ExitFunc:
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice_PROVIDERBILLING = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_PROVIDERBILLING = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Public Function UploadDocument_ARInvoice_PROVIDERBILLING(ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_ARInvoice_PROVIDERBILLING()"
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim iCnt As Integer = 0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If AddARInvoice_PROVIDERBILLING(odv.Table, sErrDesc) <> RTN_SUCCESS Then
                frmUpload.WriteToStatusScreen(False, "ERROR :: Rollback Transaction..")
                frmUpload.WriteToStatusScreen(False, "ERROR :: Failed to create Invoice for Line ::" & iCnt - 10)
                frmUpload.WriteToStatusScreen(False, "ERROR :: " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File was not successfully uploaded" & frmUpload.txtFileName.Text, sFuncName)
                If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If

            frmUpload.WriteToStatusScreen(False, "Successfully created A/R Invoice..")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            UploadDocument_ARInvoice_PROVIDERBILLING = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_ARInvoice_PROVIDERBILLING = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

End Module
