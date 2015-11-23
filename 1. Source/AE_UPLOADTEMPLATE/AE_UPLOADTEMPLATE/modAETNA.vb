Module modAETNA

    Public Function UploadDocument_AETNA(ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_AETNA()"
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
            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling Read_AETNAFile", sFuncName)
            Read_AETNAFile(sFileName, "Sheet1", bIsError, odv, sErrDesc)

            If bIsError = True Then
                frmUpload.WriteToStatusScreen(False, " ERROR :::  " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Billl To Client Excel Worksheet", sFuncName)
                WriteToLogFile("Invalid Excel Worksheet " & sFileName, sFuncName)
                MsgBox("Invalid Excel Worksheet", MsgBoxStyle.Critical, "NON-FHN3 Upload")
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
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_ARInvoice_AETNA()", sFuncName)
                If UploadDocument_ARInvoice_AETNA(odv, sErrDesc) <> RTN_SUCCESS Then
                    frmUpload.WriteToStatusScreen(False, "========================== COMPLETED WITH ERROR =======================================")
                    Throw New ArgumentException(sErrDesc)
                    'error condition.
                End If
                frmUpload.WriteToStatusScreen(False, "Successfully Uploaded A/R Invoices..")

                frmUpload.WriteToStatusScreen(False, "Uploading A/P Invoices..")
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_APInvoice_AETNA()", sFuncName)
                If UploadDocument_APInvoice_AETNA(odv, sErrDesc) <> RTN_SUCCESS Then
                    frmUpload.WriteToStatusScreen(False, "========================== COMPLETED WITH ERROR =======================================")
                    Throw New ArgumentException(sErrDesc)
                    'error condition.
                End If
                frmUpload.WriteToStatusScreen(False, "Successfully Uploaded A/P Invoices..")

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
            UploadDocument_AETNA = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_AETNA = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Private Sub Read_AETNAFile(ByVal sFileName As String, _
                               ByVal sSheet As String, _
                               ByRef bIsError As Boolean, _
                               ByRef dv As DataView, _
                               ByRef sErrdesc As String)

        Dim iHeaderRow As Integer
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "Read_AETNAFile"
        Dim sBatchNo As String = String.Empty

        iHeaderRow = 8

        dv = GetDataViewFromExcel(sFileName, sSheet, sErrdesc)


        If IsNothing(dv) Then
            bIsError = True
            Exit Sub
        End If

        If dv(iHeaderRow)(0).ToString.Trim <> "Visit No." Then
            sErrdesc = "Invalid Excel file Format - ([Visit No.] not found at Column 1"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(1).ToString.Trim <> "Visit Date" Then
            sErrdesc = "Invalid Excel file Format - [Visit Date] not found at Column 2"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(2).ToString.Trim <> "Policy No" Then
            sErrdesc = "Invalid Excel file Format - [Policy No] not found at Column 3"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(3).ToString.Trim <> "Patient Name" Then
            sErrdesc = "Invalid Excel file Format - [Patient Name] not found at Column 4"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString.Trim <> "Member Id" Then
            sErrdesc = "Invalid Excel file Format - [Member Id] not found at Column 5"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(5).ToString.Trim <> "Provider Name" Then
            sErrdesc = "Invalid Excel file Format - [Provider Name] not found at Column 6"
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

        'If dv(iHeaderRow)(7).ToString.Trim <> "Provider Invoice No." Then
        '    sErrdesc = "Invalid Excel file Format - [Provider Invoice No.] not found at Column 8"
        '    WriteToLogFile(False, sErrdesc)
        '    frmUpload.WriteToStatusScreen(False, sErrdesc)
        '    bIsError = True
        'End If


        If dv(iHeaderRow)(7).ToString.Trim <> "Diagnosis Description" Then
            sErrdesc = "Invalid Excel file Format - [Diagnosis Description] not found at Column 8"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString.Trim <> "Drug Description (Cost)" Then
            sErrdesc = "Invalid Excel file Format - [Drug Description (Cost)] not found at Column 9"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(9).ToString.Trim <> "In-house Service Description (Cost)" Then
            sErrdesc = "Invalid Excel file Format - [In-house Service Description (Cost)] not found at Column 10"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString.Trim <> "External Service Description (Cost)" Then
            sErrdesc = "Invalid Excel file Format - [External Service Description (Cost)] not found at Column 11"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(11).ToString.Trim <> "Currency" Then
            sErrdesc = "Invalid Excel file Format - [Currency] not found at Column 12"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(12).ToString.Trim <> "Consult Cost" Then
            sErrdesc = "Invalid Excel file Format - [Consult Cost] not found at Column 13"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(13).ToString.Trim <> "Drug Cost" Then
            sErrdesc = "Invalid Excel file Format - [Drug Cost] not found at Column 14"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(14).ToString.Trim <> "In-house Service Cost" Then
            sErrdesc = "Invalid Excel file Format - [Deductible / Co-insurance] not found at Column 15"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(15).ToString.Trim <> "External Service Cost" Then
            sErrdesc = "Invalid Excel file Format - [External Service Cost] not found at Column 16"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(16).ToString.Trim <> "Sub Total" Then
            sErrdesc = "Invalid Excel file Format - [Sub Total] not found at Column 17"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(17).ToString.Trim <> "Tax" Then
            sErrdesc = "Invalid Excel file Format - [Tax] not found at Column 18"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(18).ToString.Trim <> "Discount" Then
            sErrdesc = "Invalid Excel file Format - [Discount] not found at Column 19"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(21).ToString.Trim <> "Grand Total" Then
            sErrdesc = "Invalid Excel file Format - [Grand Total] not found at Column 20"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(22).ToString.Trim <> "Co-Payment" Then
            sErrdesc = "Invalid Excel file Format - [Co-Payment] not found at Column 21"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(23).ToString.Trim <> "Unclaim Amount" Then
            sErrdesc = "Invalid Excel file Format - (Unclaim Amount] not found at Column 22"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(24).ToString.Trim <> "Claim Amount (IDR)" Then
            sErrdesc = "Invalid Excel file Format - ([Cost Center-AP] not found at Column 23"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(25).ToString.Trim <> "Rate" Then
            sErrdesc = "Invalid Excel file Format - [Rate] not found at Column 24"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(24).ToString.Trim <> "Claim Amount (SGD)" Then
            sErrdesc = "Invalid Excel file Format - [Claim Amount (SGD)] not found at Column 25"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If




    End Sub

    Public Function UploadDocument_APInvoice_AETNA(ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_APInvoice_AETNA()"
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim iCnt As Integer = 0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If AddAPInvoice_AETNA(odv.Table, sErrDesc) <> RTN_SUCCESS Then
                frmUpload.WriteToStatusScreen(False, "ERROR :: Rollback Transaction..")
                frmUpload.WriteToStatusScreen(False, "ERROR :: Failed to create Invoice for Line ::" & iCnt - 10)
                frmUpload.WriteToStatusScreen(False, "ERROR :: " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File was not successfully uploaded" & frmUpload.txtFileName.Text, sFuncName)
                If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If

            frmUpload.WriteToStatusScreen(False, "Successfully created A/P Invoice..")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            UploadDocument_APInvoice_AETNA = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_APInvoice_AETNA = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Public Function UploadDocument_ARInvoice_AETNA(ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_ARInvoice_AETNA()"
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim iCnt As Integer = 0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

           
            If AddARInvoice_AETNA(odv.Table, sErrDesc) <> RTN_SUCCESS Then
                frmUpload.WriteToStatusScreen(False, "ERROR :: Rollback Transaction..")
                frmUpload.WriteToStatusScreen(False, "ERROR :: Failed to create Invoice for Line ::" & iCnt - 10)
                frmUpload.WriteToStatusScreen(False, "ERROR :: " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File was not successfully uploaded" & frmUpload.txtFileName.Text, sFuncName)
                If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If

            frmUpload.WriteToStatusScreen(False, "Successfully created A/R Invoice..")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            UploadDocument_ARInvoice_AETNA = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_ARInvoice_AETNA = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Private Function AddARInvoice_AETNA(ByVal oDt As DataTable, _
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

            sFuncName = "AddARInvoice_AETNA"


            Dim k As Integer = Microsoft.VisualBasic.InStrRev(frmUpload.txtFileName.Text, "\")
            Dim sFName As String = Microsoft.VisualBasic.Right(frmUpload.txtFileName.Text, Len(frmUpload.txtFileName.Text) - k).Trim
            Dim sFileName As String = Replace(sFName, ".xlsx", "")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            oDoc.CardCode = oDt.Rows(1).Item("F2").ToString
            oDoc.DocDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.DocDueDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.TaxDate = CDate(oDt.Rows(2).Item("F2").ToString)
            'oDoc.DocCurrency = oDt.Rows(9).Item("F13").ToString.Trim
            oDoc.NumAtCard = oDt.Rows(5).Item("F2").ToString
            'oDoc.DocRate = CDbl(Row.Item(21))
            oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

            Dim iCnt As Integer = 0

            sCostCenter = GetCostCenter(sCardCode)


            For Each LineRow As DataRow In oDt.Rows
                iCnt += 1
                If iCnt >= 10 Then
                    If IsDBNull(LineRow(19)) = True Then Exit For
                    dGrandTotal = dGrandTotal + CDbl(LineRow(24))
                    dConsultCost = dConsultCost + CDbl(LineRow(12))
                    dDrugCost = dDrugCost + CDbl(LineRow(13))
                    dInhouseServCost = dInhouseServCost + CDbl(LineRow(14))
                    dExtServCost = dExtServCost + CDbl(LineRow(15))
                    dTax = dTax + CDbl(LineRow(17))
                    dSubTotal = dSubTotal + CDbl(LineRow(16))
                    dUnClaimAmt = dUnClaimAmt + CDbl(LineRow(21))
                    dCoPayment = dCoPayment + CDbl(LineRow(20))
                End If
            Next

            ' Line 1

            oDoc.Lines.AccountCode = p_oCompDef.sAR_AETNAGL
            oDoc.Lines.ItemDescription = oDt.Rows(6).Item("F2").ToString

            oDoc.Lines.CostingCode = sCostCenter

            If dTax > 0 Then
                oDoc.Lines.VatGroup = "SO"
            Else
                oDoc.Lines.VatGroup = "ZO"
            End If

            oDoc.Lines.PriceAfterVAT = dGrandTotal
            oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConsultCost
            oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost
            oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInhouseServCost
            oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal
            oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dTax
            oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal
            oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnClaimAmt

            '' Line 2
            'If dCoPayment > 0 Then
            '    oDoc.Lines.Add()
            '    oDoc.Lines.AccountCode = "5-32020-00"
            '    oDoc.Lines.ItemDescription = "Co-Payment"
            '    oDoc.Lines.VatGroup = "ZO"
            '    oDoc.Lines.LineTotal = -1 * dCoPayment
            'End If

            '' Line 3
            'If dUnClaimAmt > 0 Then
            '    oDoc.Lines.Add()
            '    oDoc.Lines.AccountCode = "5-32020-00"
            '    oDoc.Lines.ItemDescription = "Un-Claim Amount"
            '    oDoc.Lines.VatGroup = "ZO"
            '    oDoc.Lines.LineTotal = -1 * dUnClaimAmt
            'End If

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
            AddARInvoice_AETNA = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_AETNA = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Function AddAPInvoice_AETNA(ByVal oDt As DataTable, _
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

            sFuncName = "AddAPInvoice_AETNA"


            Dim k As Integer = Microsoft.VisualBasic.InStrRev(frmUpload.txtFileName.Text, "\")
            Dim sFName As String = Microsoft.VisualBasic.Right(frmUpload.txtFileName.Text, Len(frmUpload.txtFileName.Text) - k).Trim
            Dim sFileName As String = Replace(sFName, ".xlsx", "")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            oDoc.CardCode = oDt.Rows(3).Item("F2").ToString
            oDoc.DocDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.TaxDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.DocCurrency = oDt.Rows(9).Item("F13").ToString.Trim
            oDoc.NumAtCard = oDt.Rows(5).Item("F2").ToString
            oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

            Dim iCnt As Integer = 0

            For Each Row As DataRow In oDt.Rows

                iCnt += 1
                If iCnt >= 10 Then
                    If iCnt > 10 Then
                        oDoc.Lines.Add()
                    End If

                    oDoc.Lines.AccountCode = p_oCompDef.sAP_AETNAGL
                    oDoc.Lines.ItemDescription = Row(0).ToString
                    If CDbl(Row(17)) > 0 Then
                        oDoc.Lines.VatGroup = "SI"
                    Else
                        oDoc.Lines.VatGroup = "ZI"
                    End If
                    oDoc.Lines.PriceAfterVAT = CDbl(Row(23))

                    oDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value = Row(0).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = CDate(Row(1).ToString)
                    oDoc.Lines.UserFields.Fields.Item("U_AI_PolicyNo").Value = Row(2).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = Row(3).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = Row(4).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderName").Value = Row(5).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderAddress").Value = Row(6).ToString
                    'oDoc.Lines.UserFields.Fields.Item("U_AI_InvRefNo").Value = Row(7).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = CDbl(Row(12))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = CDbl(Row(13))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = CDbl(Row(14))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = CDbl(Row(15))

                    'Cost Center
                    oDoc.Lines.CostingCode = oDt.Rows(4).Item("F2").ToString
                End If
            Next

           
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
            AddAPInvoice_AETNA = RTN_SUCCESS

        Catch ex As Exception
            AddAPInvoice_AETNA = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

End Module
