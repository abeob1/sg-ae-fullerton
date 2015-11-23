Module modBUPA

    Public Function UploadDocument_BUPA(ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_BUPA()"
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
            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadARFile", sFuncName)
            ReadBUPAFile(sFileName, "Sheet1", bIsError, odv, sErrDesc)

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
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_ARInvoice_BUPA()", sFuncName)
                If UploadDocument_ARInvoice_BUPA(odv, sErrDesc) <> RTN_SUCCESS Then
                    frmUpload.WriteToStatusScreen(False, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                    'error condition.
                End If
                frmUpload.WriteToStatusScreen(False, "Successfully Uploaded A/R Invoices..")

                frmUpload.WriteToStatusScreen(False, "Uploading A/P Invoices..")
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument_APInvoice_BUPA()", sFuncName)
                If UploadDocument_APInvoice_BUPA(odv, sErrDesc) <> RTN_SUCCESS Then
                    frmUpload.WriteToStatusScreen(False, "sErrDesc")
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
            UploadDocument_BUPA = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_BUPA = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Private Sub ReadBUPAFile(ByVal sFileName As String, _
                                ByVal sSheet As String, _
                                ByRef bIsError As Boolean, _
                                ByRef dv As DataView, _
                                ByRef sErrdesc As String)

        Dim iHeaderRow As Integer
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "ReadBUPAFile"
        Dim sBatchNo As String = String.Empty

        iHeaderRow = 4

        dv = GetDataViewFromExcel(sFileName, sSheet, sErrdesc)


        If IsNothing(dv) Then
            bIsError = True
            Exit Sub
        End If

        If dv(iHeaderRow)(0).ToString.Trim <> "No." Then
            sErrdesc = "Invalid Excel file Format - ([No.] not found at Column 1"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(1).ToString.Trim <> "Member Name" Then
            sErrdesc = "Invalid Excel file Format - ([Member Name] not found at Column 2"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(2).ToString.Trim <> "Member ID" Then
            sErrdesc = "Invalid Excel file Format - ([Member ID] not found at Column 3"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(3).ToString.Trim <> "Provider" Then
            sErrdesc = "Invalid Excel file Format - ([Provider] not found at Column 4"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString.Trim <> "Provider Address" Then
            sErrdesc = "Invalid Excel file Format - ([Provider Address] not found at Column 5"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(5).ToString.Trim <> "Visit Date" Then
            sErrdesc = "Invalid Excel file Format - ([Visit Date] not found at Column 6"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(6).ToString.Trim <> "Bill Num" Then
            sErrdesc = "Invalid Excel file Format - ([Bill Num] not found at Column 7"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(7).ToString.Trim <> "Bill Received" Then
            sErrdesc = "Invalid Excel file Format - ([Bill Received] not found at Column 8"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString.Trim <> "Surgeon Fee / Specialist Fee" Then
            sErrdesc = "Invalid Excel file Format - ([Surgeon Fee / Specialist Fee] not found at Column 9"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(9).ToString.Trim <> "Anaesthetist" Then
            sErrdesc = "Invalid Excel file Format - ([Anaesthetist] not found at Column 10"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString.Trim <> "Consultation Fee" Then
            sErrdesc = "Invalid Excel file Format - ([Consultation Fee] not found at Column 11"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(11).ToString.Trim <> "Drugs Cost" Then
            sErrdesc = "Invalid Excel file Format - ([Drugs Cost] not found at Column 12"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(12).ToString.Trim <> "Inpatient Services" Then
            sErrdesc = "Invalid Excel file Format - ([Inpatient Services] not found at Column 13"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(13).ToString.Trim <> "Discount" Then
            sErrdesc = "Invalid Excel file Format - ([Discount] not found at Column 14"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(14).ToString.Trim <> "Deductible / Co-insurance" Then
            sErrdesc = "Invalid Excel file Format - ([Deductible / Co-insurance] not found at Column 15"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(15).ToString.Trim <> "Non payable Amount" Then
            sErrdesc = "Invalid Excel file Format - ([Non payable Amount] not found at Column 16"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(16).ToString.Trim <> "Currency" Then
            sErrdesc = "Invalid Excel file Format - ([Currency] not found at Column 17"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(17).ToString.Trim <> "Total Fees Bills (MYR)" Then
            sErrdesc = "Invalid Excel file Format - ([Total Fees Bills (MYR)] not found at Column 18"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(18).ToString.Trim <> "Approved Fees" Then
            sErrdesc = "Invalid Excel file Format - ([Approved Fees] not found at Column 19"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(19).ToString.Trim <> "Conversion Date" Then
            sErrdesc = "Invalid Excel file Format - ([Conversion Date] not found at Column 20"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(20).ToString.Trim <> "Amount (SGD)" Then
            sErrdesc = "Invalid Excel file Format - ([Amount (SGD)] not found at Column 21"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(21).ToString.Trim <> "Exchange Rate" Then
            sErrdesc = "Invalid Excel file Format - (Exchange Rate] not found at Column 21"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(22).ToString.Trim <> "Cost Center-AP" Then
            sErrdesc = "Invalid Excel file Format - ([Cost Center-AP] not found at Column 22"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(23).ToString.Trim <> "Vendor Code" Then
            sErrdesc = "Invalid Excel file Format - ([Vendor Code] not found at Column 23"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If



    End Sub

    Public Function UploadDocument_APInvoice_BUPA(ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_APInvoice()"
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim iCnt As Integer = 0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oDT As DataTable
            Dim oBPDT As DataTable

            oDT = odv.Table.DefaultView.ToTable(True, "F24")
            oBPDT = oDT.Clone
            oBPDT.Clear()

            For Each row As DataRow In oDT.Rows
                If Not row(0).ToString = String.Empty And Not row(0).ToString = "Vendor Code" Then
                    oBPDT.ImportRow(row)
                End If
            Next


            For Each row As DataRow In oBPDT.Rows
                Dim oDtRows() As DataRow = odv.Table.Select("F24='" & row.Item(0).ToString & "'")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddAPInvoice()", sFuncName)
                frmUpload.WriteToStatusScreen(False, "Creating AP Invoice .....")
                If AddAPInvoice_BUPA(oDtRows, odv.Table, sErrDesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrDesc)
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            UploadDocument_APInvoice_BUPA = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_APInvoice_BUPA = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Public Function UploadDocument_ARInvoice_BUPA(ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_ARInvoice_BUPA()"
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim iCnt As Integer = 0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            For Each row As DataRow In odv.Table.Rows
                iCnt += 1
                If iCnt >= 6 Then

                    If row.Item(1).ToString = String.Empty Then Exit For

                    frmUpload.WriteToStatusScreen(False, "Creating Invoice for Line ::" & iCnt - 5)
                    If AddARInvoice_BUPA(odv.Table, row, sErrDesc) <> RTN_SUCCESS Then
                        frmUpload.WriteToStatusScreen(False, "ERROR :: Rollback Transaction..")
                        frmUpload.WriteToStatusScreen(False, "ERROR :: Failed to create Invoice for Line ::" & iCnt - 5)
                        frmUpload.WriteToStatusScreen(False, "ERROR :: " & sErrDesc)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File was not successfully uploaded" & frmUpload.txtFileName.Text, sFuncName)
                        If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    End If
                    frmUpload.WriteToStatusScreen(False, "Successfully created Invoice for Line ::" & iCnt - 5)
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            UploadDocument_ARInvoice_BUPA = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_ARInvoice_BUPA = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Private Function AddARInvoice_BUPA(ByVal oDt As DataTable, _
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
            sFuncName = "AddARInvoice_BUPA"


            Dim k As Integer = Microsoft.VisualBasic.InStrRev(frmUpload.txtFileName.Text, "\")
            Dim sFName As String = Microsoft.VisualBasic.Right(frmUpload.txtFileName.Text, Len(frmUpload.txtFileName.Text) - k).Trim
            Dim sFileName As String = Replace(sFName, ".xlsx", "")


            sCostCenter = GetCostCenter(oDt.Rows(1).Item("F2").ToString)

            If sCostCenter = String.Empty Then
                sErrDesc = "Cost Center is blank in Customer Master :: " & oDt.Rows(1).Item("F2").ToString
                frmUpload.WriteToStatusScreen(False, "ERROR :" & sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            oDoc.CardCode = oDt.Rows(1).Item("F2").ToString
            oDoc.DocDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.DocDueDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.TaxDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.DocCurrency = Row.Item(16).ToString.Trim
            oDoc.NumAtCard = Row.Item(6).ToString
            oDoc.DocRate = CDbl(Row.Item(21))
            oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName


            'Line 1 Surgeon Fee / Specialist Fee
            If Not IsDBNull(Row.Item(8)) Then
                If Convert.ToDouble(Row.Item(8)) <> 0 Then
                    bLine = True
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(8)), "Surgeon Fee / Specialist Fee")
                End If
            End If

            'Line 2 Anaesthetist
            If Not IsDBNull(Row.Item(9)) Then
                If Convert.ToDouble(Row.Item(9)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(9)), "Anaesthetist")
                End If
            End If


            'Line 3 Consultation Fee
            If Not IsDBNull(Row.Item(10)) Then
                If Convert.ToDouble(Row.Item(10)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(10)), "Consultation Fee")
                End If
            End If

            'Line 4 Drugs Cost
            If Not IsDBNull(Row.Item(11)) Then

                If CDbl(Row.Item(11)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(11)), "Drugs Cost")
                End If
            End If

            'Line 5 Inpatient Services
            If Not IsDBNull(Row.Item(12)) Then

                If CDbl(Row.Item(12)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, CDbl(Row.Item(12)), "Inpatient Services")
                End If
            End If

            'Line 6 Discount
            If Not IsDBNull(Row.Item(13)) Then

                If CDbl(Row.Item(13)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, -1 * CDbl(Row.Item(13)), "Discount")
                End If
            End If

            'Line 7 Deductible / Co-insurance
            If Not IsDBNull(Row.Item(14)) Then

                If CDbl(Row.Item(14)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, -1 * CDbl(Row.Item(14)), "Deductible / Co-insurance")
                End If
            End If

            'Line 8 Non payable Amount
            If Not IsDBNull(Row.Item(15)) Then

                If CDbl(Row.Item(15)) <> 0 Then
                    If bLine = False Then
                        bLine = True
                    Else
                        oDoc.Lines.Add()
                    End If
                    AddDocLine(oDoc, Row, sCostCenter, -1 * CDbl(Row.Item(15)), "Non payable Amount")
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
            AddARInvoice_BUPA = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_BUPA = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Sub AddDocLine(ByVal oDoc As SAPbobsCOM.Documents, ByVal Row As DataRow, _
                        ByVal sCostCenter As String, _
                        ByVal dLineTotal As Double, _
                        ByVal sDesc As String)

        'oDoc.Lines.VatGroup = "ZO"
        oDoc.Lines.CostingCode = sCostCenter
        oDoc.Lines.LineTotal = dLineTotal
        oDoc.Lines.ItemDescription = sDesc
        oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = Row.Item(1).ToString
        oDoc.Lines.UserFields.Fields.Item("U_AI_PolicyNo").Value = Row.Item(2).ToString
        oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderName").Value = Row.Item(3).ToString
        oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderAddress").Value = Row.Item(4).ToString
        oDoc.Lines.UserFields.Fields.Item("U_AI_Admitdate").Value = Row.Item(5)
        oDoc.Lines.AccountCode = p_oCompDef.sAR_BUPAGL
    End Sub

    Private Function AddAPInvoice_BUPA(ByVal oRows() As DataRow, _
                                  ByVal oDt As DataTable, _
                                  ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim bLine As Boolean = False
        Dim iCnt As Integer

        Try
            sFuncName = "AddAPInvoice_BUPA"

            'Check for duplicate VendorRef.No

            If IsVendorRefNoExists(oRows(0).Item(6).ToString) = True Then
                sErrDesc = "ERROR::" & "Vendor Ref.No : " & oRows(0).Item(6).ToString & "already exists in the system"
                frmUpload.WriteToStatusScreen(False, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Vendor ref.No already exists. Adding API failed.", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim k As Integer = Microsoft.VisualBasic.InStrRev(frmUpload.txtFileName.Text, "\")
            Dim sFName As String = Microsoft.VisualBasic.Right(frmUpload.txtFileName.Text, Len(frmUpload.txtFileName.Text) - k).Trim
            Dim sFileName As String = Replace(sFName, ".xlsx", "")


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            oDoc.CardCode = oRows(0).Item(23).ToString
            oDoc.DocDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.TaxDate = CDate(oDt.Rows(2).Item("F2").ToString)
            oDoc.NumAtCard = oRows(0).Item(6).ToString
            oDoc.DocCurrency = oRows(0).Item(16).ToString

            oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

            iCnt = 0

            For Each LineRow As DataRow In oRows
                iCnt += 1
                If iCnt > 1 Then
                    oDoc.Lines.Add()
                End If

                oDoc.Lines.AccountCode = p_oCompDef.sAP_BUPAGL
                oDoc.Lines.ItemDescription = LineRow(1).ToString
                oDoc.Lines.LineTotal = CDbl(LineRow(18))
                oDoc.Lines.CostingCode = LineRow(22).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = LineRow(5).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_EmpID").Value = LineRow(2).ToString

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AP Invoice.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                p_oCompany.GetLastError(lErrCode, sErrDesc)
                frmUpload.WriteToStatusScreen(False, "Failed to create AP Invoice .....")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding API failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            frmUpload.WriteToStatusScreen(False, "Successfully created AP Invoice .....")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddAPInvoice_BUPA = RTN_SUCCESS

        Catch ex As Exception
            AddAPInvoice_BUPA = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Public Function IsVendorRefNoExists(ByVal sVendorRef As String) As Boolean
        Dim bIsExists As Boolean = False
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = " SELECT T0.""NumAtCard"" FROM OPCH T0" & _
               " WHERE T0.""NumAtCard"" = '" & sVendorRef & "'"

        oDS = ExecuteSQLQuery(sSQL)

        If oDS.Tables(0).Rows.Count > 0 Then bIsExists = True
        Return bIsExists

    End Function

   
End Module