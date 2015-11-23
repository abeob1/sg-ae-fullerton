Module modReceiptUpload

    Private Structure ARIP_Header

        Public sDocNum As String
        Public sCustomerCode As String
        Public sReference As String
        Public sPaymentMode As String
        Public dAmount As Double
        Public sGLAccount As String
        Public sPostingDate As String
    End Structure

    Public dt_CardCode As DataTable
    Public dt_InovoiceNo As DataTable
    Public dt_GLAccount As DataTable
    Private oIPHeader As ARIP_Header
    Dim bIsError As Boolean = False

  

    Public Function UploadDocument_ReceiptUpload(ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument_ReceiptUpload()"
        Dim myfile As New System.IO.FileInfo(sFileName)
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim odv As DataView = Nothing
        Dim iCnt As Integer = 0
        Dim sQuery As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            frmUpload.WriteToStatusScreen(False, "Validating Excel file format.....")

            bIsError = False
            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadReceiptUploadFile", sFuncName)
            ReadReceiptUploadFile(sFileName, "Sheet1", bIsError, odv, sErrDesc)

            If bIsError = True Then
                frmUpload.WriteToStatusScreen(False, " ERROR :::  " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Receipt Upload Excel Worksheet", sFuncName)
                WriteToLogFile("Invalid Excel Worksheet " & sFileName, sFuncName)
                MsgBox("Invalid Excel Worksheet", MsgBoxStyle.Critical, "Receipt Upload")
                GoTo ExitFunc
            End If

            '' bIsError = False
            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling readCSVFileIP()", sFuncName)
            readCSVFileIP(sFileName, odv, sErrDesc)

            If bIsError = True Then
                frmUpload.WriteToStatusScreen(False, " ERROR :::  " & sErrDesc)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Madatory fields missing in the excel sheet.", sFuncName)
                WriteToLogFile("Madatory fields missing in the excel sheet." & sFileName, sFuncName)
                MsgBox("Madatory fields missing in the excel sheet.", MsgBoxStyle.Critical, "Receipt Upload")
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

            sQuery = "SELECT T0.""CardCode"" FROM ""OCRD"" T0 WHERE T0.""CardType"" ='C' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQueryDataTabe() for getting customer codes.", sFuncName)
            dt_CardCode = ExecuteSQLQueryDataTabe(sQuery)

            sQuery = "SELECT T0.""DocNum"", T0.""DocEntry"",T0.""CardCode"",(T0.""DocTotal""-T0.""PaidToDate"")  AS ""Amount""" & _
            ",T0.""DocCur"",( T0.""DocTotalFC""-T0.""PaidFC"")  AS ""AmountFC"" FROM OINV T0 WHERE T0.""DocStatus"" ='O'  AND IFNULL(T0.""CANCELED"",'N') ='N' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQueryDataTabe() for getting Invoice Document Number", sFuncName)
            dt_InovoiceNo = ExecuteSQLQueryDataTabe(sQuery)

            sQuery = "SELECT T0.""AcctCode"" FROM OACT T0 WHERE T0.""Postable"" ='Y'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQueryDataTabe() for getting GL Account codes.", sFuncName)
            dt_GLAccount = ExecuteSQLQueryDataTabe(sQuery)

            If p_oCompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction..", sFuncName)
                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                frmUpload.WriteToStatusScreen(False, "Uploading Incoming Payments...")
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling Process_ReceiptUpload()", sFuncName)
                If Process_ReceiptUpload(odv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then
                    frmUpload.WriteToStatusScreen(False, "sErrDesc")
                    Throw New ArgumentException(sErrDesc)
                    'error condition.
                End If

                frmUpload.WriteToStatusScreen(False, "Successfully Uploaded Incoming Payments.")

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
            UploadDocument_ReceiptUpload = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument_ReceiptUpload = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Private Sub ReadReceiptUploadFile(ByVal sFileName As String, _
                                ByVal sSheet As String, _
                                ByRef bIsError As Boolean, _
                                ByRef dv As DataView, _
                                ByRef sErrdesc As String)

        Dim iHeaderRow As Integer
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "ReadReceiptUploadFile"
        Dim sBatchNo As String = String.Empty

        iHeaderRow = 0

        dv = GetDataViewFromExcel(sFileName, sSheet, sErrdesc)


        If IsNothing(dv) Then
            bIsError = True
            Exit Sub
        End If

        If UCase(dv(iHeaderRow)(0).ToString().Trim) <> UCase("Code") Then
            sErrdesc = "Invalid Excel file Format - ([Code] not found at Column 1"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(1).ToString.Trim) <> UCase("posting date") Then
            sErrdesc = "Invalid Excel file Format - ([posting date] not found at Column 2"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(2).ToString.Trim) <> UCase("invoice number") Then
            sErrdesc = "Invalid Excel file Format - ([invoice number] not found at Column 3"
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(3).ToString.Trim <> "journal remarks" Then
            sErrdesc = "Invalid Excel file Format - ([journal remarks] not found at Column 4"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(4).ToString.Trim) <> UCase("G/L account") Then
            sErrdesc = "Invalid Excel file Format - ([G/L account] not found at Column 5"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(5).ToString.Trim) <> UCase("transfer date") Then
            sErrdesc = "Invalid Excel file Format - ([transfer date] not found at Column 6"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(6).ToString.Trim) <> UCase("Reference") Then
            sErrdesc = "Invalid Excel file Format - ([Reference] not found at Column 7"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(7).ToString.Trim) <> UCase("Amount") Then
            sErrdesc = "Invalid Excel file Format - ([Amount] not found at Column 8"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(8).ToString.Trim) <> UCase("Currency") Then
            sErrdesc = "Invalid Excel file Format - ([Currency] not found at Column 9"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(9).ToString.Trim) <> UCase("Exchange Rate") Then
            sErrdesc = "Invalid Excel file Format - ([Exchange Rate] not found at Column 10"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If


        If UCase(dv(iHeaderRow)(10).ToString.Trim) <> UCase("Receipt") Then
            sErrdesc = "Invalid Excel file Format - ([Receipt] not found at Column 11"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(11).ToString.Trim) <> UCase("Payment Method") Then
            sErrdesc = "Invalid Excel file Format - ([Payment Method] not found at Column 12"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If

        If UCase(dv(iHeaderRow)(12).ToString.Trim) <> UCase("Remarks") Then
            sErrdesc = "Invalid Excel file Format - ([Remarks] not found at Column 13"
            WriteToLogFile(False, sErrdesc)
            frmUpload.WriteToStatusScreen(False, sErrdesc)
            bIsError = True
        End If



    End Sub

    Private Function Process_ReceiptUpload(ByVal oDV As DataView, _
                                            ByVal oDICompany As SAPbobsCOM.Company, _
                                            ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oDTGRIOGrouped As DataTable = Nothing
        Dim oDTDistinct As DataTable = Nothing
        Dim oDVGIRODetl As DataView = New DataView
        Dim oDVChkDetl As DataView = New DataView
        Try

            sFuncName = "Process_ReceiptUpload()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Splitting the datas GIRO/Check", sFuncName)
            oDTGRIOGrouped = oDV.Table.DefaultView.ToTable(True, "F12")

            For intRow As Integer = 0 To oDTGRIOGrouped.Rows.Count - 1
                If Not (oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper.Trim = UCase("Payment Method")) Then
                    oDV.RowFilter = "F12= '" & oDTGRIOGrouped.Rows(intRow).Item(0).ToString & "'"

                    If oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper() = "GIRO" Then
                        oDVGIRODetl = New DataView(oDV.ToTable())
                    ElseIf oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper() = "CHECK" Then
                        oDVChkDetl = New DataView(oDV.ToTable())
                    End If

                End If
            Next
            If oDVGIRODetl.Count > 0 Then
                oDTGRIOGrouped = oDVGIRODetl.Table.DefaultView.ToTable(True, "F1")

                For intRow As Integer = 0 To oDTGRIOGrouped.Rows.Count - 1
                    oDVGIRODetl.RowFilter = "F1 = '" & oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() & "' "
                    Dim odvGIROFil As DataView = New DataView(oDVGIRODetl.ToTable())

                    oDTDistinct = odvGIROFil.Table.DefaultView.ToTable(True, "F7")

                    For iRow As Integer = 0 To oDTDistinct.Rows.Count - 1
                        odvGIROFil.RowFilter = "F7 = '" & oDTDistinct.Rows(iRow).Item(0).ToString.Trim() & "' "
                        Dim oDVFinal As DataView = New DataView(odvGIROFil.ToTable())
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AR_IncomingPayment()", sFuncName)
                        If AR_IncomingPayment(oDVFinal, oDICompany, "GIRO", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                    Next

                Next
            End If


            ''Check PAYENT
            If oDVChkDetl.Count > 0 Then
                oDTGRIOGrouped = oDVChkDetl.Table.DefaultView.ToTable(True, "F1")

                For intRow As Integer = 0 To oDTGRIOGrouped.Rows.Count - 1
                    oDVChkDetl.RowFilter = "F1 = '" & oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() & "' "
                    Dim odvGIROFil As DataView = New DataView(oDVChkDetl.ToTable())

                    oDTDistinct = odvGIROFil.Table.DefaultView.ToTable(True, "F7")

                    For iRow As Integer = 0 To oDTDistinct.Rows.Count - 1
                        odvGIROFil.RowFilter = "F7 = '" & oDTDistinct.Rows(iRow).Item(0).ToString.Trim() & "' "
                        Dim oDVFinal As DataView = New DataView(odvGIROFil.ToTable())
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AR_IncomingPayment()", sFuncName)
                        If AR_IncomingPayment(oDVFinal, oDICompany, "Check", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                    Next

                Next
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            Process_ReceiptUpload = RTN_SUCCESS

        Catch ex As Exception
            Process_ReceiptUpload = RTN_ERROR
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        End Try
    End Function

    Private Function AR_IncomingPayment(ByVal oDV As DataView, _
                                        ByVal oDICompany As SAPbobsCOM.Company, _
                                        ByVal sPaymentMode As String, _
                                        ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oIncomingPayment As SAPbobsCOM.Payments
        Dim oDistLineTable As DataTable = Nothing
        Dim dInvTotal As Double = 0.0
        Dim dDocTotal As Double = 0.0
        Dim sCardCode As String = String.Empty
        Dim sGLAccount As String = String.Empty
        Dim sInvNumber As String = String.Empty
        Dim sInvDocKey As String = String.Empty
        Dim lRetcode As Double
        Dim dBalanceAmt As Double = 0.0
        Dim sInvCardCode As String = String.Empty
        Dim sDocCurrency As String = String.Empty
        Dim sInvCur As String = String.Empty
        Dim dBalanceAmtFC As Double = 0.0
        Dim dExDifAmount As Double = 0
        Dim sFullPayment As String = String.Empty
        Dim dPayment As Double = 0.0

        Try
            sFuncName = "AR_IncomingPayment()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            sCardCode = oDV.Item(0)(0).ToString().Trim()
            sDocCurrency = oDV.Item(0)(8).ToString().Trim().ToUpper


            dt_CardCode.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
            If dt_CardCode.DefaultView.Count = 0 Then
                sErrDesc = "CardCode ::'" & sCardCode & "' provided does not exist in SAP."
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            oIncomingPayment = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
            oIncomingPayment.CardCode = CStr(sCardCode)
            oIncomingPayment.DocDate = CDate(oDV.Item(0)(1).ToString().Trim())
            oIncomingPayment.JournalRemarks = oDV.Item(0)(3).ToString().Trim()
            oIncomingPayment.Remarks = oDV.Item(0)(12).ToString().Trim()

            'If sDocCurrency <> "SGD" Then
            '    oIncomingPayment.DocCurrency = sDocCurrency
            '    If Not oDV.Item(0)(9).ToString().Trim() = String.Empty Then
            '        oIncomingPayment.DocRate = CDbl(oDV.Item(0)(9).ToString().Trim())
            '    End If
            'End If


            'Distinct basedon on invoice number:

            oDistLineTable = oDV.Table.DefaultView.ToTable(True, "F3")

            dDocTotal = 0.0
            dPayment = 0.0

            For iRow As Integer = 0 To oDistLineTable.Rows.Count - 1

                sInvNumber = oDistLineTable.Rows(iRow)(0).ToString().Trim()

                oDV.RowFilter = "F3 = '" & sInvNumber & "'"

                sFullPayment = oDV.Item(0)(10).ToString().Trim().ToUpper()

                dt_InovoiceNo.DefaultView.RowFilter = "DocNum = '" & sInvNumber & "'"
                If dt_InovoiceNo.DefaultView.Count = 0 Then
                    sErrDesc = "Invoice Number ::'" & sInvNumber & "' provided does not exist in SAP."
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    sInvDocKey = dt_InovoiceNo.DefaultView.Item(0)(1).ToString().Trim()
                    sInvCardCode = dt_InovoiceNo.DefaultView.Item(0)(2).ToString().Trim()
                    dBalanceAmt = CDbl(dt_InovoiceNo.DefaultView.Item(0)(3).ToString().Trim())
                    sInvCur = dt_InovoiceNo.DefaultView.Item(0)(4).ToString().Trim()
                    dBalanceAmtFC = CDbl(dt_InovoiceNo.DefaultView.Item(0)(5).ToString().Trim())
                End If


                If sCardCode <> sInvCardCode Then
                    sErrDesc = "Invoice Number ::'" & sInvNumber & "' and Customer Code ::'" & sCardCode & "' provided does not match in SAP."
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                sGLAccount = oDV.Item(0)(4).ToString().Trim()

                dt_GLAccount.DefaultView.RowFilter = "AcctCode = '" & sGLAccount & "'"
                If dt_GLAccount.DefaultView.Count = 0 Then
                    sErrDesc = "GL Account ::'" & sGLAccount & "' provided does not exist in SAP."
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                dInvTotal = 0.0

                For iInvCount As Integer = 0 To oDV.Count - 1
                    dInvTotal += CDbl(oDV.Item(iInvCount)(7).ToString().Trim())
                    dDocTotal += CDbl(oDV.Item(iInvCount)(7).ToString().Trim())
                Next


                'If dInvTotal > dBalanceAmt Then
                '    sErrDesc = "Amount ::'" & dInvTotal & "' provided should not exceed Balance :: '" & dBalanceAmt & "'."
                '    Call WriteToLogFile(sErrDesc, sFuncName)
                '    Throw New ArgumentException(sErrDesc)
                'End If

                oIncomingPayment.Invoices.DocEntry = sInvDocKey
                oIncomingPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice


                If sInvCur = sDocCurrency Then
                    oIncomingPayment.Invoices.SumApplied = dInvTotal
                    dPayment += dInvTotal
                Else
                    If sInvCur = "SGD" And sInvCur <> sDocCurrency Then
                        If sFullPayment = "P" Then
                            oIncomingPayment.Invoices.SumApplied = Math.Round(dInvTotal / CDbl(oDV.Item(0)(9).ToString().Trim()), 2)
                            dPayment += Math.Round(dInvTotal / CDbl(oDV.Item(0)(9).ToString().Trim()), 2)
                        Else
                            oIncomingPayment.Invoices.SumApplied = Math.Round(dDocTotal / CDbl(oDV.Item(0)(9).ToString().Trim()), 2)
                            dExDifAmount += dBalanceAmt - Math.Round(dInvTotal / CDbl(oDV.Item(0)(9).ToString().Trim()), 2)
                            dPayment += Math.Round(dDocTotal / CDbl(oDV.Item(0)(9).ToString().Trim()), 2)
                        End If
                    Else
                        If sFullPayment = "P" Then
                            oIncomingPayment.Invoices.AppliedFC = Math.Round(dInvTotal * CDbl(oDV.Item(0)(9).ToString().Trim()), 2)
                            dPayment += Math.Round(dInvTotal * CDbl(oDV.Item(0)(9).ToString().Trim()), 2)
                        Else
                            oIncomingPayment.Invoices.AppliedFC = CDbl(dBalanceAmtFC) '' Math.Round(dDocTotal * CDbl(oDV.Item(0)(9).ToString().Trim()), 2)
                            dPayment += dBalanceAmtFC
                            dExDifAmount += dBalanceAmtFC - Math.Round(dInvTotal * CDbl(oDV.Item(0)(9).ToString().Trim()), 2)
                        End If

                    End If

                End If

                oIncomingPayment.Invoices.Add()

            Next

            If sPaymentMode.ToString().ToUpper().Trim() = "GIRO" Then
                oIncomingPayment.TransferAccount = sGLAccount
                oIncomingPayment.TransferDate = CDate(oDV.Item(0)(5).ToString().Trim())
                oIncomingPayment.TransferSum = CDbl(dPayment)
                oIncomingPayment.TransferReference = oDV.Item(0)(6).ToString().Trim()
            ElseIf sPaymentMode.ToString().ToUpper().Trim() = UCase("Check") Then
                oIncomingPayment.Checks.DueDate = CDate(oDV.Item(0)(5).ToString().Trim())
                oIncomingPayment.Checks.CountryCode = p_oCompDef.sCountryCode
                oIncomingPayment.Checks.BankCode = p_oCompDef.sBankCode
                oIncomingPayment.Checks.AccounttNum = p_oCompDef.sCheckBankAccount
                oIncomingPayment.Checks.CheckSum = CDbl(dPayment)
                oIncomingPayment.CheckAccount = sGLAccount
                oIncomingPayment.CashSum = 0
                oIncomingPayment.TransferSum = 0

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming Payment... CardCode:" & sCardCode, sFuncName)

            frmUpload.WriteToStatusScreen(False, "Creating incoming payment for this Customer  :: " & sCardCode)

            oIncomingPayment.DocCurrency = sInvCur
            oIncomingPayment.DocRate = CDbl(oDV.Item(0)(9).ToString().Trim())

            lRetcode = oIncomingPayment.Add()

            If lRetcode <> 0 Then
                sErrDesc = p_oCompany.GetLastErrorDescription
                Throw New ArgumentException(sErrDesc)
            Else

                If oDV.Item(0)(10).ToString().Trim().ToUpper() = "F" Then
                    If dExDifAmount <> 0.0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Create_JE()", sFuncName)
                        If Create_JE(oDICompany, sCardCode, dExDifAmount, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If
                frmUpload.WriteToStatusScreen(False, "Successfully added incoming payment for Customer  :: " & sCardCode)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Incoiming payment added successfully. CardCode : " & sCardCode, sFuncName)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            AR_IncomingPayment = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call AddDataToTable(p_oDtReceiptLog, sCardCode, sInvNumber, sPaymentMode, sErrDesc)
            AR_IncomingPayment = RTN_ERROR
            frmUpload.WriteToStatusScreen(False, "Invoice No  :: " & sInvNumber & "  Error Message :: " & sErrDesc)
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        End Try
    End Function

    Private Function Create_JE(ByVal oDICompany As SAPbobsCOM.Company, _
                               ByVal sCustomerCode As String, ByVal dBalAmount As Double, _
                               ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oJE As SAPbobsCOM.JournalEntries
        Dim lRetCode As Double
        Dim dAmount As Double = Math.Abs(Math.Round(dBalAmount, 2))
        Try
            sFuncName = "Create_JE()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function.", sFuncName)
            oJE = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            If dBalAmount > 0 Then

                oJE.Lines.ShortName = sCustomerCode
                oJE.Lines.Debit = dAmount
                oJE.Lines.CostingCode = p_oCompDef.sJECostingCode '' "MB-AOAIA"
                oJE.Lines.Add()

                oJE.Lines.AccountCode = p_oCompDef.sExRateDiffAccount
                oJE.Lines.Credit = dAmount
                oJE.Lines.CostingCode = p_oCompDef.sJECostingCode '' "MB-AOAIA"

                oJE.Lines.Add()

            Else
                oJE.Lines.AccountCode = p_oCompDef.sExRateDiffAccount
                oJE.Lines.Debit = dAmount
                oJE.Lines.CostingCode = p_oCompDef.sJECostingCode '' "MB-AOAIA"
                oJE.Lines.Add()

                oJE.Lines.ShortName = sCustomerCode
                oJE.Lines.Credit = dAmount
                oJE.Lines.CostingCode = p_oCompDef.sJECostingCode '' "MB-AOAIA"
                oJE.Lines.Add()

            End If

            lRetCode = oJE.Add()
            If lRetCode <> 0 Then
                sErrDesc = oDICompany.GetLastErrorDescription
                Throw New ArgumentException(sErrDesc)

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            Create_JE = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Create_JE = RTN_ERROR
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        End Try
    End Function

    Private Sub readCSVFileIP(ByVal CurrFileToUpload As String, ByVal dv As DataView, _
                              ByRef sErrDesc As String)

        Dim sFuncName As String = "readCSVFileIP"
        Dim sSQL As String = String.Empty

        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            For i As Integer = 1 To dv.Count - 1
                With oIPHeader
                    .sDocNum = dv(i)(2).ToString.Trim()
                    .sCustomerCode = dv(i)(0).ToString.Trim()
                    .sGLAccount = dv(i)(4).ToString.Trim()
                    .sPaymentMode = dv(i)(10).ToString.Trim()
                    .sPostingDate = dv(i)(1).ToString.Trim()
                    .sReference = dv(i)(6).ToString.Trim()
                    .dAmount = CDbl(dv(i)(7).ToString.Trim())

                    If .sDocNum.Length = 0 Then
                        bIsError = True
                        sErrDesc = "Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Invoice Number is Mandatory"
                        Call AddDataToTable(p_oDtReceiptLog, .sCustomerCode, .sDocNum, .sPaymentMode, sErrDesc)
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Invoice Number is Mandatory", sFuncName)
                    End If

                    If .sCustomerCode.Length = 0 Then
                        bIsError = True
                        sErrDesc = "Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Customer Code is Mandatory"
                        Call AddDataToTable(p_oDtReceiptLog, .sCustomerCode, .sDocNum, .sPaymentMode, sErrDesc)
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Customer Code is Mandatory", sFuncName)
                    End If

                    If .sGLAccount.Length = 0 Then
                        bIsError = True
                        sErrDesc = "Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". GL Account Code is Mandatory"
                        Call AddDataToTable(p_oDtReceiptLog, .sCustomerCode, .sDocNum, .sPaymentMode, sErrDesc)
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". GL Account Code is Mandatory", sFuncName)
                    End If

                    If .sPaymentMode.Length = 0 Then
                        bIsError = True
                        sErrDesc = "Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Payment Mode is Mandatory"
                        Call AddDataToTable(p_oDtReceiptLog, .sCustomerCode, .sDocNum, .sPaymentMode, sErrDesc)
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Payment Mode is Mandatory", sFuncName)
                    End If

                    If .sPostingDate.Length = 0 Then
                        bIsError = True
                        sErrDesc = "Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Posting Date is Mandatory"
                        Call AddDataToTable(p_oDtReceiptLog, .sCustomerCode, .sDocNum, .sPaymentMode, sErrDesc)
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Posting Date is Mandatory", sFuncName)
                    End If

                    If .dAmount = 0 Then
                        bIsError = True
                        sErrDesc = "Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Amouunt is Mandatory"
                        Call AddDataToTable(p_oDtReceiptLog, .sCustomerCode, .sDocNum, .sPaymentMode, sErrDesc)
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Amount is Mandatory", sFuncName)
                    End If

                End With
            Next

        Catch ex As Exception
            bIsError = True
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

End Module
