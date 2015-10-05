Module modFAH

    Private isError As Boolean
    Private dtBP As DataTable
    Private dtINVHDr As DataTable
    Private dtItem As DataTable
    Private dtBranch As DataTable

    ' Company Default Structure

    Private Structure ARHeader

        Public InvoiceNumber As String
        Public ItemCode As String
        Public Quantity As Double
        Public GrossTotal As Double
        Public DocumentDate As Date
        Public CustomerCode As String
        Public CustomerName As String

    End Structure

    Private Structure IPHeader

        Public InvoiceNumber As String
        Public DocumentDate As Date
        Public CustomerCode As String
        Public CustomerName As String
        Public Amount As Double
        Public PaymentMode As String
    End Structure

    Private oARHeaderDef As ARHeader
    Private oIPHeaderDef As IPHeader
    Private InputFolderPath As String = p_oCompDef.sInboxDir

    Public Function UploadFAH(ByRef sErrdesc As String) As Long

        'Event      :   UploadMED_FAH()
        'Purpose    :   For Checking Errors & updation of data in SAP from AR CSV Files
        'Author     :   Sri
        'Date       :   24 Jun 2014

        Dim sFuncName As String = String.Empty
        Dim IsFileExist As Boolean
        Dim sFileType As String = String.Empty
        Dim oDvHdr As DataView = Nothing
        Dim oDvDtl As DataView = Nothing
        Dim sFileName As String = String.Empty
        Dim sSubFolder As String = String.Empty

        Try
            sFuncName = "UploadFAH"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo

            '=========================================== FAH File Processing Started ======================================================
            'FAH Files
            files = DirInfo.GetFiles("FAH*.csv")

            For Each File As System.IO.FileInfo In files
                IsFileExist = True

                oDvHdr = GetDataViewFromCSV(File.FullName)
                sFileName = File.Name
                sFileType = Mid(sFileName, 5, 2)

                If sFileType = "AR" Then
                    Console.WriteLine("Calling readCSVFileAR_Header for " & File.FullName & " to check error in CSV file")
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling readCSVFileAR_Header for " & File.FullName & " to check error in CSV file", sFuncName)
                    readCSVFileAR_Header(File.FullName, oDvHdr)
                    If isError = True Then
                        If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Wrong data in CSV file " & File.FullName & ".", sFuncName)
                        sErrdesc = "Invalid data in CSV file.::" & File.Name & " .Please refer to the Error log file for more details."
                        Console.WriteLine(sErrdesc)
                        AddDataToTable(p_oDtError, File.Name, "Error", sErrdesc)
                    End If

                    If isError = True Then
                        Console.WriteLine("Calling FileMoveToArchive()")
                        FileMoveToArchive(File, File.FullName, RTN_ERROR)
                        Throw New ArgumentException(sErrdesc)
                    End If

                    Console.WriteLine("Calling ProcessARFile()")
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ProcessARFile()", sFuncName)
                    If ProcessARFile(oDvHdr, sFileName, sErrdesc) <> RTN_SUCCESS Then
                        '' For Each file As System.IO.FileInfo In files
                        AddDataToTable(p_oDtError, File.Name, "Error", sErrdesc)
                        Console.WriteLine("Calling FileMoveToArchive()")
                        FileMoveToArchive(File, File.FullName, RTN_ERROR)
                        ''  Next
                        Throw New ArgumentException(sErrdesc)
                    Else
                        'Mvoe files
                        '' For Each file As System.IO.FileInfo In files
                        Console.WriteLine("Calling FileMoveToArchive()")
                        FileMoveToArchive(File, File.FullName, RTN_SUCCESS)
                        '' Next
                    End If
                ElseIf sFileType = "RC" Then

                    Console.WriteLine("Calling readCSVFileIP_Header for " & File.FullName & " to check error in CSV file")
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling readCSVFileIP_Header for " & File.FullName & " to check error in CSV file", sFuncName)
                    readCSVFileIP_Header(File.FullName, oDvHdr)
                    If isError = True Then
                        If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Wrong data in CSV file " & File.FullName & ".", sFuncName)
                        sErrdesc = "Invalid data in CSV file.::" & File.Name & " .Please refer to the Error log file for more details."
                        Console.WriteLine(sErrdesc)
                        AddDataToTable(p_oDtError, File.Name, "Error", sErrdesc)
                    End If

                    If isError = True Then
                        Console.WriteLine("Calling FileMoveToArchive()")
                        FileMoveToArchive(File, File.FullName, RTN_ERROR)
                        Throw New ArgumentException(sErrdesc)
                    End If

                    Console.WriteLine("Calling ProcessIPFile()")
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ProcessIPFile()", sFuncName)
                    If ProcessIPFile(oDvHdr, sFileName, sErrdesc) <> RTN_SUCCESS Then
                        '' For Each file As System.IO.FileInfo In files
                        AddDataToTable(p_oDtError, File.Name, "Error", sErrdesc)
                        Console.WriteLine("Calling FileMoveToArchive()")
                        FileMoveToArchive(File, File.FullName, RTN_ERROR)
                        ''  Next
                        Throw New ArgumentException(sErrdesc)
                    Else
                        'Mvoe files
                        '' For Each file As System.IO.FileInfo In files
                        Console.WriteLine("Calling FileMoveToArchive()")
                        FileMoveToArchive(File, File.FullName, RTN_SUCCESS)
                        '' Next
                    End If

                End If
            Next

            If IsFileExist = False Then
                Console.WriteLine("No FAH CSV Available for Creation")
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("No MED & FAH CSV available for updation", sFuncName)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

            UploadFAH = RTN_SUCCESS

        Catch ex As Exception
            UploadFAH = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uplodiang AR file.", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function

    Private Sub readCSVFileIP_Header(ByVal CurrFileToUpload As String, ByVal dv As DataView)

        'Event      :   readCSVFileAR_Header
        'Purpose    :   For reading of Business Partner CSV file
        'Author     :   Sri 
        'Date       :   24 NOV 2013 

        Dim sFuncName As String = "readCSVFileIP_Header_MED"
        Dim sSQL As String = String.Empty
        'Dim dt As DataTable
        Dim sErrDesc As String = String.Empty

        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            For i As Integer = 1 To dv.Count - 1
                With oIPHeaderDef
                    .InvoiceNumber = dv(i)(0).ToString
                    .DocumentDate = dv(i)(1).ToString()
                    .CustomerCode = dv(i)(2).ToString
                    .CustomerName = dv(i)(3).ToString
                    .Amount = dv(i)(13)
                    .PaymentMode = dv(i)(19).ToString

                    If .InvoiceNumber.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Invoice Number is Mandatory", sFuncName)
                    End If

                    If .CustomerCode.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Customer Code is Mandatory", sFuncName)
                    End If

                    If .CustomerName.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Customer Name is Mandatory", sFuncName)
                    End If

                    If .DocumentDate.ToString() = String.Empty Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Paid Date is Mandatory", sFuncName)
                    End If

                    If .PaymentMode.ToString() = String.Empty Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Payment Mode is Mandatory", sFuncName)
                    End If
                End With

            Next

        Catch ex As Exception
            isError = True
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Private Function ProcessIPFile(ByVal oDvHdr As DataView, _
                            ByVal sFileName As String, _
                            ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "ProcessIPFile"
        Dim oDtHdr As New DataTable
        Dim oDIComp As SAPbobsCOM.Company = Nothing
        Dim oCompanyAP As SAPbobsCOM.Company = Nothing
        Dim sSupplierCode As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim sDBCode As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Console.WriteLine("Calling ConnectToTargetCompany()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)

            sDBCode = Left(sFileName, 3)

            If ConnectToTargetCompany(oDIComp, sDBCode, sErrdesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrdesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Customer Code.", sFuncName)
            sSQL = "Select ""CardCode"" From OCRD WHERE ""CardType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtBP = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Invoice Ref.Doc NO.", sFuncName)
            sSQL = "SELECT ""DocEntry"",""U_AI_RefDocNum"" FROM OINV WHERE ""U_AI_RefDocNum"" IS NOT NULL and ""DocStatus""='O'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtINVHDr = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            Console.WriteLine("Start the SAP Transaction on Company DB :: " & oDIComp.CompanyName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDIComp.CompanyDB, sFuncName)
            oDIComp.StartTransaction()

            Console.WriteLine("Calling AddARInvoice_Item()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_Item() ", sFuncName)
            If Add_IncomingPayment(oDIComp, oDvHdr, sErrdesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrdesc)
            End If

            Console.WriteLine("Committing All Trasactions ")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit all transaction..........", sFuncName)

            If Not oDIComp Is Nothing Then
                If oDIComp.Connected = True Then
                    If oDIComp.InTransaction = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDIComp.CompanyDB, sFuncName)
                        oDIComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDIComp.CompanyDB, sFuncName)
                    oDIComp.Disconnect()
                    oDIComp = Nothing
                End If
            End If


            ProcessIPFile = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion completed successfully.", sFuncName)

        Catch ex As Exception


            If Not oDIComp Is Nothing Then
                If oDIComp.Connected = True Then
                    If oDIComp.InTransaction = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDIComp.CompanyDB, sFuncName)
                        oDIComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDIComp.CompanyDB, sFuncName)
                    oDIComp.Disconnect()
                    oDIComp = Nothing
                End If
            End If

            Console.WriteLine("Rollback All the Transactions")
            ProcessIPFile = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uploading AR File", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function

    Private Sub readCSVFileAR_Header(ByVal CurrFileToUpload As String, ByVal dv As DataView)

        'Event      :   readCSVFileAR_Header
        'Purpose    :   For reading of Business Partner CSV file
        'Author     :   Sri 
        'Date       :   24 NOV 2013 

        Dim sFuncName As String = "readCSVFileAR_Header"
        Dim sSQL As String = String.Empty
        'Dim dt As DataTable
        Dim sErrDesc As String = String.Empty

        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            For i As Integer = 1 To dv.Count - 1
                With oARHeaderDef
                    .InvoiceNumber = dv(i)(0).ToString
                    .ItemCode = dv(i)(4).ToString
                    .GrossTotal = dv(i)(6)
                    .DocumentDate = dv(i)(5)
                    .CustomerCode = dv(i)(9).ToString
                    .CustomerName = dv(i)(10).ToString

                    If .InvoiceNumber.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Invoice Number is Mandatory", sFuncName)
                    End If

                    If .ItemCode.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". ItemCode Code is Mandatory", sFuncName)
                    End If

                    If .CustomerCode.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Customer Code is Mandatory", sFuncName)
                    End If

                    If .CustomerName.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Customer Name is Mandatory", sFuncName)
                    End If

                    If .DocumentDate.ToString() = String.Empty Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Document Date is Mandatory", sFuncName)
                    End If

                End With
            Next

        Catch ex As Exception
            isError = True
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Private Function ProcessARFile(ByVal oDvHdr As DataView, _
                            ByVal sFileName As String, _
                            ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "ProcessARFile"
        Dim oDtHdr As New DataTable
        Dim oDIComp As SAPbobsCOM.Company = Nothing
        Dim oCompanyAP As SAPbobsCOM.Company = Nothing
        Dim sSupplierCode As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim sDBCode As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Console.WriteLine("Calling ConnectToTargetCompany()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)


            sDBCode = Left(sFileName, 3)

            If ConnectToTargetCompany(oDIComp, sDBCode, sErrdesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrdesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Customer Code.", sFuncName)
            sSQL = "Select ""CardCode"" From OCRD WHERE ""CardType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtBP = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Invoice Ref.Doc NO.", sFuncName)
            sSQL = "SELECT ""DocEntry"",""U_AI_RefDocNum"" FROM OINV WHERE ""U_AI_RefDocNum"" IS NOT NULL and ""DocStatus""='O'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtINVHDr = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Item Code.", sFuncName)
            sSQL = "SELECT ""ItemCode"" FROM OITM"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtItem = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Brach Code.", sFuncName)
            sSQL = "SELECT ""OcrCode"" FROM OOCR"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtBranch = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            Console.WriteLine("Start the SAP Transaction on Company DB :: " & oDIComp.CompanyName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDIComp.CompanyDB, sFuncName)
            oDIComp.StartTransaction()

            Console.WriteLine("Calling AddARInvoice_Item()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_Item() ", sFuncName)
            If AddARInvoice_Item(oDIComp, oDvHdr, sFileName, sErrdesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrdesc)
            End If


            Console.WriteLine("Committing All Trasactions ")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit all transaction..........", sFuncName)

            If Not oDIComp Is Nothing Then
                If oDIComp.Connected = True Then
                    If oDIComp.InTransaction = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDIComp.CompanyDB, sFuncName)
                        oDIComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDIComp.CompanyDB, sFuncName)
                    oDIComp.Disconnect()
                    oDIComp = Nothing
                End If
            End If


            ProcessARFile = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion completed successfully.", sFuncName)

        Catch ex As Exception


            If Not oDIComp Is Nothing Then
                If oDIComp.Connected = True Then
                    If oDIComp.InTransaction = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDIComp.CompanyDB, sFuncName)
                        oDIComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDIComp.CompanyDB, sFuncName)
                    oDIComp.Disconnect()
                    oDIComp = Nothing
                End If
            End If

            Console.WriteLine("Rollback All the Transactions")
            ProcessARFile = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uploading AR File", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function

    Private Function Add_IncomingPayment(ByVal oCompany As SAPbobsCOM.Company, ByVal oDVDetails As DataView, _
                                           ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "Add_IncomingPayment"
        Dim oDoc As SAPbobsCOM.Payments
        Dim lRetCode, lErrCode As Long
        Dim oDTDistinctHeader As DataTable = Nothing
        Dim sCardCode As String = String.Empty
        Dim sDocRefNum As String = String.Empty
        Dim sItemCode As String = String.Empty
        Dim sBranch As String = String.Empty
        Dim dtDocDate As Date = Nothing
        Dim sPaymentMode As String = String.Empty
        Dim sDocEntry As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)

            oDTDistinctHeader = oDVDetails.Table.DefaultView.ToTable(True, "F1")

            For IntHeader As Integer = 0 To oDTDistinctHeader.Rows.Count - 1

                sDocRefNum = oDTDistinctHeader.Rows(IntHeader).Item(0).ToString.Trim()

                If sDocRefNum = String.Empty Then Continue For

                Console.WriteLine("Creating A/R Service Invoice. Document No :: " & sDocRefNum)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/R Service Invoice. Document No :: " & sDocRefNum, sFuncName)

                oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                oDVDetails.RowFilter = "F1 = '" & sDocRefNum & "'"

                sCardCode = oDVDetails(0)(2).ToString.Trim()
                dtDocDate = Convert.ToDateTime(oDVDetails(0)(1).ToString.Trim())


                dtINVHDr.DefaultView.RowFilter = "U_AI_RefDocNum = '" & sDocRefNum & "'"
                If dtINVHDr.DefaultView.Count <> 0 Then
                    sDocEntry = dtINVHDr.DefaultView.Item(0)("DocEntry").ToString().Trim()
                Else
                    sErrDesc = "RefDocNum :: " & sDocRefNum & " provided Does Not Exist in SAP."
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
                If dtBP.DefaultView.Count = 0 Then
                    sErrDesc = "Cardcode :: " & sCardCode & " provided Does Not Exist in SAP."
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                oDoc.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                oDoc.CardCode = sCardCode
                oDoc.DocDate = dtDocDate

                oDoc.Invoices.DocEntry = sDocEntry
                oDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                oDoc.Invoices.Add()

                For IntRow As Integer = 0 To oDVDetails.Count - 1
                    sPaymentMode = oDVDetails(IntRow)(19).ToString().Trim()

                    If sPaymentMode.ToString().ToUpper() = "CHECK" Then
                        oDoc.Checks.DueDate = dtDocDate
                        oDoc.Checks.CheckAccount = p_oCompDef.sCheckAccount
                        ''oDoc.Checks.CheckNumber = oDVDetails(IntRow)(7).ToString().Trim()
                        oDoc.Checks.CheckSum = CDbl(oDVDetails(IntRow)(13).ToString().Trim())
                        oDoc.Checks.BankCode = p_oCompDef.sBankCode

                        oDoc.Checks.Add()
                    ElseIf sPaymentMode.ToString.ToUpper() = "CASH" Then
                        oDoc.CashAccount = p_oCompDef.sCashAccount
                        oDoc.CashSum = CDbl(oDVDetails(IntRow)(13).ToString().Trim())
                    Else
                        oDoc.TransferAccount = p_oCompDef.sTransferAccount
                        ''oDoc.TransferReference = oDVDetails(IntRow)(7).ToString().Trim()
                        oDoc.TransferSum = CDbl(oDVDetails(IntRow)(13).ToString().Trim())
                    End If
                Next

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming Payment.", sFuncName)

                oDoc.CounterReference = oDVDetails(0)(5).ToString().Trim()

                lRetCode = oDoc.Add()

                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrDesc)
                    Console.WriteLine("Adding Incoming Payment Failed. Invoice Number : " & sDocRefNum & " Error : " & sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming Payment failed.", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If
                Console.WriteLine("Incoming Payment Added Successfully. Invoice Number : " & sDocRefNum)
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Succesfully.", sFuncName)
            Add_IncomingPayment = RTN_SUCCESS

        Catch ex As Exception
            Add_IncomingPayment = RTN_ERROR
            sErrDesc = ex.Message
            Console.WriteLine("Adding AR Invoice Failed. Invoice Number : " & sDocRefNum & " Error : " & sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Function AddARInvoice_Item(ByVal oCompany As SAPbobsCOM.Company, ByVal oDVDetails As DataView, _
                                        ByVal sFileName As String, _
                                           ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "AddARInvoice_Item"
        Dim oDoc As SAPbobsCOM.Documents
        Dim lRetCode, lErrCode As Long
        Dim oDTDistinctHeader As DataTable = Nothing
        Dim sCardCode As String = String.Empty
        Dim sDocRefNum As String = String.Empty
        Dim sItemCode As String = String.Empty
        Dim sBranch As String = String.Empty
        Dim dtDocDate As DateTime = Nothing

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)

            oDTDistinctHeader = oDVDetails.Table.DefaultView.ToTable(True, "F1")

            For IntHeader As Integer = 0 To oDTDistinctHeader.Rows.Count - 1

                sDocRefNum = oDTDistinctHeader.Rows(IntHeader).Item(0).ToString.Trim
                'oDVDetails(IntHeader)(3).ToString.Trim()

                If sDocRefNum = String.Empty Then Continue For

                Console.WriteLine("Creating A/R Service Invoice. Document No :: " & sDocRefNum)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/R Service Invoice. Document No :: " & sDocRefNum, sFuncName)

                oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items


                oDVDetails.RowFilter = "F1 = '" & sDocRefNum & "'"

                sCardCode = oDVDetails(0)(9).ToString.Trim()

                dtINVHDr.DefaultView.RowFilter = "U_AI_RefDocNum = '" & sDocRefNum & "'"
                If dtINVHDr.DefaultView.Count <> 0 Then

                    Dim sDocEntry As String = dtINVHDr.DefaultView.Item(0)("DocEntry").ToString().Trim()

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CopyDocument() for Creating CreditNote. Document No :: " & sDocRefNum, sFuncName)
                    Console.WriteLine("Calling CopyDocument() for Creating CreditNote. Document No :: " & sDocRefNum)

                    If CopyDocument(SAPbobsCOM.BoObjectTypes.oInvoices, sDocEntry, SAPbobsCOM.BoObjectTypes.oCreditNotes, oCompany, sErrDesc, False) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                End If

                dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
                If dtBP.DefaultView.Count = 0 Then
                    sErrDesc = "Cardcode :: " & sCardCode & " provided Does Not Exist in SAP."
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If


                'Header Informations :

                oDoc.CardCode = sCardCode
                oDoc.NumAtCard = sDocRefNum
                oDoc.DocDate = oDVDetails(0)(5).ToString.Trim()
                oDoc.TaxDate = oDVDetails(0)(5).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_RefDocNum").Value = sDocRefNum

                oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

                oDoc.UserFields.Fields.Item("U_AI_APPLICANT").Value = oDVDetails(0)(1).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_EXAMINER").Value = oDVDetails(0)(2).ToString.Trim()

                For IntRow As Integer = 0 To oDVDetails.Count - 1

                    sItemCode = oDVDetails(IntRow)(4).ToString.Trim()
                    '' sBranch = oDVDetails(IntRow)(13).ToString.Trim()

                    dtItem.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtItem.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & "  provided Does Not Exist in SAP."
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    'dtBranch.DefaultView.RowFilter = "OcrCode = '" & sBranch & "'"
                    'If dtBranch.DefaultView.Count = 0 Then
                    '    sErrDesc = "Branch :: " & sBranch & "  provided Does Not Exist in SAP."
                    '    Console.WriteLine(sErrDesc)
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'End If

                    'Line Informations:
                    oDoc.Lines.ItemCode = sItemCode
                    oDoc.Lines.Quantity = 1
                    oDoc.Lines.LineTotal = CDbl(oDVDetails(IntRow)(6).ToString.Trim())
                    '' oDoc.Lines.UserFields.Fields.Item("U_AI_CompanyName").Value = oDVDetails(IntRow)(12).ToString.Trim()
                    oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = oDVDetails(IntRow)(5).ToString.Trim()

                    If CDbl(oDVDetails(IntRow)(7).ToString.Trim()) > 0 Then
                        oDoc.Lines.VatGroup = "G1"
                    Else
                        oDoc.Lines.VatGroup = "G3"
                    End If

                    If Not sBranch = String.Empty Then
                        oDoc.Lines.CostingCode = sBranch
                        oDoc.Lines.COGSCostingCode = sBranch
                    End If

                    oDoc.Lines.Add()

                Next

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR Invoice.", sFuncName)

                lRetCode = oDoc.Add()

                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrDesc)
                    Console.WriteLine("Adding AR Invoice Failed. Invoice Number : " & sDocRefNum & " Error : " & sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR Invoice failed.", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                Console.WriteLine("AR Invoice Added Successfully. Invoice Number : " & sDocRefNum)


            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Succesfully.", sFuncName)
            AddARInvoice_Item = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_Item = RTN_ERROR
            sErrDesc = ex.Message
            Console.WriteLine("Adding AR Invoice Failed. Invoice Number : " & sDocRefNum & " Error : " & sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Public Function CopyDocument(ByVal oBaseType As SAPbobsCOM.BoObjectTypes, _
                     ByVal oBaseEntry As Integer, ByVal oTarType As SAPbobsCOM.BoObjectTypes, _
                     ByRef oDICompany As SAPbobsCOM.Company, ByRef sErrDesc As String, ByRef IsDraft As Boolean) As Long

        Dim lRetCode As Double
        Dim sFuncName As String = String.Empty
        Dim oBaseDoc As SAPbobsCOM.Documents = _
              oDICompany.GetBusinessObject(oBaseType)
        Dim oTarDoc As SAPbobsCOM.Documents = _
                   oDICompany.GetBusinessObject(oTarType)
        Try
            sFuncName = "CopyDocument()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Console.WriteLine("Starting Function...")

            If oBaseDoc.GetByKey(oBaseEntry) Then
                'base document found, copy to target doc

                If IsDraft = False Then
                    oTarDoc = oDICompany.GetBusinessObject(oTarType)
                Else
                    oTarDoc = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                    oTarDoc.DocObjectCode = oTarType
                End If


                'todo: copy the cardcode, docduedate, lines
                oTarDoc.CardCode = oBaseDoc.CardCode
                oTarDoc.DocDueDate = oBaseDoc.DocDueDate
                oTarDoc.DocDate = oBaseDoc.DocDate
                oTarDoc.TaxDate = oBaseDoc.TaxDate
                oTarDoc.NumAtCard = oBaseDoc.NumAtCard

                oTarDoc.UserFields.Fields.Item("U_AI_RefDocNum").Value = oBaseDoc.UserFields.Fields.Item("U_AI_RefDocNum").Value
                oTarDoc.UserFields.Fields.Item("U_AI_Applicant").Value = oBaseDoc.UserFields.Fields.Item("U_AI_Applicant").Value
                oTarDoc.UserFields.Fields.Item("U_AI_Examiner").Value = oBaseDoc.UserFields.Fields.Item("U_AI_Examiner").Value

                'copy the lines
                Dim count As Integer = oBaseDoc.Lines.Count - 1
                Dim oTargetLines As SAPbobsCOM.Document_Lines = oTarDoc.Lines
                For i As Integer = 0 To count
                    If i <> 0 Then
                        oTargetLines.Add()
                    End If
                    oTargetLines.BaseType = oBaseType
                    oTargetLines.BaseEntry = oBaseEntry
                    oTargetLines.BaseLine = i

                    oTargetLines.UserFields.Fields.Item("U_AI_CompanyName").Value = oBaseDoc.Lines.UserFields.Fields.Item("U_AI_CompanyName").Value
                    oTargetLines.UserFields.Fields.Item("U_AI_VisitDate").Value = oBaseDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value

                Next

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Add Credit Note ", sFuncName)

                lRetCode = oTarDoc.Add()

                If lRetCode <> 0 Then
                    sErrDesc = oDICompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    Console.WriteLine(sErrDesc)
                    Return RTN_ERROR
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                    Console.WriteLine("Completed withh SUCCESS")
                    Return RTN_SUCCESS
                End If

            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Base Document Not Found ", sFuncName)
                Return RTN_SUCCESS
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return RTN_ERROR

        Finally
            oTarDoc = Nothing
            oBaseDoc = Nothing
        End Try
    End Function

End Module
