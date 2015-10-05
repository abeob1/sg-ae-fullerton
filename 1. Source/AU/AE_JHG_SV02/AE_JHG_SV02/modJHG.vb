Module modJHG

    Private isError As Boolean

    Private Structure ARHeader_InHouse

        Public sSODocNum As String
        Public sCustomerCode As String

    End Structure

    Private Structure ARHeader_External

        Public sSODocNum As String
        Public sCustomerCode As String
        Public sVendorCode As String

    End Structure

    Private oARHeaderDef_Inhouse As ARHeader_InHouse
    Private oARHeaderDef_External As ARHeader_External


    Private InputFolderPath As String = p_oCompDef.sInboxDir
    Private dtBP As DataTable = New DataTable
    Private dtSO As DataTable = New DataTable
    Private dtINVHDr As DataTable
    Private dtSODetail As DataTable


    Public Function UploadJHG(ByRef sErrdesc As String) As Long

        'Event      :   UploadJHG()
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
            sFuncName = "UploadJHG"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo


            files = DirInfo.GetFiles("JHG*.csv")

            For Each File As System.IO.FileInfo In files
                IsFileExist = True

                oDvHdr = GetDataViewFromCSV(File.FullName)
                sFileName = File.Name

                sFileType = Mid(sFileName, 5, 2)

                If sFileType = "AR" Then
                    Console.WriteLine("Calling readCSVFileAR_Header for " & File.FullName & " to check error in CSV file")
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling readCSVFileAR_Header for " & File.FullName & " to check error in CSV file", sFuncName)
                    readCSVFileAR_Header_InHouse(File.FullName, oDvHdr)
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
                    If ProcessARFile_InHouse(oDvHdr, sFileName, sErrdesc) <> RTN_SUCCESS Then
                        AddDataToTable(p_oDtError, File.Name, "Error", sErrdesc)
                        Console.WriteLine("Calling FileMoveToArchive()")
                        FileMoveToArchive(File, File.FullName, RTN_ERROR)
                        Throw New ArgumentException(sErrdesc)
                    Else
                        'Mvoe files

                        Console.WriteLine("Calling FileMoveToArchive()")
                        FileMoveToArchive(File, File.FullName, RTN_SUCCESS)

                    End If


                ElseIf sFileType = "PO" Then

                    Console.WriteLine("Calling readCSVFileAR_Header_External for " & File.FullName & " to check error in CSV file")
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling readCSVFileAR_Header_External for " & File.FullName & " to check error in CSV file", sFuncName)
                    readCSVFileAR_Header_External(File.FullName, oDvHdr)
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

                    Console.WriteLine("Calling ProcessARFile_External()")
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ProcessARFile()", sFuncName)
                    If ProcessARFile_External(oDvHdr, sFileName, sErrdesc) <> RTN_SUCCESS Then

                        AddDataToTable(p_oDtError, File.Name, "Error", sErrdesc)
                        Console.WriteLine("Calling FileMoveToArchive()")
                        FileMoveToArchive(File, File.FullName, RTN_ERROR)

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
                Console.WriteLine("No AR or PO CSV Available for Creation")
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("No AR CSV available for updation", sFuncName)
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

            UploadJHG = RTN_SUCCESS

        Catch ex As Exception
            UploadJHG = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uplodiang AR file.", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function

    Private Sub readCSVFileAR_Header_InHouse(ByVal CurrFileToUpload As String, ByVal dv As DataView)

        'Event      :   readCSVFileAR_Header
        'Purpose    :   For reading of Business Partner CSV file
        'Author     :   Sri 
        'Date       :   24 NOV 2013 

        Dim sFuncName As String = "readCSVFileAR_Header_InHouse"
        Dim sSQL As String = String.Empty
        'Dim dt As DataTable
        Dim sErrDesc As String = String.Empty

        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            For i As Integer = 1 To dv.Count - 1
                With oARHeaderDef_Inhouse
                    .sSODocNum = dv(i)(9).ToString
                    .sCustomerCode = dv(i)(0).ToString

                    If .sSODocNum.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". SalesOrder Number is Mandatory", sFuncName)
                    End If

                    If .sCustomerCode.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Customer Code is Mandatory", sFuncName)
                    End If

                End With
            Next

        Catch ex As Exception
            isError = True
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Private Sub readCSVFileAR_Header_External(ByVal CurrFileToUpload As String, ByVal dv As DataView)

        'Event      :   readCSVFileAR_Header
        'Purpose    :   For reading of Business Partner CSV file
        'Author     :   Sri 
        'Date       :   24 NOV 2013 

        Dim sFuncName As String = "readCSVFileAR_Header_External"
        Dim sSQL As String = String.Empty
        'Dim dt As DataTable
        Dim sErrDesc As String = String.Empty

        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            For i As Integer = 1 To dv.Count - 1
                With oARHeaderDef_External

                    .sSODocNum = dv(i)(9).ToString
                    .sCustomerCode = dv(i)(0).ToString
                    .sVendorCode = dv(i)(12).ToString

                    If .sSODocNum.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". SalesOrder Number is Mandatory", sFuncName)
                    End If

                    If .sCustomerCode.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Customer Code is Mandatory", sFuncName)
                    End If

                    If .sVendorCode.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Vendor Code is Mandatory", sFuncName)
                    End If


                End With
            Next

        Catch ex As Exception
            isError = True
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Private Function ProcessARFile_InHouse(ByVal oDvHdr As DataView, _
                            ByVal sFileName As String, _
                            ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "ProcessARFile_InHouse"
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

            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Invoice Ref.Doc NO.", sFuncName)
            'sSQL = "SELECT ""DocEntry"",""U_AI_RefDocNum"" FROM OINV WHERE ""U_AI_RefDocNum"" IS NOT NULL and ""DocStatus""='O'"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            'dtINVHDr = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Sales Order Number.", sFuncName)
            sSQL = "SELECT ""DocNum"" FROM ORDR"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtSO = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)
 

            Console.WriteLine("Start the SAP Transaction on Company DB :: " & oDIComp.CompanyName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDIComp.CompanyDB, sFuncName)
            oDIComp.StartTransaction()

            Console.WriteLine("Calling AddARInvoice_Item()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_Item() ", sFuncName)
            If AddARInvoice_InHouse(oDIComp, oDvHdr, sErrdesc) <> RTN_SUCCESS Then
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


            ProcessARFile_InHouse = RTN_SUCCESS
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
            ProcessARFile_InHouse = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uploading AR File", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function

    Private Function ProcessARFile_External(ByVal oDvHdr As DataView, _
                            ByVal sFileName As String, _
                            ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "ProcessARFile_External"
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
            sSQL = "Select ""CardCode"" From OCRD WHERE ""CardType"" IN( 'C','S')"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtBP = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Invoice Ref.Doc NO.", sFuncName)
            'sSQL = "SELECT ""DocEntry"",""U_AI_RefDocNum"" FROM OINV WHERE ""U_AI_RefDocNum"" IS NOT NULL and ""DocStatus""='O'"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            'dtINVHDr = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Sales Order Number.", sFuncName)
            sSQL = "SELECT ""DocNum"" FROM ORDR"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtSO = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)


            Console.WriteLine("Start the SAP Transaction on Company DB :: " & oDIComp.CompanyName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDIComp.CompanyDB, sFuncName)
            oDIComp.StartTransaction()

            Console.WriteLine("Calling AddARInvoice_External()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_External() ", sFuncName)
            If AddARInvoice_External(oDIComp, oDvHdr, sErrdesc) <> RTN_SUCCESS Then
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


            ProcessARFile_External = RTN_SUCCESS
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
            ProcessARFile_External = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uploading AR File", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function

    Private Function AddARInvoice_InHouse(ByVal oCompany As SAPbobsCOM.Company, ByVal oDVDetails As DataView, _
                                           ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "AddARInvoice_InHouse"

        Dim oDoc As SAPbobsCOM.Documents

        Dim lRetCode, lErrCode As Long
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim sRemarks As String = String.Empty
        Dim dCOGSAmt As Double = 0
        Dim sCOGSGLAccount As String = String.Empty
        Dim sInvGLAccount As String = String.Empty
        Dim sCostCenter As String = String.Empty


        Dim sCardCode As String = String.Empty
        Dim sDocRefNum As String = String.Empty
        Dim sSODocNum As String = String.Empty
        Dim sBranch As String = String.Empty
        Dim sClientName As String = String.Empty
        Dim oDTDistinctHeader As DataTable = Nothing
        Dim sSQL As String = String.Empty


        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)

            oDTDistinctHeader = oDVDetails.Table.DefaultView.ToTable(True, "F16")

            For IntHeader As Integer = 0 To oDTDistinctHeader.Rows.Count - 1

                sDocRefNum = oDTDistinctHeader.Rows(IntHeader).Item(0).ToString.Trim

                If sDocRefNum = String.Empty Then Continue For

                oDVDetails.RowFilter = "F16 = '" & sDocRefNum & "'"

                sSODocNum = oDVDetails(0)(9).ToString.Trim()
                sCardCode = oDVDetails(0)(0).ToString.Trim()
                sBranch = oDVDetails(0)(11).ToString.Trim()
                sClientName = oDVDetails(0)(12).ToString.Trim()

                Console.WriteLine("Creating A/R Invoice. Document No :: " & sDocRefNum)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/R Service Invoice. Document No :: " & sDocRefNum, sFuncName)

                oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices

                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items

                dtSO.DefaultView.RowFilter = "DocNum = '" & sSODocNum & "'"
                If dtSO.DefaultView.Count = 0 Then
                    sErrDesc = "SO DocNum :: " & sSODocNum & "  provided does Not Exist in SAP."
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
                If dtBP.DefaultView.Count = 0 Then
                    sErrDesc = "Cardcode :: " & sCardCode & " provided does Not Exist in SAP."
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                oDoc.CardCode = sCardCode
                oDoc.NumAtCard = sSODocNum
                oDoc.Comments = oDVDetails(0)(14).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_RefDocNum").Value = sDocRefNum
                oDoc.UserFields.Fields.Item("U_AI_APPLICANT").Value = oDVDetails(0)(2).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_POSITION").Value = oDVDetails(0)(3).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_SITE").Value = oDVDetails(0)(4).ToString.Trim()

                oDoc.UserFields.Fields.Item("U_AI_PROJECT").Value = oDVDetails(0)(5).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_VISITDATE").Value = oDVDetails(0)(6).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_ReqName").Value = oDVDetails(0)(7).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_ChitNo").Value = oDVDetails(0)(8).ToString.Trim()

                oDoc.UserFields.Fields.Item("U_AI_BOOKINGOFFICER").Value = oDVDetails(0)(10).ToString.Trim()
                ' oDoc.UserFields.Fields.Item("U_AI_EXAMINER").Value = oDVDetails(iDVRowCount)(16).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_Nature").Value = oDVDetails(0)(17).ToString.Trim()


                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for SO Details", sFuncName)

                sSQL = "SELECT T0.""ItemCode"",T0.""Quantity"",T0.""Price"",T0.""VatSum"" FROM RDR1 T0" & _
                    " INNER JOIN ORDR T1 ON T0.""DocEntry""=T1.""DocEntry"" WHERE T1.""DocNum""=" & sSODocNum & ""

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtSODetail = ExecuteSQLQueryForDT(sSQL, oCompany.CompanyDB)

                ''Line Informations:
                For iCount As Integer = 0 To dtSODetail.Rows.Count - 1

                    oDoc.Lines.ItemCode = dtSODetail.Rows(iCount)("ItemCode").ToString().Trim()
                    oDoc.Lines.Quantity = dtSODetail.Rows(iCount)("Quantity").ToString().Trim()
                    oDoc.Lines.Price = CDbl(dtSODetail.Rows(iCount)("Price").ToString().Trim())

                    If CDbl(dtSODetail.Rows(iCount)("VatSum").ToString().Trim()) > 0 Then
                        oDoc.Lines.VatGroup = "G1"
                    Else
                        oDoc.Lines.VatGroup = "G3"
                    End If

                    If Not sBranch = String.Empty Then
                        oDoc.Lines.CostingCode2 = sBranch
                        oDoc.Lines.COGSCostingCode2 = sBranch
                    End If

                    'If Not sClientName = String.Empty Then
                    '    oDoc.Lines.CostingCode5 = sClientName
                    '    oDoc.Lines.COGSCostingCode5 = sClientName
                    'End If

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

                Console.WriteLine("AR Invoice added successfully. Invoice Number : " & sDocRefNum)
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice_InHouse = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_InHouse = RTN_ERROR
            sErrDesc = ex.Message
            Console.WriteLine("Adding AR Invoice Failed. Invoice Number : " & sDocRefNum & " Error : " & sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Function AddARInvoice_External(ByVal oCompany As SAPbobsCOM.Company, ByVal oDVDetails As DataView, _
                                           ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "AddARInvoice_External"

        Dim oDoc As SAPbobsCOM.Documents
        Dim oPODoc As SAPbobsCOM.Documents

        Dim lRetCode, lErrCode As Long
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim sRemarks As String = String.Empty
        Dim dCOGSAmt As Double = 0
        Dim sCOGSGLAccount As String = String.Empty
        Dim sInvGLAccount As String = String.Empty
        Dim sCostCenter As String = String.Empty


        Dim sCardCode As String = String.Empty
        Dim sVendorCode As String = String.Empty
        Dim sDocRefNum As String = String.Empty
        Dim sSODocNum As String = String.Empty
        Dim sBranch As String = String.Empty
        Dim sClientName As String = String.Empty
        Dim oDTDistinctHeader As DataTable = Nothing

        Dim sSQL As String = String.Empty


        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)

            oDTDistinctHeader = oDVDetails.Table.DefaultView.ToTable(True, "F17")

            For IntHeader As Integer = 0 To oDTDistinctHeader.Rows.Count - 1

                sDocRefNum = oDTDistinctHeader.Rows(IntHeader).Item(0).ToString.Trim

                If (sDocRefNum = String.Empty Or sDocRefNum.ToString.ToUpper() = "MEDICAL TYPE") Then Continue For

                    oDVDetails.RowFilter = "F17 = '" & sDocRefNum & "'"

                    sSODocNum = oDVDetails(0)(9).ToString.Trim()
                    sCardCode = oDVDetails(0)(0).ToString.Trim()
                    sBranch = oDVDetails(0)(11).ToString.Trim()
                    sClientName = oDVDetails(0)(13).ToString.Trim()
                    sVendorCode = oDVDetails(0)(12).ToString.Trim()

                    Console.WriteLine("Creating A/R Invoice. Document No :: " & sDocRefNum)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/R Service Invoice. Document No :: " & sDocRefNum, sFuncName)

                    oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices

                    oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items

                    dtSO.DefaultView.RowFilter = "DocNum = '" & sSODocNum & "'"
                    If dtSO.DefaultView.Count = 0 Then
                        sErrDesc = "SO DocNum :: " & sSODocNum & "  provided Does Not Exist in SAP."
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

                    dtBP.DefaultView.RowFilter = "CardCode = '" & sVendorCode & "'"
                    If dtBP.DefaultView.Count = 0 Then
                        sErrDesc = "VendorCode :: " & sVendorCode & " provided Does Not Exist in SAP."
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    oDoc.CardCode = sCardCode
                    oDoc.NumAtCard = sSODocNum
                    oDoc.Comments = oDVDetails(0)(15).ToString.Trim()
                    oDoc.UserFields.Fields.Item("U_AI_RefDocNum").Value = sDocRefNum
                    oDoc.UserFields.Fields.Item("U_AI_APPLICANT").Value = oDVDetails(0)(2).ToString.Trim()
                    oDoc.UserFields.Fields.Item("U_AI_POSITION").Value = oDVDetails(0)(3).ToString.Trim()
                    oDoc.UserFields.Fields.Item("U_AI_SITE").Value = oDVDetails(0)(4).ToString.Trim()

                    oDoc.UserFields.Fields.Item("U_AI_PROJECT").Value = oDVDetails(0)(5).ToString.Trim()
                    oDoc.UserFields.Fields.Item("U_AI_VISITDATE").Value = oDVDetails(0)(6).ToString.Trim()
                    oDoc.UserFields.Fields.Item("U_AI_ReqName").Value = oDVDetails(0)(7).ToString.Trim()
                    oDoc.UserFields.Fields.Item("U_AI_ChitNo").Value = oDVDetails(0)(8).ToString.Trim()

                    oDoc.UserFields.Fields.Item("U_AI_BOOKINGOFFICER").Value = oDVDetails(0)(10).ToString.Trim()
                    ' oDoc.UserFields.Fields.Item("U_AI_EXAMINER").Value = oDVDetails(iDVRowCount)(16).ToString.Trim()
                    oDoc.UserFields.Fields.Item("U_AI_Nature").Value = oDVDetails(0)(18).ToString.Trim()


                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for SO Details", sFuncName)

                    sSQL = "SELECT T0.""ItemCode"",T0.""Quantity"",T0.""Price"",T0.""VatSum"" FROM RDR1 T0" & _
                        " INNER JOIN ORDR T1 ON T0.""DocEntry""=T1.""DocEntry"" WHERE T1.""DocNum""=" & sSODocNum & ""

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                    dtSODetail = ExecuteSQLQueryForDT(sSQL, oCompany.CompanyDB)

                    ''Line Informations:
                    For iCount As Integer = 0 To dtSODetail.Rows.Count - 1

                        oDoc.Lines.ItemCode = dtSODetail.Rows(iCount)("ItemCode").ToString().Trim()
                        oDoc.Lines.Quantity = dtSODetail.Rows(iCount)("Quantity").ToString().Trim()
                        oDoc.Lines.Price = CDbl(dtSODetail.Rows(iCount)("Price").ToString().Trim())

                        If CDbl(dtSODetail.Rows(iCount)("VatSum").ToString().Trim()) > 0 Then
                            oDoc.Lines.VatGroup = "G1"
                        Else
                            oDoc.Lines.VatGroup = "G3"
                        End If

                        ''  sBranch = dtSODetail.Rows(iCount)("OcrCode2").ToString().Trim()

                        If sBranch <> String.Empty Then
                            oDoc.Lines.CostingCode2 = sBranch
                            oDoc.Lines.COGSCostingCode2 = sBranch
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

                    Else


                        oPODoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

                        oPODoc.CardCode = sVendorCode
                        oPODoc.NumAtCard = sSODocNum
                        oPODoc.Comments = oDVDetails(0)(15).ToString.Trim()
                        oPODoc.UserFields.Fields.Item("U_AI_RefDocNum").Value = sDocRefNum
                        oPODoc.UserFields.Fields.Item("U_AI_APPLICANT").Value = oDVDetails(0)(2).ToString.Trim()
                        oPODoc.UserFields.Fields.Item("U_AI_POSITION").Value = oDVDetails(0)(3).ToString.Trim()
                        oPODoc.UserFields.Fields.Item("U_AI_SITE").Value = oDVDetails(0)(4).ToString.Trim()

                        oPODoc.UserFields.Fields.Item("U_AI_PROJECT").Value = oDVDetails(0)(5).ToString.Trim()
                        oPODoc.UserFields.Fields.Item("U_AI_VISITDATE").Value = oDVDetails(0)(6).ToString.Trim()
                        oPODoc.UserFields.Fields.Item("U_AI_ReqName").Value = oDVDetails(0)(7).ToString.Trim()
                        oPODoc.UserFields.Fields.Item("U_AI_ChitNo").Value = oDVDetails(0)(8).ToString.Trim()

                        oPODoc.UserFields.Fields.Item("U_AI_BOOKINGOFFICER").Value = oDVDetails(0)(10).ToString.Trim()
                        ' oDoc.UserFields.Fields.Item("U_AI_EXAMINER").Value = oDVDetails(iDVRowCount)(16).ToString.Trim()
                        oPODoc.UserFields.Fields.Item("U_AI_Nature").Value = oDVDetails(0)(18).ToString.Trim()

                        ''Line Informations:
                        For iCount As Integer = 0 To dtSODetail.Rows.Count - 1

                            oPODoc.Lines.ItemCode = dtSODetail.Rows(iCount)("ItemCode").ToString().Trim()
                            oPODoc.Lines.Quantity = dtSODetail.Rows(iCount)("Quantity").ToString().Trim()
                            ' oPODoc.Lines.Price = CDbl(dtSODetail.Rows(iCount)("Price").ToString().Trim())
                            oPODoc.Lines.UnitPrice = 0.0

                            If CDbl(dtSODetail.Rows(iCount)("VatSum").ToString().Trim()) > 0 Then
                                oPODoc.Lines.VatGroup = "G10"
                            Else
                                oPODoc.Lines.VatGroup = "G14"
                            End If

                            oPODoc.Lines.Add()

                        Next

                    End If

                    lRetCode = oPODoc.Add()

                    If lRetCode <> 0 Then
                        oCompany.GetLastError(lErrCode, sErrDesc)
                        Console.WriteLine("Adding PO Failed. Invoice Number : " & sDocRefNum & " Error : " & sErrDesc)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PO Adding failed.", sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Console.WriteLine("AR Invoice and PO Added Successfully. Invoice Number : " & sDocRefNum)

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)

            AddARInvoice_External = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_External = RTN_ERROR
            sErrDesc = ex.Message
            Console.WriteLine("Adding AR Invoice Failed. Invoice Number : " & sDocRefNum & " Error : " & sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

End Module
