
Module modCSN

  Private isError As Boolean
    Private dtBP As DataTable
    Private dtcsCnt As DataTable
    Private dtINVHDr As DataTable
    Private dtCNHDr As DataTable
    Private dtItem As DataTable
    Private dtAC As DataTable

    ' Company Default Structure

    Private Structure ARHeader

        Public InvoiceNumber As String
        Public GLCode As String
        Public Quantity As Double
        Public GrossTotal As Double
        Public DocumentDate As Date
        Public CustomerCode As String
        Public CustomerName As String

    End Structure

    Private oARHeaderDef As ARHeader
    Private InputFolderPath As String = p_oCompDef.sInboxDir

    Public Function UploadCSN(ByRef sErrdesc As String) As Long

        'Event      :   UploadCSN()
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
            sFuncName = "UploadCSN"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo

            files = DirInfo.GetFiles("CSN*.csv")

            For Each File As System.IO.FileInfo In files
                IsFileExist = True

                oDvHdr = GetDataViewFromCSV(File.FullName)
                sFileName = File.Name
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

                If IsFileExist = False Then
                    Console.WriteLine("No AR CSV Available for Creation")
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("No AR CSV available for updation", sFuncName)
                Else

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
                End If
            Next

            If IsFileExist = False Then
                Console.WriteLine("No CSV Available for Creation")
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("No AR CSV available for updation", sFuncName)
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

            UploadCSN = RTN_SUCCESS

        Catch ex As Exception
            UploadCSN = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uplodiang AR file.", sFuncName)
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
                    .GLCode = dv(i)(2).ToString
                    .Quantity = dv(i)(4)
                    .GrossTotal = dv(i)(6)
                    .DocumentDate = dv(i)(7)
                    .CustomerCode = dv(i)(10).ToString
                    .CustomerName = dv(i)(11).ToString

                    If .InvoiceNumber.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Invoice Number is Mandatory", sFuncName)
                    End If

                    If .GLCode.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". GL Code is Mandatory", sFuncName)
                    End If

                    If .CustomerCode.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Customer Code is Mandatory", sFuncName)
                    End If

                    If .CustomerName.Length = 0 Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Direct Invoicing is Mandatory", sFuncName)
                    End If

                    If .DocumentDate.ToString() = String.Empty Then
                        isError = True
                        Call WriteToLogFile("Check line no : " & CStr(i + 1) & " in " & CurrFileToUpload & ". Sent On is Mandatory", sFuncName)
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

            Console.WriteLine("Start the SAP Transaction on Company DB :: " & oDIComp.CompanyName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDIComp.CompanyDB, sFuncName)
            oDIComp.StartTransaction()

            sSQL = "Select ""CardCode"" From OCRD WHERE ""CardType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)

            dtBP = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Invoice Ref.Doc NO.", sFuncName)
            sSQL = "SELECT ""U_AI_RefDocNum"" FROM OINV WHERE ""U_AI_RefDocNum"" IS NOT NULL"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtINVHDr = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening a datatable for Account Code.", sFuncName)
            sSQL = "SELECT ""AcctCode"" FROM OACT"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtAC = ExecuteSQLQueryForDT(sSQL, oDIComp.CompanyDB)

            Console.WriteLine("Calling AddARInvoice_Service()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_Service() ", sFuncName)
            If AddARInvoice_Service(oDIComp, oDvHdr, sFileName, sErrdesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrdesc)
            End If


            Console.WriteLine("Committing All Trasactions.. ")

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

    Private Function AddARInvoice_Service(ByVal oCompany As SAPbobsCOM.Company, ByVal oDVDetails As DataView, _
                                          ByVal sFileName As String, _
                                           ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "AddARInvoice_Service"

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
        Dim sGLAccount As String = String.Empty
        Dim sBranch As String = String.Empty


        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)

            For iDVRowCount As Integer = 1 To oDVDetails.Count - 1

                sDocRefNum = oDVDetails(iDVRowCount)(0).ToString.Trim()
                sGLAccount = oDVDetails(iDVRowCount)(2).ToString.Trim()
                sCardCode = oDVDetails(iDVRowCount)(10).ToString.Trim()

                sBranch = oDVDetails(iDVRowCount)(13).ToString.Trim()
                Console.WriteLine("Creating A/R Service Invoice. Document No :: " & sDocRefNum)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/R Service Invoice. Document No :: " & sDocRefNum, sFuncName)

                oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

                dtINVHDr.DefaultView.RowFilter = "U_AI_RefDocNum = '" & sDocRefNum & "'"
                If dtINVHDr.DefaultView.Count <> 0 Then
                    sErrDesc = "Document NO :: " & sDocRefNum & " Already Exist in SAP."
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                dtAC.DefaultView.RowFilter = "AcctCode = '" & sGLAccount & "'"
                If dtAC.DefaultView.Count = 0 Then
                    sErrDesc = "AcctCode :: " & sGLAccount & "  provided Does Not Exist in SAP."
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


                'Header Informations :

                oDoc.CardCode = sCardCode
                oDoc.NumAtCard = sDocRefNum
                oDoc.DocDate = oDVDetails(iDVRowCount)(7).ToString.Trim()
                oDoc.TaxDate = oDVDetails(iDVRowCount)(7).ToString.Trim()
                oDoc.UserFields.Fields.Item("U_AI_RefDocNum").Value = sDocRefNum

                oDoc.JournalMemo = Left(oDVDetails(iDVRowCount)(3).ToString.Trim(), 50)

                oDoc.Comments = oDVDetails(iDVRowCount)(3).ToString.Trim()

                oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName


                If Len(oDVDetails(iDVRowCount)(9).ToString.Trim()) > 40 Then
                    oDoc.UserFields.Fields.Item("U_AI_INSURER").Value = Left(oDVDetails(iDVRowCount)(9).ToString.Trim(), 40)
                Else
                    oDoc.UserFields.Fields.Item("U_AI_INSURER").Value = oDVDetails(iDVRowCount)(9).ToString.Trim()
                End If
                oDoc.UserFields.Fields.Item("U_AI_Nature").Value = oDVDetails(iDVRowCount)(12).ToString.Trim()


                'Line Informations:
                oDoc.Lines.AccountCode = sGLAccount
                If oDVDetails(iDVRowCount)(3).ToString.Trim().Length > 100 Then
                    oDoc.Lines.ItemDescription = Left(oDVDetails(iDVRowCount)(3).ToString.Trim(), 100)
                Else
                    oDoc.Lines.ItemDescription = oDVDetails(iDVRowCount)(3).ToString.Trim()
                End If

                'oDoc.Lines.LineTotal = CDbl(oDVDetails(iDVRowCount)(6).ToString.Trim())

                oDoc.Lines.PriceAfterVAT = CDbl(oDVDetails(iDVRowCount)(6).ToString.Trim())

                oDoc.Lines.UserFields.Fields.Item("U_AI_BenefitType").Value = oDVDetails(iDVRowCount)(8).ToString.Trim()
                'oDoc.Lines.CostingCode2 = oDVDetails(iDVRowCount)(13).ToString.Trim()

                If CDbl(oDVDetails(iDVRowCount)(5).ToString.Trim()) > 0 Then
                    oDoc.Lines.VatGroup = "G1"
                Else
                    oDoc.Lines.VatGroup = "G2"
                End If

                If Not sBranch = String.Empty Then
                    oDoc.Lines.CostingCode = sBranch
                    oDoc.Lines.COGSCostingCode = sBranch
                End If


                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR Invoice.", sFuncName)

                lRetCode = oDoc.Add()

                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrDesc)
                    Console.WriteLine("Adding AR Invoice Failed. Invoice Number : " & sDocRefNum & " Error : " & sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR Invoice failed.", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                Console.WriteLine("AR Service Invoice Added Successfully. Invoice Number : " & sDocRefNum)

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice_Service = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_Service = RTN_ERROR
            sErrDesc = ex.Message
            Console.WriteLine("Adding AR Invoice Failed. Invoice Number : " & sDocRefNum & " Error : " & sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

End Module
