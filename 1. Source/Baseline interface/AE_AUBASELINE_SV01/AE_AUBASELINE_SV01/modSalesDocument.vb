Module modSalesDocument

    Private dtCardCode As DataTable
    Private dtGlAccount As DataTable
    Private dtTaxType As DataTable
    Private dtOcrCode5 As DataTable
    Private dtProject As DataTable

    Public Function ProcessSalesDocFile(ByVal file As System.IO.FileInfo, ByVal oDv As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessSalesDocFile"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If CompanyConnection(p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then
                Console.WriteLine("Company connected successfully")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Company connected successfully", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction() ", sFuncName)
                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                sSql = "SELECT ""CardCode"",UPPER(LEFT(""CardName"",100)) ""CardName"" FROM ""OCRD"" WHERE ""CardType"" = 'C' "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                dtCardCode = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

                sSql = "SELECT ""U_SAPGLCode"",UPPER(""Code"") AS ""Code"" FROM ""@AE_XERO_GL"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                dtGlAccount = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

                sSql = "SELECT ""U_SAPTax"",UPPER(""Code"") AS ""Code"",""Name"" FROM ""@AE_XERO_GST"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                dtTaxType = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

                sSql = "SELECT ""PrcCode"",UPPER(""PrcCode"") AS ""PrcCodeUpperCase"" FROM ""OPRC"" WHERE ""DimCode"" = 5"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                dtOcrCode5 = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

                sSql = "SELECT ""PrjCode"",UPPER(""PrjName"") AS ""PrjName"" FROM ""OPRJ"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                dtProject = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

                Dim oDtGroup As DataTable
                oDtGroup = oDv.Table.DefaultView.ToTable(True, "F34")

                oDv.RowFilter = "ISNULL(F34,'') = ''"
                If oDv.Count > 0 Then
                    sErrDesc = "Type Column in excel sheet is mandatory"
                    Console.WriteLine(sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = "" Or oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim = "TYPE") Then

                        Console.WriteLine("Filtering dataview based on TYPE column")
                        oDv.RowFilter = "F34 = '" & oDtGroup.Rows(i).Item(0).ToString() & "' "

                        Console.WriteLine("Type is " & oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim)

                        If oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim = "SALES INVOICE" Then
                            If oDv.Count > 0 Then
                                Dim oDtSalesInvoice As DataTable
                                oDtSalesInvoice = oDv.ToTable
                                Dim oDvSalesInvoice As DataView = New DataView(oDtSalesInvoice)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateSalesInvoice() ", sFuncName)
                                If CreateSalesInvoice(oDvSalesInvoice, file.Name, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        ElseIf oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim = "SALES CREDIT NOTE" Then
                            If oDv.Count > 0 Then
                                Dim oDtCreditNote As DataTable
                                oDtCreditNote = oDv.ToTable
                                Dim oDvCreditNote As DataView = New DataView(oDtCreditNote)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateSalesCreditNote() ", sFuncName)
                                If CreateSalesCreditNote(oDvCreditNote, file.Name, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    End If
                Next

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
            If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_SUCCESS)

            'Insert Success Notificaiton into Table..
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtSuccess, file.Name, "Success")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & file.FullName, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessSalesDocFile = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction() ", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_ERROR)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessSalesDocFile = RTN_ERROR
        End Try
    End Function

    Public Function Validation(ByVal oDv As DataView, ByVal sErrDesc As String) As Long
        Dim sFuncName As String = "Validation"
        Dim sContactName As String = String.Empty
        Dim sAcctCode As String = String.Empty
        Dim sTaxType As String = String.Empty
        Dim sTrackOpt1 As String = String.Empty
        Dim sTrackOpt2 As String = String.Empty
        Dim sProject As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function ", sFuncName)

            Dim oDtTest As New DataTable
            'oDtTest = oDv.ToTable 
            Dim selected As System.Data.DataTable = oDv.ToTable("Selected", False, "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10")

            Dim stringArr = selected.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
            Dim y As String = String.Join(",", stringArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

            ''Dim rows As New Array
            'Dim rows() As Array

            'rows.Add(String.Join(",", oDtTest.Rows(0).ItemArray.[Select](Function(item) item.ToString())))
            'Dim Y As String = String.Join(",", rows.Where(Function(s) Not String.IsNullOrEmpty(s)))

            '**********************************

            Dim oDtGroup As DataTable
            oDtGroup = oDv.Table.DefaultView.ToTable(True, "F11", "F34") 'F11 - Invoice Number, F34 - Type

            Dim sInvoiceNo As String = String.Empty
            Dim sType As String = String.Empty

            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = "" Or oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim = "INVOICENUMBER") Then
                    sInvoiceNo = oDtGroup.Rows(i).Item(0).ToString().ToUpper
                    sType = oDtGroup.Rows(i).Item(1).ToString().ToUpper

                    If sType.ToUpper = "SALES CREDIT NOTE" Then
                        oDv.RowFilter = "ISNULL(F35,'') = ''"

                        If oDv.Count > 0 Then
                            sErrDesc = "Base Invoice not found for creating credit note"
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                        oDv.RowFilter = Nothing
                    End If

                    oDv.RowFilter = "F11 = '" & oDtGroup.Rows(i).Item(0).ToString().ToUpper & "' AND F34 = '" & oDtGroup.Rows(i).Item(1).ToString().ToUpper & "' "
                    If oDv.Count > 0 Then
                        For j As Integer = 0 To oDv.Count - 1

                            sContactName = oDv(j)(0).ToString.Trim
                            If sContactName.Length > 100 Then
                                sContactName = sContactName.Substring(0, 100)
                            End If
                            dtCardCode.DefaultView.RowFilter = "CardName = '" & sContactName.ToUpper & "'"
                            If dtCardCode.DefaultView.Count = 0 Then
                                sErrDesc = "Customer record not found in SAP/CardName :: " & sContactName
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If

                            sAcctCode = oDv(j)(25).ToString.Trim
                            dtGlAccount.DefaultView.RowFilter = "Code = '" & sAcctCode & "'"
                            If dtGlAccount.DefaultView.Count = 0 Then
                                sErrDesc = "Mapping of account " & sAcctCode & " cannot be found in SAP"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If

                            sTaxType = oDv(j)(26).ToString.Trim
                            dtTaxType.DefaultView.RowFilter = "Code = '" & sTaxType.ToUpper() & "' "
                            If dtTaxType.DefaultView.Count = 0 Then
                                sErrDesc = "Tax code for taxtype " & sTaxType & " not found in Tax type mapping table"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If

                            sTrackOpt1 = oDv(j)(29).ToString.Trim
                            If sTrackOpt1.Length > 8 Then
                                sTrackOpt1 = sTrackOpt1.Substring(0, 8)
                            End If
                            dtOcrCode5.DefaultView.RowFilter = "PrcCodeUpperCase = '" & sTrackOpt1.ToUpper() & "' "
                            If dtOcrCode5.DefaultView.Count = 0 Then
                                sErrDesc = "Cost center " & sTrackOpt1 & " cannot be found in SAP "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If

                            sProject = oDv(j)(30).ToString.Trim
                            If sProject.Length > 100 Then
                                sProject = sProject.Substring(0, 100)
                            End If
                            dtProject.DefaultView.RowFilter = "PrjName = '" & sProject.ToUpper() & "' "
                            If dtProject.DefaultView.Count = 0 Then
                                sErrDesc = "Project " & sProject & " cannot be found in SAP"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If

                        Next

                    End If


                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Validation = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            Console.WriteLine(sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Validation = RTN_ERROR
        End Try
    End Function

    Private Function CreateSalesInvoice(ByVal oDv_Invoice As DataView, ByVal sFileName As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateSalesInvoice"
        Dim sCardName As String = String.Empty
        Dim sContactName As String = String.Empty
        Dim sAcctCode As String = String.Empty
        Dim sTaxType As String = String.Empty
        Dim sTrackOpt1 As String = String.Empty
        Dim sProject As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sInvoiceNo As String = String.Empty
        Dim sReference As String = String.Empty
        Dim sAddress As String = String.Empty
        Dim sInvoiceDate As String = String.Empty
        Dim sDueDate As String = String.Empty
        Dim sCurrency As String = String.Empty
        Dim bIsLineAdded As Boolean = False
        Dim oRecSet As SAPbobsCOM.Recordset

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            Console.WriteLine("Creating A/R Invoice")

            Dim oDtGroup As DataTable
            oDtGroup = oDv_Invoice.Table.DefaultView.ToTable(True, "F11") 'F11 - Invoice Number

            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = "" Or oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim = "INVOICENUMBER") Then
                    Console.WriteLine("Filtering rows based on invoice number " & oDtGroup.Rows(i).Item(0).ToString())

                    oDv_Invoice.RowFilter = "F11 = '" & oDtGroup.Rows(i).Item(0).ToString() & "' "

                    If oDv_Invoice.Count > 0 Then
                        Dim oDt As New DataTable
                        oDt = oDv_Invoice.ToTable()
                        Dim oDv As DataView = New DataView(oDt)

                        If oDv.Count > 0 Then
                            sContactName = oDv(0)(0).ToString.Trim
                            If sContactName.Length > 100 Then
                                sContactName = sContactName.Substring(0, 100)
                            End If
                            dtCardCode.DefaultView.RowFilter = "CardName = '" & sContactName.ToUpper & "'"
                            If dtCardCode.DefaultView.Count = 0 Then
                                sErrDesc = "Customer record not found in SAP/CardName :: " & sContactName
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sCardCode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If

                            sInvoiceNo = oDv(0)(10).ToString.Trim()
                            sReference = oDv(0)(11).ToString.Trim()
                            sInvoiceDate = oDv(0)(12).ToString.Trim
                            sDueDate = oDv(0)(13).ToString.Trim
                            sCurrency = oDv(0)(32).ToString.Trim()

                            If sInvoiceNo = "INV-003851" Then
                                MsgBox("INV-003851 Invoice NO")
                            End If


                            Dim sSQL As String
                            sSQL = "SELECT ""DocEntry"" FROM ""OINV"" WHERE ""NumAtCard"" = '" & sInvoiceNo & "' "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                            oRecSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecSet.DoQuery(sSQL)
                            If oRecSet.RecordCount > 0 Then
                                sErrDesc = "Invoice already found for reference number :: " & sInvoiceNo
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)

                            Dim iIndex As Integer = sInvoiceDate.IndexOf(" ")
                            Dim sInvoiceDate_Trimed As String
                            If iIndex > -1 Then
                                sInvoiceDate_Trimed = sInvoiceDate.Substring(0, iIndex)
                            Else
                                sInvoiceDate_Trimed = sInvoiceDate
                            End If
                            Dim iIndex1 As Integer = sInvoiceDate.IndexOf(" ")
                            Dim sDueDate_Trimed As String
                            If iIndex1 > -1 Then
                                sDueDate_Trimed = sDueDate.Substring(0, iIndex1)
                            Else
                                sDueDate_Trimed = sDueDate
                            End If
                            Dim dInvoiceDate, dDueDate As Date
                            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd/MM/yy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                            Date.TryParseExact(sInvoiceDate_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dInvoiceDate)
                            Date.TryParseExact(sDueDate_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDueDate)

                            'Dim selected As System.Data.DataTable = oDv.ToTable("Selected", False, "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10")
                            'Dim stringArr = selected.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            'sAddress = String.Join("," & Environment.NewLine, stringArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            Dim sAddr1, sAddr2, sAddr3 As String
                            'First line
                            Dim sFirstLine As System.Data.DataTable = oDv.ToTable("FirstLine", False, "F3", "F4", "F5", "F6")
                            Dim sFirstLineArr = sFirstLine.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            sAddr1 = String.Join("," & Environment.NewLine, sFirstLineArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            'Second line
                            Dim sSecondLine As System.Data.DataTable = oDv.ToTable("SecondLine", False, "F7", "F8")
                            Dim sSecondLineArr = sSecondLine.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            sAddr2 = String.Join(",", sSecondLineArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            'Third line
                            Dim sThirdLine As System.Data.DataTable = oDv.ToTable("ThirdLine", False, "F9", "F10")
                            Dim sThirdLineArr = sThirdLine.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            sAddr3 = String.Join(",", sThirdLineArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            If sAddr1 <> "" Then
                                sAddress = sAddr1
                            End If
                            If sAddr2 <> "" Then
                                If sAddress = "" Then
                                    sAddress = sAddr2
                                Else
                                    sAddress = sAddress & "," & Environment.NewLine & sAddr2
                                End If
                            End If
                            If sAddr3 <> "" Then
                                If sAddress = "" Then
                                    sAddress = sAddr3
                                Else
                                    sAddress = sAddress & "," & Environment.NewLine & sAddr3
                                End If
                            End If

                            If sInvoiceNo.Length > 100 Then
                                sInvoiceNo = sInvoiceNo.Substring(0, 100)
                            End If
                            If sAddress.Length > 254 Then
                                sAddress = sAddress.Substring(0, 254)
                            End If

                            Dim oARInvoice As SAPbobsCOM.Documents
                            oARInvoice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                            oARInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

                            oARInvoice.CardCode = sCardCode
                            oARInvoice.NumAtCard = sInvoiceNo
                            oARInvoice.DocDate = dInvoiceDate
                            oARInvoice.DocDueDate = dDueDate
                            If sReference.Length > 254 Then
                                oARInvoice.Comments = sReference.Substring(0, 254)
                            Else
                                oARInvoice.Comments = sReference
                            End If
                            If sReference.Length > 50 Then
                                oARInvoice.JournalMemo = sReference.Substring(0, 50)
                            Else
                                oARInvoice.JournalMemo = sReference
                            End If
                            oARInvoice.Address = sAddress
                            oARInvoice.DocCurrency = sCurrency
                            Dim sSent As String = String.Empty
                            sSent = oDv(0)(34).ToString.Trim.ToUpper()
                            If sSent = "SENT" Then
                                oARInvoice.UserFields.Fields.Item("U_AI_EINVOICE").Value = "Yes"
                            Else
                                oARInvoice.UserFields.Fields.Item("U_AI_EINVOICE").Value = "No"
                            End If
                            oARInvoice.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

                            Dim iLineCount As Integer = 1

                            For j As Integer = 0 To oDv.Count - 1
                                If CDbl(oDv(j)(24).ToString.Trim) = 0 And (oDv(j)(25).ToString.Trim = String.Empty) Then
                                    oDv.RowFilter = "ISNULL(F26,0) <> 0 "
                                    If oDv.Count > 0 Then
                                        sAcctCode = oDv(0)(25).ToString.Trim
                                    Else
                                        sErrDesc = "Account Code is Mandatory/Please check the excel file for invoice " & sInvoiceNo
                                        Console.WriteLine(sErrDesc)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    End If
                                    oDv.RowFilter = Nothing
                                Else
                                    sAcctCode = oDv(j)(25).ToString.Trim
                                End If
                                dtGlAccount.DefaultView.RowFilter = "Code = '" & sAcctCode.ToUpper() & "'"
                                If dtGlAccount.DefaultView.Count = 0 Then
                                    sErrDesc = "Mapping of account " & sAcctCode & " cannot be found in SAP"
                                    Console.WriteLine(sErrDesc)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sAcctCode = dtGlAccount.DefaultView.Item(0)(0).ToString().Trim()
                                End If

                                If oDv(j)(26).ToString.Trim = String.Empty Then
                                    oDv.RowFilter = "ISNULL(F27,'') <> '' "
                                    sTaxType = oDv(0)(26).ToString.Trim
                                    oDv.RowFilter = Nothing
                                Else
                                    sTaxType = oDv(j)(26).ToString.Trim
                                End If
                                dtTaxType.DefaultView.RowFilter = "Code = '" & sTaxType.ToUpper() & "' "
                                If dtTaxType.DefaultView.Count = 0 Then
                                    sErrDesc = "Tax code for taxtype " & sTaxType & " not found in Tax type mapping table"
                                    Console.WriteLine(sErrDesc)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sTaxType = dtTaxType.DefaultView.Item(0)(0).ToString().Trim()
                                End If

                                sTrackOpt1 = oDv(j)(29).ToString.Trim
                                If sTrackOpt1 <> "" Then
                                    If sTrackOpt1.Length > 8 Then
                                        sTrackOpt1 = sTrackOpt1.Substring(0, 8)
                                    End If
                                    dtOcrCode5.DefaultView.RowFilter = "PrcCodeUpperCase = '" & sTrackOpt1.ToUpper() & "' "
                                    If dtOcrCode5.DefaultView.Count = 0 Then
                                        sErrDesc = "Cost center " & sTrackOpt1 & " cannot be found in SAP "
                                        Console.WriteLine(sErrDesc)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        sTrackOpt1 = dtOcrCode5.DefaultView.Item(0)(0).ToString().Trim()
                                    End If
                                End If

                                sProject = oDv(j)(31).ToString.Trim
                                If sProject <> "" Then
                                    If sProject.Length > 100 Then
                                        sProject = sProject.Substring(0, 100)
                                    End If
                                    dtProject.DefaultView.RowFilter = "PrjName = '" & sProject.ToUpper() & "' "
                                    If dtProject.DefaultView.Count = 0 Then
                                        sErrDesc = "Project " & sProject & " cannot be found in SAP"
                                        Console.WriteLine(sErrDesc)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        sProject = dtProject.DefaultView.Item(0)(0).ToString().Trim()
                                    End If
                                End If

                                Dim sItemDesc As String = String.Empty
                                sItemDesc = oDv(j)(20).ToString.Trim()

                                If iLineCount > 1 Then
                                    oARInvoice.Lines.Add()
                                End If
                                If sItemDesc.Length > 100 Then
                                    oARInvoice.Lines.ItemDescription = sItemDesc.Substring(0, 100)
                                Else
                                    oARInvoice.Lines.ItemDescription = sItemDesc
                                End If
                                'oARInvoice.Lines.ItemDescription = sItemDesc
                                oARInvoice.Lines.UserFields.Fields.Item("U_Dscription2").Value = oDv(j)(20).ToString.Trim()
                                oARInvoice.Lines.UserFields.Fields.Item("U_AI_QTY").Value = CDbl(oDv(j)(21).ToString.Trim)
                                oARInvoice.Lines.UserFields.Fields.Item("U_AI_PRICE").Value = CDbl(oDv(j)(22).ToString.Trim)
                                oARInvoice.Lines.AccountCode = sAcctCode
                                oARInvoice.Lines.VatGroup = sTaxType
                                oARInvoice.Lines.LineTotal = CDbl(oDv(j)(24).ToString.Trim)
                                If sTrackOpt1 <> "" Then
                                    oARInvoice.Lines.CostingCode5 = sTrackOpt1
                                    oARInvoice.Lines.COGSCostingCode5 = sTrackOpt1
                                End If
                                If sProject <> "" Then
                                    oARInvoice.Lines.ProjectCode = sProject
                                End If
                                iLineCount = iLineCount + 1
                                bIsLineAdded = True
                            Next

                            If bIsLineAdded = True Then
                                If oARInvoice.Add() <> 0 Then
                                    sErrDesc = "Error while adding Invoice / " & p_oCompany.GetLastErrorDescription
                                    Console.WriteLine(sErrDesc)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    Dim iDocNo, iDocEntry As Integer
                                    p_oCompany.GetNewObjectCode(iDocEntry)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)

                                    Dim objRS As SAPbobsCOM.Recordset
                                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim sQuery As String

                                    sQuery = "SELECT ""DocNum"" FROM ""OINV"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                                    objRS.DoQuery(sQuery)
                                    If objRS.RecordCount > 0 Then
                                        iDocNo = objRS.Fields.Item("DocNum").Value
                                    End If
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Document Created successfully :: " & iDocNo, sFuncName)
                                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice Created Successfully/ Doc No is " & iDocNo, sFuncName)

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                                End If
                            End If
                        End If

                    End If
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateSalesInvoice = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateSalesInvoice = RTN_ERROR
        End Try
    End Function

    Private Function CreateSalesCreditNote(ByVal oDv_CreditNote As DataView, ByVal sFileName As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateSalesCreditNote"
        Dim sCardName As String = String.Empty
        Dim sContactName As String = String.Empty
        Dim sAcctCode As String = String.Empty
        Dim sTaxType As String = String.Empty
        Dim sTrackOpt1 As String = String.Empty
        Dim sProject As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sInvoiceNo As String = String.Empty
        Dim sReference As String = String.Empty
        Dim sAddress As String = String.Empty
        Dim sInvoiceDate As String = String.Empty
        Dim sDueDate As String = String.Empty
        Dim sCurrency As String = String.Empty
        Dim bIsLineAdded As Boolean = False
        Dim sBaseInvoice As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim dXcelTotal As Double = 0.0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            Console.WriteLine("Creating A/R Credit Note")

            oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim oDtGroup As DataTable
            oDtGroup = oDv_CreditNote.Table.DefaultView.ToTable(True, "F11") 'F11 - Invoice Number, F37 - Base Document

            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = "" Or oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim = "INVOICENUMBER") Then
                    oDv_CreditNote.RowFilter = "F11 = '" & oDtGroup.Rows(i).Item(0).ToString().ToUpper & "' "
                    If oDv_CreditNote.Count > 0 Then
                        Dim oDt As DataTable
                        oDt = oDv_CreditNote.ToTable
                        Dim oDv As DataView = New DataView(oDt)
                        If oDv.Count > 0 Then
                            oDv.RowFilter = "ISNULL(F37,'') <> ''"
                            If oDv.Count > 0 Then
                            Else
                                sErrDesc = "Base Invoice not found for creating credit note/Please check in BaseInvoice column in Excel"
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    End If
                End If
            Next


            oDv_CreditNote.RowFilter = Nothing

            oDtGroup = oDv_CreditNote.Table.DefaultView.ToTable(True, "F11") 'F11 - Invoice Number, F37 - Base Document

            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = "" Or oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim = "INVOICENUMBER") Then
                    Console.WriteLine("Filtering rows by invoice number ")

                    oDv_CreditNote.RowFilter = "F11 = '" & oDtGroup.Rows(i).Item(0).ToString().ToUpper & "' "

                    If oDv_CreditNote.Count > 0 Then
                        Dim oDt As New DataTable
                        oDt = oDv_CreditNote.ToTable
                        Dim oDv As DataView = New DataView(oDt)

                        If oDv.Count > 0 Then
                            oDv.RowFilter = "ISNULL(F37,'') <> ''"
                            If oDv.Count > 0 Then
                                sBaseInvoice = oDv(0)(36).ToString.Trim()
                            End If
                            oDv.RowFilter = Nothing

                            Dim sInvDocEntry As String = String.Empty
                            sInvoiceNo = oDv(0)(10).ToString.Trim()
                            If sBaseInvoice.Length > 100 Then
                                sBaseInvoice = sBaseInvoice.Substring(0, 100)
                            End If

                            Dim sSQL As String
                            sSQL = "SELECT ""DocEntry"" FROM ""ORIN"" WHERE ""NumAtCard"" = '" & sInvoiceNo & "' "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                            ' oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs.DoQuery(sSQL)
                            If oRs.RecordCount > 0 Then
                                sErrDesc = "Credit Note already found for reference number :: " & sInvoiceNo
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If

                            Dim sQuery As String = String.Empty
                            sQuery = "SELECT ""DocEntry"",""DocNum"" FROM ""OINV"" WHERE ""DocStatus"" = 'O' AND ""CANCELED"" = 'N' AND UPPER(""NumAtCard"") = '" & sBaseInvoice.ToUpper() & "' "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                            oRs.DoQuery(sQuery)
                            If oRs.RecordCount > 0 Then
                                sInvDocEntry = oRs.Fields.Item("DocEntry").Value
                                Console.WriteLine("Invoice Found/DocEntry is " & sInvDocEntry)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice DocEntry is " & sInvDocEntry, sFuncName)
                            Else
                                sErrDesc = "Base Invoice Number " & sBaseInvoice & " Not found in SAP"
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If

                            sReference = oDv(0)(11).ToString.Trim()
                            sInvoiceDate = oDv(0)(12).ToString.Trim
                            sDueDate = oDv(0)(13).ToString.Trim
                            sCurrency = oDv(0)(32).ToString.Trim()
                            
                            sContactName = oDv(0)(0).ToString.Trim
                            If sContactName.Length > 100 Then
                                sContactName = sContactName.Substring(0, 100)
                            End If
                            dtCardCode.DefaultView.RowFilter = "CardName = '" & sContactName.ToUpper & "'"
                            If dtCardCode.DefaultView.Count = 0 Then
                                sErrDesc = "Customer record not found in SAP/CardName :: " & sContactName
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sCardCode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If

                            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd/MM/yy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                            Dim iIndex As Integer = sInvoiceDate.IndexOf(" ")
                            Dim sInvoiceDate_Trimed As String
                            If iIndex > -1 Then
                                sInvoiceDate_Trimed = sInvoiceDate.Substring(0, iIndex)
                            Else
                                sInvoiceDate_Trimed = sInvoiceDate
                            End If
                            Dim dInvoiceDate, dDueDate As Date
                            Date.TryParseExact(sInvoiceDate_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dInvoiceDate)
                            If sDueDate <> "" Then
                                Dim iIndex1 As Integer = sDueDate.IndexOf(" ")
                                Dim sDueDate_Trimed As String
                                If iIndex1 > -1 Then
                                    sDueDate_Trimed = sDueDate.Substring(0, iIndex1)
                                Else
                                    sDueDate_Trimed = sDueDate
                                End If
                                Date.TryParseExact(sDueDate_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDueDate)
                            End If
                            'Dim selected As System.Data.DataTable = oDv.ToTable("Selected", False, "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10")
                            'Dim stringArr = selected.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            'sAddress = String.Join(",", stringArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            Dim sAddr1, sAddr2, sAddr3 As String
                            'First line
                            Dim sFirstLine As System.Data.DataTable = oDv.ToTable("FirstLine", False, "F3", "F4", "F5", "F6")
                            Dim sFirstLineArr = sFirstLine.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            sAddr1 = String.Join("," & Environment.NewLine, sFirstLineArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            'Second line
                            Dim sSecondLine As System.Data.DataTable = oDv.ToTable("SecondLine", False, "F7", "F8")
                            Dim sSecondLineArr = sSecondLine.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            sAddr2 = String.Join(",", sSecondLineArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            'Third line
                            Dim sThirdLine As System.Data.DataTable = oDv.ToTable("ThirdLine", False, "F9", "F10")
                            Dim sThirdLineArr = sThirdLine.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            sAddr3 = String.Join(",", sThirdLineArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            If sAddr1 <> "" Then
                                sAddress = sAddr1
                            End If
                            If sAddr2 <> "" Then
                                If sAddress = "" Then
                                    sAddress = sAddr2
                                Else
                                    sAddress = sAddress & "," & Environment.NewLine & sAddr2
                                End If
                            End If
                            If sAddr3 <> "" Then
                                If sAddress = "" Then
                                    sAddress = sAddr3
                                Else
                                    sAddress = sAddress & "," & Environment.NewLine & sAddr3
                                End If
                            End If

                            If sInvoiceNo.Length > 100 Then
                                sInvoiceNo = sInvoiceNo.Substring(0, 100)
                            End If
                            If sAddress.Length > 254 Then
                                sAddress = sAddress.Substring(0, 254)
                            End If

                            Dim oARCreditNote As SAPbobsCOM.Documents
                            oARCreditNote = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)

                            oARCreditNote.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

                            oARCreditNote.CardCode = sCardCode
                            oARCreditNote.NumAtCard = sInvoiceNo
                            oARCreditNote.DocDate = dInvoiceDate
                            If sDueDate <> "" Then
                                oARCreditNote.DocDueDate = dDueDate
                            End If
                            If sReference.Length > 254 Then
                                oARCreditNote.Comments = sReference.Substring(0, 254)
                            Else
                                oARCreditNote.Comments = sReference
                            End If
                            If sReference.Length > 50 Then
                                oARCreditNote.JournalMemo = sReference.Substring(0, 50)
                            Else
                                oARCreditNote.JournalMemo = sReference
                            End If
                            oARCreditNote.Address = sAddress
                            oARCreditNote.DocCurrency = sCurrency
                            Dim sSent As String = String.Empty
                            sSent = oDv(0)(34).ToString.Trim.ToUpper()
                            If sSent = "SENT" Then
                                oARCreditNote.UserFields.Fields.Item("U_AI_EINVOICE").Value = "Yes"
                            Else
                                oARCreditNote.UserFields.Fields.Item("U_AI_EINVOICE").Value = "No"
                            End If
                            oARCreditNote.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

                            For j As Integer = 0 To oDv.Count - 1
                                If CDbl(oDv(j)(24).ToString.Trim) < 0 Then
                                    dXcelTotal = dXcelTotal + (-1 * CDbl(oDv(j)(24).ToString.Trim))
                                Else
                                    dXcelTotal = dXcelTotal + (CDbl(oDv(j)(24).ToString.Trim))
                                End If
                            Next

                            Dim iLineCount As Integer = 1

                            '***************ASSIGNING VALUES TO INVOICE LINES*************
                            Dim iLine As Integer = 0
                            sSQL = "SELECT ""Dscription"",""OpenSum"" AS ""LineTotal"",""LineNum"" FROM ""INV1"" WHERE ""DocEntry"" = '" & sInvDocEntry & "' " & _
                                   " AND ""LineStatus"" = 'O' ORDER BY ""LineNum"" "
                            oRs.DoQuery(sSQL)
                            If Not (oRs.BoF And oRs.EoF) Then
                                oRs.MoveFirst()
                                Do Until oRs.EoF
                                    If dXcelTotal > 0.0 And oRs.Fields.Item("LineTotal").Value > 0.0 Then
                                        If CDbl(oDv(iLine)(24).ToString.Trim) = 0 And (oDv(iLine)(25).ToString.Trim = String.Empty) Then
                                            oDv.RowFilter = "ISNULL(F26,0) <> 0 "
                                            sAcctCode = oDv(0)(25).ToString.Trim
                                            oDv.RowFilter = Nothing
                                        Else
                                            sAcctCode = oDv(iLine)(25).ToString.Trim
                                        End If
                                        dtGlAccount.DefaultView.RowFilter = "Code = '" & sAcctCode.ToUpper() & "'"
                                        If dtGlAccount.DefaultView.Count = 0 Then
                                            sErrDesc = "Mapping of account " & sAcctCode & " cannot be found in SAP"
                                            Console.WriteLine(sErrDesc)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                            Throw New ArgumentException(sErrDesc)
                                        Else
                                            sAcctCode = dtGlAccount.DefaultView.Item(0)(0).ToString().Trim()
                                        End If

                                        If oDv(iLine)(26).ToString.Trim = String.Empty Then
                                            oDv.RowFilter = "ISNULL(F27,'') <> '' "
                                            sTaxType = oDv(0)(26).ToString.Trim
                                            oDv.RowFilter = Nothing
                                        Else
                                            sTaxType = oDv(iLine)(26).ToString.Trim
                                        End If
                                        dtTaxType.DefaultView.RowFilter = "Code = '" & sTaxType.ToUpper() & "' "
                                        If dtTaxType.DefaultView.Count = 0 Then
                                            sErrDesc = "Tax code for taxtype " & sTaxType & " not found in Tax type mapping table"
                                            Console.WriteLine(sErrDesc)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                            Throw New ArgumentException(sErrDesc)
                                        Else
                                            sTaxType = dtTaxType.DefaultView.Item(0)(0).ToString().Trim()
                                        End If

                                        sTrackOpt1 = oDv(iLine)(29).ToString.Trim
                                        If sTrackOpt1 <> "" Then
                                            If sTrackOpt1.Length > 8 Then
                                                sTrackOpt1 = sTrackOpt1.Substring(0, 8)
                                            End If
                                            dtOcrCode5.DefaultView.RowFilter = "PrcCodeUpperCase = '" & sTrackOpt1.ToUpper() & "' "
                                            If dtOcrCode5.DefaultView.Count = 0 Then
                                                sErrDesc = "Cost center " & sTrackOpt1 & " cannot be found in SAP "
                                                Console.WriteLine(sErrDesc)
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                                Throw New ArgumentException(sErrDesc)
                                            Else
                                                sTrackOpt1 = dtOcrCode5.DefaultView.Item(0)(0).ToString().Trim()
                                            End If
                                        End If

                                        sProject = oDv(iLine)(31).ToString.Trim
                                        If sProject <> "" Then
                                            If sProject.Length > 100 Then
                                                sProject = sProject.Substring(0, 100)
                                            End If
                                            dtProject.DefaultView.RowFilter = "PrjName = '" & sProject.ToUpper() & "' "
                                            If dtProject.DefaultView.Count = 0 Then
                                                sErrDesc = "Project " & sProject & " cannot be found in SAP"
                                                Console.WriteLine(sErrDesc)
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                                Throw New ArgumentException(sErrDesc)
                                            Else
                                                sProject = dtProject.DefaultView.Item(0)(0).ToString().Trim()
                                            End If
                                        End If

                                        Dim sItemDesc As String = String.Empty
                                        sItemDesc = oDv(iLine)(20).ToString.Trim()

                                        If iLineCount > 1 Then
                                            oARCreditNote.Lines.Add()
                                        End If
                                        'oARCreditNote.Lines.ItemDescription = sItemDesc
                                        If sItemDesc.Length > 100 Then
                                            oARCreditNote.Lines.ItemDescription = sItemDesc.Substring(0, 100)
                                        Else
                                            oARCreditNote.Lines.ItemDescription = sItemDesc
                                        End If
                                        oARCreditNote.Lines.UserFields.Fields.Item("U_Dscription2").Value = oDv(iLine)(20).ToString.Trim()
                                        oARCreditNote.Lines.UserFields.Fields.Item("U_AI_QTY").Value = CDbl(oDv(iLine)(21).ToString.Trim)
                                        oARCreditNote.Lines.UserFields.Fields.Item("U_AI_PRICE").Value = CDbl(oDv(iLine)(22).ToString.Trim)
                                        oARCreditNote.Lines.AccountCode = sAcctCode
                                        oARCreditNote.Lines.VatGroup = sTaxType
                                        Dim dInvLineAmt As Double = 0.0
                                        dInvLineAmt = oRs.Fields.Item("LineTotal").Value
                                        If dInvLineAmt = 0.0 Then
                                            oARCreditNote.Lines.LineTotal = 0.0
                                        ElseIf dInvLineAmt < dXcelTotal Then
                                            oARCreditNote.Lines.LineTotal = dInvLineAmt
                                            dXcelTotal = dXcelTotal - dInvLineAmt
                                        ElseIf dInvLineAmt > dXcelTotal Then
                                            oARCreditNote.Lines.LineTotal = dXcelTotal
                                            dXcelTotal = 0.0
                                        ElseIf dInvLineAmt = dXcelTotal Then
                                            oARCreditNote.Lines.LineTotal = dInvLineAmt
                                            dXcelTotal = 0.0
                                        End If

                                        If sTrackOpt1 <> "" Then
                                            oARCreditNote.Lines.CostingCode5 = sTrackOpt1
                                            oARCreditNote.Lines.COGSCostingCode5 = sTrackOpt1
                                        End If
                                        If sProject <> "" Then
                                            oARCreditNote.Lines.ProjectCode = sProject
                                        End If

                                        oARCreditNote.Lines.BaseType = "13"
                                        oARCreditNote.Lines.BaseEntry = sInvDocEntry
                                        oARCreditNote.Lines.BaseLine = oRs.Fields.Item("LineNum").Value

                                        iLineCount = iLineCount + 1
                                        iLine = iLine + 1

                                        bIsLineAdded = True
                                    End If
                                    oRs.MoveNext()
                                Loop
                            End If

                            '***************POPULATING EXTRA LINES FROM EXCEL TO CN*************
                            If oRs.RecordCount < oDv.Count Then
                                For j As Integer = iLine To oDv.Count - 1
                                    If CDbl(oDv(j)(24).ToString.Trim) = 0 And (oDv(j)(25).ToString.Trim = String.Empty) Then
                                        oDv.RowFilter = "ISNULL(F26,0) <> 0 "
                                        sAcctCode = oDv(0)(25).ToString.Trim
                                        oDv.RowFilter = Nothing
                                    Else
                                        sAcctCode = oDv(j)(25).ToString.Trim
                                    End If
                                    dtGlAccount.DefaultView.RowFilter = "Code = '" & sAcctCode.ToUpper() & "'"
                                    If dtGlAccount.DefaultView.Count = 0 Then
                                        sErrDesc = "Mapping of account " & sAcctCode & " cannot be found in SAP"
                                        Console.WriteLine(sErrDesc)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        sAcctCode = dtGlAccount.DefaultView.Item(0)(0).ToString().Trim()
                                    End If

                                    If oDv(j)(26).ToString.Trim = String.Empty Then
                                        oDv.RowFilter = "ISNULL(F27,'') <> '' "
                                        sTaxType = oDv(0)(26).ToString.Trim
                                        oDv.RowFilter = Nothing
                                    Else
                                        sTaxType = oDv(j)(26).ToString.Trim
                                    End If
                                    dtTaxType.DefaultView.RowFilter = "Code = '" & sTaxType.ToUpper() & "' "
                                    If dtTaxType.DefaultView.Count = 0 Then
                                        sErrDesc = "Tax code for taxtype " & sTaxType & " not found in Tax type mapping table"
                                        Console.WriteLine(sErrDesc)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        sTaxType = dtTaxType.DefaultView.Item(0)(0).ToString().Trim()
                                    End If

                                    sTrackOpt1 = oDv(j)(29).ToString.Trim
                                    If sTrackOpt1 <> "" Then
                                        If sTrackOpt1.Length > 8 Then
                                            sTrackOpt1 = sTrackOpt1.Substring(0, 8)
                                        End If
                                        dtOcrCode5.DefaultView.RowFilter = "PrcCodeUpperCase = '" & sTrackOpt1.ToUpper() & "' "
                                        If dtOcrCode5.DefaultView.Count = 0 Then
                                            sErrDesc = "Cost center " & sTrackOpt1 & " cannot be found in SAP "
                                            Console.WriteLine(sErrDesc)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                            Throw New ArgumentException(sErrDesc)
                                        Else
                                            sTrackOpt1 = dtOcrCode5.DefaultView.Item(0)(0).ToString().Trim()
                                        End If
                                    End If

                                    sProject = oDv(j)(31).ToString.Trim
                                    If sProject <> "" Then
                                        If sProject.Length > 100 Then
                                            sProject = sProject.Substring(0, 100)
                                        End If
                                        dtProject.DefaultView.RowFilter = "PrjName = '" & sProject.ToUpper() & "' "
                                        If dtProject.DefaultView.Count = 0 Then
                                            sErrDesc = "Project " & sProject & " cannot be found in SAP"
                                            Console.WriteLine(sErrDesc)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                            Throw New ArgumentException(sErrDesc)
                                        Else
                                            sProject = dtProject.DefaultView.Item(0)(0).ToString().Trim()
                                        End If
                                    End If

                                    Dim sItemDesc As String = String.Empty
                                    sItemDesc = oDv(j)(20).ToString.Trim()

                                    If iLineCount > 1 Then
                                        oARCreditNote.Lines.Add()
                                    End If
                                    oARCreditNote.Lines.ItemDescription = sItemDesc
                                    oARCreditNote.Lines.UserFields.Fields.Item("U_AI_QTY").Value = CDbl(oDv(j)(21).ToString.Trim)
                                    oARCreditNote.Lines.UserFields.Fields.Item("U_AI_PRICE").Value = CDbl(oDv(j)(22).ToString.Trim)
                                    oARCreditNote.Lines.AccountCode = sAcctCode
                                    oARCreditNote.Lines.VatGroup = sTaxType
                                    If CDbl(oDv(j)(24).ToString.Trim) < 0 Then
                                        oARCreditNote.Lines.LineTotal = -1 * CDbl(oDv(j)(24).ToString.Trim)
                                    Else
                                        oARCreditNote.Lines.LineTotal = CDbl(oDv(j)(24).ToString.Trim)
                                    End If
                                    If sTrackOpt1 <> "" Then
                                        oARCreditNote.Lines.CostingCode5 = sTrackOpt1
                                        oARCreditNote.Lines.COGSCostingCode5 = sTrackOpt1
                                    End If
                                    If sProject <> "" Then
                                        oARCreditNote.Lines.ProjectCode = sProject
                                    End If
                                    iLineCount = iLineCount + 1
                                    bIsLineAdded = True
                                Next

                            End If
                            '***************POPULATING THE EXTRA VALUE*************
                            'If dXcelTotal > 0.0 Then
                            '    If CDbl(oDv(0)(24).ToString.Trim) = 0 And (oDv(0)(25).ToString.Trim = String.Empty) Then
                            '        oDv.RowFilter = "ISNULL(F26,0) <> 0 "
                            '        sAcctCode = oDv(0)(25).ToString.Trim
                            '        oDv.RowFilter = Nothing
                            '    Else
                            '        sAcctCode = oDv(0)(25).ToString.Trim
                            '    End If
                            '    dtGlAccount.DefaultView.RowFilter = "Code = '" & sAcctCode.ToUpper() & "'"
                            '    If dtGlAccount.DefaultView.Count = 0 Then
                            '        sErrDesc = "Mapping of account " & sAcctCode & " cannot be found in SAP"
                            '        Console.WriteLine(sErrDesc)
                            '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                            '        Throw New ArgumentException(sErrDesc)
                            '    Else
                            '        sAcctCode = dtGlAccount.DefaultView.Item(0)(0).ToString().Trim()
                            '    End If

                            '    If oDv(0)(26).ToString.Trim = String.Empty Then
                            '        oDv.RowFilter = "ISNULL(F27,'') <> '' "
                            '        sTaxType = oDv(0)(26).ToString.Trim
                            '        oDv.RowFilter = Nothing
                            '    Else
                            '        sTaxType = oDv(0)(26).ToString.Trim
                            '    End If
                            '    dtTaxType.DefaultView.RowFilter = "Code = '" & sTaxType.ToUpper() & "' "
                            '    If dtTaxType.DefaultView.Count = 0 Then
                            '        sErrDesc = "Tax code for taxtype " & sTaxType & " not found in Tax type mapping table"
                            '        Console.WriteLine(sErrDesc)
                            '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                            '        Throw New ArgumentException(sErrDesc)
                            '    Else
                            '        sTaxType = dtTaxType.DefaultView.Item(0)(0).ToString().Trim()
                            '    End If

                            '    sTrackOpt1 = oDv(0)(29).ToString.Trim
                            '    If sTrackOpt1 <> "" Then
                            '        If sTrackOpt1.Length > 8 Then
                            '            sTrackOpt1 = sTrackOpt1.Substring(0, 8)
                            '        End If
                            '        dtOcrCode5.DefaultView.RowFilter = "PrcCodeUpperCase = '" & sTrackOpt1.ToUpper() & "' "
                            '        If dtOcrCode5.DefaultView.Count = 0 Then
                            '            sErrDesc = "Cost center " & sTrackOpt1 & " cannot be found in SAP "
                            '            Console.WriteLine(sErrDesc)
                            '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                            '            Throw New ArgumentException(sErrDesc)
                            '        Else
                            '            sTrackOpt1 = dtOcrCode5.DefaultView.Item(0)(0).ToString().Trim()
                            '        End If
                            '    End If

                            '    sProject = oDv(0)(31).ToString.Trim
                            '    If sProject <> "" Then
                            '        If sProject.Length > 100 Then
                            '            sProject = sProject.Substring(0, 100)
                            '        End If
                            '        dtProject.DefaultView.RowFilter = "PrjName = '" & sProject.ToUpper() & "' "
                            '        If dtProject.DefaultView.Count = 0 Then
                            '            sErrDesc = "Project " & sProject & " cannot be found in SAP"
                            '            Console.WriteLine(sErrDesc)
                            '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                            '            Throw New ArgumentException(sErrDesc)
                            '        Else
                            '            sProject = dtProject.DefaultView.Item(0)(0).ToString().Trim()
                            '        End If
                            '    End If

                            '    If iLineCount > 1 Then
                            '        oARCreditNote.Lines.Add()
                            '    End If
                            '    oARCreditNote.Lines.ItemDescription = "Rounding Off"
                            '    oARCreditNote.Lines.UserFields.Fields.Item("U_AI_QTY").Value = CDbl(oDv(0)(21).ToString.Trim)
                            '    oARCreditNote.Lines.UserFields.Fields.Item("U_AI_PRICE").Value = CDbl(oDv(0)(22).ToString.Trim)
                            '    oARCreditNote.Lines.AccountCode = sAcctCode
                            '    oARCreditNote.Lines.VatGroup = sTaxType
                            '    oARCreditNote.Lines.LineTotal = dXcelTotal
                            '    If sTrackOpt1 <> "" Then
                            '        oARCreditNote.Lines.CostingCode5 = sTrackOpt1
                            '        oARCreditNote.Lines.COGSCostingCode5 = sTrackOpt1
                            '    End If
                            '    If sProject <> "" Then
                            '        oARCreditNote.Lines.ProjectCode = sProject
                            '    End If
                            'End If

                            If bIsLineAdded = True Then
                                If oARCreditNote.Add() <> 0 Then
                                    sErrDesc = "Error while adding Invoice / " & p_oCompany.GetLastErrorDescription
                                    Console.WriteLine(sErrDesc)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    Dim iDocNo, iDocEntry As Integer
                                    p_oCompany.GetNewObjectCode(iDocEntry)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oARCreditNote)

                                    Dim objRS As SAPbobsCOM.Recordset
                                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    sQuery = "SELECT ""DocNum"" FROM ""ORIN"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                                    objRS.DoQuery(sQuery)
                                    If objRS.RecordCount > 0 Then
                                        iDocNo = objRS.Fields.Item("DocNum").Value
                                    End If
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Credit Note document created successfully :: " & iDocNo, sFuncName)
                                    Console.WriteLine("Credit Note document created successfully :: " & iDocNo)

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                                End If
                            End If

                        End If
                    End If
                End If
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateSalesCreditNote = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateSalesCreditNote = RTN_ERROR
        End Try
    End Function

    Private Function CreateSalesCreditNote_Backup(ByVal oDv_CreditNote As DataView, ByVal sFileName As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateSalesCreditNote"
        Dim sCardName As String = String.Empty
        Dim sContactName As String = String.Empty
        Dim sAcctCode As String = String.Empty
        Dim sTaxType As String = String.Empty
        Dim sTrackOpt1 As String = String.Empty
        Dim sProject As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sInvoiceNo As String = String.Empty
        Dim sReference As String = String.Empty
        Dim sAddress As String = String.Empty
        Dim sInvoiceDate As String = String.Empty
        Dim sDueDate As String = String.Empty
        Dim sCurrency As String = String.Empty
        Dim bIsLineAdded As Boolean = False
        Dim sBaseInvoice As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            Console.WriteLine("Creating A/R Credit Note")

            oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim oDtGroup As DataTable
            oDtGroup = oDv_CreditNote.Table.DefaultView.ToTable(True, "F11") 'F11 - Invoice Number, F37 - Base Document

            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = "" Or oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim = "INVOICENUMBER") Then
                    oDv_CreditNote.RowFilter = "F11 = '" & oDtGroup.Rows(i).Item(0).ToString().ToUpper & "' "
                    If oDv_CreditNote.Count > 0 Then
                        Dim oDt As DataTable
                        oDt = oDv_CreditNote.ToTable
                        Dim oDv As DataView = New DataView(oDt)
                        If oDv.Count > 0 Then
                            oDv.RowFilter = "ISNULL(F37,'') <> ''"
                            If oDv.Count > 0 Then
                            Else
                                sErrDesc = "Base Invoice not found for creating credit note/Please check in BaseInvoice column in Excel"
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    End If
                End If
            Next


            oDv_CreditNote.RowFilter = Nothing

            oDtGroup = oDv_CreditNote.Table.DefaultView.ToTable(True, "F11") 'F11 - Invoice Number, F37 - Base Document

            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = "" Or oDtGroup.Rows(i).Item(0).ToString.ToUpper.Trim = "INVOICENUMBER") Then
                    Console.WriteLine("Filtering rows by invoice number ")

                    oDv_CreditNote.RowFilter = "F11 = '" & oDtGroup.Rows(i).Item(0).ToString().ToUpper & "' "

                    If oDv_CreditNote.Count > 0 Then
                        Dim oDt As New DataTable
                        oDt = oDv_CreditNote.ToTable
                        Dim oDv As DataView = New DataView(oDt)

                        If oDv.Count > 0 Then
                            oDv.RowFilter = "ISNULL(F37,'') <> ''"
                            If oDv.Count > 0 Then
                                sBaseInvoice = oDv(0)(36).ToString.Trim()
                            End If
                            oDv.RowFilter = Nothing

                            Dim sInvDocEntry As String = String.Empty
                            sInvoiceNo = oDv(0)(10).ToString.Trim()
                            If sBaseInvoice.Length > 100 Then
                                sBaseInvoice = sBaseInvoice.Substring(0, 100)
                            End If

                            Dim sSQL As String
                            sSQL = "SELECT ""DocEntry"" FROM ""ORIN"" WHERE ""NumAtCard"" = '" & sInvoiceNo & "' "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                            ' oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs.DoQuery(sSQL)
                            If oRs.RecordCount > 0 Then
                                sErrDesc = "Credit Note already found for reference number :: " & sInvoiceNo
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If

                            Dim sQuery As String = String.Empty
                            sQuery = "SELECT ""DocEntry"",""DocNum"" FROM ""OINV"" WHERE ""DocStatus"" = 'O' AND ""CANCELED"" = 'N' AND UPPER(""NumAtCard"") = '" & sBaseInvoice.ToUpper() & "' "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                            oRs.DoQuery(sQuery)
                            If oRs.RecordCount > 0 Then
                                sInvDocEntry = oRs.Fields.Item("DocEntry").Value
                                Console.WriteLine("Invoice Found/DocEntry is " & sInvDocEntry)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice DocEntry is " & sInvDocEntry, sFuncName)
                            Else
                                sErrDesc = "Base Invoice Number " & sBaseInvoice & " Not found in SAP"
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If

                            sReference = oDv(0)(11).ToString.Trim()
                            sInvoiceDate = oDv(0)(12).ToString.Trim
                            sDueDate = oDv(0)(13).ToString.Trim
                            sCurrency = oDv(0)(32).ToString.Trim()

                            sContactName = oDv(0)(0).ToString.Trim
                            If sContactName.Length > 100 Then
                                sContactName = sContactName.Substring(0, 100)
                            End If
                            dtCardCode.DefaultView.RowFilter = "CardName = '" & sContactName.ToUpper & "'"
                            If dtCardCode.DefaultView.Count = 0 Then
                                sErrDesc = "Customer record not found in SAP/CardName :: " & sContactName
                                Console.WriteLine(sErrDesc)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sCardCode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If

                            Dim iIndex As Integer = sInvoiceDate.IndexOf(" ")
                            Dim sInvoiceDate_Trimed As String = sInvoiceDate.Substring(0, iIndex)
                            Dim iIndex1 As Integer = sInvoiceDate.IndexOf(" ")
                            Dim sDueDate_Trimed As String = sDueDate.Substring(0, iIndex1)
                            Dim dInvoiceDate, dDueDate As Date
                            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd/MM/yy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                            Date.TryParseExact(sInvoiceDate_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dInvoiceDate)
                            Date.TryParseExact(sDueDate_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDueDate)

                            'Dim selected As System.Data.DataTable = oDv.ToTable("Selected", False, "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10")
                            'Dim stringArr = selected.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            'sAddress = String.Join(",", stringArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            Dim sAddr1, sAddr2, sAddr3 As String
                            'First line
                            Dim sFirstLine As System.Data.DataTable = oDv.ToTable("FirstLine", False, "F3", "F4", "F5", "F6")
                            Dim sFirstLineArr = sFirstLine.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            sAddr1 = String.Join("," & Environment.NewLine, sFirstLineArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            'Second line
                            Dim sSecondLine As System.Data.DataTable = oDv.ToTable("SecondLine", False, "F7", "F8")
                            Dim sSecondLineArr = sSecondLine.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            sAddr2 = String.Join(",", sSecondLineArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            'Third line
                            Dim sThirdLine As System.Data.DataTable = oDv.ToTable("ThirdLine", False, "F9", "F10")
                            Dim sThirdLineArr = sThirdLine.Rows(0).ItemArray.[Select](Function(x) x.ToString()).ToArray()
                            sAddr3 = String.Join(",", sThirdLineArr.Where(Function(s) Not String.IsNullOrEmpty(s)))

                            If sAddr1 <> "" Then
                                sAddress = sAddr1
                            End If
                            If sAddr2 <> "" Then
                                If sAddress = "" Then
                                    sAddress = sAddr2
                                Else
                                    sAddress = sAddress & "," & Environment.NewLine & sAddr2
                                End If
                            End If
                            If sAddr3 <> "" Then
                                If sAddress = "" Then
                                    sAddress = sAddr3
                                Else
                                    sAddress = sAddress & "," & Environment.NewLine & sAddr3
                                End If
                            End If

                            If sInvoiceNo.Length > 100 Then
                                sInvoiceNo = sInvoiceNo.Substring(0, 100)
                            End If
                            If sAddress.Length > 254 Then
                                sAddress = sAddress.Substring(0, 254)
                            End If

                            Dim oARCreditNote As SAPbobsCOM.Documents
                            oARCreditNote = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)

                            oARCreditNote.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

                            oARCreditNote.CardCode = sCardCode
                            oARCreditNote.NumAtCard = sInvoiceNo
                            oARCreditNote.DocDate = dInvoiceDate
                            oARCreditNote.DocDueDate = dDueDate
                            If sReference.Length > 254 Then
                                oARCreditNote.Comments = sReference.Substring(0, 254)
                            Else
                                oARCreditNote.Comments = sReference
                            End If
                            If sReference.Length > 50 Then
                                oARCreditNote.JournalMemo = sReference.Substring(0, 50)
                            Else
                                oARCreditNote.JournalMemo = sReference
                            End If
                            oARCreditNote.Address = sAddress
                            oARCreditNote.DocCurrency = sCurrency
                            Dim sSent As String = String.Empty
                            sSent = oDv(0)(34).ToString.Trim.ToUpper()
                            If sSent = "SENT" Then
                                oARCreditNote.UserFields.Fields.Item("U_AI_EINVOICE").Value = "Yes"
                            Else
                                oARCreditNote.UserFields.Fields.Item("U_AI_EINVOICE").Value = "No"
                            End If
                            oARCreditNote.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

                            Dim iLineCount As Integer = 1

                            For j As Integer = 0 To oDv.Count - 1
                                If CDbl(oDv(j)(24).ToString.Trim) = 0 And (oDv(j)(25).ToString.Trim = String.Empty) Then
                                    oDv.RowFilter = "ISNULL(F26,0) <> 0 "
                                    sAcctCode = oDv(0)(25).ToString.Trim
                                    oDv.RowFilter = Nothing
                                Else
                                    sAcctCode = oDv(j)(25).ToString.Trim
                                End If
                                dtGlAccount.DefaultView.RowFilter = "Code = '" & sAcctCode.ToUpper() & "'"
                                If dtGlAccount.DefaultView.Count = 0 Then
                                    sErrDesc = "Mapping of account " & sAcctCode & " cannot be found in SAP"
                                    Console.WriteLine(sErrDesc)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sAcctCode = dtGlAccount.DefaultView.Item(0)(0).ToString().Trim()
                                End If

                                If oDv(j)(26).ToString.Trim = String.Empty Then
                                    oDv.RowFilter = "ISNULL(F27,'') <> '' "
                                    sTaxType = oDv(0)(26).ToString.Trim
                                    oDv.RowFilter = Nothing
                                Else
                                    sTaxType = oDv(j)(26).ToString.Trim
                                End If
                                dtTaxType.DefaultView.RowFilter = "Code = '" & sTaxType.ToUpper() & "' "
                                If dtTaxType.DefaultView.Count = 0 Then
                                    sErrDesc = "Tax code for taxtype " & sTaxType & " not found in Tax type mapping table"
                                    Console.WriteLine(sErrDesc)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sTaxType = dtTaxType.DefaultView.Item(0)(0).ToString().Trim()
                                End If

                                sTrackOpt1 = oDv(j)(29).ToString.Trim
                                If sTrackOpt1 <> "" Then
                                    If sTrackOpt1.Length > 8 Then
                                        sTrackOpt1 = sTrackOpt1.Substring(0, 8)
                                    End If
                                    dtOcrCode5.DefaultView.RowFilter = "PrcCodeUpperCase = '" & sTrackOpt1.ToUpper() & "' "
                                    If dtOcrCode5.DefaultView.Count = 0 Then
                                        sErrDesc = "Cost center " & sTrackOpt1 & " cannot be found in SAP "
                                        Console.WriteLine(sErrDesc)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        sTrackOpt1 = dtOcrCode5.DefaultView.Item(0)(0).ToString().Trim()
                                    End If
                                End If

                                sProject = oDv(j)(31).ToString.Trim
                                If sProject <> "" Then
                                    If sProject.Length > 100 Then
                                        sProject = sProject.Substring(0, 100)
                                    End If
                                    dtProject.DefaultView.RowFilter = "PrjName = '" & sProject.ToUpper() & "' "
                                    If dtProject.DefaultView.Count = 0 Then
                                        sErrDesc = "Project " & sProject & " cannot be found in SAP"
                                        Console.WriteLine(sErrDesc)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        sProject = dtProject.DefaultView.Item(0)(0).ToString().Trim()
                                    End If
                                End If

                                Dim sItemDesc As String = String.Empty
                                sItemDesc = oDv(j)(20).ToString.Trim()

                                If iLineCount > 1 Then
                                    oARCreditNote.Lines.Add()
                                End If
                                'oARCreditNote.Lines.ItemDescription = sItemDesc
                                If sItemDesc.Length > 100 Then
                                    oARCreditNote.Lines.ItemDescription = sItemDesc.Substring(0, 100)
                                Else
                                    oARCreditNote.Lines.ItemDescription = sItemDesc
                                End If
                                oARCreditNote.Lines.UserFields.Fields.Item("U_Dscription2").Value = oDv(j)(20).ToString.Trim()
                                oARCreditNote.Lines.UserFields.Fields.Item("U_AI_QTY").Value = CDbl(oDv(j)(21).ToString.Trim)
                                oARCreditNote.Lines.UserFields.Fields.Item("U_AI_PRICE").Value = CDbl(oDv(j)(22).ToString.Trim)
                                oARCreditNote.Lines.AccountCode = sAcctCode
                                oARCreditNote.Lines.VatGroup = sTaxType
                                If CDbl(oDv(j)(24).ToString.Trim) < 0 Then
                                    oARCreditNote.Lines.LineTotal = -1 * CDbl(oDv(j)(24).ToString.Trim)
                                Else
                                    oARCreditNote.Lines.LineTotal = CDbl(oDv(j)(24).ToString.Trim)
                                End If
                                If sTrackOpt1 <> "" Then
                                    oARCreditNote.Lines.CostingCode5 = sTrackOpt1
                                    oARCreditNote.Lines.COGSCostingCode5 = sTrackOpt1
                                End If
                                If sProject <> "" Then
                                    oARCreditNote.Lines.ProjectCode = sProject
                                End If

                                oARCreditNote.Lines.BaseType = "13"
                                oARCreditNote.Lines.BaseEntry = sInvDocEntry

                                sSQL = "SELECT ""LineNum"" FROM ""INV1"" WHERE ""DocEntry"" = '" & sInvDocEntry & "' AND ""Dscription"" = '" & sItemDesc & "' AND ""AcctCode"" = '" & sAcctCode & "' "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing sql " & sSQL, sFuncName)
                                oRs.DoQuery(sSQL)
                                If oRs.RecordCount > 0 Then
                                    oARCreditNote.Lines.BaseLine = oRs.Fields.Item("LineNum").Value
                                End If

                                iLineCount = iLineCount + 1
                                bIsLineAdded = True
                            Next
                            If bIsLineAdded = True Then
                                If oARCreditNote.Add() <> 0 Then
                                    sErrDesc = "Error while adding Invoice / " & p_oCompany.GetLastErrorDescription
                                    Console.WriteLine(sErrDesc)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    Dim iDocNo, iDocEntry As Integer
                                    p_oCompany.GetNewObjectCode(iDocEntry)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oARCreditNote)

                                    Dim objRS As SAPbobsCOM.Recordset
                                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    sQuery = "SELECT ""DocNum"" FROM ""ORIN"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                                    objRS.DoQuery(sQuery)
                                    If objRS.RecordCount > 0 Then
                                        iDocNo = objRS.Fields.Item("DocNum").Value
                                    End If
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Credit Note document created successfully :: " & iDocNo, sFuncName)
                                    Console.WriteLine("Credit Note document created successfully :: " & iDocNo)

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                                End If
                            End If

                        End If
                    End If
                End If
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateSalesCreditNote_Backup = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateSalesCreditNote_Backup = RTN_ERROR
        End Try
    End Function

End Module
