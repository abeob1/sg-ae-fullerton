Module modBillToClient

    Public Function Upload_BillToClient_FHG(ByVal sContractOwner As String, ByVal oDv As DataView, ByVal bServiceItem As Boolean, ByVal bIsNonPanel As Boolean, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim oDS As New DataSet
        Dim oDatasetBP As New DataSet
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim k, m As Integer
        Dim oBPArrary As New ArrayList
        Dim sBillType As String = String.Empty
        Dim oBPDT As DataTable
        Dim oBTDT As DataTable
        Dim oBPInhouseDT As DataTable
        Dim sBatchPeriodName As String = String.Empty

        Try
            sFuncName = "Upload_BillToClient_FHG"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Worksheet data.", sFuncName)
            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sBatchNo = oDv(2)(0).ToString()
            sBatchPeriod = oDv(3)(0).ToString()
            sBatchPeriodName = oDv(3)(0).ToString

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch No.", sFuncName)
            k = InStrRev(sBatchNo, ":")
            sBatchNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - k).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch No:" & sBatchNo, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch Date.", sFuncName)
            m = InStrRev(sBatchPeriod, "to")
            sBatchPeriod = Microsoft.VisualBasic.Right(sBatchPeriod, Len(sBatchPeriod) - m - 1).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Date: " & sBatchPeriod, sFuncName)

            For i As Integer = 6 To oDv.Count - 1
                If Not oBPArrary.Contains(oDv(i)(6).ToString) Then
                    If Not oDv(i)(6).ToString = String.Empty Then
                        oBPArrary.Add(oDv(i)(6).ToString)
                    End If
                End If
            Next

            oBPDT = oDv.Table.Clone
            oBPDT.Clear()

            For iRow As Integer = 0 To oBPArrary.Count - 1

                sBillType = String.Empty
                '1. Get Billing Type from OCRD
                GetBillingType(oBPArrary(iRow).ToString, sBillType)
                Dim sCName As String = oBPArrary(iRow).ToString

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting BillType for ::" & sCName, sFuncName)

                If sBillType = String.Empty Then Throw New ArgumentException("No Billing Type specified for ::" & sCName)

                '2. SELECT BILLTYPE
                'Dim BPRows() As DataRow = oDv.Table.Select("F3='" & oBPArrary(iRow).ToString & "'")

                Dim BPRows() As DataRow = oDv.Table.Select("F7='" & Replace(oBPArrary(iRow).ToString, "'", "''") & "'")

                'UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sCardName), "'", "''") & "'"

                oBPDT.Clear()
                For Each row As DataRow In BPRows
                    oBPDT.ImportRow(row)
                Next

                'Create A/R Invoice
                Select Case sBillType

                    Case "CO"
                        Console.WriteLine("Creating Document")
                        sCardName = oBPArrary(iRow).ToString
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                        If AddARInvoice_FHG(sContractOwner, sCardName, BPRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrDesc)
                        End If

                    Case "ENT"
                        Dim oEnTArray As New ArrayList
                        For Each row As DataRow In BPRows
                            If Not oEnTArray.Contains(row.Item(11).ToString.Trim) Then oEnTArray.Add(row.Item(11).ToString.Trim)
                        Next

                        For iEntRow As Integer = 0 To oEnTArray.Count - 1
                            Console.WriteLine("Creating Document")
                            Dim EntRows() As DataRow = oBPDT.Select("F12='" & Replace(oEnTArray(iEntRow).ToString.Trim, "'", "''") & "'")
                            sCardName = oEnTArray(iEntRow).ToString

                            If EntRows.Length = 0 Then
                                sErrDesc = "Entity is blank for ::" & oBPArrary(iRow).ToString
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                            If AddARInvoice_FHG(sContractOwner, sCardName, EntRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "CO-DEPT"
                        Dim oDepArray As New ArrayList
                        For Each row As DataRow In BPRows
                            If Not oDepArray.Contains(row.Item(12).ToString.Trim) Then oDepArray.Add(row.Item(12).ToString.Trim)
                        Next

                        For iDepRow As Integer = 0 To oDepArray.Count - 1
                            Console.WriteLine("Creating Document")
                            Dim DepRows() As DataRow = oBPDT.Select("F13='" & Replace(oDepArray(iDepRow).ToString.Trim, "'", "''") & "'")
                            sCardName = oBPArrary(iRow).ToString

                            If DepRows.Length = 0 Then
                                sErrDesc = "Department is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If


                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                            If AddARInvoice_FHG(sContractOwner, sCardName, DepRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "ENT-DEPT"
                        Dim oEntDeptDT As DataTable
                        oEntDeptDT = oBPDT.DefaultView.ToTable(True, "F12", "F13")

                        For Each row As DataRow In oEntDeptDT.Rows
                            Console.WriteLine("Creating Document")

                            Dim oEntDeptRows() As DataRow = oBPDT.Select("F12='" & Replace(row.Item(0).ToString.Trim, "'", "''") & "' and F13='" & Replace(row.Item(1).ToString.Trim, "'", "''") & "'")
                            sCardName = row.Item(0).ToString

                            If oEntDeptRows.Length = 0 Then
                                sErrDesc = "Entity/Department is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If


                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                            If AddARInvoice_FHG(sContractOwner, sCardName, oEntDeptRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If

                        Next

                    Case "CO-DEPT-CC"
                        Dim oDeptCCDT As DataTable
                        oDeptCCDT = oBPDT.DefaultView.ToTable(True, "F13", "F14")

                        For Each row As DataRow In oDeptCCDT.Rows
                            Console.WriteLine("Creating Document")
                            Dim oDeptCCRows() As DataRow = oBPDT.Select("F13='" & Replace(row.Item(0).ToString.Trim, "'", "''") & "' and F14='" & Replace(row.Item(1).ToString.Trim, "'", "''") & "'")
                            oBPDT.CaseSensitive = False

                            sCardName = oBPArrary(iRow).ToString

                            If oDeptCCRows.Length = 0 Then
                                sErrDesc = "Department :: " & row.Item(0).ToString.Trim & " CostCenter::" & row.Item(1).ToString.Trim & " is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                            If AddARInvoice_FHG(sContractOwner, sCardName, oDeptCCRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "ENT-DEPT-CC"

                        Dim oENTDEPCCDT As DataTable
                        oENTDEPCCDT = oBPDT.DefaultView.ToTable(True, "F12", "F13", "F14")

                        For Each row As DataRow In oENTDEPCCDT.Rows
                            Console.WriteLine("Creating Document")
                            Dim oENTDEPCCRows() As DataRow = oBPDT.Select("F12='" & Replace(row.Item(0).ToString.Trim, "'", "''") & "' and F13='" & Replace(row.Item(1).ToString.Trim, "'", "''") & "' and F14='" & Replace(row.Item(2).ToString.Trim, "'", "''") & "'")
                            sCardName = row.Item(0).ToString

                            If oENTDEPCCRows.Length = 0 Then
                                sErrDesc = "Entity/Department/CostCenter is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                            If AddARInvoice_FHG(sContractOwner, sCardName, oENTDEPCCRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "CO-CC"
                        Dim oCCArray As New ArrayList
                        For Each row As DataRow In BPRows
                            If Not oCCArray.Contains(row.Item(13).ToString.Trim) Then
                                If Not row.Item(13).ToString = String.Empty Then
                                    oCCArray.Add(row.Item(13).ToString.Trim)
                                End If
                            End If
                        Next
                        For iCCRow As Integer = 0 To oCCArray.Count - 1
                            Console.WriteLine("Creating Document")
                            Dim CCRows() As DataRow = oBPDT.Select("F14='" & Replace(oCCArray(iCCRow).ToString.Trim, "'", "''") & "'")
                            sCardName = oBPArrary(iRow).ToString

                            If CCRows.Length = 0 Then
                                sErrDesc = "Cost Center is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                            If AddARInvoice_FHG(sContractOwner, sCardName, CCRows, sBatchNo, sBatchPeriod, bServiceItem, False, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "ENT-CC"
                        Dim oEntCCDT As DataTable
                        oEntCCDT = oBPDT.DefaultView.ToTable(True, "F12", "F14")

                        For Each row As DataRow In oEntCCDT.Rows
                            Console.WriteLine("Creating Document")
                            Dim oEntCCRows() As DataRow = oBPDT.Select("F12='" & Replace(row.Item(0).ToString.Trim, "'", "''") & "' and F14='" & Replace(row.Item(1).ToString.Trim, "'", "''") & "'")
                            sCardName = row.Item(0).ToString

                            If oEntCCRows.Length = 0 Then
                                sErrDesc = "Entity/Cost center is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                            If AddARInvoice_FHG(sContractOwner, sCardName, oEntCCRows, sBatchNo, sBatchPeriod, bServiceItem, False, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "CO-BEN"
                        Dim oBenArray As New ArrayList
                        For Each row As DataRow In BPRows
                            If Not oBenArray.Contains(row.Item(20).ToString.Trim) Then oBenArray.Add(row.Item(20).ToString.Trim)
                        Next

                        For iBenRow As Integer = 0 To oBenArray.Count - 1

                            Console.WriteLine("Creating Document")
                            Dim BenRows() As DataRow = oBPDT.Select("F21='" & Replace(oBenArray(iBenRow).ToString.Trim, "'", "''") & "'")
                            sCardName = oBPArrary(iRow).ToString

                            If BenRows.Length = 0 Then
                                sErrDesc = "Benefit type is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            Dim oProvider As New ArrayList
                            Dim BenProvRows() As DataRow = Nothing

                            For Each row As DataRow In BenRows
                                If Not oProvider.Contains(UCase(row.Item(20).ToString.Trim)) Then oProvider.Add(UCase(row.Item(20).ToString.Trim)) 'Provider Name
                            Next

                            oBTDT = oBPDT.Clone
                            oBTDT.Clear()
                            oBPInhouseDT = oBPDT.Clone
                            oBPInhouseDT.Clear()

                            For iProvRow As Integer = 0 To oProvider.Count - 1
                                Dim sProviderName As String = oProvider(iProvRow).ToString
                                'Check Inhouse clinic? if Yes further splic into by Provider
                                BenProvRows = oBPDT.Select("F21='" & Replace(oBenArray(iBenRow).ToString.Trim, "'", "''") & "' and F22='" & Replace(sProviderName.Trim, "'", "''") & "'")

                                If IsInhouseClinic(sProviderName) = True Then
                                    For Each row As DataRow In BenProvRows
                                        oBPInhouseDT.ImportRow(row)
                                    Next
                                Else
                                    For Each row As DataRow In BenProvRows
                                        oBTDT.ImportRow(row)
                                    Next
                                End If

                                'If IsInhouseClinic(sProviderName) = True Then
                                '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                                '    If AddARInvoice_FHG(sCardName, BenProvRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                '        Throw New ArgumentException(sErrDesc)
                                '    End If
                                'Else
                                '    For Each row As DataRow In BenProvRows
                                '        oBTDT.ImportRow(row)
                                '    Next
                                'End If

                            Next

                            Dim BenFinalRow() As DataRow = oBTDT.Select("F7='" & Replace(sCardName, "'", "''") & "'")
                            If BenFinalRow.Length > 0 Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                                If AddARInvoice_FHG(sContractOwner, sCardName, BenFinalRow, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                    Throw New ArgumentException(sErrDesc)
                                End If
                            End If

                            Dim BenInhouseRow() As DataRow = oBPInhouseDT.Select("F7='" & Replace(sCardName, "'", "''") & "'")
                            If BenInhouseRow.Length > 0 Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_FHG()", sFuncName)
                                If AddARInvoice_FHG(sContractOwner, sCardName, BenInhouseRow, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                    Throw New ArgumentException(sErrDesc)
                                End If
                            End If
                        Next

                    Case "VISIT NO."
                        Console.WriteLine("Creating Document")
                        sCardName = oBPArrary(iRow).ToString
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_VisitNo()", sFuncName)

                        For i As Integer = 6 To oDv.Count - 1
                            Dim VisitNoRows() As DataRow = oDv.Table.Select("F1='" & oDv(i)(0).ToString & "'")
                            If AddARInvoice_VisitNo(sCardName, VisitNoRows, sBatchNo, sBatchPeriod, sBatchPeriodName, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "VISIT NO-SERVICEFEE"
                        Console.WriteLine("Creating Document")
                        sCardName = oBPArrary(iRow).ToString
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_VisitNo_ServiceFee()", sFuncName)

                        Dim oVisitNoArray As New ArrayList
                        For Each row As DataRow In BPRows
                            If Not oVisitNoArray.Contains(row.Item(0).ToString.Trim) Then oVisitNoArray.Add(row.Item(0).ToString.Trim)
                        Next

                        For iVisRow As Integer = 0 To oVisitNoArray.Count - 1
                            Dim VisitNoRows() As DataRow = oDv.Table.Select("F1='" & oVisitNoArray(iVisRow).ToString.Trim & "'")
                            If AddARInvoice_VisitNo_ServiceFee(sCardName, VisitNoRows, sBatchNo, sBatchPeriod, sBatchPeriodName, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next


                End Select
            Next


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            Upload_BillToClient_FHG = RTN_SUCCESS

        Catch ex As Exception
            Upload_BillToClient_FHG = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        Finally

        End Try
    End Function

    Private Function AddARInvoice_FHG(ByVal sContractOwner As String, _
                                        ByVal sCardName As String, _
                                        ByVal oBPRow() As DataRow, _
                                        ByVal sBatchNo As String, _
                                        ByVal sBatchPeriod As String, _
                                        ByVal bServiceItem As Boolean, _
                                        ByVal bIsNonPanel As Boolean, _
                                        ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim bSO As Boolean = False
        Dim sName1 As String = "MERCER (SINGAPORE) PTE LTD"
        Dim sInvRef As String = String.Empty

        Try
            sFuncName = "AddARInvoice_FHG"

            If CheckBP_FHG(sCardName, sCardCode, "C", sInvRef, sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If

            sCostCenter = GetCostCenter(sCardCode)

            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter

            If CheckCreateSO(sCardCode) = True And bIsNonPanel = True Then
                Console.WriteLine("Creating Sales Order..")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIapi SO Object", sFuncName)
                oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                bSO = True
                oDoc.DocDueDate = CDate(sBatchPeriod)
            Else
                Console.WriteLine("Creating A/R Invoice..")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
                oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            End If

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            'oDoc.DocDueDate = CDate(sBatchPeriod)
            oDoc.TaxDate = CDate(sBatchPeriod)
            oDoc.NumAtCard = "Batch No. : " & sBatchNo
            oDoc.Comments = sBatchPeriod
            oDoc.ImportFileNum = sBatchNo

            For Each row As DataRow In oBPRow
                iCnt += 1
                If iCnt > 1 Then
                    oDoc.Lines.Add()
                End If

                If row.Item(19).ToString = "FFS" Then
                    If bIsNonPanel = True Then
                        oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCodeNonPanel
                    Else
                        oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCode
                    End If
                ElseIf row.Item(19).ToString = "CAP" Then
                    If Not sContractOwner = sName1 Then oDoc.Lines.ItemCode = p_oCompDef.sCAPItemCode '"MERCER (SINGAPORE) PTE LTD"
                ElseIf row.Item(19).ToString = "3FS" Then
                    If bIsNonPanel = True Then
                        oDoc.Lines.ItemCode = p_oCompDef.s3FSItemCodeNonPanel
                    Else
                        oDoc.Lines.ItemCode = p_oCompDef.s3FSItemCode
                    End If
                Else
                    sErrDesc = "No Contract Type :: " & row.Item(12).ToString & " found in SAP. Please check the contract type."
                    Throw New ArgumentException(sErrDesc)
                End If

                oDoc.Lines.Quantity = 1

                If UCase(sInvRef) = "SUB-TOTAL" And row.Item(20).ToString = "GP Panel" Then
                    oDoc.Lines.VatGroup = "SO"
                    oDoc.Lines.Price = CDbl(row.Item(29))  ' Sub-Total
                Else
                    If bIsNonPanel = True Then
                        If CDbl(row.Item(30)) > 0 Then
                            oDoc.Lines.VatGroup = "SO"
                        Else
                            oDoc.Lines.VatGroup = "ZO"
                        End If
                        oDoc.Lines.PriceAfterVAT = CDbl(row.Item(33))
                    Else
                        oDoc.Lines.VatGroup = "SO"
                        If CDbl(row.Item(30)) > 0 Then
                            oDoc.Lines.PriceAfterVAT = CDbl(row.Item(33))
                        Else
                            oDoc.Lines.Price = CDbl(row.Item(33))
                        End If
                    End If
                End If

                If bSO = True Then
                    oDoc.Lines.SerialNum = GetSeriesNum("FSO")
                Else
                    oDoc.Lines.SerialNum = GetSeriesNum("FVM")
                End If

                oDoc.Lines.COGSCostingCode = sCostCenter
                oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = row.Item(16).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = row.Item(1)
                oDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value = row.Item(0).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = row.Item(17).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = CDbl(row.Item(25))
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = CDbl(row.Item(26))
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = CDbl(row.Item(27))
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = CDbl(row.Item(28))
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = CDbl(row.Item(29))
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = CDbl(row.Item(30))
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = CDbl(row.Item(31))
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = CDbl(row.Item(32))
                oDoc.Lines.UserFields.Fields.Item("U_AI_CostCenter").Value = row.Item(13).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_CompanyName").Value = row.Item(6).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderName").Value = row.Item(21).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_EmpName").Value = row.Item(14).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_EmpID").Value = row.Item(15).ToString

                oDoc.Lines.UserFields.Fields.Item("U_AI_BenefitType").Value = row.Item(20).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_Department").Value = row.Item(12).ToString

                oDoc.UserFields.Fields.Item("U_AI_BillToDept").Value = row.Item(12).ToString
                oDoc.UserFields.Fields.Item("U_AI_BillToCCenter").Value = row.Item(13).ToString

                '*****CHANGES
                If Not IsDBNull(row.Item(2)) = True Then
                    If Not row.Item(2).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = row.Item(2)
                End If

                If Not IsDBNull(row.Item(3)) = True Then
                    If Not row.Item(3).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = row.Item(3)
                End If

                If Not IsDBNull(row.Item(4)) = True Then
                    If Not row.Item(4).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = row.Item(4)
                End If

                oDoc.Lines.UserFields.Fields.Item("U_AI_BrokerCaseNo").Value = row.Item(8).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_InsuranceRefNo").Value = row.Item(9).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExternalRefNo").Value = row.Item(10).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_LineItem").Value = row.Item(22).ToString


                If Not IsDBNull(row.Item(23)) = True Then
                    If Not row.Item(23).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_QTY").Value = row.Item(23)
                End If

                '*****
            Next


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding ARI.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                p_oCompany.GetLastError(lErrCode, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding API failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim iARInvNo As Integer
            p_oCompany.GetNewObjectCode(iARInvNo)
            'AddDataToTable(p_oDtReport, "AR", iARInvNo, sCardCode, String.Empty)
            If bSO = True Then
                Console.WriteLine("Successfully Created Sales Order ::" & iARInvNo)
            Else
                Console.WriteLine("Successfully Created A/R Invoice ::" & iARInvNo)
            End If



            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice_FHG = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_FHG = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Public Function Upload_BillToClient_GJ(ByVal oDv As DataView, ByVal bServiceItem As Boolean, ByVal bIsNonPanel As Boolean, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim oDS As New DataSet
        Dim oDatasetBP As New DataSet
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim k, m As Integer
        Dim oBPArrary As New ArrayList
        Dim sBillType As String = String.Empty
        Dim oBPDT As DataTable
        Dim oBTDT As DataTable
        Dim oBPInhouseDT As DataTable

        Try
            sFuncName = "Upload_BillToClient_GJ"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Worksheet data.", sFuncName)
            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sBatchNo = oDv(2)(0).ToString()
            sBatchPeriod = oDv(3)(0).ToString()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch No.", sFuncName)
            k = InStrRev(sBatchNo, ":")
            sBatchNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - k).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch No:" & sBatchNo, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch Date.", sFuncName)
            m = InStrRev(sBatchPeriod, "to")
            sBatchPeriod = Microsoft.VisualBasic.Right(sBatchPeriod, Len(sBatchPeriod) - m - 1).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Date: " & sBatchPeriod, sFuncName)

            For i As Integer = 6 To oDv.Count - 1
                If Not oBPArrary.Contains(oDv(i)(3).ToString) Then
                    If Not oDv(i)(3).ToString = String.Empty Then
                        oBPArrary.Add(oDv(i)(3).ToString)
                    End If
                End If
            Next

            oBPDT = oDv.Table.Clone
            oBPDT.Clear()

            For iRow As Integer = 0 To oBPArrary.Count - 1

                sBillType = String.Empty
                '1. Get Billing Type from OCRD
                GetBillingType(oBPArrary(iRow).ToString, sBillType)
                Dim sCName As String = oBPArrary(iRow).ToString

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting BillType for ::" & sCName, sFuncName)

                If sBillType = String.Empty Then Throw New ArgumentException("No Billing Type specified for ::" & sCName)

                '2. SELECT BILLTYPE
                'Dim BPRows() As DataRow = oDv.Table.Select("F3='" & oBPArrary(iRow).ToString & "'")

                Dim BPRows() As DataRow = oDv.Table.Select("F4='" & Replace(oBPArrary(iRow).ToString, "'", "''") & "'")

                'UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sCardName), "'", "''") & "'"

                oBPDT.Clear()
                For Each row As DataRow In BPRows
                    oBPDT.ImportRow(row)
                Next

                'Create A/R Invoice
                Select Case sBillType

                    Case "CO"
                        Console.WriteLine("Creating Document")
                        sCardName = oBPArrary(iRow).ToString
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                        If AddARInvoice_GJ(sCardName, BPRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrDesc)
                        End If

                    Case "ENT"
                        Dim oEnTArray As New ArrayList
                        For Each row As DataRow In BPRows
                            If Not oEnTArray.Contains(row.Item(5).ToString) Then oEnTArray.Add(row.Item(5).ToString)
                        Next

                        For iEntRow As Integer = 0 To oEnTArray.Count - 1
                            Console.WriteLine("Creating Document")
                            Dim EntRows() As DataRow = oBPDT.Select("F6='" & Replace(oEnTArray(iEntRow).ToString, "'", "''") & "'")
                            sCardName = oEnTArray(iEntRow).ToString

                            If EntRows.Length = 0 Then
                                sErrDesc = "Entity is blank for ::" & oBPArrary(iRow).ToString
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                            If AddARInvoice_GJ(sCardName, EntRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "CO-DEPT"
                        Dim oDepArray As New ArrayList
                        For Each row As DataRow In BPRows
                            If Not oDepArray.Contains(row.Item(6).ToString) Then oDepArray.Add(row.Item(6).ToString)
                        Next

                        For iDepRow As Integer = 0 To oDepArray.Count - 1
                            Console.WriteLine("Creating Document")
                            Dim DepRows() As DataRow = oBPDT.Select("F7='" & Replace(oDepArray(iDepRow).ToString, "'", "''") & "'")
                            sCardName = oBPArrary(iRow).ToString

                            If DepRows.Length = 0 Then
                                sErrDesc = "Department is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If


                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                            If AddARInvoice_GJ(sCardName, DepRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "ENT-DEPT"
                        Dim oEntDeptDT As DataTable
                        oEntDeptDT = oBPDT.DefaultView.ToTable(True, "F6", "F7")

                        For Each row As DataRow In oEntDeptDT.Rows
                            Console.WriteLine("Creating Document")

                            Dim oEntDeptRows() As DataRow = oBPDT.Select("F6='" & Replace(row.Item(0).ToString, "'", "''") & "' and F7='" & Replace(row.Item(1).ToString, "'", "''") & "'")
                            sCardName = row.Item(0).ToString

                            If oEntDeptRows.Length = 0 Then
                                sErrDesc = "Entity/Department is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If


                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                            If AddARInvoice_GJ(sCardName, oEntDeptRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If

                        Next

                    Case "CO-DEPT-CC"
                        Dim oDeptCCDT As DataTable
                        oDeptCCDT = oBPDT.DefaultView.ToTable(True, "F7", "F8")

                        For Each row As DataRow In oDeptCCDT.Rows
                            Console.WriteLine("Creating Document")
                            Dim oDeptCCRows() As DataRow = oBPDT.Select("F7='" & Replace(row.Item(0).ToString, "'", "''") & "' and F8='" & Replace(row.Item(1).ToString, "'", "''") & "'")
                            sCardName = oBPArrary(iRow).ToString

                            If oDeptCCRows.Length = 0 Then
                                sErrDesc = "Department/CostCenter is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                            If AddARInvoice_GJ(sCardName, oDeptCCRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "ENT-DEPT-CC"

                        Dim oENTDEPCCDT As DataTable
                        oENTDEPCCDT = oBPDT.DefaultView.ToTable(True, "F6", "F7", "F8")

                        For Each row As DataRow In oENTDEPCCDT.Rows
                            Console.WriteLine("Creating Document")
                            Dim oENTDEPCCRows() As DataRow = oBPDT.Select("F6='" & Replace(row.Item(0).ToString, "'", "''") & "' and F7='" & Replace(row.Item(1).ToString, "'", "''") & "' and F8='" & Replace(row.Item(2).ToString, "'", "''") & "'")
                            sCardName = row.Item(0).ToString

                            If oENTDEPCCRows.Length = 0 Then
                                sErrDesc = "Entity/Department/CostCenter is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                            If AddARInvoice_GJ(sCardName, oENTDEPCCRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "CO-CC"
                        Dim oCCArray As New ArrayList
                        For Each row As DataRow In BPRows
                            If Not oCCArray.Contains(row.Item(7).ToString) Then
                                If Not row.Item(7).ToString = String.Empty Then
                                    oCCArray.Add(row.Item(7).ToString)
                                End If
                            End If
                        Next
                        For iCCRow As Integer = 0 To oCCArray.Count - 1
                            Console.WriteLine("Creating Document")
                            Dim CCRows() As DataRow = oBPDT.Select("F8='" & Replace(oCCArray(iCCRow).ToString, "'", "''") & "'")
                            sCardName = oBPArrary(iRow).ToString

                            If CCRows.Length = 0 Then
                                sErrDesc = "Cost Center is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                            If AddARInvoice_GJ(sCardName, CCRows, sBatchNo, sBatchPeriod, bServiceItem, False, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "ENT-CC"
                        Dim oEntCCDT As DataTable
                        oEntCCDT = oBPDT.DefaultView.ToTable(True, "F6", "F8")

                        For Each row As DataRow In oEntCCDT.Rows
                            Console.WriteLine("Creating Document")
                            Dim oEntCCRows() As DataRow = oBPDT.Select("F6='" & Replace(row.Item(0).ToString, "'", "''") & "' and F8='" & Replace(row.Item(1).ToString, "'", "''") & "'")
                            sCardName = row.Item(0).ToString

                            If oEntCCRows.Length = 0 Then
                                sErrDesc = "Entity/Cost center is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                            If AddARInvoice_GJ(sCardName, oEntCCRows, sBatchNo, sBatchPeriod, bServiceItem, False, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Next

                    Case "CO-BEN"
                        Dim oBenArray As New ArrayList
                        For Each row As DataRow In BPRows
                            If Not oBenArray.Contains(row.Item(14).ToString.Trim) Then oBenArray.Add(row.Item(14).ToString.Trim)
                        Next

                        For iBenRow As Integer = 0 To oBenArray.Count - 1

                            Console.WriteLine("Creating Document")
                            Dim BenRows() As DataRow = oBPDT.Select("F15='" & Replace(oBenArray(iBenRow).ToString, "'", "''") & "'")
                            sCardName = oBPArrary(iRow).ToString

                            If BenRows.Length = 0 Then
                                sErrDesc = "Benefit type is blank for ::" & sCardName
                                Throw New ArgumentException(sErrDesc)
                            End If

                            Dim oProvider As New ArrayList
                            Dim BenProvRows() As DataRow = Nothing

                            For Each row As DataRow In BenRows
                                If Not oProvider.Contains(UCase(row.Item(15).ToString.Trim)) Then oProvider.Add(UCase(row.Item(15).ToString.Trim)) 'Provider Name
                            Next

                            oBTDT = oBPDT.Clone
                            oBTDT.Clear()
                            oBPInhouseDT = oBPDT.Clone
                            oBPInhouseDT.Clear()

                            For iProvRow As Integer = 0 To oProvider.Count - 1
                                Dim sProviderName As String = oProvider(iProvRow).ToString
                                'Check Inhouse clinic? if Yes further splic into by Provider
                                BenProvRows = oBPDT.Select("F15='" & Replace(oBenArray(iBenRow).ToString, "'", "''") & "' and F16='" & Replace(sProviderName, "'", "''") & "'")

                                If IsInhouseClinic(sProviderName) = True Then
                                    For Each row As DataRow In BenProvRows
                                        oBPInhouseDT.ImportRow(row)
                                    Next
                                Else
                                    For Each row As DataRow In BenProvRows
                                        oBTDT.ImportRow(row)
                                    Next
                                End If

                                'If IsInhouseClinic(sProviderName) = True Then
                                '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                                '    If AddARInvoice_GJ(sCardName, BenProvRows, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                '        Throw New ArgumentException(sErrDesc)
                                '    End If
                                'Else
                                '    For Each row As DataRow In BenProvRows
                                '        oBTDT.ImportRow(row)
                                '    Next
                                'End If

                            Next

                            Dim BenFinalRow() As DataRow = oBTDT.Select("F4='" & Replace(sCardName, "'", "''") & "'")
                            If BenFinalRow.Length > 0 Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                                If AddARInvoice_GJ(sCardName, BenFinalRow, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                    Throw New ArgumentException(sErrDesc)
                                End If
                            End If

                            Dim BenInhouseRow() As DataRow = oBPInhouseDT.Select("F4='" & Replace(sCardName, "'", "''") & "'")
                            If BenInhouseRow.Length > 0 Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice_GJ()", sFuncName)
                                If AddARInvoice_GJ(sCardName, BenInhouseRow, sBatchNo, sBatchPeriod, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                                    Throw New ArgumentException(sErrDesc)
                                End If
                            End If
                        Next

                End Select
            Next


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            Upload_BillToClient_GJ = RTN_SUCCESS

        Catch ex As Exception
            Upload_BillToClient_GJ = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        Finally

        End Try
    End Function

    Private Function AddARInvoice_GJ(ByVal sCardName As String, _
                                       ByVal oBPRow() As DataRow, _
                                       ByVal sBatchNo As String, _
                                       ByVal sBatchPeriod As String, _
                                       ByVal bServiceItem As Boolean, _
                                       ByVal bIsNonPanel As Boolean, _
                                       ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim bSO As Boolean = False
        Dim bhasRows As Boolean = False
        Try
            sFuncName = "AddARInvoice_GJ"

            If CheckBP(sCardName, sCardCode, "C", sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If

            sCostCenter = GetCostCenter(sCardCode)

            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter

            If CheckCreateSO(sCardCode) = True And bIsNonPanel = True Then
                Console.WriteLine("Creating Sales Order..")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIapi SO Object", sFuncName)
                oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                bSO = True
                oDoc.DocDueDate = CDate(sBatchPeriod)
            Else
                Console.WriteLine("Creating A/R Invoice..")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
                oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            End If

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            'oDoc.DocDueDate = CDate(sBatchPeriod)
            oDoc.TaxDate = CDate(sBatchPeriod)
            oDoc.NumAtCard = "Batch No. : " & sBatchNo
            oDoc.Comments = sBatchPeriod
            oDoc.ImportFileNum = sBatchNo

            For Each row As DataRow In oBPRow
                If row.Item(13).ToString = "FFS" Then
                    bhasRows = True
                    iCnt += 1
                    If iCnt > 1 Then
                        oDoc.Lines.Add()
                    End If

                    If row.Item(13).ToString = "FFS" Then
                        oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCode
                    ElseIf row.Item(13).ToString = "CAP" Then
                        oDoc.Lines.ItemCode = p_oCompDef.sCAPItemCode
                    Else
                        sErrDesc = "No Contract Type :: " & row.Item(12).ToString & " found in SAP. Please check the contract type."
                        Throw New ArgumentException(sErrDesc)
                    End If

                    oDoc.Lines.Quantity = 1

                    If bIsNonPanel = True Then
                        If CDbl(row.Item(22)) > 0 Then
                            oDoc.Lines.VatGroup = "SO"
                        Else
                            oDoc.Lines.VatGroup = "ZO"
                        End If
                        oDoc.Lines.PriceAfterVAT = CDbl(row.Item(25))
                    Else
                        oDoc.Lines.VatGroup = "SO"
                        If CDbl(row.Item(22)) > 0 Then
                            oDoc.Lines.PriceAfterVAT = CDbl(row.Item(25))
                        Else
                            oDoc.Lines.Price = CDbl(row.Item(25))
                        End If
                    End If

                    If bSO = True Then
                        oDoc.Lines.SerialNum = GetSeriesNum("FSO")
                    Else
                        oDoc.Lines.SerialNum = GetSeriesNum("FVM")
                    End If

                    oDoc.Lines.COGSCostingCode = sCostCenter
                    oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = row.Item(10).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = row.Item(1)
                    oDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value = row.Item(0).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = row.Item(11).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = CDbl(row.Item(17))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = CDbl(row.Item(18))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = CDbl(row.Item(19))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = CDbl(row.Item(20))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = CDbl(row.Item(21))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = CDbl(row.Item(22))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = CDbl(row.Item(23))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = CDbl(row.Item(24))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_CostCenter").Value = row.Item(7).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_CompanyName").Value = row.Item(2).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderName").Value = row.Item(15).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_EmpName").Value = row.Item(8).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_EmpID").Value = row.Item(9).ToString

                    oDoc.Lines.UserFields.Fields.Item("U_AI_BenefitType").Value = row.Item(14).ToString

                    oDoc.Lines.UserFields.Fields.Item("U_AI_Department").Value = row.Item(6).ToString

                    oDoc.UserFields.Fields.Item("U_AI_BillToDept").Value = row.Item(6).ToString
                    oDoc.UserFields.Fields.Item("U_AI_BillToCCenter").Value = row.Item(7).ToString

                End If

            Next
            If bhasRows = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding ARI.", sFuncName)
                lRetCode = oDoc.Add
                If lRetCode <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding API failed.", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                Dim iARInvNo As Integer
                p_oCompany.GetNewObjectCode(iARInvNo)
                'AddDataToTable(p_oDtReport, "AR", iARInvNo, sCardCode, String.Empty)
                If bSO = True Then
                    Console.WriteLine("Successfully Created Sales Order ::" & iARInvNo)
                Else
                    Console.WriteLine("Successfully Created A/R Invoice ::" & iARInvNo)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice_GJ = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_GJ = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Public Function AddARInvoice_VisitNo(ByVal sCardName As String, _
                                     ByVal oBPRow() As DataRow, _
                                     ByVal sBatchNo As String, _
                                     ByVal sBatchPeriod As String, _
                                     ByVal sBatchPeriodName As String, _
                                     ByVal bServiceItem As Boolean, _
                                     ByVal bIsNonPanel As Boolean, _
                                     ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim bSO As Boolean = False

        Try
            sFuncName = "AddARInvoice_VisitNo"

            If CheckBP(sCardName, sCardCode, "C", sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If

            sCostCenter = GetCostCenter(sCardCode)

            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter

            If CheckCreateSO(sCardCode) = True And bIsNonPanel = True Then
                Console.WriteLine("Creating Sales Order..")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIapi SO Object", sFuncName)
                oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                bSO = True
            Else
                Console.WriteLine("Creating A/R Invoice..")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
                oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            End If

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            oDoc.DocDueDate = CDate(sBatchPeriod)
            oDoc.TaxDate = CDate(sBatchPeriod)
            oDoc.NumAtCard = "Batch No. : " & sBatchNo
            oDoc.ImportFileNum = sBatchNo


            For Each row As DataRow In oBPRow
                iCnt += 1
                If iCnt > 1 Then
                    oDoc.Lines.Add()
                End If

                ''oDoc.Comments = "Visit No: " & row.Item(0).ToString     ' Show Visit No.

                If row.Item(19).ToString = "FFS" Then
                    oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCode
                ElseIf row.Item(19).ToString = "CAP" Then
                    oDoc.Lines.ItemCode = p_oCompDef.sCAPItemCode
                Else
                    sErrDesc = "No Contract Type :: " & row.Item(19).ToString & " found in SAP. Please check the contract type."
                    Throw New ArgumentException(sErrDesc)
                End If

                oDoc.Lines.Quantity = 1

                If GetInvoiceRef(sCardName) = "SUB-TOTAL" And UCase(row.Item(20).ToString) = "GP PANEL" Then
                    oDoc.Lines.VatGroup = "SO"
                    oDoc.Lines.Price = CDbl(row.Item(29))  ' Sub-Total
                Else
                    If bIsNonPanel = True Then
                        If CDbl(row.Item(30)) > 0 Then
                            oDoc.Lines.VatGroup = "SO"
                        Else
                            oDoc.Lines.VatGroup = "ZO"
                        End If
                        oDoc.Lines.PriceAfterVAT = CDbl(row.Item(33))
                    Else
                        oDoc.Lines.VatGroup = "SO"
                        If CDbl(row.Item(30)) > 0 Then
                            oDoc.Lines.PriceAfterVAT = CDbl(row.Item(33))
                        Else
                            oDoc.Lines.Price = CDbl(row.Item(33))
                        End If
                    End If
                End If

                If bSO = True Then
                    oDoc.Lines.SerialNum = GetSeriesNum("FSO")
                Else
                    oDoc.Lines.SerialNum = GetSeriesNum("FVM")
                End If

                oDoc.Lines.COGSCostingCode = sCostCenter
                oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = row.Item(16).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = row.Item(1)
                oDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value = row.Item(0).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = row.Item(17).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = CDbl(row.Item(25))
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = CDbl(row.Item(26))
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = CDbl(row.Item(27))
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = CDbl(row.Item(28))
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = CDbl(row.Item(29))
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = CDbl(row.Item(30))
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = CDbl(row.Item(31))
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = CDbl(row.Item(32))
                oDoc.Lines.UserFields.Fields.Item("U_AI_CostCenter").Value = row.Item(13).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_CompanyName").Value = row.Item(6).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderName").Value = row.Item(21).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_EmpName").Value = row.Item(14).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_EmpID").Value = row.Item(15).ToString

                oDoc.Lines.UserFields.Fields.Item("U_AI_BenefitType").Value = row.Item(20).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_Department").Value = row.Item(12).ToString

                oDoc.UserFields.Fields.Item("U_AI_BillToDept").Value = row.Item(12).ToString
                oDoc.UserFields.Fields.Item("U_AI_BillToCCenter").Value = row.Item(13).ToString

                oDoc.UserFields.Fields.Item("U_AI_Nature").Value = sBatchPeriodName

                '*****CHANGES
                If Not IsDBNull(row.Item(2)) = True Then
                    If Not row.Item(2).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = row.Item(2)
                End If

                If Not IsDBNull(row.Item(3)) = True Then
                    If Not row.Item(3).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = row.Item(3)
                End If

                If Not IsDBNull(row.Item(4)) = True Then
                    If Not row.Item(4).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = row.Item(4)
                End If

                oDoc.Lines.UserFields.Fields.Item("U_AI_BrokerCaseNo").Value = row.Item(8).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_InsuranceRefNo").Value = row.Item(9).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExternalRefNo").Value = row.Item(10).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_LineItem").Value = row.Item(22).ToString


                If Not IsDBNull(row.Item(23)) = True Then
                    If Not row.Item(23).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_QTY").Value = row.Item(23)
                End If

                '*****


                'Update Remarks with Visit date & Patient Name
                oDoc.Comments = "Visit No :" & row.Item(0).ToString & vbCrLf & "Visit Date : " & row.Item(1).ToString & vbCrLf & row.Item(16).ToString

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding ARI.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                p_oCompany.GetLastError(lErrCode, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding ARI failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim iARInvNo As Integer
            p_oCompany.GetNewObjectCode(iARInvNo)
            'AddDataToTable(p_oDtReport, "AR", iARInvNo, sCardCode, String.Empty).
            If bSO = True Then
                Console.WriteLine("Successfully Created Sales Order ::" & iARInvNo)
            Else
                Console.WriteLine("Successfully Created A/R Invoice ::" & iARInvNo)
            End If



            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice_VisitNo = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_VisitNo = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Public Function AddARInvoice_VisitNo_ServiceFee(ByVal sCardName As String, _
                                     ByVal oBPRow() As DataRow, _
                                     ByVal sBatchNo As String, _
                                     ByVal sBatchPeriod As String, _
                                     ByVal sBatchPeriodName As String, _
                                     ByVal bServiceItem As Boolean, _
                                     ByVal bIsNonPanel As Boolean, _
                                     ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim bSO As Boolean = False

        Try
            sFuncName = "AddARInvoice_VisitNo_ServiceFee"

            If CheckBP(sCardName, sCardCode, "C", sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If

            sCostCenter = GetCostCenter(sCardCode)

            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter

            If CheckCreateSO(sCardCode) = True And bIsNonPanel = True Then
                Console.WriteLine("Creating Sales Order..")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIapi SO Object", sFuncName)
                oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                bSO = True
            Else
                Console.WriteLine("Creating A/R Invoice..")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
                oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            End If

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            oDoc.DocDueDate = CDate(sBatchPeriod)
            oDoc.TaxDate = CDate(sBatchPeriod)
            oDoc.NumAtCard = "Batch No. : " & sBatchNo
            oDoc.ImportFileNum = sBatchNo


            For Each row As DataRow In oBPRow
                iCnt += 1
                If iCnt > 1 Then
                    oDoc.Lines.Add()
                End If

                ''oDoc.Comments = "Visit No: " & row.Item(0).ToString     ' Show Visit No.

                If row.Item(19).ToString = "FFS" Then
                    oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCode
                ElseIf row.Item(19).ToString = "CAP" Then
                    oDoc.Lines.ItemCode = p_oCompDef.sCAPItemCode
                Else
                    sErrDesc = "No Contract Type :: " & row.Item(19).ToString & " found in SAP. Please check the contract type."
                    Throw New ArgumentException(sErrDesc)
                End If

                oDoc.Lines.Quantity = 1

                If GetInvoiceRef(sCardName) = "SUB-TOTAL" And UCase(row.Item(20).ToString) = "GP PANEL" Then
                    oDoc.Lines.VatGroup = "SO"
                    oDoc.Lines.Price = CDbl(row.Item(29))  ' Sub-Total
                Else
                    If bIsNonPanel = True Then
                        If CDbl(row.Item(30)) > 0 Then
                            oDoc.Lines.VatGroup = "SO"
                        Else
                            oDoc.Lines.VatGroup = "ZO"
                        End If
                        oDoc.Lines.PriceAfterVAT = CDbl(row.Item(33))
                    Else
                        oDoc.Lines.VatGroup = "SO"
                        If CDbl(row.Item(30)) > 0 Then
                            oDoc.Lines.PriceAfterVAT = CDbl(row.Item(33))
                        Else
                            oDoc.Lines.Price = CDbl(row.Item(33))
                        End If
                    End If
                End If

                If bSO = True Then
                    oDoc.Lines.SerialNum = GetSeriesNum("FSO")
                Else
                    oDoc.Lines.SerialNum = GetSeriesNum("FVM")
                End If

                oDoc.Lines.COGSCostingCode = sCostCenter
                oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = row.Item(16).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = row.Item(1)
                oDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value = row.Item(0).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = row.Item(17).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = CDbl(row.Item(25))
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = CDbl(row.Item(26))
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = CDbl(row.Item(27))
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = CDbl(row.Item(28))
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = CDbl(row.Item(29))
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = CDbl(row.Item(30))
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = CDbl(row.Item(31))
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = CDbl(row.Item(32))
                oDoc.Lines.UserFields.Fields.Item("U_AI_CostCenter").Value = row.Item(13).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_CompanyName").Value = row.Item(6).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderName").Value = row.Item(21).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_EmpName").Value = row.Item(14).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_EmpID").Value = row.Item(15).ToString

                oDoc.Lines.UserFields.Fields.Item("U_AI_BenefitType").Value = row.Item(20).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_Department").Value = row.Item(12).ToString

                oDoc.UserFields.Fields.Item("U_AI_BillToDept").Value = row.Item(12).ToString
                oDoc.UserFields.Fields.Item("U_AI_BillToCCenter").Value = row.Item(13).ToString

                oDoc.UserFields.Fields.Item("U_AI_Nature").Value = sBatchPeriodName

                '*****CHANGES
                If Not IsDBNull(row.Item(2)) = True Then
                    If Not row.Item(2).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = row.Item(2)
                End If

                If Not IsDBNull(row.Item(3)) = True Then
                    If Not row.Item(3).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = row.Item(3)
                End If

                If Not IsDBNull(row.Item(4)) = True Then
                    If Not row.Item(4).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = row.Item(4)
                End If

                oDoc.Lines.UserFields.Fields.Item("U_AI_BrokerCaseNo").Value = row.Item(8).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_InsuranceRefNo").Value = row.Item(9).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExternalRefNo").Value = row.Item(10).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_LineItem").Value = row.Item(22).ToString


                If Not IsDBNull(row.Item(23)) = True Then
                    If Not row.Item(23).ToString = String.Empty Then oDoc.Lines.UserFields.Fields.Item("U_AI_QTY").Value = row.Item(23)
                End If

                '*****


                'Update Remarks with Visit date & Patient Name
                oDoc.Comments = "Visit No :" & row.Item(0).ToString & vbCrLf & "Visit Date : " & row.Item(1).ToString & vbCrLf & row.Item(16).ToString

            Next

            ' ************  Add Service Fee. ********************************
            Dim dServiceFee As Double = 0

            For Each dtRow As DataRow In oBPRow
                dServiceFee = dServiceFee + CDbl(dtRow.Item(34))
            Next

            If dServiceFee > 0 Then
                oDoc.Lines.Add()
                oDoc.Lines.ItemCode = p_oCompDef.sNonStockItem
                oDoc.Lines.Quantity = 1
                oDoc.Lines.Price = dServiceFee
                oDoc.Lines.COGSCostingCode = sCostCenter
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding ARI.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                p_oCompany.GetLastError(lErrCode, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding ARI failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim iARInvNo As Integer
            p_oCompany.GetNewObjectCode(iARInvNo)
            'AddDataToTable(p_oDtReport, "AR", iARInvNo, sCardCode, String.Empty).
            If bSO = True Then
                Console.WriteLine("Successfully Created Sales Order ::" & iARInvNo)
            Else
                Console.WriteLine("Successfully Created A/R Invoice ::" & iARInvNo)
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice_VisitNo_ServiceFee = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_VisitNo_ServiceFee = RTN_ERROR
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function


End Module
