
Module modProcess

    Public Sub Start()

        Dim sFuncName As String = "Start()"
        Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
        Dim files() As System.IO.FileInfo
        Dim sErrdesc As String = String.Empty

        Try
            files = DirInfo.GetFiles("Batch*.xlsx")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Uploadfiles()", sFuncName)
            sGJDBName = String.Empty

            Uploadfiles(files)
            'Send Error Email if Datable has rows.
            If p_oDtError.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Error()", sFuncName)
                EmailTemplate_Error()
            End If
            p_oDtError.Rows.Clear()

            'Send Success Email if Datable has rows..
            If p_oDtSuccess.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Success()", sFuncName)
                EmailTemplate_Success()
            End If
            p_oDtSuccess.Rows.Clear()

            'Send SMS failure email if datatable has rows


           
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in upload", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try

    End Sub

    Private Sub Uploadfiles(ByVal files() As System.IO.FileInfo)

        Dim sFuncName As String = "Uploadfiles()"
        Dim sErrDesc As String = String.Empty
        Dim bIsFilesExist As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function..", sFuncName)

            p_oDtSuccess = CreateDataTable("FileName", "Status")
            p_oDtError = CreateDataTable("FileName", "Status", "ErrDesc")
            p_oDtReport = CreateDataTable("Type", "DocEntry", "BPCode", "Owner")
            p_oDtSMS = CreateDataTable("DocEntry", "MobileNo", "Amount")

            For Each File As System.IO.FileInfo In files
                sErrDesc = ""
                bIsFilesExist = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File name is : " & File.Name.ToUpper, sFuncName)

                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadDocument_GetDBName()", sFuncName)
                If ReadDocument_GetDBName(File.FullName, sErrDesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException("Unable to Read Document to get the DB Name..")
                End If

                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ConnectToCompany()", sFuncName)
                If ConnectToCompany(p_oCompany, sErrDesc) <> RTN_SUCCESS Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Unable to connect to SAP.", sFuncName)
                    EmailTemplate_GeneralError("Unable to connect to SAP." & " " & sErrDesc)
                    Throw New ArgumentException("Unable to connect to SAP.")
                End If

                If p_oCompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument()", sFuncName)
                    If UploadDocument(File.FullName, sErrDesc) <> RTN_SUCCESS Then
                        'Insert Error Description into Table
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                        AddDataToTable(p_oDtError, File.Name, "Error", sErrDesc)
                        'error condition
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Moving " & File.Name & " to " & p_oCompDef.sSuccessDir, sFuncName)
                        Dim UploadedFileName As String = Mid(File.Name, 1, File.Name.Length - 5) & "_" & Now.ToString("yyyyMMddhhmmss") & ".txt"
                        File.MoveTo(p_oCompDef.sFailDir & "\" & Replace(UploadedFileName, ".txt", ".xlsx"))

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File was not successfully uploaded" & File.FullName, sFuncName)
                        If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction()", sFuncName)
                        If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Moving " & File.Name & " to " & p_oCompDef.sSuccessDir, sFuncName)
                        Dim UploadedFileName As String = Mid(File.Name, 1, File.Name.Length - 5) & "_" & Now.ToString("yyyyMMddhhmmss") & ".txt"
                        File.MoveTo(p_oCompDef.sSuccessDir & "\" & Replace(UploadedFileName, ".txt", ".xlsx"))

                        'Insert Success Notificaiton into Table..
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                        AddDataToTable(p_oDtSuccess, File.Name, "Success")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & File.FullName, sFuncName)

                        ''Send SMS
                        'If p_oDtSMS.Rows.Count > 0 Then
                        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendSMS()", sFuncName)
                        '    If SendSMS(p_oDtSMS, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        'End If
                    End If

                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Disconnecting from SAP Databases", sFuncName)
                    p_oCompany.Disconnect()
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Disconnected from SAP Databases", sFuncName)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Unable to connect to SAP.", sFuncName)
                    Throw New ArgumentException("Unable to connect to SAP.")
                End If
            Next File


            If bIsFilesExist = False Then If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No files found to upload in INUPUT Folder.", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Funcation complete successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in upload setup", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Public Function UploadDocument(ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument()"
        Dim myfile As New System.IO.FileInfo(sFileName)
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim sPatientType As String = String.Empty


        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            sFileType = sFileName

            p_sPatientType = String.Empty

            Dim k As Integer = InStrRev(sFileType, "(")
            sFileType = Microsoft.VisualBasic.Right(sFileType, Len(sFileType) - k).Trim
            sFileType = "(" & Replace(sFileType, ".xlsx", "").Trim

            p_sPatientType = Left(Right(sFileType, 3), 2)

            If UCase(sFileType) = "(NON-PANEL OP)" Or UCase(sFileType) = "(NON-PANEL IP)" Then
                sSheet1 = "Bill to Client"
                sSheet2 = "Reimburse to Member"
                Dim oDv1 As DataView = Nothing
                Dim oDv2 As DataView = Nothing

                If p_sPatientType = "OP" Then
                    p_sPatientType = "non-panel outpatient"
                Else
                    p_sPatientType = "non-panel inpatient"
                End If

                bIsError = False
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadBillToClient", sFuncName)
                ReadBillToClient(sFileName, sSheet1, bIsError, oDv1, sErrDesc)

                If bIsError = True Then
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Billl To Client Excel Worksheet", sFuncName)
                    WriteToLogFile("Invalid Billl To Client Excel Worksheet " & sFileName, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                bIsError = False
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadReimburse_Member", sFuncName)
                ReadReimburse_Member(sFileName, sSheet2, bIsError, oDv2)

                If bIsError = True Then
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Reimburse to Provider Excel Worksheet", sFuncName)
                    sErrDesc = "Invalid Reimburse to Provider Excel Worksheet " & sFileName
                    WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction..", sFuncName)
                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessBillToClient()", sFuncName)
                If oDv1.Count > 6 Then
                    If Not oDv1(6)(0).ToString = String.Empty Then
                        If ProcessBillToClient(sFileName, sSheet1, True, oDv1, True, sErrDesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                End If
                If oDv2.Count > 6 Then
                    If Not oDv2(6)(0).ToString = String.Empty Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessReimburse_Member()", sFuncName)
                        If ProcessReimburse_Member(sFileName, sSheet2, oDv2, sErrDesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                End If

            ElseIf UCase(sFileType) = "(PANEL OP)" Or UCase(sFileType) = "(PANEL IP)" Then

                sSheet1 = "Bill to Client"
                sSheet2 = "Reimburse to Provider"
                Dim oDv3 As DataView = Nothing
                Dim oDv4 As DataView = Nothing

                If p_sPatientType = "OP" Then
                    p_sPatientType = "panel outpatient"
                Else
                    p_sPatientType = "panel inpatient"
                End If

                bIsError = False
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadBillToClient", sFuncName)
                ReadBillToClient(sFileName, sSheet1, bIsError, oDv3, sErrDesc)

                If bIsError = True Then
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Billl To Client Excel Worksheet", sFuncName)
                    WriteToLogFile("Invalid Billl To Client Excel Worksheet " & sFileName, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                bIsError = False

                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadReimburse_Provider()", sFuncName)
                ReadReimburse_Provider(sFileName, sSheet2, bIsError, oDv4)

                If bIsError = True Then
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Reimburse to Provider Excel Worksheet", sFuncName)
                    sErrDesc = "Invalid Reimburse to Provider Excel Worksheet " & sFileName
                    WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction..", sFuncName)
                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If oDv3.Count > 6 Then
                    If Not oDv3(6)(0).ToString = String.Empty Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessBillToClient()", sFuncName)
                        If ProcessBillToClient(sFileName, sSheet1, False, oDv3, False, sErrDesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                End If

                If oDv4.Count > 6 Then
                    If Not oDv4(6)(0).ToString = String.Empty Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessReimburse_Provider()", sFuncName)
                        If ProcessReimburse_Provider(sFileName, sSheet2, oDv4, sErrDesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                End If
            Else
                sErrDesc = "File name is Invalid. Please check the file name ::" & sFileName
                WriteToLogFile(sErrDesc, sFuncName)
                If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
                UploadDocument = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Public Function ReadDocument_GetDBName(ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "ReadDocument_GetDBName()"
        Dim myfile As New System.IO.FileInfo(sFileName)
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim sGJName As String = String.Empty
        Dim sDBName As String = String.Empty

        Try
            'Added logic to check if Contract Owner , if GJ connect to GJ DB else FHG_LIVE
            '25-10-2014

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            sFileType = sFileName
            Dim k As Integer = InStrRev(sFileType, "(")
            sFileType = Microsoft.VisualBasic.Right(sFileType, Len(sFileType) - k).Trim
            sFileType = "(" & Replace(sFileType, ".xlsx", "").Trim

            If UCase(sFileType) = "(NON-PANEL OP)" Or UCase(sFileType) = "(NON-PANEL IP)" Then
                sSheet1 = "Bill to Client"
                sSheet2 = "Reimburse to Member"
                Dim oDv1 As DataView = Nothing
                Dim oDv2 As DataView = Nothing

                ReadContractOwner(oDv1, sSheet1, sFileName, sDBName) 'Bill to client
                ReadContractOwner(oDv2, sSheet2, sFileName, sDBName) 'Reimburse to member
                If Not sDBName = String.Empty Then
                    p_oCompDef.sSAPDBName = sDBName
                    p_oCompDef.sCheckBankAccount = p_oCompDef.sGJ_CheckBankAccount
                    p_oCompDef.sCheckBankCode = p_oCompDef.sGJ_CheckBankCode
                    p_oCompDef.sCheckGLAccount = p_oCompDef.sGJ_CheckGLAccount
                    p_oCompDef.sGIROGLAccount = p_oCompDef.sGJ_GIROGLAccount

                    p_oCompDef.sCAPGLCode = p_oCompDef.sGJ_CAPGLCode
                    p_oCompDef.sFFSGLCode = p_oCompDef.sGJ_FFSGLCode

                End If

            ElseIf UCase(sFileType) = "(PANEL OP)" Or UCase(sFileType) = "(PANEL IP)" Then
                sSheet1 = "Bill to Client"
                sSheet2 = "Reimburse to Provider"
                Dim oDv3 As DataView = Nothing
                Dim oDv4 As DataView = Nothing

                ReadContractOwner(oDv3, sSheet1, sFileName, sDBName) 'Bill to client
                ReadContractOwner(oDv4, sSheet2, sFileName, sDBName) 'Reimburse to Provider
                If Not sDBName = String.Empty Then
                    p_oCompDef.sSAPDBName = sDBName
                    p_oCompDef.sCheckBankAccount = p_oCompDef.sGJ_CheckBankAccount
                    p_oCompDef.sCheckBankCode = p_oCompDef.sGJ_CheckBankCode
                    p_oCompDef.sCheckGLAccount = p_oCompDef.sGJ_CheckGLAccount
                    p_oCompDef.sGIROGLAccount = p_oCompDef.sGJ_GIROGLAccount

                    p_oCompDef.sCAPGLCode = p_oCompDef.sGJ_CAPGLCode
                    p_oCompDef.sFFSGLCode = p_oCompDef.sGJ_FFSGLCode

                End If
            Else
                sErrDesc = "File name is Invalid. Please check the file name ::" & sFileName
                WriteToLogFile(sErrDesc, sFuncName)
                If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            ReadDocument_GetDBName = RTN_SUCCESS

        Catch ex As Exception
            ReadDocument_GetDBName = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Public Function CheckBP(ByVal sBPName As String, ByRef sBPCode As String, ByVal sType As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "CheckBP()"
        Dim oRS As SAPbobsCOM.Recordset
        Dim RS As SAPbobsCOM.Recordset
        Dim sSql As String = String.Empty
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim lRetCode, lErrCode As Long

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oBP = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If sGJDBName = "Gethin-Jones Medical Practice Pte Ltd" Then
                sSql = "Select ""CardCode"" from ""OCRD"" Where ""CardType""='" & sType.Trim & "' and UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sBPName.Trim), "'", "''") & "'"
                oRS.DoQuery(sSql)
                If oRS.EoF Then
                    sErrDesc = "BP code doesn't exists in SAP for :: " & sBPName.Trim
                    Throw New ArgumentException(sErrDesc)
                Else
                    sBPCode = oRS.Fields.Item(0).Value.ToString
                    GoTo ExitSuc
                End If
            End If

            sSql = "Select ""CardCode"" from ""OCRD"" Where LEFT(""CardCode"",1)='M' and ""CardType""='" & sType.Trim & "' and UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sBPName.Trim), "'", "''") & "'"
            oRS.DoQuery(sSql)

            If oRS.EoF Then
                oBP.CardName = UCase(sBPName.Trim)
                If sType = "C" Then
                    oBP.CardType = SAPbobsCOM.BoCardTypes.cCustomer
                    oBP.Series = GetSeriesNum(p_oCompDef.sCustBPSeriesName)
                Else
                    oBP.CardType = SAPbobsCOM.BoCardTypes.cSupplier
                    oBP.Series = GetSeriesNum(p_oCompDef.sVenBPSeriesName)
                    oBP.GroupCode = GetBPGroupNum("MBMS-Panel GP")

                    oBP.Properties(1) = SAPbobsCOM.BoYesNoEnum.tYES

                    Dim iCnt As Integer
                    RS.DoQuery("SELECT ""PayMethCod"" FROM OPYM WHERE ""Type"" IN('O')")
                    If RS.RecordCount > 0 Then
                        RS.MoveFirst()
                        While RS.EoF = False
                            iCnt += 1
                            oBP.BPPaymentMethods.PaymentMethodCode = RS.Fields.Item("PayMethCod").Value
                            oBP.BPPaymentMethods.SetCurrentLine(iCnt - 1)
                            oBP.BPPaymentMethods.Add()
                            RS.MoveNext()
                        End While
                    End If
                    oBP.DebitorAccount = "2-12100-00"
                   
                End If

                lRetCode = oBP.Add
                If lRetCode <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding BP failed.", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If
                p_oCompany.GetNewObjectCode(sBPCode)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)
            Else
                sBPCode = oRS.Fields.Item(0).Value.ToString
            End If
ExitSuc:
            CheckBP = RTN_SUCCESS
        Catch ex As Exception
            CheckBP = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        Finally
            oRS = Nothing
        End Try
    End Function

    Public Function CheckBP_FHG(ByVal sBPName As String, ByRef sBPCode As String, ByVal sType As String, ByRef sInvRef As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "CheckBP_FHG()"
        Dim oRS As SAPbobsCOM.Recordset
        Dim RS As SAPbobsCOM.Recordset
        Dim sSql As String = String.Empty
        Dim oBP As SAPbobsCOM.BusinessPartners


        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oBP = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If sGJDBName = "Gethin-Jones Medical Practice Pte Ltd" Then
                sSql = "Select ""CardCode"" from ""OCRD"" Where ""CardType""='" & sType.Trim & "' and UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sBPName.Trim), "'", "''") & "'"
                oRS.DoQuery(sSql)
                If oRS.EoF Then
                    sErrDesc = "BP code doesn't exists in SAP for :: " & sBPName.Trim
                    Throw New ArgumentException(sErrDesc)
                Else
                    sBPCode = oRS.Fields.Item(0).Value.ToString
                    GoTo ExitSuc
                End If
            End If

            sSql = "Select ""CardCode"",""U_AI_INVAMTREF"" from ""OCRD"" Where LEFT(""CardCode"",1)='M' and ""CardType""='" & sType.Trim & "' and UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sBPName.Trim), "'", "''") & "'"
            oRS.DoQuery(sSql)

            If oRS.EoF Then
                sErrDesc = "BP code doesn't exists in SAP for :: " & sBPName.Trim
                Throw New ArgumentException(sErrDesc)
            Else
                sBPCode = oRS.Fields.Item(0).Value.ToString
                sInvRef = oRS.Fields.Item(1).Value.ToString
                GoTo ExitSuc
            End If
ExitSuc:
            CheckBP_FHG = RTN_SUCCESS
        Catch ex As Exception
            CheckBP_FHG = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        Finally
            oRS = Nothing
        End Try

    End Function


    Public Function GetCostCenter(ByVal sCardCode As String) As String
        Dim sCostCenter As String = String.Empty
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = "SELECT ""U_AI_DefaultCostCent"" FROM OCRD where ""CardCode""='" & sCardCode & "'"
        oDS = ExecuteSQLQuery(sSQL)
        If oDS.Tables(0).Rows.Count > 0 Then sCostCenter = oDS.Tables(0).Rows(0).Item(0).ToString

        Return sCostCenter

    End Function

    Public Function GetCostCenterByCardName(ByVal sCardName As String) As String
        Dim sCostCenter As String = String.Empty
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = "SELECT ""U_AI_DefaultCostCent"" FROM OCRD where UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sCardName), "'", "''") & "'"
        oDS = ExecuteSQLQuery(sSQL)
        If oDS.Tables(0).Rows.Count > 0 Then sCostCenter = oDS.Tables(0).Rows(0).Item(0).ToString

        Return sCostCenter

    End Function

    Public Function GetCostCenterByCardName_MV(ByVal sCardName As String) As String
        Dim sCostCenter As String = String.Empty
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = "SELECT ""U_AI_DefaultCostCent"" FROM OCRD where LEFT(""CardCode"",1)='M' and UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sCardName), "'", "''") & "'"
        oDS = ExecuteSQLQuery(sSQL)
        If oDS.Tables(0).Rows.Count > 0 Then sCostCenter = oDS.Tables(0).Rows(0).Item(0).ToString

        Return sCostCenter

    End Function


    Public Function GetSeriesNum(ByVal sSeriesName As String) As Integer

        Dim iSeriesNum As Integer = Nothing
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = "select ""Series"" from NNM1 where ""SeriesName""='" & sSeriesName & "'"

        oDS = ExecuteSQLQuery(sSQL)
        If oDS.Tables(0).Rows.Count > 0 Then iSeriesNum = oDS.Tables(0).Rows(0).Item(0)

        Return iSeriesNum

    End Function

    Private Function GetBPGroupNum(ByVal sGroupName As String) As Integer

        Dim iGroupNum As Integer = Nothing
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = "select ""GroupCode"" from OCRG  where ""GroupName""='" & sGroupName & "'"

        oDS = ExecuteSQLQuery(sSQL)
        If oDS.Tables(0).Rows.Count > 0 Then iGroupNum = oDS.Tables(0).Rows(0).Item(0)

        Return iGroupNum

    End Function

    Private Function GetDfltGLAccount(ByVal sBankCode As String, ByVal sBankAcct As String) As String

        Dim sGLAcct As String = String.Empty
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = "select ""GLAccount"" from DSC1 where ""BankCode""='" & sBankCode & "' and ""Account""='" & sBankAcct & "'"

        oDS = ExecuteSQLQuery(sSQL)
        If oDS.Tables(0).Rows.Count > 0 Then sGLAcct = oDS.Tables(0).Rows(0).Item(0)

        Return sGLAcct

    End Function


#Region "Bill to Client"

    Private Sub ReadBillToClient(ByVal sFileName As String, _
                                 ByVal sSheet As String, _
                                 ByRef bIsError As Boolean, _
                                 ByRef dv As DataView, _
                                 ByRef sErrdesc As String)

        Dim iHeaderRow As Integer
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "ReadBillToClient"
        Dim sBatchNo As String = String.Empty
        Dim k As Integer

        iHeaderRow = 5

        dv = GetDataViewFromExcel(sFileName, sSheet)

        If IsNothing(dv) Then
            Exit Sub
        End If

        sBatchNo = dv(2)(0).ToString()

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch No.", sFuncName)
        k = InStrRev(sBatchNo, ":")
        sBatchNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - k).Trim
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch No:" & sBatchNo, sFuncName)
        'Check if Batch No already exists

        If IsBatchNoExists(sBatchNo) = True Then
            sErrdesc = "Batch No. " & sBatchNo & " already uploaded in SAP. Please check the upload File ::" & sFileName
            WriteToLogFile(False, sErrdesc)
            bIsError = True
            Exit Sub
        End If

        If dv(iHeaderRow)(0).ToString <> "Visit No." Then
            sErrdesc = "Invalid Excel file Format - ([Visit No.] not found at Column 1"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(1).ToString <> "Visit Date" Then
            sErrdesc = "Invalid Excel file Format - ([Visit Date] not found at Column 2"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(2).ToString <> "Admission Date" Then
            sErrdesc = "Invalid Excel file Format - ([Admission Date] not found at Column 3"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(3).ToString <> "Discharged Date" Then
            sErrdesc = "Invalid Excel file Format - ([Discharged Date] not found at Column 4"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString <> "Injured Date" Then
            sErrdesc = "Invalid Excel file Format - ([Injured Date] not found at Column 5"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(5).ToString <> "Company Code" Then
            sErrdesc = "Invalid Excel file Format - ([Company Code] not found at Column 6"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(6).ToString <> "Company Name" Then
            sErrdesc = "Invalid Excel file Format - ([Company Name] not found at Column 7"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(7).ToString <> "Broker Name" Then
            sErrdesc = "Invalid Excel file Format - ([Broker Name] not found at Column 8"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString <> "Broker Case No." Then
            sErrdesc = "Invalid Excel file Format - ([Broker Case No.] not found at Column 9"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(9).ToString <> "Insurer Ref No." Then
            sErrdesc = "Invalid Excel file Format - ([Insurer Ref No.] not found at Column 10"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString <> "External Ref No." Then
            sErrdesc = "Invalid Excel file Format - ([External Ref No.] not found at Column 11"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(11).ToString <> "Entity" Then
            sErrdesc = "Invalid Excel file Format - ([Entity] not found at Column 12"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(12).ToString <> "Department" Then
            sErrdesc = "Invalid Excel file Format - ([Department] not found at Column 13"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(13).ToString <> "Cost Centre" Then
            sErrdesc = "Invalid Excel file Format - ([Cost Centre] not found at Column 14"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(14).ToString <> "Employee Name" Then
            sErrdesc = "Invalid Excel file Format - ([Employee Name] not found at Column 15"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(15).ToString <> "Employee ID No." Then
            sErrdesc = "Invalid Excel file Format - ([Employee ID No.] not found at Column 16"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(16).ToString <> "Patient Name" Then
            sErrdesc = "Invalid Excel file Format - ([Patient Name] not found at Column 17"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(17).ToString <> "Patient ID No." Then
            sErrdesc = "Invalid Excel file Format - ([Patient ID No.] not found at Column 18"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(18).ToString <> "Patient Member Type" Then
            sErrdesc = "Invalid Excel file Format - ([Patient Member Type] not found at Column 19"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(19).ToString <> "Contract Type" Then
            sErrdesc = "Invalid Excel file Format - ([Contract Type] not found at Column 20"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(20).ToString <> "Benefit Type" Then
            sErrdesc = "Invalid Excel file Format - ([Benefit Type] not found at Column 21"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(21).ToString <> "Provider Name" Then
            sErrdesc = "Invalid Excel file Format - ([Provider Name] not found at Column 22"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(22).ToString <> "Line Item" Then
            sErrdesc = "Invalid Excel file Format - ([Line Item] not found at Column 23"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(23).ToString <> "Qty." Then
            sErrdesc = "Invalid Excel file Format - ([Qty.] not found at Column 24"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(24).ToString <> "Currency" Then
            sErrdesc = "Invalid Excel file Format - ([Currency] not found at Column 25"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(25).ToString <> "Consult Cost" Then
            sErrdesc = "Invalid Excel file Format - ([Consult Cost] not found at Column 26"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(26).ToString <> "Drug Cost" Then
            sErrdesc = "Invalid Excel file Format - ([Drug Cost] not found at Column 27"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(27).ToString <> "In-house Service Cost" Then
            sErrdesc = "Invalid Excel file Format - ([In-house Service Cost] not found at Column 28"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(28).ToString <> "External Service Cost" Then
            sErrdesc = "Invalid Excel file Format - ([External Service Cost ] not found at Column 29"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(29).ToString <> "Sub-total" Then
            sErrdesc = "Invalid Excel file Format - ([Sub-total] not found at Column 30"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(30).ToString <> "Tax" Then
            sErrdesc = "Invalid Excel file Format - ([Tax] not found at Column 31"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(31).ToString <> "Grand Total" Then
            sErrdesc = "Invalid Excel file Format - ([Grand Total] not found at Column 32"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(32).ToString <> "Unclaim Amt." Then
            sErrdesc = "Invalid Excel file Format - ([Unclaim Amt.] not found at Column 33"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(33).ToString <> "Claim Amt." Then
            sErrdesc = "Invalid Excel file Format - ([Claim Amt.] not found at Column 34"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(34).ToString <> "Service Fee" Then
            sErrdesc = "Invalid Excel file Format - ([Claim Amt.] not found at Column 35"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(35).ToString <> "Remarks for Client" Then
            sErrdesc = "Invalid Excel file Format - ([Remarks for Client] not found at Column 36"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If



    End Sub

    Private Function ProcessBillToClient(ByVal sFileName As String, _
                                         ByVal sSheet As String, _
                                         ByVal bServiceItem As Boolean, _
                                         ByVal oDv As DataView, _
                                         ByVal bIsNonPanel As Boolean, _
                                         ByRef sErrDesc As String) As Long

        Dim IsError As Boolean = False
        Dim sFuncName As String = "ProcessBillToClient"

        Try

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling AddARIvoice_BillToClient()", sFuncName)
            If AddARIvoice_BillToClient(oDv, bServiceItem, bIsNonPanel, sErrDesc) <> RTN_SUCCESS Then
                If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Error in Adding AR Invoice - Bill to Clinet", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Function Completed successfully.", sFuncName)
            ProcessBillToClient = RTN_SUCCESS
        Catch ex As Exception
            ProcessBillToClient = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        End Try

    End Function

    Private Function AddARIvoice_BillToClient(ByVal oDv As DataView, ByVal bServiceItem As Boolean, ByVal bIsNonPanel As Boolean, ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "AddARIvoice_BillToClient"
        Dim sCardName As String = String.Empty
        Dim sCardCode As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Contract Owner Name.", sFuncName)
            sCardName = oDv(0)(0).ToString
            Dim K As Integer = InStrRev(sCardName, ":")
            sCardName = Microsoft.VisualBasic.Right(sCardName, Len(sCardName) - K).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Card Name:" & sCardName, sFuncName)

            Dim sGJName As String = "Gethin-Jones Medical Practice Pte Ltd"

            If UCase(sCardName.Trim) = UCase(sGJName) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Upload_BillToClient_GJ()", sFuncName)
                If Upload_BillToClient_GJ(oDv, bServiceItem, bIsNonPanel, sErrdesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrdesc)
                End If
            ElseIf Left(UCase(sCardName.Trim), 5) = "AETNA" Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice()", sFuncName)
                If Upload_BillToClient_FHG(sCardName, oDv, bServiceItem, bIsNonPanel, sErrdesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrdesc)
                End If
            ElseIf CheckContractOwner(sCardName) = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Upload_BillToClient_FHG()", sFuncName)
                If Upload_BillToClient_FHG(sCardName, oDv, bServiceItem, bIsNonPanel, sErrdesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrdesc)
                End If
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CheckBP()", sFuncName)
                If CheckBP(sCardName, sCardCode, "C", sErrdesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrdesc)
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddARInvoice()", sFuncName)
                If AddARInvoice(oDv, sCardName, bServiceItem, bIsNonPanel, sErrdesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrdesc)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)
            AddARIvoice_BillToClient = RTN_SUCCESS

        Catch ex As Exception
            AddARIvoice_BillToClient = RTN_ERROR
            sErrdesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrdesc, sFuncName)
        End Try
    End Function

    Private Function AddARInvoice(ByVal oDv As DataView, _
                                  ByVal sCardName As String, _
                                  ByVal bServiceItem As Boolean, _
                                  ByVal bIsNonPanel As Boolean, _
                                  ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim sRemarks As String = String.Empty
        Dim k, j As Integer
        Dim sCostCenter As String = String.Empty
        Dim dGrossTotal As Double = 0
        Dim dConsultCost As Double = 0
        Dim dDrugCost As Double = 0
        Dim dInhouseServiceCost As Double = 0
        Dim dExternalServiceCost As Double = 0
        Dim dSubTotal As Double = 0
        Dim dTax As Double = 0
        Dim dGrandTotal As Double = 0
        Dim dUnclaimAmt As Double = 0
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim bHasLines As Boolean = False
        Dim bNoTaxLines As Boolean = False
        Dim bNoTaxLines_3FS As Boolean = False

        Dim bHasLines_3FS As Boolean
        Dim iCnt_3FS As Integer
        Dim dGrossTotal_3FS As Double = 0
        Dim dConsultCost_3FS As Double = 0
        Dim dDrugCost_3FS As Double = 0
        Dim dInhouseServiceCost_3FS As Double = 0
        Dim dExternalServiceCost_3FS As Double = 0
        Dim dSubTotal_3FS As Double = 0
        Dim dTax_3FS As Double = 0
        Dim dGrandTotal_3FS As Double = 0
        Dim dUnclaimAmt_3FS As Double = 0


        Try

            If CheckBP(sCardName, sCardCode, "C", sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If

            sCostCenter = GetCostCenter(sCardCode)

            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sBatchNo = oDv(2)(0).ToString()
            sBatchPeriod = oDv(3)(0).ToString()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch No.", sFuncName)
            k = InStrRev(sBatchNo, ":")
            sBatchNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - k).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch No:" & sBatchNo, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch Date.", sFuncName)
            j = InStrRev(sBatchPeriod, "to")
            sBatchPeriod = Microsoft.VisualBasic.Right(sBatchPeriod, Len(sBatchPeriod) - j - 1).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Date: " & sBatchPeriod, sFuncName)

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            'oDoc.DocDueDate = CDate(sBatchPeriod)
            oDoc.TaxDate = CDate(sBatchPeriod)
            oDoc.NumAtCard = oDv(2)(0).ToString()
            oDoc.Comments = sBatchPeriod
            oDoc.ImportFileNum = sBatchNo

            Dim sContractType As String = String.Empty


            For i As Integer = 6 To oDv.Count - 1
                If oDv(i)(19).ToString = "FFS" Then
                    sContractType = oDv(i)(19).ToString
                    If CDbl(oDv(i)(30)) > 0 Then
                        bHasLines = True
                        iCnt += 1
                        dGrossTotal = dGrossTotal + CDbl(oDv(i)(33))
                        dConsultCost = dConsultCost + CDbl(oDv(i)(25))
                        dDrugCost = dDrugCost + CDbl(oDv(i)(26))
                        dInhouseServiceCost = dInhouseServiceCost + CDbl(oDv(i)(27))
                        dExternalServiceCost = dExternalServiceCost + CDbl(oDv(i)(28))
                        dSubTotal = dSubTotal + CDbl(oDv(i)(29))
                        dTax = dTax + CDbl(oDv(i)(30))
                        dGrandTotal = dGrandTotal + CDbl(oDv(i)(31))
                        dUnclaimAmt = dUnclaimAmt + +CDbl(oDv(i)(32))
                    End If
                End If
            Next

            If bHasLines = True Then

                If bIsNonPanel = True Then
                    oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCodeNonPanel
                Else
                    oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCode
                End If



                oDoc.Lines.Quantity = iCnt
                oDoc.Lines.COGSCostingCode = sCostCenter

                If bIsNonPanel = True Then
                    If dTax > 0 Then
                        oDoc.Lines.VatGroup = "SO"
                        oDoc.Lines.LineTotal = dGrossTotal * 100 / 107
                    Else
                        oDoc.Lines.VatGroup = "ZO"
                        oDoc.Lines.LineTotal = dGrossTotal
                    End If

                Else
                    oDoc.Lines.VatGroup = "SO"
                    If dTax > 0 Then
                        oDoc.Lines.LineTotal = dGrossTotal * 100 / 107
                    Else
                        oDoc.Lines.LineTotal = dGrossTotal
                    End If
                End If


                oDoc.Lines.SerialNum = GetSeriesNum("FVM")
                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConsultCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInhouseServiceCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = dExternalServiceCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dTax
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnclaimAmt

                '*****CHANGES
                'oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = CDate(oDv(6)(2))
                'oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = CDate(oDv(6)(3))
                'oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = CDate(oDv(6)(4))

                oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = oDv(6)(2)
                oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = oDv(6)(3)
                oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = oDv(6)(4)

                oDoc.Lines.UserFields.Fields.Item("U_AI_BrokerCaseNo").Value = oDv(6)(8).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_InsuranceRefNo").Value = oDv(6)(9).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExternalRefNo").Value = oDv(6)(10).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_LineItem").Value = oDv(6)(22).ToString
                If Not oDv(6)(23).ToString = String.Empty Then
                    oDoc.Lines.UserFields.Fields.Item("U_AI_QTY").Value = CDbl(oDv(6)(23))
                End If

                '*****

            Else
                bHasLines = False
            End If

            '======================Add No Tax Lines FFS ===============================

            dGrossTotal = 0
            dConsultCost = 0
            dDrugCost = 0
            dInhouseServiceCost = 0
            dExternalServiceCost = 0
            dSubTotal = 0
            dTax = 0
            dGrandTotal = 0
            dUnclaimAmt = 0
            iCnt = 0

            sContractType = String.Empty

            For i As Integer = 6 To oDv.Count - 1
                If oDv(i)(19).ToString = "FFS" Then
                    If CDbl(oDv(i)(30)) = 0 Then
                        bNoTaxLines = True
                        iCnt += 1
                        dGrossTotal = dGrossTotal + CDbl(oDv(i)(33))
                        dConsultCost = dConsultCost + CDbl(oDv(i)(25))
                        dDrugCost = dDrugCost + CDbl(oDv(i)(26))
                        dInhouseServiceCost = dInhouseServiceCost + CDbl(oDv(i)(27))
                        dExternalServiceCost = dExternalServiceCost + CDbl(oDv(i)(28))
                        dSubTotal = dSubTotal + CDbl(oDv(i)(29))
                        dTax = dTax + CDbl(oDv(i)(30))
                        dGrandTotal = dGrandTotal + CDbl(oDv(i)(31))
                        dUnclaimAmt = dUnclaimAmt + +CDbl(oDv(i)(32))
                    End If
                End If
            Next

            If bNoTaxLines = True Then
                If bHasLines = True Then
                    oDoc.Lines.Add()
                End If

                If bIsNonPanel = True Then
                    oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCodeNonPanel
                Else
                    oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCode
                End If

                oDoc.Lines.Quantity = iCnt
                oDoc.Lines.COGSCostingCode = sCostCenter


                If bIsNonPanel = True Then
                    If dTax > 0 Then
                        oDoc.Lines.VatGroup = "SO"
                        oDoc.Lines.LineTotal = dGrossTotal * 100 / 107
                    Else
                        oDoc.Lines.VatGroup = "ZO"
                        oDoc.Lines.LineTotal = dGrossTotal
                    End If

                Else
                    oDoc.Lines.VatGroup = "SO"
                    If dTax > 0 Then
                        oDoc.Lines.LineTotal = dGrossTotal * 100 / 107
                    Else
                        oDoc.Lines.LineTotal = dGrossTotal
                    End If
                End If


                oDoc.Lines.SerialNum = GetSeriesNum("FVM")
                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConsultCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInhouseServiceCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = dExternalServiceCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dTax
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnclaimAmt

                '*****CHANGES
                'oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = CDate(oDv(6)(2))
                'oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = CDate(oDv(6)(3))
                'oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = CDate(oDv(6)(4))

                oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = oDv(6)(2)
                oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = oDv(6)(3)
                oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = oDv(6)(4)


                oDoc.Lines.UserFields.Fields.Item("U_AI_BrokerCaseNo").Value = oDv(6)(8).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_InsuranceRefNo").Value = oDv(6)(9).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExternalRefNo").Value = oDv(6)(10).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_LineItem").Value = oDv(6)(22).ToString
                If Not oDv(6)(23).ToString = String.Empty Then
                    oDoc.Lines.UserFields.Fields.Item("U_AI_QTY").Value = CDbl(oDv(6)(23))
                End If

                '*****

            End If

            ' *************************   3FS     ****************************************************************************************************************************

            For i As Integer = 6 To oDv.Count - 1
                If oDv(i)(19).ToString = "3FS" Then
                    sContractType = oDv(i)(19).ToString
                    If CDbl(oDv(i)(30)) > 0 Then
                        bHasLines_3FS = True
                        iCnt_3FS += 1
                        dGrossTotal_3FS = dGrossTotal_3FS + CDbl(oDv(i)(33))
                        dConsultCost_3FS = dConsultCost_3FS + CDbl(oDv(i)(25))
                        dDrugCost_3FS = dDrugCost_3FS + CDbl(oDv(i)(26))
                        dInhouseServiceCost_3FS = dInhouseServiceCost_3FS + CDbl(oDv(i)(27))
                        dExternalServiceCost_3FS = dExternalServiceCost_3FS + CDbl(oDv(i)(28))
                        dSubTotal_3FS = dSubTotal_3FS + CDbl(oDv(i)(29))
                        dTax_3FS = dTax_3FS + CDbl(oDv(i)(30))
                        dGrandTotal_3FS = dGrandTotal_3FS + CDbl(oDv(i)(31))
                        dUnclaimAmt_3FS = dUnclaimAmt_3FS + +CDbl(oDv(i)(32))
                    End If
                End If
            Next

            If bHasLines_3FS = True Then

                If bHasLines = True Or bNoTaxLines = True Then oDoc.Lines.Add()

                If bIsNonPanel = True Then
                    oDoc.Lines.ItemCode = p_oCompDef.s3FSItemCodeNonPanel
                Else
                    oDoc.Lines.ItemCode = p_oCompDef.s3FSItemCode
                End If

                oDoc.Lines.Quantity = iCnt_3FS
                oDoc.Lines.COGSCostingCode = sCostCenter

                If bIsNonPanel = True Then
                    If dTax > 0 Then
                        oDoc.Lines.VatGroup = "SO"
                        oDoc.Lines.LineTotal = dGrossTotal_3FS * 100 / 107
                    Else
                        oDoc.Lines.VatGroup = "ZO"
                        oDoc.Lines.LineTotal = dGrossTotal_3FS
                    End If

                Else
                    oDoc.Lines.VatGroup = "SO"
                    If dTax > 0 Then
                        oDoc.Lines.LineTotal = dGrossTotal_3FS * 100 / 107
                    Else
                        oDoc.Lines.LineTotal = dGrossTotal_3FS
                    End If
                End If


                oDoc.Lines.SerialNum = GetSeriesNum("FVM")
                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConsultCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInhouseServiceCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = dExternalServiceCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dTax_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnclaimAmt_3FS

                '*****CHANGES
                'oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = CDate(oDv(6)(2))
                'oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = CDate(oDv(6)(3))
                'oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = CDate(oDv(6)(4))

                oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = oDv(6)(2)
                oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = oDv(6)(3)
                oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = oDv(6)(4)

                oDoc.Lines.UserFields.Fields.Item("U_AI_BrokerCaseNo").Value = oDv(6)(8).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_InsuranceRefNo").Value = oDv(6)(9).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExternalRefNo").Value = oDv(6)(10).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_LineItem").Value = oDv(6)(22).ToString
                If Not oDv(6)(23).ToString = String.Empty Then
                    oDoc.Lines.UserFields.Fields.Item("U_AI_QTY").Value = CDbl(oDv(6)(23))
                End If
                '*****

            Else
                bHasLines_3FS = False
            End If


            '=======================            Add No Tax Lines 3FS     ===============================

            dGrossTotal_3FS = 0
            dConsultCost_3FS = 0
            dDrugCost_3FS = 0
            dInhouseServiceCost_3FS = 0
            dExternalServiceCost_3FS = 0
            dSubTotal_3FS = 0
            dTax_3FS = 0
            dGrandTotal_3FS = 0
            dUnclaimAmt_3FS = 0
            iCnt_3FS = 0

            sContractType = String.Empty

            For i As Integer = 6 To oDv.Count - 1
                If oDv(i)(19).ToString = "3FS" Then
                    If CDbl(oDv(i)(30)) = 0 Then
                        bNoTaxLines_3FS = True
                        iCnt_3FS += 1
                        dGrossTotal_3FS = dGrossTotal_3FS + CDbl(oDv(i)(33))
                        dConsultCost_3FS = dConsultCost_3FS + CDbl(oDv(i)(25))
                        dDrugCost_3FS = dDrugCost_3FS + CDbl(oDv(i)(26))
                        dInhouseServiceCost_3FS = dInhouseServiceCost_3FS + CDbl(oDv(i)(27))
                        dExternalServiceCost_3FS = dExternalServiceCost_3FS + CDbl(oDv(i)(28))
                        dSubTotal_3FS = dSubTotal_3FS + CDbl(oDv(i)(29))
                        dTax_3FS = dTax_3FS + CDbl(oDv(i)(30))
                        dGrandTotal_3FS = dGrandTotal_3FS + CDbl(oDv(i)(31))
                        dUnclaimAmt_3FS = dUnclaimAmt_3FS + +CDbl(oDv(i)(32))
                    End If
                End If
            Next

            If bNoTaxLines_3FS = True Then
                If bHasLines_3FS = True Or bHasLines = True Or bNoTaxLines = True Then
                    oDoc.Lines.Add()
                End If

                If bIsNonPanel = True Then
                    oDoc.Lines.ItemCode = p_oCompDef.s3FSItemCodeNonPanel
                Else
                    oDoc.Lines.ItemCode = p_oCompDef.s3FSItemCode
                End If

                oDoc.Lines.Quantity = iCnt_3FS
                oDoc.Lines.COGSCostingCode = sCostCenter

                If bIsNonPanel = True Then
                    If dTax_3FS > 0 Then
                        oDoc.Lines.VatGroup = "SO"
                        oDoc.Lines.LineTotal = dGrossTotal_3FS * 100 / 107
                    Else
                        oDoc.Lines.VatGroup = "ZO"
                        oDoc.Lines.LineTotal = dGrossTotal_3FS
                    End If

                Else
                    oDoc.Lines.VatGroup = "SO"
                    If dTax > 0 Then
                        oDoc.Lines.LineTotal = dGrossTotal_3FS * 100 / 107
                    Else
                        oDoc.Lines.LineTotal = dGrossTotal_3FS
                    End If
                End If

                oDoc.Lines.SerialNum = GetSeriesNum("FVM")
                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConsultCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInhouseServiceCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = dExternalServiceCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dTax_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnclaimAmt_3FS

                '*****CHANGES
                'oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = CDate(oDv(6)(2))
                'oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = CDate(oDv(6)(3))
                'oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = CDate(oDv(6)(4))

                oDoc.Lines.UserFields.Fields.Item("U_AI_AdmissionDate").Value = oDv(6)(2)
                oDoc.Lines.UserFields.Fields.Item("U_AI_DischargeDate").Value = oDv(6)(3)
                oDoc.Lines.UserFields.Fields.Item("U_AI_InjuryDate").Value = oDv(6)(4)


                oDoc.Lines.UserFields.Fields.Item("U_AI_BrokerCaseNo").Value = oDv(6)(8).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_InsuranceRefNo").Value = oDv(6)(9).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_ExternalRefNo").Value = oDv(6)(10).ToString
                oDoc.Lines.UserFields.Fields.Item("U_AI_LineItem").Value = oDv(6)(22).ToString
                If Not oDv(6)(23).ToString = String.Empty Then
                    oDoc.Lines.UserFields.Fields.Item("U_AI_QTY").Value = CDbl(oDv(6)(23))
                End If
                '*****

            End If


            ' ************************************************** END 3FS *****************************************************************************************************
            ' Add Service Fee.

            Dim oDtSVFee As DataTable
            oDtSVFee = oDv.Table.DefaultView.ToTable(True, "F35")

            For h As Integer = 0 To oDtSVFee.Rows.Count - 1
                With oDtSVFee.Rows(h)
                    If IsNumeric(.Item(0)) And h > 1 Then
                        If CDbl(.Item(0)) > 0 Then
                            Dim DepRows() As DataRow = oDv.Table.Select("F35='" & .Item(0).ToString.Trim & "'")
                            If DepRows.Length > 0 Then
                                oDoc.Lines.Add()
                                oDoc.Lines.ItemCode = p_oCompDef.sNonStockItem
                                oDoc.Lines.Quantity = DepRows.Length
                                oDoc.Lines.Price = CDbl(.Item(0))
                                oDoc.Lines.COGSCostingCode = sCostCenter
                            End If
                        End If
                    End If
                End With
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

            For i As Integer = 6 To oDv.Count - 1

                sSQL = "insert into ""AI_TB02_BILLTOCLIENT"" values(" & iARInvNo & ",'" & oDv(i)(0).ToString & "','" & CDate(oDv(i)(1)).ToString("yyyyMMdd") & "','" & Replace(oDv(i)(6).ToString, "'", "''") & "'" & _
                      " ,'" & Replace(oDv(i)(5).ToString, "'", "''") & "','" & oDv(i)(7).ToString & "','" & Replace(oDv(i)(11).ToString, "'", "''") & "','" & Replace(oDv(i)(12).ToString, "'", "''") & "'" & _
                      " ,'" & Replace(oDv(i)(13).ToString, "'", "''") & "','" & Replace(oDv(i)(14).ToString, "'", "''") & "','" & oDv(i)(15).ToString & "','" & Replace(oDv(i)(16).ToString, "'", "''") & "'" & _
                      " ,'" & Replace(oDv(i)(17).ToString, "'", "''") & "','" & oDv(i)(18).ToString & "','" & oDv(i)(19).ToString & "','" & oDv(i)(20).ToString & "'" & _
                      " ,'" & Replace(oDv(i)(21).ToString, "'", "''") & "','" & oDv(i)(24).ToString & "'," & CDbl(oDv(i)(25)) & "," & CDbl(oDv(i)(26)) & "," & CDbl(oDv(i)(27)) & "," & CDbl(oDv(i)(28)) & "," & CDbl(oDv(i)(29)) & _
                      "," & CDbl(oDv(i)(30)) & "," & CDbl(oDv(i)(31)) & "," & CDbl(oDv(i)(32)) & "," & CDbl(oDv(i)(33)) & "," & CDbl(oDv(i)(34)) & ",'" & oDv(i)(35).ToString & "'" & _
                      ",'" & p_sPatientType & "','" & Replace(oDv(3)(0).ToString(), "Batch Period : ", "") & "','" & sBatchNo.Trim & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)"

                oRS.DoQuery(sSQL)

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Function AddARInvoice_Backup(ByVal oDv As DataView, _
                                  ByVal sCardName As String, _
                                  ByVal bServiceItem As Boolean, _
                                  ByVal bIsNonPanel As Boolean, _
                                  ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim sRemarks As String = String.Empty
        Dim k, j As Integer
        Dim sCostCenter As String = String.Empty
        Dim dGrossTotal As Double = 0
        Dim dConsultCost As Double = 0
        Dim dDrugCost As Double = 0
        Dim dInhouseServiceCost As Double = 0
        Dim dExternalServiceCost As Double = 0
        Dim dSubTotal As Double = 0
        Dim dTax As Double = 0
        Dim dGrandTotal As Double = 0
        Dim dUnclaimAmt As Double = 0
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing

        Try

            If CheckBP(sCardName, sCardCode, "C", sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If

            sCostCenter = GetCostCenter(sCardCode)

            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sBatchNo = oDv(2)(0).ToString()
            sBatchPeriod = oDv(3)(0).ToString()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch No.", sFuncName)
            k = InStrRev(sBatchNo, ":")
            sBatchNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - k).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch No:" & sBatchNo, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch Date.", sFuncName)
            j = InStrRev(sBatchPeriod, "to")
            sBatchPeriod = Microsoft.VisualBasic.Right(sBatchPeriod, Len(sBatchPeriod) - j - 1).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Date: " & sBatchPeriod, sFuncName)

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            oDoc.DocDueDate = CDate(sBatchPeriod)
            oDoc.TaxDate = CDate(sBatchPeriod)
            oDoc.NumAtCard = oDv(2)(0).ToString()
            oDoc.Comments = sBatchPeriod
            oDoc.ImportFileNum = sBatchNo


            For i As Integer = 6 To oDv.Count - 1
                If oDv(i)(13).ToString = "FFS" Then
                    iCnt += 1
                    dGrossTotal = dGrossTotal + CDbl(oDv(i)(25))
                    dConsultCost = dConsultCost + CDbl(oDv(i)(17))
                    dDrugCost = dDrugCost + CDbl(oDv(i)(18))
                    dInhouseServiceCost = dInhouseServiceCost + CDbl(oDv(i)(19))
                    dExternalServiceCost = dExternalServiceCost + CDbl(oDv(i)(20))
                    dSubTotal = dSubTotal + CDbl(oDv(i)(21))
                    dTax = dTax + CDbl(oDv(i)(22))
                    dGrandTotal = dGrandTotal + CDbl(oDv(i)(23))
                    dUnclaimAmt = dUnclaimAmt + +CDbl(oDv(i)(24))
                End If
            Next

            oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCode
            oDoc.Lines.Quantity = iCnt
            oDoc.Lines.COGSCostingCode = sCostCenter

            'If bIsNonPanel = True Then
            '    If UCase(sCardName) = UCase("AIA Singapore Pte. Ltd.") Then
            '        oDoc.Lines.VatGroup = "ZO"
            '        oDoc.Lines.LineTotal = dGrossTotal
            '    Else
            '        oDoc.Lines.VatGroup = "SO"
            '        oDoc.Lines.LineTotal = dGrossTotal * 100 / 107
            '    End If
            'Else

            If dTax > 0 Then
                oDoc.Lines.VatGroup = "SO"
                oDoc.Lines.LineTotal = dGrossTotal * 100 / 107
            Else
                oDoc.Lines.VatGroup = "ZO"
                oDoc.Lines.LineTotal = dGrossTotal

            End If
            'End If

            oDoc.Lines.SerialNum = GetSeriesNum("FVM")
            oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConsultCost
            oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost
            oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInhouseServiceCost
            oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = dExternalServiceCost
            oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal
            oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dTax
            oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal
            oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnclaimAmt

            'If bIsNonPanel = True Then
            '    If UCase(sCardName) = UCase("AIA Singapore Pte. Ltd.") Then
            '        'Add Service Fee Item (Quantity: no. of lines , Price: Service Fee)
            '        If bServiceItem = True Then
            '            oDoc.Lines.Add()
            '            oDoc.Lines.ItemCode = p_oCompDef.sNonStockItem
            '            oDoc.Lines.Quantity = iCnt
            '            oDoc.Lines.Price = p_oCompDef.dServiceFee
            '            oDoc.Lines.COGSCostingCode = sCostCenter
            '        End If
            '    End If
            'End If

            ' Add Service Fee

            Dim oDtSVFee As DataTable
            oDtSVFee = oDv.Table.DefaultView.ToTable(True, "F27")

            For h As Integer = 0 To oDtSVFee.Rows.Count - 1
                With oDtSVFee.Rows(h)
                    If IsNumeric(.Item(0)) And h > 1 Then
                        If CDbl(.Item(0)) > 0 Then
                            Dim DepRows() As DataRow = oDv.Table.Select("F27='" & .Item(0).ToString.Trim & "'")
                            If DepRows.Length > 0 Then
                                oDoc.Lines.Add()
                                oDoc.Lines.ItemCode = p_oCompDef.sNonStockItem
                                oDoc.Lines.Quantity = DepRows.Length
                                oDoc.Lines.Price = CDbl(.Item(0))
                                oDoc.Lines.COGSCostingCode = sCostCenter
                            End If
                        End If
                    End If
                End With
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

            For i As Integer = 6 To oDv.Count - 1

                sSQL = "insert into ""AI_TB02_BILLTOCLIENT"" values(" & iARInvNo & ",'" & oDv(i)(0).ToString & "','" & CDate(oDv(i)(1)).ToString("yyyyMMdd") & "','" & Replace(oDv(i)(2).ToString, "'", "''") & "'" & _
                        " ,'" & Replace(oDv(i)(3).ToString, "'", "''") & "','" & oDv(i)(4).ToString & "','" & Replace(oDv(i)(5).ToString, "'", "''") & "','" & oDv(i)(6).ToString & "'" & _
                        " ,'" & oDv(i)(7).ToString & "','" & oDv(i)(8).ToString & "','" & oDv(i)(9).ToString & "','" & Replace(oDv(i)(10).ToString, "'", "''") & "'" & _
                        " ,'" & oDv(i)(11).ToString & "','" & oDv(i)(12).ToString & "','" & oDv(i)(13).ToString & "','" & oDv(i)(14).ToString & "'" & _
                        " ,'" & Replace(oDv(i)(15).ToString, "'", "''") & "','" & oDv(i)(16).ToString & "'," & CDbl(oDv(i)(17)) & "," & CDbl(oDv(i)(18)) & "," & CDbl(oDv(i)(19)) & "," & CDbl(oDv(i)(20)) & "," & CDbl(oDv(i)(21)) & _
                        "," & CDbl(oDv(i)(22)) & "," & CDbl(oDv(i)(23)) & "," & CDbl(oDv(i)(24)) & "," & CDbl(oDv(i)(25)) & "," & CDbl(oDv(i)(26)) & ",'" & oDv(i)(27).ToString & "'" & _
                        ",'" & p_sPatientType & "','" & Replace(oDv(3)(0).ToString(), "Batch Period : ", "") & "','" & sBatchNo.Trim & " ')"
                oRS.DoQuery(sSQL)

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice_Backup = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice_Backup = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Function AddSalesOrder(ByVal oDv As DataView, _
                                  ByVal sCardName As String, _
                                  ByVal bServiceItem As Boolean, _
                                  ByVal bIsNonPanel As Boolean, _
                                  ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim sRemarks As String = String.Empty
        Dim k, j As Integer
        Dim sCostCenter As String = String.Empty
        Dim dGrossTotal As Double = 0
        Dim dConsultCost As Double = 0
        Dim dDrugCost As Double = 0
        Dim dInhouseServiceCost As Double = 0
        Dim dExternalServiceCost As Double = 0
        Dim dSubTotal As Double = 0
        Dim dTax As Double = 0
        Dim dGrandTotal As Double = 0
        Dim dUnclaimAmt As Double = 0

        Try
            sFuncName = "AddSalesOrder"

            If CheckBP(sCardName, sCardCode, "C", sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If

            sCostCenter = GetCostCenter(sCardCode)

            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI SO Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

            sBatchNo = oDv(2)(0).ToString()
            sBatchPeriod = oDv(3)(0).ToString()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch No.", sFuncName)
            k = InStrRev(sBatchNo, ":")
            sBatchNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - k).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch No:" & sBatchNo, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch Date.", sFuncName)
            j = InStrRev(sBatchPeriod, "to")
            sBatchPeriod = Microsoft.VisualBasic.Right(sBatchPeriod, Len(sBatchPeriod) - j - 1).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Date: " & sBatchPeriod, sFuncName)

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            oDoc.DocDueDate = CDate(sBatchPeriod)
            oDoc.TaxDate = CDate(sBatchPeriod)
            oDoc.NumAtCard = oDv(2)(0).ToString()
            oDoc.Comments = sBatchPeriod
            oDoc.ImportFileNum = sBatchNo


            For i As Integer = 6 To oDv.Count - 1
                If oDv(i)(13).ToString = "FFS" Then
                    iCnt += 1
                    dGrossTotal = dGrossTotal + CDbl(oDv(i)(24))
                    dConsultCost = dConsultCost + CDbl(oDv(i)(16))
                    dDrugCost = dDrugCost + CDbl(oDv(i)(17))
                    dInhouseServiceCost = dInhouseServiceCost + CDbl(oDv(i)(18))
                    dExternalServiceCost = dExternalServiceCost + CDbl(oDv(i)(19))
                    dSubTotal = dSubTotal + CDbl(oDv(i)(20))
                    dTax = dTax + CDbl(oDv(i)(21))
                    dGrandTotal = dGrandTotal + CDbl(oDv(i)(22))
                    dUnclaimAmt = dUnclaimAmt + +CDbl(oDv(i)(23))
                End If
            Next

            For Each row As DataRow In oDv.Table.Rows
                iCnt += 1
                If iCnt = 6 Then
                    If iCnt > 7 Then
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
                    oDoc.Lines.PriceAfterVAT = CDbl(row.Item(24))

                    If bIsNonPanel = True Then
                        oDoc.Lines.VatGroup = "SO"
                    Else
                        If CDbl(row.Item(21)) > 0 Then
                            oDoc.Lines.VatGroup = "SO"
                        Else
                            oDoc.Lines.VatGroup = "ZO"
                        End If
                    End If

                    oDoc.Lines.SerialNum = GetSeriesNum("FSO")
                    oDoc.Lines.COGSCostingCode = sCostCenter
                    oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = row.Item(10).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = row.Item(1)
                    oDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value = row.Item(0).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = row.Item(11).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = CDbl(row.Item(16))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = CDbl(row.Item(17))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = CDbl(row.Item(18))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ExtServCost").Value = CDbl(row.Item(19))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = CDbl(row.Item(20))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = CDbl(row.Item(21))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = CDbl(row.Item(22))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = CDbl(row.Item(23))
                    oDoc.Lines.UserFields.Fields.Item("U_AI_CostCenter").Value = row.Item(7).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_CompanyName").Value = row.Item(2).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_ProviderName").Value = row.Item(14).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_EmpName").Value = row.Item(8).ToString
                    oDoc.Lines.UserFields.Fields.Item("U_AI_EmpID").Value = row.Item(9).ToString

                    oDoc.Lines.UserFields.Fields.Item("U_AI_Department").Value = row.Item(6).ToString

                    oDoc.UserFields.Fields.Item("U_AI_BillToDept").Value = row.Item(6).ToString
                    oDoc.UserFields.Fields.Item("U_AI_BillToCCenter").Value = row.Item(7).ToString


                End If

            Next



            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding ARI.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                p_oCompany.GetLastError(lErrCode, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding ARI failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If


            AddSalesOrder = RTN_SUCCESS

        Catch ex As Exception
            AddSalesOrder = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function


#End Region

#Region "ReimbursetoProvider"

    Private Function ProcessReimburse_Provider(ByVal sFileName As String, ByVal sSheet As String, ByVal oDv As DataView, ByRef sErrdesc As String) As Long

        Dim IsError As Boolean = False
        Dim sFuncName As String = "ProcessReimburse_Provider"

        Try
            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling Upload_ReimbursetoProvider_WorkSheet()", sFuncName)
            If Upload_ReimbursetoProvider_WorkSheet(oDv, sErrdesc) <> RTN_SUCCESS Then
                If RollBackTransaction(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Error in uploading Reimburse to Provider", sFuncName)
                Throw New ArgumentException(sErrdesc)
            End If

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Function completed successfully.", sFuncName)
            ProcessReimburse_Provider = RTN_SUCCESS

        Catch ex As Exception
            ProcessReimburse_Provider = RTN_ERROR
            If RollBackTransaction(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        End Try

    End Function

    Private Sub ReadReimburse_Provider(ByVal sFileName As String, _
                                 ByVal sSheet As String, _
                                 ByRef bIsError As Boolean, _
                                 ByRef dv As DataView)

        Dim iHeaderRow As Integer
        Dim sErrDesc As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "ReadReimburse_Provider"
        iHeaderRow = 5

        dv = GetDataViewFromExcel(sFileName, sSheet)

        If dv(iHeaderRow)(0).ToString.Trim <> "Visit No." Then
            sErrDesc = "Invalid Excel file Format - ([Visit No.] not found at Column 1"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(1).ToString.Trim <> "Visit Date" Then
            sErrDesc = "Invalid Excel file Format - ([Visit Date] not found at Column 2"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(2).ToString.Trim <> "Invoice No." Then
            sErrDesc = "Invalid Excel file Format - ([Invoice No.] not found at Column 3"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(3).ToString.Trim <> "Company Code" Then
            sErrDesc = "Invalid Excel file Format - ([Company Code] not found at Column 4"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString.Trim <> "Company Name" Then
            sErrDesc = "Invalid Excel file Format - ([Company Name] not found at Column 5"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If


        If dv(iHeaderRow)(5).ToString.Trim <> "Broker Name" Then
            sErrDesc = "Invalid Excel file Format - ([Broker Name] not found at Column 5"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If


        If dv(iHeaderRow)(6).ToString.Trim <> "Patient Name" Then
            sErrDesc = "Invalid Excel file Format - ([Patient Name] not found at Column 6"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(7).ToString.Trim <> "Patient ID No." Then
            sErrDesc = "Invalid Excel file Format - ([Patient ID No.] not found at Column 7"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString.Trim <> "Patient Member Type" Then
            sErrDesc = "Invalid Excel file Format - ([Patient Member Type] not found at Column 8"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(9).ToString.Trim <> "Contract Type" Then
            sErrDesc = "Invalid Excel file Format - ([Contract Type] not found at Column 9"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString.Trim <> "Benefit Type" Then
            sErrDesc = "Invalid Excel file Format - ([Benefit Type] not found at Column 10"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If


        If dv(iHeaderRow)(11).ToString.Trim <> "Provider Code" Then
            sErrDesc = "Invalid Excel file Format - ([Provider Code] not found at Column 11"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If


        If dv(iHeaderRow)(12).ToString.Trim <> "Provider Name" Then
            sErrDesc = "Invalid Excel file Format - ([Provider Name] not found at Column 13"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(13).ToString.Trim <> "Provider Address Line 1" Then
            sErrDesc = "Invalid Excel file Format - ([Provider Address Line 1] not found at Column 14"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(14).ToString.Trim <> "Provider Address Line 2" Then
            sErrDesc = "Invalid Excel file Format - ([Provider Address Line 2] not found at Column 15"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(15).ToString.Trim <> "Provider Address Line 3" Then
            sErrDesc = "Invalid Excel file Format - ([Provider Address Line 3] not found at Column 16"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(16).ToString.Trim <> "Provider Country" Then
            sErrDesc = "Invalid Excel file Format - ([Provider Country] not found at Column 17"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(17).ToString.Trim <> "Provider Postal Code" Then
            sErrDesc = "Invalid Excel file Format - ([Provider Postal Code] not found at Column 18"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(18).ToString.Trim <> "Mobile No." Then
            sErrDesc = "Invalid Excel file Format - ([Mobile No.] not found at Column 19"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(19).ToString.Trim <> "Email Address" Then
            sErrDesc = "Invalid Excel file Format - ([Email Address] not found at Column 20"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(20).ToString.Trim <> "Currency" Then
            sErrDesc = "Invalid Excel file Format - ([Currency] not found at Column 21"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(21).ToString.Trim <> "Consult Cost" Then
            sErrDesc = "Invalid Excel file Format - ([Consult Cost] not found at Column 22"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(22).ToString.Trim <> "Drug Cost" Then
            sErrDesc = "Invalid Excel file Format - ([Drug Cost] not found at Column 23"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(23).ToString.Trim <> "In-house Service Cost" Then
            sErrDesc = "Invalid Excel file Format - ([In-house Service Cost] not found at Column 24"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(24).ToString.Trim <> "Sub-total" Then
            sErrDesc = "Invalid Excel file Format - ([Sub-total] not found at Column 25"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(25).ToString.Trim <> "Tax" Then
            sErrDesc = "Invalid Excel file Format - ([Tax] not found at Column 26"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(26).ToString.Trim <> "Grand Total" Then
            sErrDesc = "Invalid Excel file Format - ([Grand Total] not found at Column 27"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(27).ToString.Trim <> "Unclaim Amt." Then
            sErrDesc = "Invalid Excel file Format - ([Unclaim Amt.] not found at Column 28"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(28).ToString.Trim <> "Claim Amt." Then
            sErrDesc = "Invalid Excel file Format - ([Claim Amt.] not found at Column 29"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(29).ToString.Trim <> "Fee 1" Then
            sErrDesc = "Invalid Excel file Format - ([Fee 1] not found at Column 30"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(30).ToString.Trim <> "Fee 2" Then
            sErrDesc = "Invalid Excel file Format - ([Fee 2] not found at Column 31"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(31).ToString.Trim <> "Fee 3" Then
            sErrDesc = "Invalid Excel file Format - ([Fee 3] not found at Column 32"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If


        If dv(iHeaderRow)(32).ToString.Trim <> "TPA Fee" Then
            sErrDesc = "Invalid Excel file Format - ([TPA Fee] not found at Column 33"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(33).ToString.Trim <> "TPA Fee Tax" Then
            sErrDesc = "Invalid Excel file Format - ([TPA Fee Tax] not found at Column 34"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(34).ToString.Trim <> "TPA Fee Total" Then
            sErrDesc = "Invalid Excel file Format - ([TPA Fee Total] not found at Column 35"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(35).ToString.Trim <> "Adjustment Amt." Then
            sErrDesc = "Invalid Excel file Format - ([Adjustment Amt.] not found at Column 36"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(36).ToString.Trim <> "Reimbursement Amt." Then
            sErrDesc = "Invalid Excel file Format - ([Reimbursement Amt.] not found at Column 37"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(37).ToString.Trim <> "Payment Mode" Then
            sErrDesc = "Invalid Excel file Format - ([Payment Mode] not found at Column 38"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(38).ToString.Trim <> "Payee Name" Then
            sErrDesc = "Invalid Excel file Format - ([Payee Name] not found at Column 39"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(39).ToString.Trim <> "Payee Account No." Then
            sErrDesc = "Invalid Excel file Format - ([Payee Account No.] not found at Column 40"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(40).ToString <> "Remarks for Provider" Then
            sErrDesc = "Invalid Excel file Format - ([Remarks for Provider] not found at Column 41"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

    End Sub

    Private Function Upload_ReimbursetoProvider_WorkSheet(ByVal oDv As DataView, _
                                                          ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim oDS As New DataSet
        Dim oDatasetBP As New DataSet
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim k, m, r As Integer
        Dim sCostCenter As String = String.Empty
        Dim oBPArrary As New ArrayList
        Dim sCompanyName As String = String.Empty

        Try
            sFuncName = "Upload_ReimbursetoProvider_WorkSheet"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function..", sFuncName)
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

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Contract Owner Name.", sFuncName)
            sCardName = oDv(0)(0).ToString()
            k = InStrRev(sCardName, ":")
            sCardName = Microsoft.VisualBasic.Right(sCardName, Len(sCardName) - k).Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Company Name.", sFuncName)
            sCompanyName = oDv(1)(0).ToString()
            r = InStrRev(sCompanyName, ":")
            sCompanyName = Microsoft.VisualBasic.Right(sCompanyName, Len(sCompanyName) - r).Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Card Name:" & sCardName, sFuncName)

            For i As Integer = 6 To oDv.Count - 1
                If Not oBPArrary.Contains(oDv(i)(12).ToString) Then oBPArrary.Add(oDv(i)(12).ToString)
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CheckBP()", sFuncName)
            If CheckBP(sCardName, sCardCode, "C", sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If

            Dim sName As String = "Fullerton Healthcare Group Pte Ltd"
            Dim sGJName As String = "Gethin-Jones Medical Practice Pte Ltd"

            If UCase(sCardName.Trim) = UCase(sName) Or UCase(sCardName.Trim) = UCase(sGJName) Then
                sCostCenter = GetCostCenterByCardName_MV(sCompanyName)
            Else
                sCostCenter = GetCostCenter(sCardCode)
            End If


            Dim oDoc As SAPbobsCOM.Documents = Nothing
            Dim oAPCRNDoc As SAPbobsCOM.Documents = Nothing

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("####################### Total No. of Providers :: " & oBPArrary.Count, sFuncName)
            Dim x As Integer = 0
            Dim iRow As Integer
            For iRow = 0 To oBPArrary.Count - 1
                If IsProviderBatchNoExists(oDv(2)(0).ToString(), oBPArrary(iRow).ToString) = False Then
                    x += 1
                    Dim BPRows() As DataRow = oDv.Table.Select("F13='" & Replace(oBPArrary(iRow).ToString, "'", "''") & "'")
                    sCardName = oBPArrary(iRow).ToString
                    If BPRows.Length > 0 Then
                        Dim iAPInvNum As Integer
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddAPInvoice()", sFuncName)
                        Console.WriteLine("Creating " & x & " out of " & oBPArrary.Count & " AP Invoice for:: " & sCardName)
                        Console.WriteLine("Total Rows ::" & BPRows.Length)

                        iAPInvNum = 0
                        If AddAPInvoice_UDT(oDoc, sCardName, BPRows, sBatchNo, sBatchPeriod, sCostCenter, iAPInvNum, sErrDesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrDesc)
                        End If

                        Console.WriteLine("Successfully created " & x & " out of " & oBPArrary.Count & " AP Invoice for:: " & sCardName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AP Invoice created Succesfully. : " & iAPInvNum, sFuncName)

                        Console.WriteLine("Creating " & x & " out of " & oBPArrary.Count & " AP Credit Memo for:: " & sCardName)

                        If iAPInvNum > 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddAPCreditMemo_TPA()", sFuncName)
                            If AddAPCreditMemo_TPA(oDoc, oAPCRNDoc, iAPInvNum, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        End If

                        Console.WriteLine("Successfully created " & x & " out of " & oBPArrary.Count & " AP Invoice for:: " & sCardName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AP CreditMemo created Succesfully for Provider::" & sCardName, sFuncName)

                    End If
                    GC.Collect()
                End If
            Next


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            Upload_ReimbursetoProvider_WorkSheet = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            Upload_ReimbursetoProvider_WorkSheet = RTN_ERROR

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        Finally

        End Try
    End Function

    Private Function AddAPInvoice_UDT(ByVal oDoc As SAPbobsCOM.Documents, _
                                ByVal sCardName As String, _
                                ByVal oBPRows() As DataRow, _
                                ByVal sBatchNo As String, _
                                ByVal sBatchPeriod As String, _
                                ByVal sCostCenter As String, _
                                ByRef iAPInvNo As Integer, _
                                ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet
        Dim bIsPymtNoBlank As Boolean = False
        Dim iCode As Integer = 0
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim dTotalAmt As Double = 0
        Dim dConnsultCost As Double = 0
        Dim dDrugCost As Double = 0
        Dim dInHouseServCost As Double = 0
        Dim dSubTotal As Double = 0
        Dim dGSTAmt As Double = 0
        Dim dGrandTotal As Double = 0
        Dim dUnClaimAmt As Double = 0
        Dim dTPAFee As Double = 0
        Dim dTPAFeeTax As Double = 0
        Dim dTPAFeeTotal As Double = 0

        Dim sFee1 As String = String.Empty
        Dim sFee2 As String = String.Empty
        Dim sFee3 As String = String.Empty

        Dim dTotalAmt_CAP As Double = 0
        Dim dConnsultCost_CAP As Double = 0
        Dim dDrugCost_CAP As Double = 0
        Dim dInHouseServCost_CAP As Double = 0
        Dim dSubTotal_CAP As Double = 0
        Dim dGSTAmt_CAP As Double = 0
        Dim dGrandTotal_CAP As Double = 0
        Dim dUnClaimAmt_CAP As Double = 0
        Dim dTPAFee_CAP As Double = 0
        Dim dTPAFeeTax_CAP As Double = 0
        Dim dTPAFeeTotal_CAP As Double = 0
        Dim iCnt_CAP As Integer = 0
        Dim bRowsCAP As Boolean = False
        Dim bRowsFFS As Boolean = False

        Dim sFee1_CAP As String = String.Empty
        Dim sFee2_CAP As String = String.Empty
        Dim sFee3_CAP As String = String.Empty

        Dim dTotalAmt_3FS As Double = 0
        Dim dConnsultCost_3FS As Double = 0
        Dim dDrugCost_3FS As Double = 0
        Dim dInHouseServCost_3FS As Double = 0
        Dim dSubTotal_3FS As Double = 0
        Dim dGSTAmt_3FS As Double = 0
        Dim dGrandTotal_3FS As Double = 0
        Dim dUnClaimAmt_3FS As Double = 0
        Dim dTPAFee_3FS As Double = 0
        Dim dTPAFeeTax_3FS As Double = 0
        Dim dTPAFeeTotal_3FS As Double = 0
        Dim iCnt_3FS As Integer = 0
        Dim bRows3FS As Boolean = False

        Dim sFee1_3FS As String = String.Empty
        Dim sFee2_3FS As String = String.Empty
        Dim sFee3_3FS As String = String.Empty

        Try

            sFuncName = "AddAPInvoice_UDT"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Staring Function..", sFuncName)

            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/P Invoice for Provider :: " & sCardName, sFuncName)

            If CheckBP(sCardName, sCardCode, "S", sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            'oDoc.DocDueDate = CDate(sBatchPeriod)
            oDoc.TaxDate = CDate(sBatchPeriod)
            oDoc.NumAtCard = "Batch No. : " & sBatchNo
            oDoc.Comments = sBatchPeriod
            oDoc.ImportFileNum = sBatchNo


            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter
            bIsPymtNoBlank = False

            For Each row As DataRow In oBPRows
                If row.Item(9).ToString = "FFS" Then    'Contract Type
                    iCnt += 1
                    bRowsFFS = True
                    dTotalAmt = dTotalAmt + CDbl(row.Item(28))     'Claim Amount
                    dConnsultCost = Math.Round(dConnsultCost, 4) + Math.Round(CDbl(row.Item(21)), 4)    'Consult Cost
                    dDrugCost = dDrugCost + CDbl(row.Item(22))  ' Drug Cost
                    dInHouseServCost = dInHouseServCost + +CDbl(row.Item(23))   'InHouse Service cost
                    dSubTotal = dSubTotal + CDbl(row.Item(24))  ' Subtotal
                    dGSTAmt = dGSTAmt + CDbl(row.Item(25))  ' Tax
                    dGrandTotal = dGrandTotal + CDbl(row.Item(26))  ' GrandTotal
                    dUnClaimAmt = dUnClaimAmt + CDbl(row.Item(27))  ' Un ClaimAmt

                    sFee1 = row.Item(29).ToString  'Fee1
                    sFee2 = row.Item(30).ToString  'Fee2
                    sFee3 = row.Item(31).ToString  'Fee3

                    dTPAFee = dTPAFee + CDbl(row.Item(32)) ' TPA Fee
                    dTPAFeeTax = dTPAFeeTax + CDbl(row.Item(33)) ' TPA Fee Tax
                    dTPAFeeTotal = dTPAFeeTotal + CDbl(row.Item(34)) ' TPA Fee Total

                    If row.Item(36).ToString = String.Empty Then bIsPymtNoBlank = True 'Payee Account No

                ElseIf row.Item(9).ToString = "CAP" Then     'Contract Type
                    iCnt_CAP += 1
                    bRowsCAP = True
                    dTotalAmt_CAP = dTotalAmt_CAP + CDbl(row.Item(28))  'Claim Amount
                    dConnsultCost_CAP = dConnsultCost_CAP + CDbl(row.Item(21))  'Consult Cost
                    dDrugCost_CAP = dDrugCost_CAP + CDbl(row.Item(22))  ' Drug Cost
                    dInHouseServCost_CAP = dInHouseServCost_CAP + CDbl(row.Item(23))     'InHouse Service cost
                    dSubTotal_CAP = dSubTotal_CAP + CDbl(row.Item(24))  ' Subtotal
                    dGSTAmt_CAP = dGSTAmt_CAP + CDbl(row.Item(25))  'Tax
                    dGrandTotal_CAP = dGrandTotal_CAP + CDbl(row.Item(26))  'Grand Total
                    dUnClaimAmt_CAP = dUnClaimAmt_CAP + CDbl(row.Item(27))  'Un-Claim Amt

                    sFee1_CAP = row.Item(29).ToString  'Fee1
                    sFee2_CAP = row.Item(30).ToString  'Fee2
                    sFee3_CAP = row.Item(31).ToString  'Fee3

                    dTPAFee_CAP = dTPAFee_CAP + CDbl(row.Item(32)) ' TPA Fee
                    dTPAFeeTax_CAP = dTPAFeeTax_CAP + CDbl(row.Item(33)) 'TPA Fee Tax
                    dTPAFeeTotal_CAP = dTPAFeeTotal_CAP + CDbl(row.Item(34)) 'TPA Fee Total


                ElseIf row.Item(9).ToString = "3FS" Then
                    iCnt_3FS += 1
                    bRows3FS = True
                    dTotalAmt_3FS = dTotalAmt_3FS + CDbl(row.Item(28))  'Claim Amount
                    dConnsultCost_3FS = dConnsultCost_3FS + CDbl(row.Item(21))  'Consult Cost
                    dDrugCost_3FS = dDrugCost_3FS + CDbl(row.Item(22))  ' Drug Cost
                    dInHouseServCost_3FS = dInHouseServCost_3FS + CDbl(row.Item(23))     'InHouse Service cost
                    dSubTotal_3FS = dSubTotal_3FS + CDbl(row.Item(24))  ' Subtotal
                    dGSTAmt_3FS = dGSTAmt_3FS + CDbl(row.Item(25))  'Tax
                    dGrandTotal_3FS = dGrandTotal_3FS + CDbl(row.Item(26))  'Grand Total
                    dUnClaimAmt_3FS = dUnClaimAmt_3FS + CDbl(row.Item(27))  'Un-Claim Amt

                    sFee1_3FS = row.Item(29).ToString  'Fee1
                    sFee2_3FS = row.Item(30).ToString  'Fee2
                    sFee3_3FS = row.Item(31).ToString  'Fee3

                    dTPAFee_3FS = dTPAFee_3FS + CDbl(row.Item(32)) ' TPA Fee

                    dTPAFeeTax_3FS = dTPAFeeTax_3FS + CDbl(row.Item(33)) 'TPA Fee Tax
                    dTPAFeeTotal_3FS = dTPAFeeTotal_3FS + CDbl(row.Item(34)) 'TPA Fee Total
                Else
                    sErrDesc = "No Contract type:: " & row.Item(8).ToString & " found in SAP. Please check the contract type."
                    Throw New ArgumentException(sErrDesc)
                End If
            Next

            '***************  To check Zero Amount Claims ( Create A/P Invoice with Zero Amount )

            If dTotalAmt <= 0 And dTotalAmt_CAP = 0 And dTotalAmt_3FS = 0 Then
                Dim iInvNo As Integer
                If AddAPInvoice_ZeroAmount(oDoc, sCardName, oBPRows, sBatchNo, sBatchPeriod, sCostCenter, iInvNo, sErrDesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrDesc)
                End If

                If iInvNo > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddAPCreditMemo_TPA_ZeroClaimAmt()", sFuncName)
                    If AddAPCreditMemo_TPA_ZeroClaimAmt(oDoc, iInvNo, sErrDesc) <> RTN_SUCCESS Then
                        Throw New ArgumentException(sErrDesc)
                    End If
                End If
                GoTo ExitFunc
            End If


            '***************** End

            Dim dAddExtraLine As Boolean = False

            If dTotalAmt > 0 Or bRowsFFS = True Then
                oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCode
                oDoc.Lines.Quantity = 1
                oDoc.Lines.UserFields.Fields.Item("U_AI_NoOfVisits").Value = iCnt

                oDoc.Lines.COGSCostingCode = sCostCenter

                If dGSTAmt > 0 Then
                    oDoc.Lines.VatGroup = "SI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt
                Else
                    oDoc.Lines.VatGroup = "ZI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt
                End If

                If bIsPymtNoBlank = True Then
                    oDoc.PaymentMethod = "Check"
                End If

                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConnsultCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInHouseServCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dGSTAmt
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnClaimAmt
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value = dTPAFee
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value = dTPAFeeTax
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value = dTPAFeeTotal

                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value = sFee1
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value = sFee2
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value = sFee3

                dAddExtraLine = True
            End If

            ' CAP
            If dAddExtraLine = True And (dTotalAmt_CAP > 0 Or bRowsCAP = True) Then
                oDoc.Lines.Add()
            End If

            If dTotalAmt_CAP > 0 Or bRowsCAP = True Then
                oDoc.Lines.ItemCode = p_oCompDef.sCAPItemCode
                oDoc.Lines.Quantity = 1
                oDoc.Lines.UserFields.Fields.Item("U_AI_NoOfVisits").Value = iCnt_CAP

                oDoc.Lines.COGSCostingCode = sCostCenter
                If dGSTAmt_CAP > 0 Then
                    oDoc.Lines.VatGroup = "SI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt_CAP
                Else
                    oDoc.Lines.VatGroup = "ZI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt_CAP
                End If

                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConnsultCost_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInHouseServCost_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dGSTAmt_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnClaimAmt_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value = dTPAFee_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value = dTPAFeeTax_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value = dTPAFeeTotal_CAP

                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value = sFee1_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value = sFee2_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value = sFee3_CAP
                dAddExtraLine = True
            End If

            '3FS
            If dAddExtraLine = True And (dTotalAmt_3FS > 0 Or bRows3FS = True) Then
                oDoc.Lines.Add()
            End If

            If dTotalAmt_3FS > 0 Or bRows3FS = True Then
                oDoc.Lines.ItemCode = p_oCompDef.s3FSItemCode
                oDoc.Lines.Quantity = 1
                oDoc.Lines.UserFields.Fields.Item("U_AI_NoOfVisits").Value = iCnt_3FS

                oDoc.Lines.COGSCostingCode = sCostCenter
                If dGSTAmt_3FS > 0 Then
                    oDoc.Lines.VatGroup = "SI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt_3FS
                Else
                    oDoc.Lines.VatGroup = "ZI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt_3FS
                End If

                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConnsultCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInHouseServCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dGSTAmt_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnClaimAmt_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value = dTPAFee_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value = dTPAFeeTax_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value = dTPAFeeTotal_3FS

                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value = sFee1_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value = sFee2_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value = sFee3_3FS

                dAddExtraLine = True
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AP Inovice to SAP.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                p_oCompany.GetLastError(lErrCode, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding API failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            p_oCompany.GetNewObjectCode(iAPInvNo)
            AddDataToTable(p_oDtReport, "AP", iAPInvNo, sCardCode, String.Empty)

            ' CHANGE: ClaimAmt Column insert, changed from Reimbursement amt(29) to ClaimAmt(25) 11/06/2014
            ' Change SIA ( Excel template columns order changed)
            '1. DocEntry
            '2. U_AI_PatientName (6)
            '3. U_AI_VisitDate (1)
            '4. U_AI_ProviderName (12)
            '5. U_AI_VisitNo (0)
            '6. U_AI_ICNo (7)
            '7. U_AI_ConsultCost (21)
            '8. U_AI_DrugCost (22)
            '9. U_AI_InHouseServCost (23)
            '10. U_AI_SubTotal (24)
            '11. U_AI_GSTAmount (25)
            '12. U_AI_UncliamedAmount (27)
            '13. U_AI_GrandTotal (26)
            '14. U_AI_TPAFee (29)
            '15. U_AI_TPAFeeTax (30)
            '16. U_AI_TPAFeeTotal (31)
            '17. U_AI_ClaimAmt (28)
            '18. BatchNo
            '19. U_AI_Fee1
            '20. U_AI_Fee2
            '21. U_AI_Fee3

            'Update FHN3 Provider Code in BP Master..
            Dim sQuery As String = "UPDATE OCRD SET ""U_AI_FHN3Code""='" & oBPRows(0).Item(11).ToString & "' WHERE ""CardCode""='" & sCardCode & "'"
            Console.WriteLine(sQuery)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sQuery, sFuncName)
            oRS.DoQuery(sQuery)

            For Each row As DataRow In oBPRows
                sSQL = "insert into ""AI_TB01_PROVIDERS"" values(" & iAPInvNo & ",'" & Replace(row.Item(6).ToString, "'", "''") & "'" & _
                        " ,'" & CDate(row.Item(1)).ToString("yyyyMMdd") & "','" & Replace(sCardName, "'", "''") & "','" & row.Item(0) & "','" & row.Item(7) & "'" & _
                        "," & row.Item(21) & "," & row.Item(22) & "," & row.Item(23) & "," & row.Item(24) & "," & row.Item(25) & _
                        "," & row.Item(27) & "," & row.Item(26) & "," & row.Item(32) & "," & row.Item(33) & "," & row.Item(34) & "," & row.Item(28) & ",'" & sBatchNo & "','" & row.Item(29).ToString & "','" & row.Item(30).ToString & "','" & row.Item(31).ToString & "')"
                oRS.DoQuery(sSQL)
            Next

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc)
ExitFunc:
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully created A/P Invoice for Provider :: " & sCardName, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddAPInvoice_UDT = RTN_SUCCESS

        Catch ex As Exception
            AddAPInvoice_UDT = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Failed to create A/P Invoice for Provider :: " & sCardName, sFuncName)
        Finally
            GC.Collect()
        End Try
    End Function

    Private Function AddAPInvoice_ZeroAmount(ByVal oDoc As SAPbobsCOM.Documents, _
                               ByVal sCardName As String, _
                               ByVal oBPRows() As DataRow, _
                               ByVal sBatchNo As String, _
                               ByVal sBatchPeriod As String, _
                               ByVal sCostCenter As String, _
                               ByRef iAPInvNo As Integer, _
                               ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet
        Dim bIsPymtNoBlank As Boolean = False
        Dim iCode As Integer = 0
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim dTotalAmt As Double = 0
        Dim dConnsultCost As Double = 0
        Dim dDrugCost As Double = 0
        Dim dInHouseServCost As Double = 0
        Dim dSubTotal As Double = 0
        Dim dGSTAmt As Double = 0
        Dim dGrandTotal As Double = 0
        Dim dUnClaimAmt As Double = 0
        Dim dTPAFee As Double = 0
        Dim dTPAFeeTax As Double = 0
        Dim dTPAFeeTotal As Double = 0
        Dim sFee1 As String = String.Empty
        Dim sFee2 As String = String.Empty
        Dim sFee3 As String = String.Empty

        Dim dTotalAmt_CAP As Double = 0
        Dim dConnsultCost_CAP As Double = 0
        Dim dDrugCost_CAP As Double = 0
        Dim dInHouseServCost_CAP As Double = 0
        Dim dSubTotal_CAP As Double = 0
        Dim dGSTAmt_CAP As Double = 0
        Dim dGrandTotal_CAP As Double = 0
        Dim dUnClaimAmt_CAP As Double = 0
        Dim dTPAFee_CAP As Double = 0
        Dim dTPAFeeTax_CAP As Double = 0
        Dim dTPAFeeTotal_CAP As Double = 0
        Dim dFee1_CAP As Double = 0
        Dim dFee2_CAP As Double = 0
        Dim dFee3_CAP As Double = 0
        Dim iCnt_CAP As Integer = 0
        Dim dCAP, dFFS, d3FS As Boolean
        Dim sFee1_CAP As String = String.Empty
        Dim sFee2_CAP As String = String.Empty
        Dim sFee3_CAP As String = String.Empty


        Dim dTotalAmt_3FS As Double = 0
        Dim dConnsultCost_3FS As Double = 0
        Dim dDrugCost_3FS As Double = 0
        Dim dInHouseServCost_3FS As Double = 0
        Dim dSubTotal_3FS As Double = 0
        Dim dGSTAmt_3FS As Double = 0
        Dim dGrandTotal_3FS As Double = 0
        Dim dUnClaimAmt_3FS As Double = 0
        Dim dTPAFee_3FS As Double = 0
        Dim dTPAFeeTax_3FS As Double = 0
        Dim dTPAFeeTotal_3FS As Double = 0
        Dim dFee1_3FS As Double = 0
        Dim dFee2_3FS As Double = 0
        Dim dFee3_3FS As Double = 0
        Dim iCnt_3FS As Integer = 0
        Dim bRows3FS As Boolean = False

        Dim sFee1_3FS As String = String.Empty
        Dim sFee2_3FS As String = String.Empty
        Dim sFee3_3FS As String = String.Empty

        Try

            sFuncName = "AddAPInvoice_ZeroAmount"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Staring Function..", sFuncName)

            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/P Invoice for Provider :: " & sCardName, sFuncName)

            'oDoc.CardCode = sCardCode
            'oDoc.DocDate = CDate(sBatchPeriod)
            'oDoc.DocDueDate = CDate(sBatchPeriod)
            'oDoc.TaxDate = CDate(sBatchPeriod)
            'oDoc.NumAtCard = "Batch No. : " & sBatchNo
            'oDoc.Comments = sBatchPeriod
            'oDoc.ImportFileNum = sBatchNo


            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter
            bIsPymtNoBlank = False

            dFFS = False
            dCAP = False
            d3FS = False

            For Each row As DataRow In oBPRows
                If row.Item(9).ToString = "FFS" Then
                    iCnt += 1
                    'dTotalAmt = Math.Round(dTotalAmt, 4) + Math.Round(CDbl(row.Item(24)), 4)
                    dTotalAmt = dTotalAmt + CDbl(row.Item(28))
                    dConnsultCost = Math.Round(dConnsultCost, 4) + Math.Round(CDbl(row.Item(21)), 4)
                    dDrugCost = dDrugCost + CDbl(row.Item(22))
                    dInHouseServCost = dInHouseServCost + +CDbl(row.Item(23))
                    dSubTotal = dSubTotal + CDbl(row.Item(24))
                    dGSTAmt = dGSTAmt + CDbl(row.Item(25))
                    dGrandTotal = dGrandTotal + CDbl(row.Item(26))
                    dUnClaimAmt = dUnClaimAmt + CDbl(row.Item(28))

                    'dFee1 = dFee1 + CDbl(row.Item(29))  'Fee1
                    'dFee2 = dFee2 + CDbl(row.Item(30))  'Fee2
                    'dFee3 = dFee3 + CDbl(row.Item(31))  'Fee3

                    sFee1 = row.Item(29).ToString  'Fee1
                    sFee2 = row.Item(30).ToString  'Fee2
                    sFee3 = row.Item(31).ToString  'Fee3


                    dTPAFee = dTPAFee + CDbl(row.Item(32))
                    dTPAFeeTax = dTPAFeeTax + CDbl(row.Item(33))
                    dTPAFeeTotal = dTPAFeeTotal + CDbl(row.Item(34))

                    If row.Item(36).ToString = String.Empty Then bIsPymtNoBlank = True
                    dFFS = True

                ElseIf row.Item(9).ToString = "CAP" Then
                    dCAP = True
                    iCnt_CAP += 1
                    dTotalAmt_CAP = dTotalAmt_CAP + CDbl(row.Item(28))
                    dConnsultCost_CAP = dConnsultCost_CAP + CDbl(row.Item(21))
                    dDrugCost_CAP = dDrugCost_CAP + CDbl(row.Item(22))
                    dInHouseServCost_CAP = dInHouseServCost_CAP + CDbl(row.Item(23))
                    dSubTotal_CAP = dSubTotal_CAP + CDbl(row.Item(24))
                    dGSTAmt_CAP = dGSTAmt_CAP + CDbl(row.Item(25))
                    dGrandTotal_CAP = dGrandTotal_CAP + CDbl(row.Item(26))
                    dUnClaimAmt_CAP = dUnClaimAmt_CAP + CDbl(row.Item(27))

                    'dFee1_CAP = dFee1_CAP + CDbl(row.Item(29))  'Fee1
                    'dFee2_CAP = dFee2_CAP + CDbl(row.Item(30))  'Fee2
                    'dFee3_CAP = dFee3_CAP + CDbl(row.Item(31))  'Fee3

                    sFee1_CAP = row.Item(29).ToString  'Fee1
                    sFee2_CAP = row.Item(30).ToString  'Fee2
                    sFee3_CAP = row.Item(31).ToString  'Fee3


                    dTPAFee_CAP = dTPAFee_CAP + CDbl(row.Item(32))
                    dTPAFeeTax_CAP = dTPAFeeTax_CAP + CDbl(row.Item(33))
                    dTPAFeeTotal_CAP = dTPAFeeTotal_CAP + CDbl(row.Item(34))

                ElseIf row.Item(9).ToString = "3FS" Then
                    d3FS = True
                    iCnt_3FS += 1
                    dTotalAmt_3FS = dTotalAmt_3FS + CDbl(row.Item(28))
                    dConnsultCost_3FS = dConnsultCost_3FS + CDbl(row.Item(21))
                    dDrugCost_3FS = dDrugCost_3FS + CDbl(row.Item(22))
                    dInHouseServCost_3FS = dInHouseServCost_3FS + CDbl(row.Item(23))
                    dSubTotal_3FS = dSubTotal_3FS + CDbl(row.Item(24))
                    dGSTAmt_3FS = dGSTAmt_3FS + CDbl(row.Item(25))
                    dGrandTotal_3FS = dGrandTotal_3FS + CDbl(row.Item(26))
                    dUnClaimAmt_3FS = dUnClaimAmt_3FS + CDbl(row.Item(27))

                    'dFee1_3FS = dFee1_3FS + CDbl(row.Item(29))  'Fee1
                    'dFee2_3FS = dFee2_3FS + CDbl(row.Item(30))  'Fee2
                    'dFee3_3FS = dFee3_3FS + CDbl(row.Item(31))  'Fee3

                    sFee1_3FS = row.Item(29).ToString  'Fee1
                    sFee2_3FS = row.Item(30).ToString  'Fee2
                    sFee3_3FS = row.Item(31).ToString  'Fee3


                    dTPAFee_3FS = dTPAFee_3FS + CDbl(row.Item(32))
                    dTPAFeeTax_3FS = dTPAFeeTax_3FS + CDbl(row.Item(33))
                    dTPAFeeTotal_3FS = dTPAFeeTotal_3FS + CDbl(row.Item(34))
                Else
                    sErrDesc = "No Contract type:: " & row.Item(9).ToString & " found in SAP. Please check the contract type."
                    Throw New ArgumentException(sErrDesc)
                End If
            Next

            Dim dAddExtraLine As Boolean = False

            If dFFS = True Then

                oDoc.Lines.ItemCode = p_oCompDef.sFFSItemCode
                oDoc.Lines.Quantity = 1
                oDoc.Lines.UserFields.Fields.Item("U_AI_NoOfVisits").Value = iCnt

                oDoc.Lines.COGSCostingCode = sCostCenter

                If dGSTAmt > 0 Then
                    oDoc.Lines.VatGroup = "SI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt
                Else
                    oDoc.Lines.VatGroup = "ZI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt
                End If

                If bIsPymtNoBlank = True Then
                    oDoc.PaymentMethod = "Check"
                End If

                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConnsultCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInHouseServCost
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dGSTAmt
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnClaimAmt

                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value = sFee1
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value = sFee2
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value = sFee3

                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value = dTPAFee
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value = dTPAFeeTax
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value = dTPAFeeTotal
                dAddExtraLine = True
            End If

            If dAddExtraLine = True And dCAP = True Then
                oDoc.Lines.Add()
            End If


            If dCAP = True Then
                oDoc.Lines.ItemCode = p_oCompDef.sCAPItemCode
                oDoc.Lines.Quantity = 1
                oDoc.Lines.UserFields.Fields.Item("U_AI_NoOfVisits").Value = iCnt_CAP

                oDoc.Lines.COGSCostingCode = sCostCenter
                If dGSTAmt_CAP > 0 Then
                    oDoc.Lines.VatGroup = "SI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt_CAP
                Else
                    oDoc.Lines.VatGroup = "ZI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt_CAP
                End If

                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConnsultCost_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInHouseServCost_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dGSTAmt_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnClaimAmt_CAP

                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value = sFee1_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value = sFee2_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value = sFee3_CAP

                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value = dTPAFee_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value = dTPAFeeTax_CAP
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value = dTPAFeeTotal_CAP
                dAddExtraLine = True
            End If

            If dAddExtraLine = True And d3FS = True Then
                oDoc.Lines.Add()
            End If

            If d3FS = True Then
                oDoc.Lines.ItemCode = p_oCompDef.s3FSItemCode
                oDoc.Lines.Quantity = 1
                oDoc.Lines.UserFields.Fields.Item("U_AI_NoOfVisits").Value = iCnt_3FS

                oDoc.Lines.COGSCostingCode = sCostCenter
                If dGSTAmt_3FS > 0 Then
                    oDoc.Lines.VatGroup = "SI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt_3FS
                Else
                    oDoc.Lines.VatGroup = "ZI"
                    oDoc.Lines.PriceAfterVAT = dTotalAmt_3FS
                End If

                oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = dConnsultCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = dDrugCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_InHouseServCost").Value = dInHouseServCost_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = dSubTotal_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = dGSTAmt_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = dGrandTotal_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_UncliamedAmount").Value = dUnClaimAmt_3FS

                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value = sFee1_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value = sFee2_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value = sFee3_3FS

                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value = dTPAFee_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value = dTPAFeeTax_3FS
                oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value = dTPAFeeTotal_3FS
                dAddExtraLine = True
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AP Inovice to SAP.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                p_oCompany.GetLastError(lErrCode, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding API failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            p_oCompany.GetNewObjectCode(iAPInvNo)

            ' CHANGE: ClaimAmt Column insert, changed from Reimbursement amt(29) to ClaimAmt(25) 11/06/2014
            'For Each row As DataRow In oBPRows
            '    sSQL = "insert into ""AI_TB01_PROVIDERS"" values(" & iAPInvNo & ",'" & Replace(row.Item(6).ToString, "'", "''") & "'" & _
            '            " ,'" & CDate(row.Item(1)).ToString("yyyyMMdd") & "','" & Replace(sCardName, "'", "''") & "','" & row.Item(0) & "','" & row.Item(6) & "'" & _
            '            "," & row.Item(19) & "," & row.Item(20) & "," & row.Item(21) & "," & row.Item(22) & "," & row.Item(23) & _
            '            "," & row.Item(25) & "," & row.Item(24) & "," & row.Item(27) & "," & row.Item(28) & "," & row.Item(29) & "," & row.Item(26) & ",'" & sBatchNo & "')"
            '    oRS.DoQuery(sSQL)
            '    'ExecuteSQLNonQuery(sSQL)
            'Next

            For Each row As DataRow In oBPRows
                'sSQL = "insert into ""AI_TB01_PROVIDERS"" values(" & iAPInvNo & ",'" & Replace(row.Item(6).ToString, "'", "''") & "'" & _
                '         " ,'" & CDate(row.Item(1)).ToString("yyyyMMdd") & "','" & Replace(sCardName, "'", "''") & "','" & row.Item(0) & "','" & row.Item(7) & "'" & _
                '         "," & row.Item(21) & "," & row.Item(22) & "," & row.Item(23) & "," & row.Item(24) & "," & row.Item(25) & _
                '         "," & row.Item(27) & "," & row.Item(26) & "," & row.Item(29) & "," & row.Item(30) & "," & row.Item(31) & "," & row.Item(28) & ",'" & sBatchNo & "," & row.Item(29) & "," & row.Item(30) & "," & row.Item(31) & "')"

                sSQL = "insert into ""AI_TB01_PROVIDERS"" values(" & iAPInvNo & ",'" & Replace(row.Item(6).ToString, "'", "''") & "'" & _
                        " ,'" & CDate(row.Item(1)).ToString("yyyyMMdd") & "','" & Replace(sCardName, "'", "''") & "','" & row.Item(0) & "','" & row.Item(7) & "'" & _
                        "," & row.Item(21) & "," & row.Item(22) & "," & row.Item(23) & "," & row.Item(24) & "," & row.Item(25) & _
                        "," & row.Item(27) & "," & row.Item(26) & "," & row.Item(32) & "," & row.Item(33) & "," & row.Item(34) & "," & row.Item(28) & ",'" & sBatchNo & "','" & row.Item(29).ToString & "','" & row.Item(30).ToString & "','" & row.Item(31).ToString & "')"

                oRS.DoQuery(sSQL)
                'ExecuteSQLNonQuery(sSQL)
            Next


            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully created A/P Invoice for Provider :: " & sCardName, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddAPInvoice_ZeroAmount = RTN_SUCCESS

        Catch ex As Exception
            AddAPInvoice_ZeroAmount = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Failed to create A/P Invoice for Provider :: " & sCardName, sFuncName)
        Finally
            GC.Collect()
        End Try
    End Function

    Private Function AddAPCreditMemo_TPA(ByVal oDoc As SAPbobsCOM.Documents, _
                                        ByVal oAPCRNDoc As SAPbobsCOM.Documents, _
                                        ByVal iAPInvNo As Integer, _
                                        ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        'Dim oDoc As SAPbobsCOM.Document
        'Dim oAPCRNDoc As SAPbobsCOM.Documents
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet
        Dim dGrandTotal As Double = 0
        Dim bAddDoc As Boolean = False
        Try

            sFuncName = "AddAPCreditMemo_TPA"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI APCRN Object", sFuncName)

            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
            oAPCRNDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)


            If oDoc.GetByKey(iAPInvNo) = True Then
                oAPCRNDoc.CardCode = oDoc.CardCode
                oAPCRNDoc.DocDate = oDoc.DocDate
                oAPCRNDoc.DocDueDate = oDoc.DocDueDate
                oAPCRNDoc.TaxDate = oDoc.TaxDate
                oAPCRNDoc.NumAtCard = oDoc.NumAtCard
                oAPCRNDoc.Comments = oDoc.Comments
                oAPCRNDoc.ImportFileNum = oDoc.ImportFileNum

                bAddDoc = False
                For i As Integer = 0 To oDoc.Lines.Count - 1
                    oDoc.Lines.SetCurrentLine(i)
                    'Check if Line has TPA Fee Amount
                    If oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value > 0 Then
                        bAddDoc = True
                        iCnt += 1
                        If iCnt > 1 Then
                            oAPCRNDoc.Lines.Add()
                        End If

                        'oAPCRNDoc.Lines.SetCurrentLine(i)

                        oAPCRNDoc.Lines.SetCurrentLine(iCnt - 1)

                        oDoc.Lines.SetCurrentLine(i)

                        oAPCRNDoc.Lines.ItemCode = oDoc.Lines.ItemCode
                        oAPCRNDoc.Lines.Quantity = 1

                        oAPCRNDoc.Lines.VatGroup = "ZI" 'NO GST..
                        oAPCRNDoc.Lines.TaxTotal = 0
                        oAPCRNDoc.Lines.PriceAfterVAT = (oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value) * 100 / 107


                        oAPCRNDoc.Lines.BaseEntry = oDoc.DocEntry
                        oAPCRNDoc.Lines.BaseLine = oDoc.Lines.LineNum
                        oAPCRNDoc.Lines.BaseType = 18

                        oAPCRNDoc.Lines.COGSCostingCode = oDoc.Lines.COGSCostingCode

                        oDoc.Lines.SerialNum = GetSeriesNum("FPM")

                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_PayeeAcNo").Value

                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value

                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value
                        oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value

                        dGrandTotal = dGrandTotal + CDbl(oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value)
                    End If
                Next

                If bAddDoc = True Then
                    'Add TPA Fee Item
                    oAPCRNDoc.Lines.Add()
                    oAPCRNDoc.Lines.ItemCode = p_oCompDef.sTPAItemCode
                    oAPCRNDoc.Lines.Quantity = 1
                    oAPCRNDoc.Lines.PriceAfterVAT = dGrandTotal
                    oAPCRNDoc.Lines.TaxOnly = SAPbobsCOM.BoYesNoEnum.tYES
                    oAPCRNDoc.Lines.VatGroup = "SO1"
                    oAPCRNDoc.Lines.COGSCostingCode = p_oCompDef.sDefaultCostCenter


                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding APCredit Noted.", sFuncName)
                    lRetCode = oAPCRNDoc.Add
                    If lRetCode <> 0 Then
                        p_oCompany.GetLastError(lErrCode, sErrDesc)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding APCredit Note failed.", sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                End If

                Dim iAPCRNInvNo As Integer
                p_oCompany.GetNewObjectCode(iAPCRNInvNo)
                AddDataToTable(p_oDtReport, "TPA", iAPCRNInvNo, sCardCode, String.Empty)

            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddAPCreditMemo_TPA = RTN_SUCCESS

        Catch ex As Exception
            AddAPCreditMemo_TPA = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        Finally
            GC.Collect()
        End Try
    End Function

    Private Function AddAPCreditMemo_TPA_ZeroClaimAmt(ByVal oDoc As SAPbobsCOM.Documents, _
                                                      ByVal iAPInvNo As Integer, _
                                                      ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        'Dim oDoc As SAPbobsCOM.Documents
        Dim oAPCRNDoc As SAPbobsCOM.Documents
        Dim iCnt As Integer
        Dim lRetCode, lErrCode As Long
        Dim sRemarks As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet
        Dim dGrandTotal As Double = 0

        Try

            sFuncName = "AddAPCreditMemo_TPA_ZeroClaimAmt"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI APCRN Object", sFuncName)

            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
            oAPCRNDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)


            If oDoc.GetByKey(iAPInvNo) = True Then
                oAPCRNDoc.CardCode = oDoc.CardCode
                oAPCRNDoc.DocDate = oDoc.DocDate
                oAPCRNDoc.DocDueDate = oDoc.DocDueDate
                oAPCRNDoc.TaxDate = oDoc.TaxDate
                oAPCRNDoc.NumAtCard = oDoc.NumAtCard
                oAPCRNDoc.Comments = oDoc.Comments
                oAPCRNDoc.ImportFileNum = oDoc.ImportFileNum


                For i As Integer = 0 To oDoc.Lines.Count - 1
                    iCnt += 1
                    If iCnt > 1 Then
                        oAPCRNDoc.Lines.Add()
                    End If

                    oAPCRNDoc.Lines.SetCurrentLine(i)
                    oDoc.Lines.SetCurrentLine(i)

                    oAPCRNDoc.Lines.ItemCode = oDoc.Lines.ItemCode
                    oAPCRNDoc.Lines.Quantity = 1

                    oAPCRNDoc.Lines.VatGroup = "ZI" 'NO GST
                    oAPCRNDoc.Lines.TaxTotal = 0
                    oAPCRNDoc.Lines.PriceAfterVAT = (oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value) * 100 / 107

                    'oAPCRNDoc.Lines.BaseEntry = oDoc.DocEntry
                    'oAPCRNDoc.Lines.BaseLine = oDoc.Lines.LineNum
                    'oAPCRNDoc.Lines.BaseType = 18

                    oAPCRNDoc.Lines.COGSCostingCode = oDoc.Lines.COGSCostingCode

                    oDoc.Lines.SerialNum = GetSeriesNum("FPM")

                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_PatientName").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_VisitDate").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_VisitNo").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_ICNo").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_ConsultCost").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_DrugCost").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_SubTotal").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_GSTAmount").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_GrandTotal").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_PayeeAcNo").Value

                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFee").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTax").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value

                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_Fee1").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_Fee2").Value
                    oAPCRNDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value = oDoc.Lines.UserFields.Fields.Item("U_AI_Fee3").Value

                    dGrandTotal = dGrandTotal + CDbl(oDoc.Lines.UserFields.Fields.Item("U_AI_TPAFeeTotal").Value)
                Next

                'Add TPA Fee Item
                oAPCRNDoc.Lines.Add()
                oAPCRNDoc.Lines.ItemCode = p_oCompDef.sTPAItemCode
                oAPCRNDoc.Lines.Quantity = 1
                oAPCRNDoc.Lines.PriceAfterVAT = dGrandTotal
                oAPCRNDoc.Lines.TaxOnly = SAPbobsCOM.BoYesNoEnum.tYES
                oAPCRNDoc.Lines.VatGroup = "SO1"
                oAPCRNDoc.Lines.COGSCostingCode = p_oCompDef.sDefaultCostCenter


                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding APCredit Noted.", sFuncName)
                lRetCode = oAPCRNDoc.Add
                If lRetCode <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding APCredit Note failed.", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddAPCreditMemo_TPA_ZeroClaimAmt = RTN_SUCCESS

        Catch ex As Exception
            AddAPCreditMemo_TPA_ZeroClaimAmt = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        Finally
            GC.Collect()
        End Try
    End Function


#End Region

#Region "Reimburse to Member"

    Private Function ProcessReimburse_Member(ByVal sFileName As String, ByVal sSheet As String, ByVal oDv As DataView, ByRef sErrdesc As String) As Long

        Dim IsError As Boolean = False
        Dim sFuncName As String = "ProcessReimburse_Member"

        Try
            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadReimburse_Member", sFuncName)
            ReadReimburse_Member(sFileName, sSheet, IsError, oDv)

            If IsError = True Then
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Reimburse to Provider Excel Worksheet", sFuncName)
                sErrdesc = "Invalid Reimburse to Provider Excel Worksheet " & sFileName
                WriteToLogFile(sErrdesc, sFuncName)
                Throw New ArgumentException(sErrdesc)
            Else
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling Upload_ReimbursetoProvider_WorkSheet()", sFuncName)
                If Upload_ReimbursetoMember_WorkSheet(oDv, sErrdesc) <> RTN_SUCCESS Then
                    If RollBackTransaction(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Error in uploading Reimburse to Provider", sFuncName)
                    Throw New ArgumentException(sErrdesc)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Function completed successfully.", sFuncName)
            ProcessReimburse_Member = RTN_SUCCESS

        Catch ex As Exception
            ProcessReimburse_Member = RTN_ERROR
            If RollBackTransaction(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        End Try

    End Function

    Private Sub ReadReimburse_Member(ByVal sFileName As String, _
                                ByVal sSheet As String, _
                                ByRef bIsError As Boolean, _
                                ByRef dv As DataView)

        Dim iHeaderRow As Integer
        Dim sErrDesc As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "ReadReimburse_Member"


        iHeaderRow = 5

        dv = GetDataViewFromExcel(sFileName, sSheet)

        If IsNothing(dv) Then Exit Sub

        If dv(iHeaderRow)(0).ToString <> "Visit No." Then
            sErrDesc = "Invalid Excel file Format - ([Visit No.] not found at Column 1"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(1).ToString <> "Visit Date" Then
            sErrDesc = "Invalid Excel file Format - ([Visit Date] not found at Column 2"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(2).ToString <> "Invoice No." Then
            sErrDesc = "Invalid Excel file Format - ([Invoice No.] not found at Column 3"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If


        If dv(iHeaderRow)(3).ToString <> "Company Code" Then
            sErrDesc = "Invalid Excel file Format - ([Company Code] not found at Column 4"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString <> "Company Name" Then
            sErrDesc = "Invalid Excel file Format - ([Company Name] not found at Column 5"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If


        If dv(iHeaderRow)(5).ToString <> "Broker Name" Then
            sErrDesc = "Invalid Excel file Format - ([Broker Name] not found at Column 6"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If


        If dv(iHeaderRow)(6).ToString <> "Employee Name" Then
            sErrDesc = "Invalid Excel file Format - ([Employee Name] not found at Column 7"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(7).ToString <> "Employee ID No." Then
            sErrDesc = "Invalid Excel file Format - ([Employee ID No.] not found at Column 8"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString <> "Claimant Name" Then
            sErrDesc = "Invalid Excel file Format - ([Claimant Name] not found at Column 9"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(9).ToString <> "Claimant ID No." Then
            sErrDesc = "Invalid Excel file Format - ([Claimant ID No.] not found at Column 10"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString <> "Claimant Member Type" Then
            sErrDesc = "Invalid Excel file Format - ([Claimant Member Type] not found at Column 11"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(11).ToString <> "Contract Type" Then
            sErrDesc = "Invalid Excel file Format - ([Contract Type] not found at Column 12"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(12).ToString <> "Benefit Type" Then
            sErrDesc = "Invalid Excel file Format - ([Benefit Type] not found at Column 13"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(13).ToString <> "Provider Name" Then
            sErrDesc = "Invalid Excel file Format - ([Provider Name] not found at Column 14"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(14).ToString <> "Member Address Line 1" Then
            sErrDesc = "Invalid Excel file Format - ([Member Address Line 1] not found at Column 15"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(15).ToString <> "Member Address Line 2" Then
            sErrDesc = "Invalid Excel file Format - ([Member Address Line 2] not found at Column 16"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(16).ToString <> "Member Address Line 3" Then
            sErrDesc = "Invalid Excel file Format - ([Member Address Line 3] not found at Column 17"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(17).ToString <> "Member Country" Then
            sErrDesc = "Invalid Excel file Format - ([Member Country] not found at Column 18"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(18).ToString <> "Member Postal Code" Then
            sErrDesc = "Invalid Excel file Format - ([Member Postal Code] not found at Column 19"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(19).ToString <> "Mobile No." Then
            sErrDesc = "Invalid Excel file Format - ([Mobile No.] not found at Column 20"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(20).ToString <> "Email Address" Then
            sErrDesc = "Invalid Excel file Format - ([Email Address] not found at Column 21"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(21).ToString <> "Currency" Then
            sErrDesc = "Invalid Excel file Format - ([Currency] not found at Column 22"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(22).ToString <> "Consult Cost" Then
            sErrDesc = "Invalid Excel file Format - ([Consult Costt] not found at Column 23"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(23).ToString <> "Drug Cost" Then
            sErrDesc = "Invalid Excel file Format - ([Drug Cost] not found at Column 24"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(24).ToString <> "In-house Service Cost" Then
            sErrDesc = "Invalid Excel file Format - ([In-house Service Cost] not found at Column 25"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(25).ToString <> "Other Cost" Then
            sErrDesc = "Invalid Excel file Format - ([Other Cost] not found at Column 26"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(26).ToString <> "Sub-total" Then
            sErrDesc = "Invalid Excel file Format - ([Sub-total] not found at Column 27"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(27).ToString <> "Tax" Then
            sErrDesc = "Invalid Excel file Format - ([Tax] not found at Column 28"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(28).ToString <> "Grand Total" Then
            sErrDesc = "Invalid Excel file Format - ([Grand Total] not found at Column 29"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(29).ToString <> "Unclaim Amt." Then
            sErrDesc = "Invalid Excel file Format - ([Unclaim Amt.] not found at Column 30"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(30).ToString <> "Claim Amt." Then
            sErrDesc = "Invalid Excel file Format - ([Claim Amt.] not found at Column 31"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(31).ToString <> "Reimbursement Amt." Then
            sErrDesc = "Invalid Excel file Format - ([Reimbursement Amt.] not found at Column 32"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(32).ToString <> "Payment Mode" Then
            sErrDesc = "Invalid Excel file Format - ([Payment Mode] not found at Column 33"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(33).ToString <> "Payee Name" Then
            sErrDesc = "Invalid Excel file Format - ([Payee Name] not found at Column 34"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(34).ToString <> "Payee Account No." Then
            sErrDesc = "Invalid Excel file Format - ([Payee Account No.] not found at Column 35"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If

        If dv(iHeaderRow)(35).ToString <> "Remarks for Member" Then
            sErrDesc = "Invalid Excel file Format - ([Remarks for Member] not found at Column 36"
            WriteToLogFile(sErrDesc, sFuncName)
            bIsError = True
        End If


    End Sub

    Private Function Upload_ReimbursetoMember_WorkSheet(ByVal oDv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim oDS As New DataSet
        Dim oDatasetBP As New DataSet
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim k, m, r As Integer
        Dim sCostCenter As String = String.Empty
        Dim oPatArrary As New ArrayList
        Dim sCompanyName As String = String.Empty

        Try
            sFuncName = "Upload_ReimbursetoMember_WorkSheet"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function..", sFuncName)
            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sBatchNo = oDv(2)(0).ToString()
            sBatchPeriod = oDv(3)(0).ToString()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch Date.", sFuncName)
            m = InStrRev(sBatchPeriod, "to")
            sBatchPeriod = Microsoft.VisualBasic.Right(sBatchPeriod, Len(sBatchPeriod) - m - 1).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Date: " & sBatchPeriod, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Contract Owner Name.", sFuncName)
            sCardName = oDv(0)(0).ToString()
            k = InStrRev(sCardName, ":")
            sCardName = Microsoft.VisualBasic.Right(sCardName, Len(sCardName) - k).Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Card Name:" & sCardName, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Company Name.", sFuncName)
            sCompanyName = oDv(1)(0).ToString()
            r = InStrRev(sCompanyName, ":")
            sCompanyName = Microsoft.VisualBasic.Right(sCompanyName, Len(sCompanyName) - r).Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CheckBP()", sFuncName)
            If CheckBP(sCardName, sCardCode, "C", sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If


            'Chnaged from Claimant Name to Employee Name
            'Changed : split outgoing payment based on Employee Name + Payee Ac.No
            '*** 18/09/2014 ************
            'Change : Group all Payments with Payment Mode = CPF Medisave and create one Outgoing payment ( with Payename as CPF Board)
            '       : Split oDV into two tables one with EmployeeName + Payee Ac.No and Other with Payment mode CPF Medisave
            'Added Mobile NO. at column 20 in template
            '******** End ************

            Dim oDtCPF As DataTable
            Dim oDtEP As DataTable
            oDtEP = oDv.Table.Clone
            oDtCPF = oDv.Table.Clone
            oDtEP.Clear()
            oDtCPF.Clear()


            For Each row As DataRow In oDv.Table.Rows
                If Not (row.Item(32).ToString = "CPF Medisave" Or row.Item(32).ToString = "CPF Medishield") Then
                    oDtEP.ImportRow(row)
                End If
            Next

            For Each row As DataRow In oDv.Table.Rows
                If row.Item(32).ToString = "CPF Medisave" Or row.Item(32).ToString = "CPF Medishield" Then
                    oDtCPF.ImportRow(row)
                End If
            Next

            Dim oDT As DataTable
            oDT = oDtEP.DefaultView.ToTable(True, "F7", "F35")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Total Patient Count..." & oDT.Rows.Count, sFuncName)
            Dim DBRows() As DataRow

            For Each row As DataRow In oDT.Rows
                If row.Item(0).ToString = String.Empty And row.Item(1).ToString = String.Empty Then
                    'DO NOTHING
                Else
                    If Not row.Item(0).ToString.Trim = "Employee Name" And Not row.Item(1).ToString.Trim = "Payee Account No." Then
                        'Dim s As String = "F7='" & row.Item(0).ToString.Trim & "' and F34='" & row.Item(1).ToString.Trim & "'"
                        If row.Item(1).ToString.Trim = String.Empty Then
                            DBRows = oDtEP.Select("F7='" & Replace(row.Item(0).ToString.Trim, "'", "''") & "' and (F35 = '' or F35 IS NULL)")
                        Else
                            DBRows = oDtEP.Select("F7='" & Replace(row.Item(0).ToString.Trim, "'", "''") & "' and F35='" & row.Item(1).ToString.Trim & "'")
                        End If
                        If DBRows.Length > 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateOutgoingPayment()", sFuncName)
                            If CreateOutgoingPayment(DBRows, row.Item(0).ToString, sBatchNo, sBatchPeriod, sCardCode, sCardName, sCompanyName, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                        Else
                            sErrDesc = "Failed to created document for ::" & row.Item(0).ToString.Trim
                            Throw (New ArgumentException(sErrDesc))
                        End If
                    End If
                End If
            Next

            ''Payment "CPF Medisave" - Create Outgoing Payment with Payee as CPF
            If oDtCPF.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateOutgoingPayment()", sFuncName)
                If CreateOutgoingPayment_CPF(oDtCPF, sBatchNo, sBatchPeriod, sCardCode, sCardName, sCompanyName, sErrDesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrDesc)
                End If



            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            Upload_ReimbursetoMember_WorkSheet = RTN_SUCCESS

        Catch ex As Exception
            Upload_ReimbursetoMember_WorkSheet = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        Finally

        End Try
    End Function

    Private Function CreateOutgoingPayment(ByVal dtRows() As DataRow, _
                                            ByVal sName As String, _
                                            ByVal sBatchNo As String, _
                                            ByVal sBatchPeriod As String, _
                                            ByVal sCardCode As String, _
                                            ByVal sCardName As String, _
                                            ByVal sCompanyName As String, _
                                            ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "CreateOutgoingPayment"
        Dim sCostCenter As String
        Dim oPayment As SAPbobsCOM.IPayments
        Dim iCnt, k As Integer
        Dim lRetCode, lErrCode As Long
        Dim dGrandTotal As Double = 0
        Dim bIsPymtNoBlank As Boolean = False
        Dim sDfltBankCode As String = String.Empty
        Dim sdfltBankAcct As String = String.Empty
        Dim sBNo As Integer
        Dim sMobileNo As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Outgoing Payment for ::" & sName, sFuncName)

            oPayment = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
            oPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount

            Dim sBPName, sBPName1 As String

            sBPName = "Gethin-Jones Medical Practice Pte Ltd"
            sBPName1 = "Fullerton Healthcare Group Pte Ltd"

            If UCase(sGJDBName.Trim) = UCase(sBPName) Then
                sCostCenter = "CAP"
            ElseIf UCase(sCardName.Trim) = UCase(sBPName1) Then
                sCostCenter = GetCostCenterByCardName(sCompanyName)
            Else
                sCostCenter = GetCostCenter(sCardCode)
            End If

            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter

            GetDefaultBankDetails(sDfltBankCode, sdfltBankAcct)

            k = InStrRev(sBatchNo, ":")
            sBNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - k).Trim


            oPayment.DocDate = CDate(sBatchPeriod)
            oPayment.TaxDate = CDate(sBatchPeriod)
            oPayment.Remarks = "Batch-" & sBatchNo
            oPayment.CounterReference = sBNo

            'Assign mobile No & Batch No
            sMobileNo = dtRows(0).Item(19).ToString
            oPayment.UserFields.Fields.Item("U_AI_MobileNo").Value = sMobileNo
            oPayment.UserFields.Fields.Item("U_AI_BatchNo").Value = CStr(sBNo)

            'Assign contract owner
            oPayment.UserFields.Fields.Item("U_AI_CtrOwner").Value = sCardName.Trim

            '****************  Broker code ***********************

            Dim sSQL As String = String.Empty
            Dim sDeliveryMode As String = String.Empty
            Dim sBrokerName As String = String.Empty
            Dim sBrokerCode As String = String.Empty
            Dim iCount As Integer

            If Left(UCase(sCardName.ToString.Trim), 3).ToString = "AIA" Then
                sSQL = "SELECT * FROM ""@AE_BROKERSETUP"""
                Dim oDS As New DataSet
                oDS = ExecuteSQLQuery(sSQL)

                For Each row As DataRow In oDS.Tables(0).Rows
                    Dim i As Integer = row.Item("U_AE_ColumnNo")
                    Dim sValue As String()
                    sValue = dtRows(0).Item(i - 1).ToString.Split

                    Dim sBrkCode As String()
                    sBrkCode = row.Item("U_AE_BrokerCode").ToString.Split

                    iCount = 0
                    For iLine As Integer = 0 To sValue.Length - 1
                        For iBkline As Integer = 0 To sBrkCode.Length - 1
                            If String.Compare(sValue(iLine).ToString.Trim, sBrkCode(iBkline).ToString.Trim, True) = 0 Then
                                iCount += 1
                                Exit For
                            End If
                        Next
                    Next

                    If iCount = sBrkCode.Length Then
                        sDeliveryMode = row.Item("U_AE_DelMode").ToString
                        sBrokerName = row.Item("U_AE_BrokerName").ToString
                        sBrokerCode = row.Item("U_AE_BrokerCode").ToString
                        Exit For
                    End If

                Next

                If sDeliveryMode = String.Empty Then sDeliveryMode = "M"
                oPayment.UserFields.Fields.Item("U_AE_DelMode").Value = sDeliveryMode
                oPayment.UserFields.Fields.Item("U_AE_BrokerName").Value = sBrokerName
                oPayment.UserFields.Fields.Item("U_AE_BrokerCode").Value = sBrokerCode
            Else

                oPayment.UserFields.Fields.Item("U_AE_DelMode").Value = "M"
                oPayment.UserFields.Fields.Item("U_AE_BrokerName").Value = String.Empty
                oPayment.UserFields.Fields.Item("U_AE_BrokerCode").Value = String.Empty

            End If


            If dtRows(0).Item(32).ToString = "Private Shield" Then
                oPayment.UserFields.Fields.Item("U_AE_DelMode").Value = "S"
                oPayment.UserFields.Fields.Item("U_AE_BrokerName").Value = dtRows(0).Item(33).ToString
                oPayment.UserFields.Fields.Item("U_AE_BrokerCode").Value = dtRows(0).Item(33).ToString
            End If


            '-- end --


            For Each row As DataRow In dtRows
                If CDbl(row.Item(30)) > 0 Then
                    iCnt += 1
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing Line No. : " & iCnt, sFuncName)
                    If iCnt > 1 Then
                        oPayment.AccountPayments.Add()
                    End If

                    If row.Item(11).ToString = "FFS" Then
                        oPayment.AccountPayments.AccountCode = p_oCompDef.sFFSGLCode
                    ElseIf row.Item(11).ToString = "CAP" Then
                        oPayment.AccountPayments.AccountCode = p_oCompDef.sCAPGLCode
                    ElseIf row.Item(11).ToString = "3FS" Then
                        oPayment.AccountPayments.AccountCode = p_oCompDef.s3FSGLCode
                    End If

                    If row.Item(34).ToString = String.Empty Then
                        bIsPymtNoBlank = True
                        oPayment.Address = row.Item(33).ToString
                        oPayment.CardName = row.Item(33).ToString
                    Else
                        oPayment.Address = sName
                    End If

                    'oPayment.AccountPayments.GrossAmount = CDbl(row.Item(28))
                    If IsDBNull(row.Item(31)) Then
                        sErrdesc = "No Reimbursement Amount found for Employee Name:: " & sName
                        Throw New ArgumentException(sErrdesc)
                    End If

                    oPayment.AccountPayments.GrossAmount = CDbl(row.Item(31))

                    'If sGJDBName = "Gethin-Jones Medical Practice Pte Ltd" Then
                    '    oPayment.AccountPayments.ProfitCenter2 = sCostCenter.Trim
                    'Else
                    '    oPayment.AccountPayments.ProfitCenter = sCostCenter
                    'End If

                    oPayment.AccountPayments.ProfitCenter = sCostCenter

                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_PatientName").Value = row.Item(8).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_VisitDate").Value = row.Item(1)
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ProviderName").Value = row.Item(13).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_CostCenter").Value = row.Item(10).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_VisitNo").Value = row.Item(0).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ICNo").Value = row.Item(9).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_DrugCost").Value = CDbl(row.Item(23))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_InHouseServCost").Value = CDbl(row.Item(24))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ExtServCost").Value = CDbl(row.Item(25))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_SubTotal").Value = CDbl(row.Item(26))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_GSTAmount").Value = CDbl(row.Item(27))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_UnclaimedAmount").Value = CDbl(row.Item(29))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_GrandTotal").Value = CDbl(row.Item(28))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ContractType").Value = row.Item(11).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_BenefitType").Value = row.Item(12).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_PayeeName").Value = row.Item(33).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = row.Item(34).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_RemarkMember").Value = row.Item(35).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_CompanyName").Value = row.Item(4).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_EmpName").Value = row.Item(6).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_EmpID").Value = row.Item(7).ToString

                    'New fileds added for Generation of Bank file
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD1").Value = row.Item(14).ToString 'Member Address Line 1
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD2").Value = row.Item(15).ToString 'Member Address Line 2
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD3").Value = row.Item(16).ToString 'Member Address Line 3
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD4").Value = row.Item(18).ToString 'Member Postal Code
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD5").Value = row.Item(17).ToString 'Member Country
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_INVOICENO").Value = row.Item(2).ToString 'Invoice No
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_CompanyCode").Value = row.Item(3).ToString

                    dGrandTotal = dGrandTotal + CDbl(row.Item(31))
                End If
            Next

            Dim sGLAccount As String = GetDfltGLAccount(sDfltBankCode, sdfltBankAcct)
            If sGLAccount = String.Empty Then
                sErrdesc = "No GL Account found for Bank Code::" & sDfltBankCode & " and BankAccount::" & sdfltBankAcct
                Throw New ArgumentException(sErrdesc)
            End If

            If dGrandTotal > 0 Then
                If bIsPymtNoBlank = True Then
                    If Left(sCardName.Trim, 3) = "AIA" Then
                        oPayment.Checks.BankCode = p_oCompDef.sDBS_CheckBankCode
                        oPayment.Checks.AccounttNum = p_oCompDef.sDBS_CheckBankAccount
                        oPayment.Checks.CountryCode = "SG"
                        oPayment.Checks.CheckSum = dGrandTotal
                        oPayment.CheckAccount = p_oCompDef.sDBS_CheckGLAccount
                    ElseIf Left(sCardName.Trim, 3) = "AON" Then
                        oPayment.Checks.BankCode = p_oCompDef.sDBS_AONCheckBankCode
                        oPayment.Checks.AccounttNum = p_oCompDef.sDBS_AONCheckBankAccount
                        oPayment.Checks.CountryCode = "SG"
                        oPayment.Checks.CheckSum = dGrandTotal
                        oPayment.CheckAccount = p_oCompDef.sDBS_AONCheckGLAccount
                    Else
                        oPayment.Checks.BankCode = p_oCompDef.sCheckBankCode
                        oPayment.Checks.AccounttNum = p_oCompDef.sCheckBankAccount
                        oPayment.Checks.CheckSum = dGrandTotal
                        oPayment.CheckAccount = p_oCompDef.sCheckGLAccount
                    End If


                    oPayment.CashSum = 0
                    oPayment.TransferSum = 0
                Else
                    oPayment.TransferSum = dGrandTotal
                    oPayment.CashSum = 0

                    If sCardName = "AIA Singapore Pte. Ltd." Then
                        oPayment.TransferAccount = p_oCompDef.sGIROGLAccountAIA
                    Else
                        oPayment.TransferAccount = p_oCompDef.sGIROGLAccount
                    End If
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Outgoing Payment Document for ::" & sName, sFuncName)
                lRetCode = oPayment.Add

                If lRetCode <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrdesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming paymnet failed.", sFuncName)
                    Throw New ArgumentException(sErrdesc)
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully created Outgoing Payment for ::" & sName, sFuncName)

                'Add Mobile No.,DocEntry and Amount to Datatable for GIRO Payment
                'If bIsPymtNoBlank = False Then
                '    Dim iPymtNo As Integer
                '    p_oCompany.GetNewObjectCode(iPymtNo)
                '    AddDataToTable(p_oDtSMS, iPymtNo, sMobileNo, dGrandTotal)
                'End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            CreateOutgoingPayment = RTN_SUCCESS

        Catch ex As Exception
            CreateOutgoingPayment = RTN_ERROR
            sErrdesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrdesc, sFuncName)
        Finally

        End Try
    End Function

    Private Function CreateOutgoingPayment_CPF(ByVal oDt As DataTable, _
                                           ByVal sBatchNo As String, _
                                           ByVal sBatchPeriod As String, _
                                           ByVal sCardCode As String, _
                                           ByVal sCardName As String, _
                                           ByVal sCompanyName As String, _
                                           ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "CreateOutgoingPayment_CPF"
        Dim sCostCenter As String
        Dim oPayment As SAPbobsCOM.IPayments
        Dim iCnt, k As Integer
        Dim lRetCode, lErrCode As Long
        Dim dGrandTotal As Double = 0
        Dim bIsPymtNoBlank As Boolean = False
        Dim sDfltBankCode As String = String.Empty
        Dim sdfltBankAcct As String = String.Empty
        Dim sBNo As Integer
        Dim sName As String
        Dim sMobileNo As String = String.Empty
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sName = "CPF Board"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Outgoing Payment for ::" & sName, sFuncName)

            oPayment = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
            oPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount

            Dim sBPName, sBPName1 As String

            sBPName = "Gethin-Jones Medical Practice Pte Ltd"
            sBPName1 = "Fullerton Healthcare Group Pte Ltd"

            If UCase(sGJDBName.Trim) = UCase(sBPName) Then
                sCostCenter = "CAP"
            ElseIf UCase(sCardName.Trim) = UCase(sBPName1) Then
                sCostCenter = GetCostCenterByCardName(sCompanyName)
            Else
                sCostCenter = GetCostCenter(sCardCode)
            End If

            If sCostCenter = String.Empty Then sCostCenter = p_oCompDef.sDefaultCostCenter


            GetDefaultBankDetails(sDfltBankCode, sdfltBankAcct)

            k = InStrRev(sBatchNo, ":")
            sBNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - k).Trim


            oPayment.DocDate = CDate(sBatchPeriod)
            oPayment.TaxDate = CDate(sBatchPeriod)
            oPayment.Remarks = "Batch-" & sBatchNo
            oPayment.CounterReference = sBNo

            'Assign mobile No & Batch No
            sMobileNo = oDt.Rows(0).Item(19).ToString().Trim()
            oPayment.UserFields.Fields.Item("U_AI_MobileNo").Value = sMobileNo
            oPayment.UserFields.Fields.Item("U_AI_BatchNo").Value = CStr(sBNo)

            'Assign contract owner
            oPayment.UserFields.Fields.Item("U_AI_CtrOwner").Value = sCardName.Trim

            'Broker/Insurer
            oPayment.UserFields.Fields.Item("U_AE_DelMode").Value = "S"
            oPayment.UserFields.Fields.Item("U_AE_BrokerCode").Value = "CPF"
            oPayment.UserFields.Fields.Item("U_AE_BrokerName").Value = "CPF BOARD"


            For Each row As DataRow In oDt.Rows
                If CDbl(row.Item(30)) > 0 Then
                    iCnt += 1
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing Line No. : " & iCnt, sFuncName)
                    If iCnt > 1 Then
                        oPayment.AccountPayments.Add()
                    End If

                    If row.Item(11).ToString = "FFS" Then
                        oPayment.AccountPayments.AccountCode = p_oCompDef.sFFSGLCode
                    ElseIf row.Item(11).ToString = "CAP" Then
                        oPayment.AccountPayments.AccountCode = p_oCompDef.sCAPGLCode
                    ElseIf row.Item(11).ToString = "3FS" Then
                        oPayment.AccountPayments.AccountCode = p_oCompDef.s3FSGLCode
                    End If

                    oPayment.Address = sName
                    oPayment.CardName = sName

                    'oPayment.AccountPayments.GrossAmount = CDbl(row.Item(28))
                    If IsDBNull(row.Item(31)) Then
                        sErrdesc = "No Reimbursement Amount found for Employee Name:: " & sName
                        Throw New ArgumentException(sErrdesc)
                    End If

                    oPayment.AccountPayments.GrossAmount = CDbl(row.Item(31))

                    'If sGJDBName = "Gethin-Jones Medical Practice Pte Ltd" Then
                    '    oPayment.AccountPayments.ProfitCenter2 = sCostCenter
                    'Else
                    '    oPayment.AccountPayments.ProfitCenter = sCostCenter
                    'End If

                    oPayment.AccountPayments.ProfitCenter = sCostCenter

                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_PatientName").Value = row.Item(8).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_VisitDate").Value = row.Item(1)
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ProviderName").Value = row.Item(13).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_CostCenter").Value = row.Item(10).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_VisitNo").Value = row.Item(0).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ICNo").Value = row.Item(9).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_DrugCost").Value = CDbl(row.Item(23))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_InHouseServCost").Value = CDbl(row.Item(24))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ExtServCost").Value = CDbl(row.Item(25))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_SubTotal").Value = CDbl(row.Item(26))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_GSTAmount").Value = CDbl(row.Item(27))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_UnclaimedAmount").Value = CDbl(row.Item(29))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_GrandTotal").Value = CDbl(row.Item(28))
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ContractType").Value = row.Item(11).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_BenefitType").Value = row.Item(12).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_PayeeName").Value = row.Item(33).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = row.Item(34).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_RemarkMember").Value = row.Item(35).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_CompanyName").Value = row.Item(4).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_EmpName").Value = row.Item(6).ToString
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_EmpID").Value = row.Item(7).ToString

                    'New fileds added for Generation of Bank file
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD1").Value = row.Item(14).ToString 'Member Address Line 1
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD2").Value = row.Item(15).ToString 'Member Address Line 2
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD3").Value = row.Item(16).ToString 'Member Address Line 3
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD4").Value = row.Item(18).ToString 'Member Postal Code
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_ADD5").Value = row.Item(17).ToString 'Member Country
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_INVOICENO").Value = row.Item(2).ToString 'Invoice No
                    oPayment.AccountPayments.UserFields.Fields.Item("U_AI_CompanyCode").Value = row.Item(3).ToString

                    dGrandTotal = dGrandTotal + CDbl(row.Item(31))
                End If
            Next

            Dim sGLAccount As String = GetDfltGLAccount(sDfltBankCode, sdfltBankAcct)
            If sGLAccount = String.Empty Then
                sErrdesc = "No GL Account found for Bank Code::" & sDfltBankCode & " and BankAccount::" & sdfltBankAcct
                Throw New ArgumentException(sErrdesc)
            End If

            If dGrandTotal > 0 Then
                If Left(sCardName.Trim, 3) = "AIA" Then
                    oPayment.Checks.BankCode = p_oCompDef.sDBS_CheckBankCode
                    oPayment.Checks.AccounttNum = p_oCompDef.sDBS_CheckBankAccount
                    oPayment.Checks.CountryCode = "SG"
                    oPayment.Checks.CheckSum = dGrandTotal
                    oPayment.CheckAccount = p_oCompDef.sDBS_CheckGLAccount
                ElseIf Left(sCardName.Trim, 3) = "AON" Then
                    oPayment.Checks.BankCode = p_oCompDef.sDBS_AONCheckBankCode
                    oPayment.Checks.AccounttNum = p_oCompDef.sDBS_AONCheckBankAccount
                    oPayment.Checks.CountryCode = "SG"
                    oPayment.Checks.CheckSum = dGrandTotal
                    oPayment.CheckAccount = p_oCompDef.sDBS_AONCheckGLAccount
                Else
                    oPayment.Checks.BankCode = p_oCompDef.sCheckBankCode
                    oPayment.Checks.AccounttNum = p_oCompDef.sCheckBankAccount
                    oPayment.Checks.CheckSum = dGrandTotal
                    oPayment.CheckAccount = p_oCompDef.sCheckGLAccount
                End If



                oPayment.CashSum = 0
                oPayment.TransferSum = 0

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Outgoing Payment Document for ::" & sName, sFuncName)
                lRetCode = oPayment.Add

                If lRetCode <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrdesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming paymnet failed.", sFuncName)
                    Throw New ArgumentException(sErrdesc)
                Else
                    'Add Mobile No.,DocEntry and Amount to Datatable for GIRO Payment
                    'If bIsPymtNoBlank = False Then
                    '    Dim iPymtNo As Integer
                    '    p_oCompany.GetNewObjectCode(iPymtNo)
                    '    AddDataToTable(p_oDtSMS, iPymtNo, sMobileNo, dGrandTotal)
                    'End If

                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully created Outgoing Payment for ::" & sName, sFuncName)

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            CreateOutgoingPayment_CPF = RTN_SUCCESS

        Catch ex As Exception
            CreateOutgoingPayment_CPF = RTN_ERROR
            sErrdesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrdesc, sFuncName)
        Finally

        End Try
    End Function

#End Region

End Module
