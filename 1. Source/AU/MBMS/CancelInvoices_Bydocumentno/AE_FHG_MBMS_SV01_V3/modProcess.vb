
Module modProcess

    Public Sub Start()

        Dim sFuncName As String = "Start()"
        Dim sErrdesc As String = String.Empty

        Try

            WriteToStatusScreen(False, "=============================== Cancelling Batch No:: " & CancelDocuments.txtDocFrom.Text & "  ===========================")

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ConnectToCompany()", sFuncName)
            If ConnectToCompany(p_oCompany, sErrdesc) <> RTN_SUCCESS Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Unable to connect to SAP.", sFuncName)
                Throw New ArgumentException("Unable to connect to SAP.")
            End If

            'If CancelOutgoingPayment(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

            If CancelDocuments.DocType.SelectedItem = "---Select---" Then
                sErrdesc = "Select a document to Cancel"
                WriteToStatusScreen(False, sErrdesc)
                Throw New ArgumentException(sErrdesc)
            ElseIf CancelDocuments.DocType.SelectedItem = "AP Invoice" Then
                'WriteToStatusScreen(False, "============= Start - Cancelling AP CN =================")

                'If CancelAPCreditNotes(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

                'WriteToStatusScreen(False, "============= END - Cancelling AP CN =================")

                WriteToStatusScreen(False, "============= Start - Cancelling AP Invoices =================")

                If CancelAPInvoices(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

                WriteToStatusScreen(False, "============= END - Cancelling AP Invoices =================")

            ElseIf CancelDocuments.DocType.SelectedItem = "AR Invoice" Then
                'WriteToStatusScreen(False, "============= Start - Cancelling A/R CN =================")

                'If CancelARCreditNotes(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

                'WriteToStatusScreen(False, "============= END - Cancelling A/R CN =================")

                WriteToStatusScreen(False, "============= Start - Cancelling A/R Invoices =================")

                If CancelARInvoices(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

                WriteToStatusScreen(False, "============= END - Cancelling A/R Invoices =================")
            End If



            WriteToStatusScreen(False, "=============  COMPLETED =================")


        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in upload", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try

    End Sub



#Region "AP"

    Private Function CancelOutgoingPayment(ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "CancelOutgoingPayment"
        Dim oPayment As SAPbobsCOM.IPayments
        Dim lRetCode, lErrCode As Long
        Dim oDS As New DataSet

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)


            oPayment = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
            oPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount

            Dim sSQL As String = "SELECT T0.""DocNum"",T0.""DocEntry"" FROM OVPM T0 WHERE T0.""Canceled""='N' and T0.""CounterRef"" ='6200'"

            oDS = ExecuteSQLQuery(sSQL)

            For Each row As DataRow In oDS.Tables(0).Rows

                oPayment.GetByKey(row.Item("DocEntry"))

                lRetCode = oPayment.Cancel

                If lRetCode <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrdesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming paymnet failed.", sFuncName)
                    Throw New ArgumentException(sErrdesc)
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            CancelOutgoingPayment = RTN_SUCCESS

        Catch ex As Exception
            CancelOutgoingPayment = RTN_ERROR
            sErrdesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrdesc, sFuncName)
        Finally

        End Try
    End Function

    Private Function CancelAPCreditNotes(ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "CancelAPCreditNotes"
        Dim lErrCode As Long
        Dim oDS As New DataSet
        Dim oAPCRNDoc As SAPbobsCOM.Documents

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oAPCRNDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)

            Dim sSQL As String = "SELECT T0.""DocNum"",T0.""DocEntry"" FROM ORPC T0 WHERE T0.""DocNum"">=" & CancelDocuments.txtDocFrom.Text & " AND T0.""DocNum""<=" & CancelDocuments.txtDocTo.Text & " AND ""CANCELED""='N'"

            oDS = ExecuteSQLQuery(sSQL)
            Dim iCount As Integer = 0
            For Each row As DataRow In oDS.Tables(0).Rows

                WriteToStatusScreen(False, "============ Cancel AP CN :: " & row.Item("DocNum") & " =================")

                oAPCRNDoc.GetByKey(row.Item("DocEntry"))
                Dim oCancelDoc As SAPbobsCOM.Documents = oAPCRNDoc.CreateCancellationDocument()
                oAPCRNDoc.DocDate = "20161130"

                If oCancelDoc.Add() <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrdesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming paymnet failed.", sFuncName)
                    Throw New ArgumentException(sErrdesc)
                Else
                    iCount = iCount + 1
                    Call WriteToLogFile_Debug("Updated count value is " & iCount, sFuncName)
                End If
            Next

            WriteToStatusScreen(False, "============= TOTAL COUNT :: " & iCount & " =================")
            Call WriteToLogFile_Debug("Completed AP CN", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            CancelAPCreditNotes = RTN_SUCCESS

        Catch ex As Exception
            CancelAPCreditNotes = RTN_ERROR
            sErrdesc = ex.Message
            WriteToStatusScreen(False, "============= ERROR ::" & sErrdesc & "  =================")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrdesc, sFuncName)
        Finally

        End Try
    End Function

    Private Function CancelAPInvoices(ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "CancelAPInvoices"
        Dim lErrCode As Long
        Dim oDS As New DataSet
        Dim oAPCRNDoc As SAPbobsCOM.Documents

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oAPCRNDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            Dim sSQL As String = "SELECT T0.""DocNum"",T0.""DocEntry"" FROM OPCH T0 WHERE T0.""DocNum"">=" & CancelDocuments.txtDocFrom.Text & " AND T0.""DocNum""<=" & CancelDocuments.txtDocTo.Text & " AND ""CANCELED""='N'"

            oDS = ExecuteSQLQuery(sSQL)
            Dim iCount As Integer = 0

            For Each row As DataRow In oDS.Tables(0).Rows
                Dim sDocEntry As String = row.Item("DocEntry")
                oAPCRNDoc.GetByKey(row.Item("DocEntry"))

                WriteToStatusScreen(False, "============ Cancel AP Invoice :: " & row.Item("DocNum") & " =================")

                Dim oCancelDoc As SAPbobsCOM.Documents = oAPCRNDoc.CreateCancellationDocument()

                If oCancelDoc.Add() <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrdesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming paymnet failed.", sFuncName)
                    Throw New ArgumentException(sErrdesc)
                Else
                    iCount = iCount + 1
                    Call WriteToLogFile_Debug("Updated count value is " & iCount, sFuncName)
                End If

            Next

            WriteToStatusScreen(False, "============= TOTAL COUNT :: " & iCount & " =================")
            Call WriteToLogFile_Debug("Completed AP CN", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            CancelAPInvoices = RTN_SUCCESS

        Catch ex As Exception
            CancelAPInvoices = RTN_ERROR
            sErrdesc = ex.Message
            WriteToStatusScreen(False, "============= ERROR ::" & sErrdesc & "  =================")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrdesc, sFuncName)
        Finally

        End Try
    End Function

#End Region

#Region "AR"


    Private Function CancelARCreditNotes(ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "CancelARCreditNotes"
        Dim lErrCode As Long
        Dim oDS As New DataSet
        Dim oARCRNDoc As SAPbobsCOM.Documents

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oARCRNDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)

            Dim sSQL As String = "SELECT T0.""DocNum"",T0.""DocEntry"" FROM ORIN T0 WHERE T0.""DocNum"">=" & CancelDocuments.txtDocFrom.Text & " AND T0.""DocNum""<=" & CancelDocuments.txtDocTo.Text & " AND ""CANCELED""='N'"

            oDS = ExecuteSQLQuery(sSQL)
            Dim iCount As Integer = 0
            For Each row As DataRow In oDS.Tables(0).Rows

                WriteToStatusScreen(False, "============ Cancel A/R CN :: " & row.Item("DocNum") & " =================")

                oARCRNDoc.GetByKey(row.Item("DocEntry"))
                Dim oCancelDoc As SAPbobsCOM.Documents = oARCRNDoc.CreateCancellationDocument()

                If oCancelDoc.Add() <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrdesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming paymnet failed.", sFuncName)
                    Throw New ArgumentException(sErrdesc)
                Else
                    iCount = iCount + 1
                    Call WriteToLogFile_Debug("Updated count value is " & iCount, sFuncName)
                End If
            Next

            WriteToStatusScreen(False, "============= TOTAL COUNT :: " & iCount & " =================")

            Call WriteToLogFile_Debug("Completed A/R CN", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            CancelARCreditNotes = RTN_SUCCESS

        Catch ex As Exception
            CancelARCreditNotes = RTN_ERROR
            sErrdesc = ex.Message
            WriteToStatusScreen(False, "============= ERROR ::" & sErrdesc & "  =================")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrdesc, sFuncName)
        Finally

        End Try
    End Function

    Private Function CancelARInvoices(ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "CancelARInvoices"
        Dim lErrCode As Long
        Dim oDS As New DataSet
        Dim oARInvoice As SAPbobsCOM.Documents

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oARInvoice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            Dim sSQL As String = "SELECT T0.""DocNum"",T0.""DocEntry"" FROM OINV T0 WHERE T0.""DocNum"">=" & CancelDocuments.txtDocFrom.Text & " AND T0.""DocNum""<=" & CancelDocuments.txtDocTo.Text & " AND ""CANCELED""='N'"

            oDS = ExecuteSQLQuery(sSQL)
            Dim iCount As Integer = 0

            For Each row As DataRow In oDS.Tables(0).Rows
                Dim sDocEntry As String = row.Item("DocEntry")
                oARInvoice.GetByKey(row.Item("DocEntry"))

                WriteToStatusScreen(False, "============ Cancel AR Invoice :: " & row.Item("DocNum") & " =================")

                Dim oCancelDoc As SAPbobsCOM.Documents = oARInvoice.CreateCancellationDocument()

                If oCancelDoc.Add() <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrdesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Incoming paymnet failed.", sFuncName)
                    Throw New ArgumentException(sErrdesc)
                Else
                    iCount = iCount + 1
                    Call WriteToLogFile_Debug("Updated count value is " & iCount, sFuncName)
                End If

            Next

            WriteToStatusScreen(False, "============= TOTAL COUNT :: " & iCount & " =================")
            Call WriteToLogFile_Debug("Completed AR Invoices", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function with SUCCESS", sFuncName)
            CancelARInvoices = RTN_SUCCESS

        Catch ex As Exception
            CancelARInvoices = RTN_ERROR
            sErrdesc = ex.Message
            WriteToStatusScreen(False, "============= ERROR ::" & sErrdesc & "  =================")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrdesc, sFuncName)
        Finally

        End Try
    End Function

#End Region

    Public Sub WriteToStatusScreen(ByVal Clear As Boolean, ByVal msg As String)
        If Clear Then
            CancelDocuments.txtStatusMsg.Text = ""
        End If
        CancelDocuments.txtStatusMsg.HideSelection = True
        CancelDocuments.txtStatusMsg.Text &= msg & vbCrLf
        CancelDocuments.txtStatusMsg.SelectAll()
        CancelDocuments.txtStatusMsg.ScrollToCaret()
        CancelDocuments.txtStatusMsg.Refresh()
    End Sub


End Module
