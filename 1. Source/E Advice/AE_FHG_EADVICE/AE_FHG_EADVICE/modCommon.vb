Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Data.Common
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.IO

Module modCommon

    ' Company Default Structure
    Public Structure CompanyDefault

        Public sServer As String
        Public sLicenceServer As String
        Public iServerLanguage As Integer
        Public iServerType As Integer
        Public sSAPDBName As String
        Public sSAPUser As String
        Public sSAPPwd As String
        Public sDBUser As String
        Public sDBPwd As String
        Public sDSN As String

        Public sInboxDir As String
        Public sSuccessDir As String
        Public sFailDir As String
        Public sLogPath As String
        Public sDebug As String

        Public sEmailFrom As String
        Public sEmailTo As String
        Public sEmailSubject As String
        Public sSMTPServer As String
        Public sSMTPPort As String
        Public sSMTPUser As String
        Public sSMTPPassword As String

        Public sCreditNoteGL As String
        Public sNonStockItem As String
        Public sFFSItemCode As String
        Public sCAPItemCode As String
        Public sTPAItemCode As String
        Public sFFSGLCode As String
        Public sCAPGLCode As String
        Public sDefaultCostCenter As String
        Public dServiceFee As Double
        Public sCustBPSeriesName As String
        Public sVenBPSeriesName As String

        Public sReportPDFPath As String
        Public sReportDSN As String
        Public sReportsPath As String

        Public sCheckGLAccount As String
        Public sGIROGLAccount As String
        Public sCheckBankAccount As String
        Public sCheckBankCode As String
        Public sDBList As String

    End Structure

    ' Return Value Variable Control
    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    ' Debug Value Variable Control
    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    ' Global variables group
    Public p_iDebugMode As Int16 = DEBUG_ON
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16
    Public p_oCompDef As CompanyDefault
    Public p_dProcessing As DateTime
    Public p_oDtSuccess As DataTable
    Public p_oDtError As DataTable
    Public p_oDtReport As DataTable
    Public p_SyncDateTime As String
    Public oCompany As SAPbobsCOM.Company

    Public oCompList As New Dictionary(Of String, String)
    Public oCompListSOA As New Dictionary(Of String, String)

    Public PaymentAdvice_F As New frmPaymentadvice
    Public SOA_F As New frmSOA
    Public Eadvice_F As New frmEadvice
    '' Public CFL_F As New CFL
    Public sCFL As String = String.Empty
    Public sCFL1 As String = String.Empty
    Public sCompanyName As String = String.Empty





    Public Function ExecuteSQLQuery(ByVal sQuery As String) As DataSet

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sri
        ' Date          : 
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty

        Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
        Dim oCon As New Odbc.OdbcConnection(sConstr)
        Dim oCmd As New Odbc.OdbcCommand
        Dim oDs As New DataSet

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connection " & sConstr, sFuncName)

        Try
            sFuncName = "ExecuteQuery()"
            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(oCmd)
            da.Fill(oDs)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
        Return oDs
    End Function

    Public Function GetSystemIntializeInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemIntializeInfo()
        '   Purpose     :   This function will be providing to proceed the initializing 
        '                   variable control during the system start-up
        '               
        '   Parameters  :   ByRef oCompDef As CompanyDefault
        '                       oCompDef =  set the Company Default structure
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   October 2013
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Try

            sFuncName = "GetSystemIntializeInfo()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCompDef.sServer = String.Empty
            oCompDef.iServerLanguage = 3
            oCompDef.iServerType = 7            ' 6 = SQL 2008, 7 = 2012, 9 = HANA
            oCompDef.sSAPUser = String.Empty
            oCompDef.sSAPPwd = String.Empty

            oCompDef.sInboxDir = String.Empty
            oCompDef.sSuccessDir = String.Empty
            oCompDef.sFailDir = String.Empty
            oCompDef.sLogPath = String.Empty
            oCompDef.sDebug = String.Empty


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenceServer")) Then
                oCompDef.sLicenceServer = ConfigurationManager.AppSettings("LicenceServer")
            End If

            ''If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
            ''    oCompDef.sSAPDBName = ConfigurationManager.AppSettings("SAPDBName")
            ''End If

            oCompDef.sSAPDBName = ""

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.sSAPUser = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.sSAPPwd = ConfigurationManager.AppSettings("SAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DSN")) Then
                oCompDef.sDSN = ConfigurationManager.AppSettings("DSN")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sLogPath = IO.Directory.GetCurrentDirectory
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.sDebug = ConfigurationManager.AppSettings("Debug")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailFrom")) Then
                oCompDef.sEmailFrom = ConfigurationManager.AppSettings("EmailFrom")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailTo")) Then
                oCompDef.sEmailTo = ConfigurationManager.AppSettings("EmailTo")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailSubject")) Then
                oCompDef.sEmailSubject = ConfigurationManager.AppSettings("EmailSubject")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPServer")) Then
                oCompDef.sSMTPServer = ConfigurationManager.AppSettings("SMTPServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPort")) Then
                oCompDef.sSMTPPort = ConfigurationManager.AppSettings("SMTPPort")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPUser")) Then
                oCompDef.sSMTPUser = ConfigurationManager.AppSettings("SMTPUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DebugValue")) Then
                oCompDef.sSMTPPassword = ConfigurationManager.AppSettings("DebugValue")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ReportDSN")) Then
                oCompDef.sReportDSN = ConfigurationManager.AppSettings("ReportDSN")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ReportPDFPath")) Then
                oCompDef.sReportPDFPath = ConfigurationManager.AppSettings("ReportPDFPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ReportsPath")) Then
                oCompDef.sReportsPath = ConfigurationManager.AppSettings("ReportsPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBList")) Then
                oCompDef.sDBList = ConfigurationManager.AppSettings("DBList")
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

    Public Function SendEmailNotification(ByVal sfileName As String, _
                                          ByVal sSenderEmail As String, _
                                          ByVal sBody As String, _
                                          ByVal sSubject As String, _
                                          ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oSmtpServer As New SmtpClient()
        Dim oMail As New MailMessage
        Dim sEmailAddress As String = String.Empty

        Try
            sFuncName = "SendEmailNotification()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)

            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
            '--------- Message Content in HTML tags

            oSmtpServer.Credentials = New Net.NetworkCredential(p_oCompDef.sSMTPUser, p_oCompDef.sSMTPPassword)
            oSmtpServer.Port = p_oCompDef.sSMTPPort
            oSmtpServer.Host = p_oCompDef.sSMTPServer
            oSmtpServer.EnableSsl = True
            oMail.From = New MailAddress(p_oCompDef.sEmailFrom)

            Dim sSendTo As String() = sSenderEmail.Split(";")

            'oMail.To.Add(sSenderEmail)

            For Each sEmailAddress In sSendTo
                oMail.To.Add(sEmailAddress)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Email Address" & ":  " & sEmailAddress, sFuncName)
            Next

            oMail.Attachments.Add(New Attachment(sfileName))
            oMail.Subject = sSubject
            'oMail.Body = Mail_Body(sBody)
            oMail.Body = sBody
            oMail.IsBodyHtml = True

            frmSOA.WriteToStatusScreen(False, "Sending Email to :: " & sSenderEmail)

            oSmtpServer.Send(oMail)
            oMail.Dispose()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)
            frmSOA.WriteToStatusScreen(False, "Email Has been sent successfully :: " & sSenderEmail)
            SendEmailNotification = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = "Please check the email Address." & "Fail to send email : " & sSenderEmail & ":: " & ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            SendEmailNotification = RTN_ERROR
        Finally
            File.Delete(sfileName)
        End Try

    End Function

    Public Function GeneratingPDF_WithDocumentRange(ByVal iDocFrom As Integer, ByVal iDocTo As Integer, ByVal tablename As String, ByVal FileNAme As String, ByVal sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDS As New DataSet
        Dim sTargetFileName As String = String.Empty
        Dim sRptFileName As String = String.Empty

        Try
            sFuncName = "GeneratingPDF_WithDocumentRange()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            For imjs As Integer = iDocFrom To iDocTo

                sSQL = "SELECT T0.E_Mail, T0.CardCode FROM OCRD T0 WHERE T0.[CardCode] = (SELECT T0.[CardCode] FROM " & tablename & " T0 WHERE T0.[DocNum] = '" & imjs & "')"
                frmSOA.WriteToStatusScreen(False, "Executing Taxinvoice Document No. :: " & imjs)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                oDS = ExecuteSQLQuery(sSQL)
                If oDS.Tables(0).Rows.Count > 0 Then
                    If Not String.IsNullOrEmpty(oDS.Tables(0).Rows(0).Item("E_Mail").ToString) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                        frmSOA.WriteToStatusScreen(False, "Generating PDF for " & FileNAme)
                        sTargetFileName = FileNAme & "_DocNum_ " & imjs & "_AE_" & oDS.Tables(0).Rows(0).Item("CardCode").ToString & ".pdf"
                        sTargetFileName = p_oCompDef.sReportsPath & "\" & sTargetFileName
                        sRptFileName = p_oCompDef.sReportPDFPath & "\AE_RP003_TaxInvoice.rpt"

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExportToPDF" & sSQL, sFuncName)
                        frmSOA.WriteToStatusScreen(False, "Attempting Function ExportToPDF  :: ")
                        'If ExportToPDF(imjs, sTargetFileName, sRptFileName, sErrDesc) <> RTN_SUCCESS Then
                        '    Throw New ArgumentException(sErrDesc)
                        'End If
                        frmSOA.WriteToStatusScreen(False, "Successfully  generated PDF :: " & sTargetFileName)
                        ' If SendEmailNotification(sTargetFileName, oDS.Tables(0).Rows(0).Item("E_Mail").ToString, "", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email address is blank for the Customer " & oDS.Tables(0).Rows(0).Item("CardCode").ToString, sFuncName)
                        frmSOA.WriteToStatusScreen(False, "Email address is blank for the Customer :: " & oDS.Tables(0).Rows(0).Item("CardCode").ToString)
                    End If
                End If
            Next

            GeneratingPDF_WithDocumentRange = RTN_SUCCESS

        Catch ex As Exception
            frmSOA.WriteToStatusScreen(False, "Error Message  :: " & ex.Message)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GeneratingPDF_WithDocumentRange = RTN_ERROR
        End Try


    End Function

    Public Function GeneratingPDF_WithBatchNo(ByVal sBatchNo As String, ByVal DocTable As String, ByVal LineTable As String, ByVal FileName As String, ByVal sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDS As New DataSet
        Dim oDSDocNum As New DataSet
        Dim sTargetFileName As String = String.Empty
        Dim sRptFileName As String = String.Empty

        Try
            sFuncName = "GeneratingPDF_WithBatchNo()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            sSQL = "SELECT T0.[DocNum] FROM " & DocTable & " T0  INNER JOIN " & LineTable & " T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[ItemCode] = (SELECT T0.[ItemCode] FROM OIBT T0 WHERE T0.[BatchNum] = '" & sBatchNo & "') group by T0.[DocNum]"
            oDSDocNum = ExecuteSQLQuery(sSQL)
            If oDSDocNum.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In oDSDocNum.Tables(0).Rows
                    sSQL = "SELECT T0.E_Mail, T0.CardCode FROM OCRD T0 WHERE T0.[CardCode] = (SELECT T0.[CardCode] FROM " & DocTable & " T0 WHERE T0.[DocNum] = '" & row.Item(0).ToString & "')"
                    frmSOA.WriteToStatusScreen(False, "Executing Taxinvoice Document No. :: " & row.Item(0).ToString)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    oDS = ExecuteSQLQuery(sSQL)
                    If oDS.Tables(0).Rows.Count > 0 Then
                        If Not String.IsNullOrEmpty(oDS.Tables(0).Rows(0).Item("E_Mail").ToString) Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                            frmSOA.WriteToStatusScreen(False, "Generating PDF for " & FileName)
                            sTargetFileName = FileName & "_DocNum_ " & row.Item(0).ToString & "_AE_" & oDS.Tables(0).Rows(0).Item("CardCode").ToString & ".pdf"
                            sTargetFileName = p_oCompDef.sReportsPath & "\" & sTargetFileName
                            sRptFileName = p_oCompDef.sReportPDFPath & "\AE_RP003_TaxInvoice.rpt"

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExportToPDF" & sSQL, sFuncName)
                            frmSOA.WriteToStatusScreen(False, "Attempting Function ExportToPDF  :: ")
                            'If ExportToPDF(row.Item(0).ToString, sTargetFileName, sRptFileName, sErrDesc) <> RTN_SUCCESS Then
                            '    Throw New ArgumentException(sErrDesc)
                            'End If
                            frmSOA.WriteToStatusScreen(False, "Successfully  generated PDF :: " & sTargetFileName)
                            'If SendEmailNotification(sTargetFileName, oDS.Tables(0).Rows(0).Item("E_Mail").ToString, "", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email address is blank for the Customer  - Invoice No." & row.Item(0).ToString, sFuncName)
                            frmSOA.WriteToStatusScreen(False, "Email address is blank for the Customer in this Invoice No. :: " & row.Item(0).ToString)
                        End If
                    End If
                Next
            Else

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invalid Batch Number. " & sBatchNo, sFuncName)
                frmSOA.WriteToStatusScreen(False, "Invalid Batch Number :: " & sBatchNo)
            End If

            GeneratingPDF_WithBatchNo = RTN_SUCCESS

        Catch ex As Exception
            frmSOA.WriteToStatusScreen(False, "Error Message  :: " & ex.Message)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GeneratingPDF_WithBatchNo = RTN_ERROR
        End Try


    End Function

    Public Function GeneratingPDF_WithBatchNoAndDocNum(ByVal iDocNumFrom As Integer, ByVal iDocNumTo As Integer, ByVal sBatchNo As String, ByVal DocTable As String, ByVal LineTable As String, ByVal FileName As String, ByVal sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDS As New DataSet
        Dim oDSDocNum As New DataSet
        Dim sTargetFileName As String = String.Empty
        Dim sRptFileName As String = String.Empty

        Try
            sFuncName = "GeneratingPDF_WithBatchNoAndDocNum()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            sSQL = "SELECT T0.[DocNum] FROM " & DocTable & " T0  INNER JOIN " & LineTable & " T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[ItemCode] = (SELECT T0.[ItemCode] FROM OIBT T0 WHERE T0.[BatchNum] = '" & sBatchNo & "')  and T0.docnum >= '" & iDocNumFrom & "' and T0.docnum <= '" & iDocNumTo & "' group by T0.[DocNum]"
            oDSDocNum = ExecuteSQLQuery(sSQL)
            If oDSDocNum.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In oDSDocNum.Tables(0).Rows
                    sSQL = "SELECT T0.E_Mail, T0.CardCode FROM OCRD T0 WHERE T0.[CardCode] = (SELECT T0.[CardCode] FROM " & DocTable & " T0 WHERE T0.[DocNum] = '" & row.Item(0).ToString & "')"
                    frmSOA.WriteToStatusScreen(False, "Executing Taxinvoice Document No. :: " & row.Item(0).ToString)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    oDS = ExecuteSQLQuery(sSQL)
                    If oDS.Tables(0).Rows.Count > 0 Then
                        If Not String.IsNullOrEmpty(oDS.Tables(0).Rows(0).Item("E_Mail").ToString) Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                            frmSOA.WriteToStatusScreen(False, "Generating PDF for " & FileName)
                            sTargetFileName = FileName & "_DocNum_ " & row.Item(0).ToString & "_AE_" & oDS.Tables(0).Rows(0).Item("CardCode").ToString & ".pdf"
                            sTargetFileName = p_oCompDef.sReportsPath & "\" & sTargetFileName
                            sRptFileName = p_oCompDef.sReportPDFPath & "\AE_RP003_TaxInvoice.rpt"

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExportToPDF" & sSQL, sFuncName)
                            frmSOA.WriteToStatusScreen(False, "Attempting Function ExportToPDF  :: ")
                            'If ExportToPDF(row.Item(0).ToString, sTargetFileName, sRptFileName, sErrDesc) <> RTN_SUCCESS Then
                            '    Throw New ArgumentException(sErrDesc)
                            'End If
                            frmSOA.WriteToStatusScreen(False, "Successfully  generated PDF :: " & sTargetFileName)
                            'If SendEmailNotification(sTargetFileName, oDS.Tables(0).Rows(0).Item("E_Mail").ToString, "", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email address is blank for the Customer  - Invoice No." & row.Item(0).ToString, sFuncName)
                            frmSOA.WriteToStatusScreen(False, "Email address is blank for the Customer in this Invoice No. :: " & row.Item(0).ToString)
                        End If
                    End If
                Next
            Else

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("This Batch No. " & sBatchNo & " does not present in the Document range " & iDocNumFrom & " - " & iDocNumTo, sFuncName)
                frmSOA.WriteToStatusScreen(False, "This Batch No. " & sBatchNo & " does not present in the Document range :: " & iDocNumFrom & " - " & iDocNumTo)
            End If

            GeneratingPDF_WithBatchNoAndDocNum = RTN_SUCCESS

        Catch ex As Exception
            frmSOA.WriteToStatusScreen(False, "Error Message  :: " & ex.Message)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GeneratingPDF_WithBatchNoAndDocNum = RTN_ERROR
        End Try


    End Function

    Public Function GeneratingPDF_WithCardCodeRange(ByVal CardCodeFrom As String, ByVal CardCodeTo As String, ByVal _Date As Date, ByVal sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDS As New DataSet
        Dim oDSCardCode As New DataSet
        Dim sTargetFileName As String = String.Empty
        Dim sRptFileName As String = String.Empty

        Try
            sFuncName = "GeneratingPDF_WithBatchNo()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            sSQL = "SELECT T0.""CardCode"", T0.""CardName"" FROM " & """" & frmSOA.SCompany.Text & """" & ".OCRD T0 WHERE T0.""CardCode"" between '" & CardCodeFrom & "' and '" & CardCodeTo & "'"
            oDSCardCode = ExecuteSQLQuery(sSQL)
            If oDSCardCode.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In oDSCardCode.Tables(0).Rows
                    sSQL = "SELECT T0.""E_Mail"", T0.""CardCode"" FROM " & """" & frmSOA.SCompany.Text & """" & ".OCRD T0 WHERE T0.""CardCode"" = '" & row.Item(0).ToString & "'"
                    frmSOA.WriteToStatusScreen(False, "Executing SOA for BP :: " & row.Item(0).ToString)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    oDS = ExecuteSQLQuery(sSQL)
                    If oDS.Tables(0).Rows.Count > 0 Then
                        If Not String.IsNullOrEmpty(oDS.Tables(0).Rows(0).Item("E_Mail").ToString) Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                            frmSOA.WriteToStatusScreen(False, "Generating PDF for Statement of Accounts")
                            sTargetFileName = "SOA_" & row.Item(1).ToString & "_AE_" & Format(_Date, "dd/MM/yyyy") & ".pdf"
                            sTargetFileName = p_oCompDef.sReportPDFPath & "\" & sTargetFileName
                            sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_SOA.rpt"

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExportToPDF" & sSQL, sFuncName)
                            frmSOA.WriteToStatusScreen(False, "Attempting Function ExportToPDF  :: ")
                            If ExportToPDF(CardCodeFrom, CardCodeTo, _Date, sTargetFileName, sRptFileName, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                            frmSOA.WriteToStatusScreen(False, "Successfully  generated PDF :: " & sTargetFileName)
                            'If SendEmailNotification(sTargetFileName, oDS.Tables(0).Rows(0).Item("E_Mail").ToString, "", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email address is blank for the Customer " & row.Item(0).ToString, sFuncName)
                            frmSOA.WriteToStatusScreen(False, "Email address is blank for the Customer :: " & row.Item(0).ToString)
                        End If
                    End If
                Next
            Else

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Range has been found ...... ", sFuncName)
                frmSOA.WriteToStatusScreen(False, "No Range has been found ......")
            End If

            GeneratingPDF_WithCardCodeRange = RTN_SUCCESS

        Catch ex As Exception
            frmSOA.WriteToStatusScreen(False, "Error Message  :: " & ex.Message)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GeneratingPDF_WithCardCodeRange = RTN_ERROR
        End Try


    End Function

    Public Function GeneratingPDF_WithCardCode(ByVal CardCode As String, ByVal _Date As Date, ByVal sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDS As New DataSet
        Dim oDSCardCode As New DataSet
        Dim sTargetFileName As String = String.Empty
        Dim sRptFileName As String = String.Empty

        Try
            sFuncName = "GeneratingPDF_WithCardCode()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            sSQL = "SELECT T0.""E_Mail"", T0.""CardCode"" FROM OCRD T0 WHERE T0.[CardCode] = '" & CardCode & "'"
            frmSOA.WriteToStatusScreen(False, "Executing SOA for BP :: " & CardCode)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
            oDS = ExecuteSQLQuery(sSQL)
            If oDS.Tables(0).Rows.Count > 0 Then
                If Not String.IsNullOrEmpty(oDS.Tables(0).Rows(0).Item("E_Mail").ToString) Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                    frmSOA.WriteToStatusScreen(False, "Generating PDF for Statement of Accounts")
                    sTargetFileName = "SOA_" & CardCode & "_AE_" & Format(_Date, "dd/MM/yyyy") & ".pdf"
                    sTargetFileName = p_oCompDef.sReportsPath & "\" & sTargetFileName
                    sRptFileName = p_oCompDef.sReportPDFPath & "\AE_RP003_TaxInvoice.rpt"

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExportToPDF" & sSQL, sFuncName)
                    frmSOA.WriteToStatusScreen(False, "Attempting Function ExportToPDF  :: ")
                    If ExportToPDF(CardCode, CardCode, _Date, sTargetFileName, sRptFileName, sErrDesc) <> RTN_SUCCESS Then
                        Throw New ArgumentException(sErrDesc)
                    End If
                    frmSOA.WriteToStatusScreen(False, "Successfully  generated PDF :: " & sTargetFileName)
                    ' If SendEmailNotification(sTargetFileName, oDS.Tables(0).Rows(0).Item("E_Mail").ToString, "", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email address is blank for the Customer " & CardCode, sFuncName)
                    frmSOA.WriteToStatusScreen(False, "Email address is blank for the Customer :: " & CardCode)
                End If
            End If

            GeneratingPDF_WithCardCode = RTN_SUCCESS

        Catch ex As Exception
            frmSOA.WriteToStatusScreen(False, "Error Message  :: " & ex.Message)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GeneratingPDF_WithCardCode = RTN_ERROR
        End Try


    End Function

  

    ''Public Function Customer_CFL(ByVal sCondition As String, ByVal sCompanyName As String, ByRef sErrDesc As String) As Long

    ''    Try
    ''        Dim sFuncName As String = String.Empty
    ''        Dim sSQL As String = String.Empty
    ''        Dim oDS As New DataSet

    ''        sFuncName = "CustomerCFLCommonFunction() " & sCondition
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

    ''        If frmSOA.DocType.Text = "Payment Advice" Then
    ''            sSQL = "SELECT T0.""CardCode"", T0.""CardName"" FROM " & """" & sCompanyName & """" & ".OCRD T0 WHERE T0.""CardType"" = 'S' and T0.""CardCode"" <> '' ORDER BY T0.""CardName"""
    ''        Else
    ''            sSQL = "SELECT T0.""CardCode"", T0.""CardName"" FROM " & """" & sCompanyName & """" & ".OCRD T0 WHERE T0.""CardType"" = 'C' and T0.""CardCode"" <> '' ORDER BY T0.""CardName"""
    ''        End If
    ''        'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
    ''        oDS = ExecuteSQLQuery(sSQL)
    ''        CFL.CFL_Customer.DataSource = oDS.Tables(0)
    ''        CFL.CFL_Customer.Columns(1).Width = 300
    ''        CFL.TextBox1.Text = sCondition
    ''        '  CFL.CFL_Customer.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)
    ''        Customer_CFL = RTN_SUCCESS
    ''        CFL.Show()
    ''    Catch ex As Exception
    ''        Customer_CFL = RTN_ERROR
    ''    End Try
    ''End Function

    ''Public Function CustomerGroup_CFL(ByVal sCondition As String, ByVal sCompanyName As String, ByRef sErrDesc As String) As Long

    ''    Try
    ''        Dim sFuncName As String = String.Empty
    ''        Dim sSQL As String = String.Empty
    ''        Dim oDS As New DataSet
    ''        CFL_F = New CFL
    ''        sFuncName = "CustomerCFLCommonFunction() " & sCondition
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

    ''        sSQL = "SELECT T0.""GroupCode"",T0.""GroupName"" FROM " & """" & sCompanyName & """" & " .OCRG T0 WHERE T0.""GroupType"" = 'C'"
    ''        'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
    ''        oDS = ExecuteSQLQuery(sSQL)
    ''        ''CFL.CFL_Customer.DataSource = oDS.Tables(0)
    ''        ''CFL.CFL_Customer.Columns(1).Width = 300
    ''        ''CFL.TextBox1.Text = sCondition
    ''        ' ''  CFL.CFL_Customer.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)
    ''        ''CustomerGroup_CFL = RTN_SUCCESS
    ''        ''CFL.Show

    ''        CFL_F.CFL_Customer.DataSource = oDS.Tables(0)
    ''        CFL_F.CFL_Customer.Columns(1).Width = 300
    ''        CFL_F.TextBox1.Text = sCondition
    ''        '  CFL.CFL_Customer.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)
    ''        CustomerGroup_CFL = RTN_SUCCESS
    ''        ''CFL_F.Show()
    ''    Catch ex As Exception
    ''        CustomerGroup_CFL = RTN_ERROR
    ''    End Try
    ''End Function

    Public Function Mail_Body(ByVal sBody_msg As String) As String

        Try
            Dim sBody As String = String.Empty

            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
            sBody = sBody & " Dear Valued Customer,<br /><br />"
            sBody = sBody & p_SyncDateTime & " <br /><br />"
            sBody = sBody & " " & sBody_msg
            ' sBody = sBody & sErrDesc & "<br /><br />"
            sBody = sBody & "<br/> Note: This email message is computer generated and it will be used internal purpose usage only.<div/>"

            Return sBody
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return ex.Message
        End Try
    End Function

    Public Sub Write_TextFile_BPList(ByVal sBPList(,) As String, ByVal sDef As String)
        Try
            Dim irow As Integer
            Dim sPath As String = p_oCompDef.sReportPDFPath & "\"
            Dim sFileName As String = "BPList_WithoutEmail.txt"
            Dim sbuffer As String = String.Empty

            If File.Exists(sPath & sFileName) Then
                Try
                    File.Delete(sPath & sFileName)
                Catch ex As Exception
                End Try
            End If

            Dim sw As StreamWriter = New StreamWriter(sPath & sFileName)
            ' Add some text to the file.
            sw.WriteLine("")
            sw.WriteLine("List of the BP`s without Email address ")
            sw.WriteLine("")
            sw.WriteLine("")

            If sDef = "1" Then
                sw.WriteLine("Card Code           Card Name                                ")
                sw.WriteLine("=============================================================")
                sw.WriteLine(" ")

                For irow = 1 To UBound(sBPList, 1)
                    sw.WriteLine(sBPList(irow, 0).ToString.PadRight(20, " "c) & sBPList(irow, 1).ToString.PadRight(40, " "c))
                Next irow

                sw.WriteLine(" ")
                sw.WriteLine("===============================================================")
            Else
                sw.WriteLine("Card Code           Card Name                                Invoice No.")
                sw.WriteLine("========================================================================")
                sw.WriteLine(" ")

                For irow = 1 To UBound(sBPList, 1)
                    sw.WriteLine(sBPList(irow, 0).ToString.PadRight(20, " "c) & sBPList(irow, 1).ToString.PadRight(40, " "c) & " " & sBPList(irow, 2))
                Next irow

                sw.WriteLine(" ")
                sw.WriteLine("========================================================================")
            End If

            
            sw.Close()
            Process.Start(sPath & sFileName)


        Catch ex As Exception

        End Try

    End Sub

    Public Function ExecuteSQLNonQuery(ByVal sQuery As String) As Integer

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sri
        ' Date          : Nov 2013
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty
        ' Dim sConstr As String = "SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"
        Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
        Dim oCon As New Odbc.OdbcConnection(sConstr)
        Dim oCmd As New Odbc.OdbcCommand
        Dim oDs As New DataSet
        Try
            sFuncName = "ExecuteSQLNonQuery()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)

            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            If oCon.State = ConnectionState.Closed Then
                oCon.Open()
            End If
            oCmd.CommandTimeout = 0
            oCmd.ExecuteNonQuery()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            If Not oCon Is Nothing Then
                oCon.Close()
                oCon.Dispose()
            End If
        End Try

    End Function

    Public Sub GetMaxCode(ByRef iCode As Integer)
        Dim sSQL As String
        Dim oDS As New DataSet
        sSQL = "select MAX(TO_INT(""Code"")) as code from """ & sCompanyName & """ .""@SOAEMAILLOG"""
        oDS = ExecuteSQLQuery(sSQL)

        If oDS.Tables(0).Rows.Count > 0 Then
            If Not IsDBNull(oDS.Tables(0).Rows(0).Item(0)) = True Then
                iCode = oDS.Tables(0).Rows(0).Item(0) + 1
            Else
                iCode = 0 + 1
            End If
        End If
    End Sub

    Public Sub FrmPaymentadvice_Show()
        Try
            Dim bCloseApp As Boolean = False
            PaymentAdvice_F = New frmPaymentadvice
            PaymentAdvice_F.MdiParent = frmEadvice
            PaymentAdvice_F.Show()
            bCloseApp = (PaymentAdvice_F.DialogResult = DialogResult.Abort)
            PaymentAdvice_F = Nothing

            If bCloseApp Then Application.Exit()


        Catch ex As Exception

        End Try
    End Sub

    Public Sub FrmSOA_Show()
        Try
            Dim bCloseApp As Boolean = False
            SOA_F = New frmSOA
            SOA_F.MdiParent = frmEadvice
            SOA_F.Show()
            bCloseApp = (SOA_F.DialogResult = DialogResult.Abort)
            SOA_F = Nothing

            If bCloseApp Then Application.Exit()

        Catch ex As Exception

        End Try
    End Sub


End Module
