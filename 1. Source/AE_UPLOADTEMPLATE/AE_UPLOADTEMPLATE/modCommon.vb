Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Data.Common
Imports System.IO


Module modCommon

    ' Company Default Structure
    Public Structure CompanyDefault
        Public sDBName As String
        Public sServer As String
        Public sLicenceServer As String
        Public iServerLanguage As Integer
        Public iServerType As Integer
        Public sSAPUser As String
        Public sSAPPwd As String
        Public sSAPDBName As String
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

        Public sAR_BUPAGL As String
        Public sAP_BUPAGL As String
        Public sAR_AETNAGL As String
        Public sAP_AETNAGL As String

        Public sCountryCode As String
        Public sBankCode As String
        Public sCheckBankAccount As String
        Public sExRateDiffAccount As String
        Public sJECostingCode As String

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
    Public p_oCompany As SAPbobsCOM.Company

    Public p_oDtReceiptLog As DataTable



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

            oCompDef.sDBName = String.Empty
            oCompDef.sServer = String.Empty
            oCompDef.iServerLanguage = 3
            oCompDef.iServerType = 9
            oCompDef.sSAPUser = String.Empty
            oCompDef.sSAPPwd = String.Empty
            oCompDef.sSAPDBName = String.Empty
            oCompDef.sCountryCode = String.Empty
            oCompDef.sBankCode = String.Empty
            oCompDef.sCheckBankAccount = String.Empty
            oCompDef.sSuccessDir = String.Empty
            oCompDef.sExRateDiffAccount = String.Empty
            oCompDef.sJECostingCode = String.Empty

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenceServer")) Then
                oCompDef.sLicenceServer = ConfigurationManager.AppSettings("LicenceServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
                oCompDef.sSAPDBName = ConfigurationManager.AppSettings("SAPDBName")
            End If

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

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SuccessDir")) Then
                oCompDef.sSuccessDir = ConfigurationManager.AppSettings("SuccessDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("AR_BUPAGL")) Then
                oCompDef.sAR_BUPAGL = ConfigurationManager.AppSettings("AR_BUPAGL")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("AP_BUPAGL")) Then
                oCompDef.sAP_BUPAGL = ConfigurationManager.AppSettings("AP_BUPAGL")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("AR_AETNAGL")) Then
                oCompDef.sAR_AETNAGL = ConfigurationManager.AppSettings("AR_AETNAGL")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("AP_AETNAGL")) Then
                oCompDef.sAP_AETNAGL = ConfigurationManager.AppSettings("AP_AETNAGL")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CountryCode")) Then
                oCompDef.sCountryCode = ConfigurationManager.AppSettings("CountryCode")
            End If
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("BankCode")) Then
                oCompDef.sBankCode = ConfigurationManager.AppSettings("BankCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CheckBankAccount")) Then
                oCompDef.sCheckBankAccount = ConfigurationManager.AppSettings("CheckBankAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ExRateDiffAccount")) Then
                oCompDef.sExRateDiffAccount = ConfigurationManager.AppSettings("ExRateDiffAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("JECostingCode")) Then
                oCompDef.sJECostingCode = ConfigurationManager.AppSettings("JECostingCode")
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

    Public Function GetDataViewFromExcel(ByVal CurrFileToUpload As String, ByVal sSheet As String, ByRef sErrdesc As String) As DataView

        'Event      :   GetDataViewFromExcel
        'Purpose    :   For reading of CSV file
        'Author     :   Sri 
        'Date       :   22 Nov 2013 

        Dim k As Integer
        Dim sFileExtension As String = String.Empty
        Dim sConnectionString As String

        k = Microsoft.VisualBasic.InStrRev(CurrFileToUpload, ".")

        sFileExtension = Microsoft.VisualBasic.Right(CurrFileToUpload, Len(CurrFileToUpload) - k).Trim

        If sFileExtension = "xls" Then
            sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CurrFileToUpload & ";Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"""
        Else
            sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"""
        End If

        Dim objConn As New System.Data.OleDb.OleDbConnection(sConnectionString)
        Dim da As OleDb.OleDbDataAdapter
        Dim dt As DataTable
        Dim dv As DataView
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "GetDataViewFromExcel"
            'Open Data Adapter to Read from Text file
            da = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & sSheet & "$]", objConn)
            dt = New DataTable("NonFHN3")

            'Fill dataset using dataadapter
            da.Fill(dt)
            dv = New DataView(dt)
            Return dv

        Catch ex As Exception
            sErrdesc = ex.Message
            Return Nothing
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try

    End Function

    Public Function ConnectToCompany(ByRef oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String, Optional ByVal sDBName As String = "") As Long
        ' **********************************************************************************
        '   Function    :   ConnectToCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   October 2013
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Try
            sFuncName = "ConnectToCompany()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
            oCompany = New SAPbobsCOM.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)

            oCompany.Server = p_oCompDef.sServer
            ' oCompany.Server = "192.168.11.35"

            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB

            oCompany.CompanyDB = p_oCompDef.sSAPDBName
            oCompany.UserName = p_oCompDef.sSAPUser
            oCompany.Password = p_oCompDef.sSAPPwd

            oCompany.LicenseServer = p_oCompDef.sLicenceServer

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            oCompany.UseTrusted = False

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
            iRetValue = oCompany.Connect()

            If iRetValue <> 0 Then
                oCompany.GetLastError(iErrCode, sErrDesc)

                sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                    oCompany.CompanyDB, System.Environment.NewLine, _
                                vbTab, sErrDesc)

                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConnectToCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectToCompany = RTN_ERROR
        End Try
    End Function

    Public Function StartTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    StartTransaction()
        '   Purpose    :    Start DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :   Sri
        '   Date       :   29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "StartTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oCompany.InTransaction Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Found hanging transaction.Rolling it back.", sFuncName)
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            p_oCompany.StartTransaction()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            StartTransaction = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            StartTransaction = RTN_ERROR
        End Try

    End Function

    Public Function RollBackTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    RollBackTransaction()
        '   Purpose    :    Roll Back DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Sri
        '   Date       :    29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "RollBackTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oCompany.InTransaction Then
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for rollback", sFuncName)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            RollBackTransaction = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            RollBackTransaction = RTN_ERROR
        Finally
            GC.Collect()
        End Try

    End Function

    Public Function CommitTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    CommitTransaction()
        '   Purpose    :    Commit DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc=Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Sri
        '   Date       :    29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "CommitTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oCompany.InTransaction Then
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for commit", sFuncName)
            End If

            CommitTransaction = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            CommitTransaction = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
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

    Public Function IsFileNameExists(ByVal sType As String, ByVal sFileName As String) As Boolean
        Dim bIsExists As Boolean = False
        Dim sSQL As String
        Dim oDS As New DataSet

        If sType = "AP" Then
            sSQL = " SELECT ""U_AI_APARUploadName"" from OPCH " & _
                   " WHERE ""U_AI_APARUploadName"" = '" & sFileName & "'"

            oDS = ExecuteSQLQuery(sSQL)

            If oDS.Tables(0).Rows.Count > 0 Then bIsExists = True

        End If

        If sType = "AR" Then
            sSQL = " SELECT ""U_AI_APARUploadName"" from OINV " & _
                   " WHERE ""U_AI_APARUploadName"" = '" & sFileName & "'"

            oDS = ExecuteSQLQuery(sSQL)

            If oDS.Tables(0).Rows.Count > 0 Then bIsExists = True

        End If

        Return bIsExists

    End Function

    Public Function ExecuteSQLQueryDataTabe(ByVal sQuery As String) As DataTable

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
        Return oDs.Tables(0)
    End Function

    Public Function CreateDataTable(ByVal ParamArray oColumnName() As String) As System.Data.DataTable

        Dim oDataTable As New System.Data.DataTable

        Dim oDataColumn As System.Data.DataColumn

        For i As Integer = LBound(oColumnName) To UBound(oColumnName)
            oDataColumn = New System.Data.DataColumn()
            oDataColumn.DataType = Type.GetType("System.String")
            oDataColumn.ColumnName = oColumnName(i).ToString
            oDataTable.Columns.Add(oDataColumn)
        Next

        Return oDataTable

    End Function

    Public Sub AddDataToTable(ByVal oDt As System.Data.DataTable, ByVal ParamArray sColumnValue() As String)
        Dim oRow As System.Data.DataRow = Nothing
        oRow = oDt.NewRow()
        For i As Integer = LBound(sColumnValue) To UBound(sColumnValue)
            oRow(i) = sColumnValue(i).ToString
        Next
        oDt.Rows.Add(oRow)
    End Sub

    Public Function Write_TextFile(ByVal oDT_FinalResult As System.Data.DataTable, _
                                       ByVal sFileName As String, _
                                       ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim sPath As String = String.Empty

        Try

            sFuncName = "Write_TextFile()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            'If sType.ToString.ToUpper().Trim() = "COA" Then
            '    sPath = System.Windows.Forms.Application.StartupPath & "\" & sFileName
            'ElseIf sType.ToString.ToUpper().Trim() = "OUSR" Then
            '    sPath = System.Windows.Forms.Application.StartupPath & "\" & sFileName
            'End If

            sPath = System.Windows.Forms.Application.StartupPath & "\" & sFileName

            If File.Exists(sPath) Then
                Try
                    File.Delete(sPath)
                Catch ex As Exception
                End Try
            End If

            Dim sw As StreamWriter = New StreamWriter(sPath)


            sw.WriteLine(oDT_FinalResult.Columns(0).ColumnName.ToString().Trim() & _
                         oDT_FinalResult.Columns(1).ColumnName.ToString().Trim().PadLeft(40, " "c) & _
                         oDT_FinalResult.Columns(2).ColumnName.ToString().Trim().PadLeft(40, " "c) & _
                         oDT_FinalResult.Columns(3).ColumnName.ToString().Trim().PadLeft(40, " "c))


            sw.WriteLine("===========================================================================================================================================")
            sw.WriteLine("                                                                                                                                           ")

            ' Add some text to the file.

            For imjs As Integer = 0 To oDT_FinalResult.Rows.Count - 1

                sw.WriteLine(oDT_FinalResult.Rows(imjs).Item(0).ToString.Trim() & _
                             oDT_FinalResult.Rows(imjs).Item(1).ToString.Trim().PadLeft(40 + (oDT_FinalResult.Columns(1).ColumnName.ToString().Trim().Length), " "c) & _
                             oDT_FinalResult.Rows(imjs).Item(2).ToString.Trim().PadLeft(20 + (oDT_FinalResult.Columns(1).ColumnName.ToString().Trim().Length), " "c) & _
                             oDT_FinalResult.Rows(imjs).Item(3).ToString.Trim().PadLeft(110 + (oDT_FinalResult.Columns(3).ColumnName.ToString().Trim().Length), " "c))

            Next imjs

            sw.WriteLine("                                                                                                                                             ")
            sw.WriteLine("=============================================================================================================================================")
            sw.WriteLine("                                                                                                                                             ")

            sw.Close()
            'If scheck = "Y" Then
            '    Process.Start(sPAth & sFileName)
            'End If

            Process.Start(sPath)

            Write_TextFile = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

        Catch ex As Exception
            Write_TextFile = RTN_ERROR
            sErrDesc = ex.Message

            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

End Module


