Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Data.Common
Imports System.Threading

Module modCommon

#Region "Connection Object [Connect to DI Company]"

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

            oCompDef.sInboxDir = String.Empty
            oCompDef.sSuccessDir = String.Empty
            oCompDef.sFailDir = String.Empty
            oCompDef.sLogPath = String.Empty
            oCompDef.sDebug = String.Empty

            oCompDef.p_sSMSUserName = String.Empty
            oCompDef.p_sSMSPassword = String.Empty
            oCompDef.p_sSMSFrom = String.Empty
            oCompDef.p_sGIROSMS = String.Empty
            oCompDef.p_sCheckSMS = String.Empty

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

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DSN")) Then
                oCompDef.sDSN = ConfigurationManager.AppSettings("DSN")
            End If

            ' folder
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("InboxDir")) Then
                oCompDef.sInboxDir = ConfigurationManager.AppSettings("InboxDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SuccessDir")) Then
                oCompDef.sSuccessDir = ConfigurationManager.AppSettings("SuccessDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FailDir")) Then
                oCompDef.sFailDir = ConfigurationManager.AppSettings("FailDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sLogPath = ConfigurationManager.AppSettings("LogPath")
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

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPassword")) Then
                oCompDef.sSMTPPassword = ConfigurationManager.AppSettings("SMTPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CreditNoteGL")) Then
                oCompDef.sCreditNoteGL = ConfigurationManager.AppSettings("CreditNoteGL")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("NonStockItem")) Then
                oCompDef.sNonStockItem = ConfigurationManager.AppSettings("NonStockItem")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FFSItemCode")) Then
                oCompDef.sFFSItemCode = ConfigurationManager.AppSettings("FFSItemCode")
            End If


            ' New Items
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FFSItemCodeNonPanel")) Then
                oCompDef.sFFSItemCodeNonPanel = ConfigurationManager.AppSettings("FFSItemCodeNonPanel")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("3FSItemCode")) Then
                oCompDef.s3FSItemCode = ConfigurationManager.AppSettings("3FSItemCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("3FSItemCodeNonPanel")) Then
                oCompDef.s3FSItemCodeNonPanel = ConfigurationManager.AppSettings("3FSItemCodeNonPanel")
            End If

            ' ****************************


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CAPItemCode")) Then
                oCompDef.sCAPItemCode = ConfigurationManager.AppSettings("CAPItemCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("TPAItemCode")) Then
                oCompDef.sTPAItemCode = ConfigurationManager.AppSettings("TPAItemCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FFSGLCode")) Then
                oCompDef.sFFSGLCode = ConfigurationManager.AppSettings("FFSGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CAPGLCode")) Then
                oCompDef.sCAPGLCode = ConfigurationManager.AppSettings("CAPGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("3FSGLCode")) Then
                oCompDef.s3FSGLCode = ConfigurationManager.AppSettings("3FSGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CAPGLCode")) Then
                oCompDef.sCAPGLCode = ConfigurationManager.AppSettings("CAPGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DefaultCostCenter")) Then
                oCompDef.sDefaultCostCenter = ConfigurationManager.AppSettings("DefaultCostCenter")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ServiceFee")) Then
                oCompDef.dServiceFee = CDbl(ConfigurationManager.AppSettings("ServiceFee"))
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CustBPSeriesName")) Then
                oCompDef.sCustBPSeriesName = ConfigurationManager.AppSettings("CustBPSeriesName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("VenBPSeriesName")) Then
                oCompDef.sVenBPSeriesName = ConfigurationManager.AppSettings("VenBPSeriesName")
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

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CheckGLAccount")) Then
                oCompDef.sCheckGLAccount = ConfigurationManager.AppSettings("CheckGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GIROGLAccount")) Then
                oCompDef.sGIROGLAccount = ConfigurationManager.AppSettings("GIROGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CheckBankAccount")) Then
                oCompDef.sCheckBankAccount = ConfigurationManager.AppSettings("CheckBankAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CheckBankCode")) Then
                oCompDef.sCheckBankCode = ConfigurationManager.AppSettings("CheckBankCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GIROGLAccountAIA")) Then
                oCompDef.sGIROGLAccountAIA = ConfigurationManager.AppSettings("GIROGLAccountAIA")
            End If

            'GJ Start
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_CheckGLAccount")) Then
                oCompDef.sGJ_CheckGLAccount = ConfigurationManager.AppSettings("GJ_CheckGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_GIROGLAccount")) Then
                oCompDef.sGJ_GIROGLAccount = ConfigurationManager.AppSettings("GJ_GIROGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_CheckBankAccount")) Then
                oCompDef.sGJ_CheckBankAccount = ConfigurationManager.AppSettings("GJ_CheckBankAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_CheckBankCode")) Then
                oCompDef.sGJ_CheckBankCode = ConfigurationManager.AppSettings("GJ_CheckBankCode")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_FFSGLCode")) Then
                oCompDef.sGJ_FFSGLCode = ConfigurationManager.AppSettings("GJ_FFSGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_CAPGLCode")) Then
                oCompDef.sGJ_CAPGLCode = ConfigurationManager.AppSettings("GJ_CAPGLCode")
            End If

            'GJ End

            'SMS Credential
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMSUserName")) Then
                oCompDef.p_sSMSUserName = ConfigurationManager.AppSettings("SMSUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMSPassword")) Then
                oCompDef.p_sSMSPassword = ConfigurationManager.AppSettings("SMSPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMSFrom")) Then
                oCompDef.p_sSMSFrom = ConfigurationManager.AppSettings("SMSFrom")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GIROSMS")) Then
                oCompDef.p_sGIROSMS = ConfigurationManager.AppSettings("GIROSMS")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CHECKSMS")) Then
                oCompDef.p_sCheckSMS = ConfigurationManager.AppSettings("CHECKSMS")
            End If

            'dbs

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBS_CheckGLAccount")) Then
                oCompDef.sDBS_CheckGLAccount = ConfigurationManager.AppSettings("DBS_CheckGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBS_CheckBankAccount")) Then
                oCompDef.sDBS_CheckBankAccount = ConfigurationManager.AppSettings("DBS_CheckBankAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBS_CheckBankCode")) Then
                oCompDef.sDBS_CheckBankCode = ConfigurationManager.AppSettings("DBS_CheckBankCode")
            End If

            'dbs AON

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBS_AONCheckGLAccount")) Then
                oCompDef.sDBS_AONCheckGLAccount = ConfigurationManager.AppSettings("DBS_AONCheckGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBS_AONCheckBankAccount")) Then
                oCompDef.sDBS_AONCheckBankAccount = ConfigurationManager.AppSettings("DBS_AONCheckBankAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBS_CheckBankCode")) Then
                oCompDef.sDBS_AONCheckBankCode = ConfigurationManager.AppSettings("DBS_CheckBankCode")
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

#End Region

    Public Function ExecuteSQLQuery(ByVal sQuery As String) As DataSet

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sri
        ' Date          : Nov 2013
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty
        'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

        Dim oCmd As New Odbc.OdbcCommand
        Dim oDs As New DataSet
        Dim oDbProviderFactoryObject As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim oCon As DbConnection = oDbProviderFactoryObject.CreateConnection()

        Try
            sFuncName = "ExecuteQuery()"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)
            'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
            oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName

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

    Public Function ExecuteProcedureForDataSet(ByVal spName As String, ByVal ParamArray parameters() As SqlParameter) As DataSet

        '**************************************************************
        ' Function      : ExecuteProcedureForDataSet
        ' Purpose       : Execute Procedure and returns Dataset
        ' Parameters    : ByVal spName - string command Text
        '               : Byval parameters ParamArray 
        ' Author        : Sri
        ' Date          : March 07 2008
        ' Change        :
        '**************************************************************

        'p_oCompDef

        'Dim sConstr As String = "Data Source=" & ConfigurationManager.AppSettings("Server") & ";Initial Catalog=" & ConfigurationManager.AppSettings("DBName") & ";User ID=" & ConfigurationManager.AppSettings("SQLUser") & "; Password=" & ConfigurationManager.AppSettings("SQLPwd")

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "ExecuteProcedureForDataSet()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)

            oCmd.CommandType = CommandType.StoredProcedure
            oCmd.CommandText = spName
            oCmd.Connection = oCon
            If parameters.Length > 0 Then
                Dim p As SqlParameter
                For Each p In parameters
                    If Not p Is Nothing Then
                        oCmd.Parameters.Add(p)
                    End If
                Next
            End If
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDs)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Throw New Exception(ex.Message)
        Finally
            If Not oCon Is Nothing Then
                oCon.Close()
                oCon.Dispose()
            End If
        End Try
        Return oDs

    End Function

    Public Function ExecuteProcedureForNonQuery(ByVal spName As String, ByVal ParamArray parameters() As SqlParameter) As Integer

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand

        Try
            oCmd.CommandType = CommandType.StoredProcedure
            oCmd.CommandText = spName
            oCmd.Connection = oCon
            If oCon.State = ConnectionState.Closed Then
                oCon.Open()
            End If
            If parameters.Length > 0 Then
                Dim p As SqlParameter
                For Each p In parameters
                    If Not p Is Nothing Then
                        oCmd.Parameters.Add(p)
                    End If
                Next
            End If
            oCmd.CommandTimeout = 0
            oCmd.ExecuteNonQuery()

        Catch ex As Exception
            Throw ex
        Finally
            If Not oCon Is Nothing Then
                oCon.Close()
                oCon.Dispose()
            End If
        End Try
    End Function

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

    Public Function CreateDataTable(ByVal ParamArray oColumnName() As String) As DataTable
        Dim oDataTable As DataTable = New DataTable()

        Dim oDataColumn As DataColumn

        For i As Integer = LBound(oColumnName) To UBound(oColumnName)
            oDataColumn = New DataColumn()
            oDataColumn.DataType = Type.GetType("System.String")
            oDataColumn.ColumnName = oColumnName(i).ToString
            oDataTable.Columns.Add(oDataColumn)
        Next

        Return oDataTable

    End Function

    Public Sub AddDataToTable(ByVal oDt As DataTable, ByVal ParamArray sColumnValue() As String)
        Dim oRow As DataRow = Nothing
        oRow = oDt.NewRow()
        For i As Integer = LBound(sColumnValue) To UBound(sColumnValue)
            oRow(i) = sColumnValue(i).ToString
        Next
        oDt.Rows.Add(oRow)
    End Sub

    Public Function CreateTable_ReimbursetoProvider(ByVal sTableName As String, ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "CreateTable_ReimbursetoProvider()"
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "CALL " & """" & p_oCompDef.sSAPDBName & """" & "." & "AE_SP001_IsTableExists('" & sTableName & "','" & p_oCompDef.sSAPDBName & "')"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL : " & sSQL, sFuncName)
            oRS.DoQuery(sSQL)

            sSQL = " CREATE COLUMN TABLE " & """" & p_oCompDef.sSAPDBName & """" & "." & """" & sTableName & """" & "(""VISITNO"" NVARCHAR(50),""VISITDATE"" DATE,""COMPANYNAME"" NVARCHAR(100),"
            sSQL = sSQL & """COMPANYCODE"" NVARCHAR(100),""PATIENTNAME"" NVARCHAR(100),""PATIENTID"" NVARCHAR(100),""PATIENTMEMBERTYPE"" NVARCHAR(100),"
            sSQL = sSQL & """CONTRACTTYPE"" NVARCHAR(50),""BENEFITTYPE"" NVARCHAR(100),""PROVIDERNAME"" NVARCHAR(250),""ADDRESSLINE1"" NVARCHAR(250),"
            sSQL = sSQL & """ADDRESSLINE2"" NVARCHAR(250),""ADDRESSLINE3"" NVARCHAR(250),""PROVIDERCOUNTRY"" NVARCHAR(100),""POSTALCODE"" NVARCHAR(20),"
            sSQL = sSQL & """EMAIL"" NVARCHAR(150),""CURRENCY"" NVARCHAR(10),""CONSULTCOST"" DECIMAL(34,0) CS_FIXED,""DRUGCOST"" DECIMAL(34,0) CS_FIXED,"
            sSQL = sSQL & """SERVICECOST"" DECIMAL(34,0) CS_FIXED,""SUBTOTAL"" DECIMAL(34,0) CS_FIXED,""TAX"" DECIMAL(34,0) CS_FIXED,""GRADTOTAL"" DECIMAL(34,0) CS_FIXED,"
            sSQL = sSQL & """UNCLAIMAMT"" DECIMAL(34,0) CS_FIXED,""CLAIMAMT"" DECIMAL(34,0) CS_FIXED,""TPAFEE"" DECIMAL(34,0) CS_FIXED,""TPAFEETAX"" DECIMAL(34,0) CS_FIXED,"
            sSQL = sSQL & """TPAFEETOTAL"" DECIMAL(34,0) CS_FIXED,""REIMBURSEMENTAMT"" DECIMAL(34,0) CS_FIXED,""PAYMENTMODE"" NVARCHAR(100),""PAYEENAME"" NVARCHAR(250),"
            sSQL = sSQL & """PAYEEACNO"" NVARCHAR(200),""REMARKS"" NVARCHAR(500))"


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL : " & sSQL, sFuncName)
            oRS.DoQuery(sSQL)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed successfully.", sFuncName)
            CreateTable_ReimbursetoProvider = RTN_SUCCESS

        Catch ex As Exception
            sErrdesc = ex.Message
            CreateTable_ReimbursetoProvider = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Error Creating Table:" & ex.Message, sFuncName)
        Finally
            oRS = Nothing
        End Try

    End Function

    Public Function CreateTable_ReimbursetoMember(ByVal sTableName As String, ByRef sErrdesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset

        Try
            sFuncName = "CreateTable_ReimbursetoMember()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "CALL " & "AE_SP001_IsTableExists('" & sTableName & "','" & p_oCompDef.sSAPDBName & "')"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL : " & sSQL, sFuncName)
            oRS.DoQuery(sSQL)

            sSQL = " CREATE COLUMN TABLE " & """" & sTableName & """" & "(""VISITNO"" NVARCHAR(50),""VISITDATE"" DATE,""INVOICENO"" NVARCHAR(100),""COMPANYNAME"" NVARCHAR(100),"
            sSQL = sSQL & """COMPANYCODE"" NVARCHAR(100),""EMPLOYEENAME"" NVARCHAR(100),""EMPLOYEEID"" NVARCHAR(100),""CLAIMANTNAME"" NVARCHAR(100), ""CLAIMANTID"" NVARCHAR(100),"
            sSQL = sSQL & """CLAIMANTMEMBERTYPE"" NVARCHAR(100),""CONTRACTTYPE"" NVARCHAR(50),""BENEFITTYPE"" NVARCHAR(100),""PROVIDERNAME"" NVARCHAR(250),""ADDRESSLINE1"" NVARCHAR(250),"
            sSQL = sSQL & """ADDRESSLINE2"" NVARCHAR(250),""ADDRESSLINE3"" NVARCHAR(250),""MEMBERCOUNTRY"" NVARCHAR(100),""POSTALCODE"" NVARCHAR(20),"
            sSQL = sSQL & """EMAIL"" NVARCHAR(150),""CURRENCY"" NVARCHAR(10),""CONSULTCOST"" DECIMAL(34,0) CS_FIXED,""DRUGCOST"" DECIMAL(34,0) CS_FIXED,"
            sSQL = sSQL & """SERVICECOST"" DECIMAL(34,0) CS_FIXED,""OTHERCOST"" DECIMAL(34,0) CS_FIXED,""SUBTOTAL"" DECIMAL(34,0) CS_FIXED,""TAX"" DECIMAL(34,0) CS_FIXED,""GRADTOTAL"" DECIMAL(34,0) CS_FIXED,"
            sSQL = sSQL & """UNCLAIMAMT"" DECIMAL(34,0) CS_FIXED,""CLAIMAMT"" DECIMAL(34,0) CS_FIXED,""REIMBURSEMENTAMT"" DECIMAL(34,0) CS_FIXED,""PAYMENTMODE"" NVARCHAR(100),"
            sSQL = sSQL & """PAYEENAME"" NVARCHAR(250),""PAYEEACNO"" NVARCHAR(200),""REMARKS"" NVARCHAR(500))"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL : " & sSQL, sFuncName)
            oRS.DoQuery(sSQL)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed successfully.", sFuncName)
            CreateTable_ReimbursetoMember = RTN_SUCCESS

        Catch ex As Exception
            sErrdesc = ex.Message
            CreateTable_ReimbursetoMember = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Error Creating Table:" & ex.Message, sFuncName)
        Finally
            oRS = Nothing
        End Try

    End Function

    Public Function CreateTable_BillToClientFHG(ByVal sTableName As String, _
                                                ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset

        Try
            sFuncName = "CreateTable_BillToClientFHG"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "CALL " & "AE_SP001_IsTableExists('" & sTableName & "','" & p_oCompDef.sSAPDBName & "')"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL : " & sSQL, sFuncName)
            oRS.DoQuery(sSQL)

            sSQL = " CREATE COLUMN TABLE " & """" & sTableName & """" & "(""VISITNO"" NVARCHAR(50),""VISITDATE"" DATE,""COMPANYNAME"" NVARCHAR(100),"
            sSQL = sSQL & """ENTITY"" NVARCHAR(250),""DEPARTMENT"" NVARCHAR(100),""COSTCENTER"" NVARCHAR(100),""EMPLOYEENAME"" NVARCHAR(250),"
            sSQL = sSQL & """EMPLOYEEID"" NVARCHAR(100),""PATIENTNAME"" NVARCHAR(150),""PATIENTID"" NVARCHAR(100),""PATIENTMEMBERTYPE"" NVARCHAR(100),"
            sSQL = sSQL & """CONTRACTTYPE"" NVARCHAR(50),""PROVIDERNAME"" NVARCHAR(250),""CURRENCY"" NVARCHAR(10),""CONSULTCOST"" DECIMAL(34,0) CS_FIXED,""DRUGCOST"" DECIMAL(34,0) CS_FIXED,""INHOUSESERVICECOST"" DECIMAL(34,0) CS_FIXED,"
            sSQL = sSQL & """EXTERNALSERVICECOST"" DECIMAL(34,0) CS_FIXED,""SUBTOTAL"" DECIMAL(34,0) CS_FIXED,""TAX"" DECIMAL(34,0) CS_FIXED,""GRADTOTAL"" DECIMAL(34,0) CS_FIXED,"
            sSQL = sSQL & """UNCLAIMAMT"" DECIMAL(34,0) CS_FIXED,""CLAIMAMT"" DECIMAL(34,0) CS_FIXED)"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL : " & sSQL, sFuncName)
            oRS.DoQuery(sSQL)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed successfully.", sFuncName)
            CreateTable_BillToClientFHG = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            CreateTable_BillToClientFHG = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Error Creating Table:" & ex.Message, sFuncName)
        Finally
            oRS = Nothing
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

    Public Function GetDataViewFromExcel(ByVal CurrFileToUpload As String, ByVal sSheet As String) As DataView

        'Event      :   GetDataViewFromExcel
        'Purpose    :   For reading of CSV file
        'Author     :   Sri 
        'Date       :   22 Nov 2013 

        'Dim sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CurrFileToUpload & ";Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"""

        Dim sConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties=""Excel 12.0;HDR=NO;IMEX=1"""


        Dim objConn As New System.Data.OleDb.OleDbConnection(sConnectionString)
        Dim da As OleDb.OleDbDataAdapter
        Dim dt As DataTable
        Dim dv As DataView
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "GetDataViewFromExcel"
            'Open Data Adapter to Read from Text file
            da = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & sSheet & "$]", objConn)
            dt = New DataTable("BilltoClient")

            'Fill dataset using dataadapter
            da.Fill(dt)
            dv = New DataView(dt)
            Return dv

        Catch ex As Exception
            Return Nothing
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try

    End Function

    Public Function IsBatchNoExists(ByVal sBatchNo As String) As Boolean

        Dim bIsExists As Boolean = False
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = "SELECT CAST(""ImportEnt"" AS VARCHAR(20)) AS ""Batch"" from OINV WHERE ""ImportEnt""=" & sBatchNo & _
               " UNION ALL " & _
               " SELECT CAST(""ImportEnt"" AS VARCHAR(20)) AS ""Batch"" from OPCH WHERE ""ImportEnt""=" & sBatchNo & _
               " UNION ALL " & _
               " SELECT CAST(""ImportEnt"" AS VARCHAR(20)) AS ""Batch"" from ORIN WHERE ""ImportEnt""=" & sBatchNo & _
               " UNION ALL " & _
               " SELECT CAST(""ImportEnt"" AS VARCHAR(20)) AS ""Batch"" from ORPC WHERE ""ImportEnt""=" & sBatchNo & _
               " UNION ALL " & _
              " SELECT CAST(""CounterRef"" AS VARCHAR(20)) As ""Batch"" from OVPM WHERE ""CounterRef""='" & sBatchNo & "'"

        oDS = ExecuteSQLQuery(sSQL)

        If oDS.Tables(0).Rows.Count > 0 Then bIsExists = True
        Return bIsExists

    End Function

    Public Function IsProviderBatchNoExists(ByVal sBatchNo As String, ByVal sProviderName As String) As Boolean
        Dim bIsExists As Boolean = False
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = " SELECT ""NumAtCard"" from OPCH " & _
               " WHERE ""NumAtCard"" = '" & "Batch-" & sBatchNo & "' and UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sProviderName), "'", "''") & "'"

        oDS = ExecuteSQLQuery(sSQL)

        If oDS.Tables(0).Rows.Count > 0 Then bIsExists = True
        Return bIsExists

    End Function

    Public Sub GetDefaultBankDetails(ByRef sDfltBankCode As String, ByRef sDfltBankAcct As String)
        Dim sSQL As String
        Dim oDS As New DataSet
        sSQL = "select ""DflBnkCode"",""DflBnkAcct"" from OADM"
        oDS = ExecuteSQLQuery(sSQL)

        If oDS.Tables(0).Rows.Count > 0 Then
            sDfltBankCode = oDS.Tables(0).Rows(0).Item("DflBnkCode").ToString
            sDfltBankAcct = oDS.Tables(0).Rows(0).Item("DflBnkAcct").ToString
        End If


    End Sub

    Public Sub GetBillingType(ByVal sCardName As String, ByRef sBillType As String)
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = "Select ifnull(""U_AI_MBMSBilling"",'CO') AS U_AI_MBMSBilling,""CardCode"" from ""OCRD"" Where UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sCardName), "'", "''") & "'"
        oDS = ExecuteSQLQuery(sSQL)

        For Each row As DataRow In oDS.Tables(0).Rows
            If Left(row.Item("CardCode").ToString, 1) = "M" Then
                sBillType = row.Item("U_AI_MBMSBilling").ToString
                Exit For
            Else
                sBillType = row.Item("U_AI_MBMSBilling").ToString
            End If
        Next
    End Sub

    Public Function IsInhouseClinic(ByVal sProviderName As String) As Boolean
        Dim sSQL As String
        Dim oDS As New DataSet
        Dim bInhouse As Boolean

        sSQL = "Select ifnull(""U_AI_INHOUSECLINIC"",'N') AS U_AI_INHOUSECLINIC,""CardCode"" from ""OCRD"" Where UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sProviderName), "'", "''") & "'"
        oDS = ExecuteSQLQuery(sSQL)

        For Each row As DataRow In oDS.Tables(0).Rows
            If row.Item("U_AI_INHOUSECLINIC").ToString = "Y" Then
                bInhouse = True
            Else
                bInhouse = False
            End If
        Next

        Return bInhouse

    End Function


    Public Function CheckCreateSO(ByVal sCardCode As String) As Boolean
        Dim sSQL As String
        Dim oDS As New DataSet
        Dim bSO As Boolean = False
        sSQL = "Select ifnull(""U_AI_MBMS"",'INV') AS CheckSO,""CardCode"" from ""OCRD"" Where ifnull(""U_AI_MBMS"",'INV') = 'SO' AND ""CardCode""='" & sCardCode & "'"
        oDS = ExecuteSQLQuery(sSQL)

        For Each row As DataRow In oDS.Tables(0).Rows
            If row.Item("CheckSO").ToString = "SO" Then
                bSO = True
            Else
                bSO = False
            End If
        Next

        Return bSO

    End Function

    Public Sub GetMaxCode(ByRef iCode As Integer)
        Dim sSQL As String
        Dim oDS As New DataSet
        sSQL = "select MAX(TO_INT(""Code"")) as code from ""@AI_TB02_PROVIDERS"""
        oDS = ExecuteSQLQuery(sSQL)

        If oDS.Tables(0).Rows.Count > 0 Then
            If Not IsDBNull(oDS.Tables(0).Rows(0).Item(0)) = True Then
                iCode = oDS.Tables(0).Rows(0).Item(0)
            Else
                iCode = 0
            End If
        End If
    End Sub

    Public Function GetDBName(ByVal sName As String) As String
        'get db name
        Dim sSQL As String
        Dim oDS As New DataSet
        Dim sFuncName As String = "GetDBName"
        Dim sDBName As String = String.Empty

        sSQL = "SELECT * FROM ""@AI_TB01_COMPANYDATA""  WHERE ""U_DBNAME"" ='" & sName & "'"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)

        oDS = ExecuteSQLQuery(sSQL)

        If oDS.Tables(0).Rows.Count > 0 Then
            sDBName = oDS.Tables(0).Rows(0).Item("Name").ToString
        End If

        Return sDBName
    End Function

    Public Sub ReadContractOwner(ByVal oDv As DataView, ByVal sSheet As String, ByVal sFileName As String, ByRef sDBName As String)

        Dim sGJName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim k As Integer
        Dim sFuncName As String = "ReadContractOwner"

        oDv = GetDataViewFromExcel(sFileName, sSheet)
        If IsNothing(oDv) Then
            Exit Sub
        End If

        sGJName = oDv(0)(0).ToString()

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading ContractOwner.", sFuncName)
        k = InStrRev(sGJName, ":")
        sGJName = Microsoft.VisualBasic.Right(sGJName, Len(sGJName) - k).Trim
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Contract Owner:" & sGJName, sFuncName)

        If sGJName = "Gethin-Jones Medical Practice Pte Ltd" Then
            sGJDBName = sGJName
            sDBName = GetDBName(sGJName)
        End If
    End Sub

    Function SendSMS(ByVal dtSMSHeader As DataTable, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oWebClient As New System.Net.WebClient
        Dim sUserName As String = String.Empty
        Dim sPassword As String = String.Empty
        Dim sMobileNo As String = String.Empty
        Dim sFrom As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sDocNum As String = String.Empty
        Dim sAmount As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim oDTSMS As DataTable = New DataTable
        Dim sReturnMsg As String = String.Empty
        Dim sStatusUrl As String = String.Empty
        Dim sStatusQStr As String = String.Empty
        Dim sMessageID As String = String.Empty

        Try

            sFuncName = "SendSMS()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sUserName = p_oCompDef.p_sSMSUserName
            sPassword = p_oCompDef.p_sSMSPassword
            sFrom = p_oCompDef.p_sSMSFrom

            oDTSMS = CreateDataTable("DocNum", "MobileNo", "MessgaeID")

            For iRow As Integer = 0 To dtSMSHeader.Rows.Count - 1

                sDocNum = dtSMSHeader.Rows(iRow)(0).ToString().Trim()
                sMobileNo = dtSMSHeader.Rows(iRow)(1).ToString().Trim()
                sAmount = dtSMSHeader.Rows(iRow)(2).ToString().Trim()

                If sMobileNo.ToString() = String.Empty Then
                    Call WriteToLogFile("Check Document Number : " & sDocNum & ". Mobile Number is Blank!", sFuncName)
                    Continue For

                End If

                Dim sUrl As String = "http://mx.fortdigital.net/http/send-message?username={0}&password={1}&to=%2B65{2}&from={3}&message={4}"

                sMessage = String.Format(p_oCompDef.p_sGIROSMS, sAmount)

                Dim QStr As String = String.Empty
                QStr = String.Format(sUrl, sUserName, sPassword, sMobileNo, sFrom, sMessage)
                oWebClient.Encoding = System.Text.Encoding.ASCII
                oWebClient.UseDefaultCredentials = False

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sending SMS for Mobile NO." & sMobileNo, sFuncName)

                sReturnMsg = oWebClient.DownloadString(QStr)

                AddDataToTable(oDTSMS, sDocNum, sMobileNo, Mid(sReturnMsg, 5, sReturnMsg.Length - 8))

            Next

            Thread.Sleep(30000)

            For iRow As Integer = 0 To oDTSMS.Rows.Count - 1

                ''GETTING THE STATUS FOR SEND SMS
                sDocNum = oDTSMS.Rows(iRow)(0).ToString().Trim()
                sMobileNo = oDTSMS.Rows(iRow)(1).ToString().Trim()
                sMessageID = oDTSMS.Rows(iRow)(2).ToString().Trim()

                sStatusUrl = "http://mx.fortdigital.net/http/request-status-update?username={0}&password={1}&message-id={2}"

                sStatusQStr = String.Format(sStatusUrl, sUserName, sPassword, sMessageID)
                oWebClient.Encoding = System.Text.Encoding.ASCII
                oWebClient.UseDefaultCredentials = False

                ' System.Diagnostics.EventLog.WriteEntry("AE_FHG_SMSNotification", "SMS Status")

                Dim sStatus As String = String.Empty

                sStatus = oWebClient.DownloadString(sStatusQStr)

                If Mid(sStatus, 1, 7).ToString().ToUpper() = "SUCCESS" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Sent SMS to mobile No." & sMobileNo, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery for Update the Flag. Mobile NO." & sMobileNo, sFuncName)
                    sQuery = "UPDATE OVPM SET ""U_AI_SentSMS""='Y' WHERE ""DocNum""='" & sDocNum & "'"
                    ExecuteSQLNonQuery(sQuery)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sStatus & ".Mobile No." & sMobileNo, sFuncName)
                End If

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SendSMS = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SendSMS = RTN_ERROR

        End Try

    End Function

    Public Function CheckContractOwner(ByVal sValue As String) As Boolean

        Dim sSQL As String
        Dim oDS As New DataSet
        Dim bCtrOwner As Boolean = False

        sSQL = "SELECT * FROM ""@AI_CONTRACTOWNER"" where UCASE(""U_Contractowner"")='" & Replace(Microsoft.VisualBasic.UCase(sValue), "'", "''") & "'"

        oDS = ExecuteSQLQuery(sSQL)
        If oDS.Tables(0).Rows.Count > 0 Then
            bCtrOwner = True
        End If

        Return bCtrOwner

    End Function

    Public Sub GetBrokerCode(ByVal sValue As String, ByRef sBrokerName As String, ByRef sDelMode As String)
        Dim sSQL As String
        Dim oDS As New DataSet

        sSQL = "SELECT * FROM ""@AE_BROKERSETUP"""
        oDS = ExecuteSQLQuery(sSQL)

        For Each row As DataRow In oDS.Tables(0).Rows
            If row.Item("U_AE_BrokerCode").ToString.Contains(sValue) Then
                sBrokerName = row.Item("U_AE_BrokerName").ToString
                sDelMode = row.Item("U_AE_DelMode").ToString
            Else
                sDelMode = "M"
            End If
        Next
    End Sub

    Public Function GetInvoiceRef(ByVal sValue As String) As String
        Dim sSQL As String
        Dim oDS As New DataSet
        Dim sInvRef As String = String.Empty

        sSQL = "Select ""CardCode"",""U_AI_INVAMTREF"" from ""OCRD"" Where UCASE(""CardName"")='" & Replace(Microsoft.VisualBasic.UCase(sValue.Trim), "'", "''") & "'"
        oDS = ExecuteSQLQuery(sSQL)
        If oDS.Tables(0).Rows.Count > 0 Then
            sInvRef = oDS.Tables(0).Rows(0).Item("U_AI_INVAMTREF").ToString
        End If

        Return sInvRef

    End Function

End Module
