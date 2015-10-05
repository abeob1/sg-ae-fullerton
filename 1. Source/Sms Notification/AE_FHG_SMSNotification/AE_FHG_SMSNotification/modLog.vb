Option Explicit On

Imports System.IO
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Data.Common

Module modLog

    Public Structure CompanyDefault

        Public p_sServerName As String
        Public p_sLicServerName As String
        Public p_sDataBaseName As String

        Public p_sDBUserName As String
        Public p_sDBPassword As String
        Public p_sSAPUserName As String
        Public p_sSAPPassword As String
        Public p_sSQLType As String

        Public p_sLogDir As String
        Public p_sDebug As String

        Public p_sSMSUserName As String
        Public p_sSMSPassword As String
        Public p_sSMSFrom As String
        Public p_sGIROSMS As String
        Public p_sCheckSMS As String
        Public p_sSendURL As String
        Public p_sStatusURL As String

    End Structure


    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    ' Debug Value Variable Control
    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    ' Global variables group
    Public p_iDebugMode As Int16
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16
    Public p_oCompDef As CompanyDefault

    ' Public p_oCompany As SAPbobsCOM.Company
    Public p_oDataView As DataView = Nothing
    Public P_sConString As String = String.Empty
    Public p_sSAPConnString As String = String.Empty
    Public p_oCompany As SAPbobsCOM.Company
    Public P_sCardCode As String = String.Empty

    '***************************************
    'Name       :   modLog
    'Descrption :   Contains function for log errors and Application related information
    'Author     :   JOHN
    'Created    :   MAY 2014
    '***************************************

    Private Const MAXFILESIZE_IN_MB As Int16 = 5 '(2 MB)
    Private Const LOG_FILE_ERROR As String = "ErrorLog"
    Private Const LOG_FILE_ERROR_ARCH As String = "ErrorLog_"
    Private Const LOG_FILE_DEBUG As String = "DebugLog"
    Private Const LOG_FILE_DEBUG_ARCH As String = "DebugLog_"
    Private Const FILE_SIZE_CHECK_ENABLE As Int16 = 1
    Private Const FILE_SIZE_CHECK_DISABLE As Int16 = 0
    Public Const sTitle As String = "SMS Notification"

    Public Function WriteToLogFile(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long

        ' **********************************************************************************
        '   Function   :    WriteToLogFile()
        '   Purpose    :    This function checks if given input file name exists or not
        '
        '   Parameters :    ByVal strErrText As String
        '                       strErrText = Text to be written to the log
        '                   ByVal intLogType As Integer
        '                       intLogType = Log type (1 - Log ; 2 - Error ; 0 - None)
        '                   ByVal strSourceName As String
        '                       strSourceName = Function name calling this function
        '                   Optional ByVal intCheckFileForDelete As Integer
        '                       intCheckFileForDelete = Flag to indicate if file size need to be checked before logging (0 - No check ; 1 - Check)
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    JOHN
        '   Date       :    MAY 2014
        ' **********************************************************************************

        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String = String.Empty
        Dim strArchFileName As String = String.Empty
        Dim strTempString As String = String.Empty
        Dim lngFileSizeInMB As Double

        Try
            strTempString = Space(IIf(Len(strSourceName) > 30, 0, 30 - Len(strSourceName)))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText


            strFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_ERROR & ".log"
            strArchFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_ERROR_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"


            'strFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_ERROR & ".log"
            'strArchFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_ERROR_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        'If intCheckDeleteDebugLog=1 then remove all debug_log file
                        If p_iDeleteDebugLog = 1 Then
                            For Each sFileName As String In Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), LOG_FILE_DEBUG_ARCH & "*")
                                File.Delete(sFileName)
                            Next
                        End If
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function

    Public Function WriteToLogFile_Debug(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long
        ' **********************************************************************************
        '   Function   :    WriteToLogFile_Debug()
        '   Purpose    :    This function checks if given input file name exists or not
        '
        '   Parameters :    ByVal strErrText As String
        '                       strErrText = Text to be written to the log
        '                   ByVal intLogType As Integer
        '                       intLogType = Log type (1 - Log ; 2 - Error ; 0 - None)
        '                   ByVal strSourceName As String
        '                       strSourceName = Function name calling this function
        '                   Optional ByVal intCheckFileForDelete As Integer
        '                       intCheckFileForDelete = Flag to indicate if file size need to be checked before logging (0 - No check ; 1 - Check)
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    JOHN
        '   Date       :    MAY 2013
        '   Changes    : 
        '                   
        ' **********************************************************************************

        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String = String.Empty
        Dim strArchFileName As String = String.Empty
        Dim strTempString As String = String.Empty
        Dim lngFileSizeInMB As Double
        Dim iFileCount As Integer = 0

        Try
            strTempString = Space(IIf(Len(strSourceName) > 30, 0, 30 - Len(strSourceName)))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText

            strFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_DEBUG & ".log"
            strArchFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_DEBUG_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"


            'strFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_DEBUG & ".log"
            'strArchFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_DEBUG_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        'If intCheckDeleteDebugLog=1 then remove all debug_log file
                        If p_iDeleteDebugLog = 1 Then
                            For Each sFileName As String In Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), LOG_FILE_DEBUG_ARCH & "*")
                                File.Delete(sFileName)
                            Next
                        End If
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile_Debug = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile_Debug = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function

    Public Function GetSystemIntializeInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemIntializeInfo()
        '   Purpose     :   This function will be providing information about the initialing variables
        '               
        '   Parameters  :   ByRef oCompDef As CompanyDefault
        '                       oCompDef =  set the Company Default structure
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   Srinivasan
        '   Date        :   JUNE 2015
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim sSqlstr As String = String.Empty
        Try

            sFuncName = "GetSystemIntializeInfo()"
            '' frmSendSMS.WriteToStatusScreen(False, "Starting Function", sFuncName, sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)


            oCompDef.p_sServerName = String.Empty
            oCompDef.p_sLicServerName = String.Empty
            oCompDef.p_sDBUserName = String.Empty
            oCompDef.p_sDBPassword = String.Empty
            oCompDef.p_sSQLType = String.Empty

            oCompDef.p_sDataBaseName = String.Empty
            oCompDef.p_sSAPUserName = String.Empty
            oCompDef.p_sSAPPassword = String.Empty

            oCompDef.p_sLogDir = String.Empty
            oCompDef.p_sDebug = String.Empty

            oCompDef.p_sSMSUserName = String.Empty
            oCompDef.p_sSMSPassword = String.Empty
            oCompDef.p_sSMSFrom = String.Empty
            oCompDef.p_sGIROSMS = String.Empty
            oCompDef.p_sCheckSMS = String.Empty
            oCompDef.p_sSendURL = String.Empty
            oCompDef.p_sStatusURL = String.Empty


            P_sCardCode = String.Empty


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.p_sServerName = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenseServer")) Then
                oCompDef.p_sLicServerName = ConfigurationManager.AppSettings("LicenseServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
                oCompDef.p_sDataBaseName = ConfigurationManager.AppSettings("SAPDBName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.p_sSAPUserName = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.p_sSAPPassword = ConfigurationManager.AppSettings("SAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.p_sDBUserName = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.p_sDBPassword = ConfigurationManager.AppSettings("DBPwd")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SQLType")) Then
                oCompDef.p_sSQLType = ConfigurationManager.AppSettings("SQLType")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CardCode")) Then
                P_sCardCode = ConfigurationManager.AppSettings("CardCode")
            End If


            ' folder
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogDir")) Then
                oCompDef.p_sLogDir = ConfigurationManager.AppSettings("LogDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.p_sDebug = ConfigurationManager.AppSettings("Debug")
                If p_oCompDef.p_sDebug.ToUpper = "ON" Then
                    p_iDebugMode = 1
                Else
                    p_iDebugMode = 0
                End If
            Else
                p_iDebugMode = 0
            End If


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

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SendURL")) Then
                oCompDef.p_sSendURL = ConfigurationManager.AppSettings("SendURL")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("StatusURL")) Then
                oCompDef.p_sStatusURL = ConfigurationManager.AppSettings("StatusURL")
            End If

            ''frmSendSMS.WriteToStatusScreen(False, "Completed with SUCCESS", sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            '' frmSendSMS.WriteToStatusScreen(False, "Completed with ERROR. Error :" & sErrDesc, sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

    Function Get_DataSet(ByVal sQueryString As String, ByVal sConnString As String, ByRef sErrDesc As String) As DataSet

        Dim oDataSet As DataSet
        Dim oDataAdapter As SqlDataAdapter
        Dim oDataTable As DataTable
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "Get_DataSet()"

            frmSendSMS.WriteToStatusScreen(False, "Starting Function", sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query : " & sQueryString, sFuncName)

            oDataSet = New DataSet

            'To Get the Dataset Based on the Query String :

            oDataAdapter = New SqlDataAdapter(sQueryString, sConnString)
            oDataTable = New DataTable
            oDataAdapter.Fill(oDataTable)
            oDataSet.Tables.Add(oDataTable)

            If oDataSet.Tables(0).Rows.Count > 0 Then
                frmSendSMS.WriteToStatusScreen(False, "Completed with SUCCESS", sFuncName, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                Return oDataSet
            Else
                frmSendSMS.WriteToStatusScreen(False, "Completed with SUCCESS", sFuncName, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                Return Nothing
            End If


        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            frmSendSMS.WriteToStatusScreen(False, "Completed with ERROR. Error : " & sErrDesc, sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return Nothing
        End Try
    End Function

    Public Function ExecuteSQLQueryForDT(ByVal sQuery As String, ByVal sSAPDBName As String) As DataTable

        '**************************************************************
        ' Function      : ExecuteSQLQueryForDT
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Srinvasan
        ' Date          : JULY 2015
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty

        Dim oCmd As New Odbc.OdbcCommand
        Dim oDs As New DataSet
        Dim oDbProviderFactoryObject As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim oCon As DbConnection = oDbProviderFactoryObject.CreateConnection()

        Try
            sFuncName = "ExecuteSQLQueryForDT()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)
            oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.p_sDBUserName & ";PWD=" & p_oCompDef.p_sDBPassword & ";SERVERNODE=" & p_oCompDef.p_sServerName & ";CS=" & sSAPDBName

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing Query : " & sQuery, sFuncName)

            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(oCmd)
            da.Fill(oDs)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
        Return oDs.Tables(0)
    End Function

    Public Function ExecuteSQLNonQuery(ByVal sQuery As String) As Integer

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sriniasan
        ' Date          : Jul 2015
        ' Change        :
        '**************************************************************


        Dim sFuncName As String = String.Empty

        Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.p_sDBUserName & ";PWD=" & p_oCompDef.p_sDBPassword & ";SERVERNODE=" & p_oCompDef.p_sServerName & ";CS=" & p_oCompDef.p_sDataBaseName
        Dim oCon As New Odbc.OdbcConnection(sConstr)
        Dim oCmd As New Odbc.OdbcCommand
        Dim oDs As New DataSet
        Try
            sFuncName = "ExecuteSQLNonQuery()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing Query : " & sQuery, sFuncName)

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

End Module
