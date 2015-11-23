Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Data.Common
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.IO
Imports System.Threading

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

        Public sLogPath As String
        Public sDebug As String

        Public sOrgID As String
        Public sSenderName As String

        Public p_sSuccessDir As String
        Public p_sPaymentDir As String

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
    Private oRecordSet As SAPbobsCOM.Recordset
    Public oCompList As New Dictionary(Of String, String)
    Public sTemplateType As String = String.Empty



    Public Function ExecuteSQLQuery(ByVal sQuery As String) As DataTable

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

            oCompDef.sSAPDBName = String.Empty
            oCompDef.sLogPath = String.Empty
            oCompDef.sDebug = String.Empty
            oCompDef.p_sSuccessDir = String.Empty
            oCompDef.p_sPaymentDir = String.Empty


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

            oCompDef.sLogPath = IO.Directory.GetCurrentDirectory

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.sDebug = ConfigurationManager.AppSettings("Debug")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("OrganizationID")) Then
                oCompDef.sOrgID = ConfigurationManager.AppSettings("OrganizationID")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SenderName")) Then
                oCompDef.sSenderName = ConfigurationManager.AppSettings("SenderName")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("PaymentDir")) Then
                oCompDef.p_sPaymentDir = ConfigurationManager.AppSettings("PaymentDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SuccessDir")) Then
                oCompDef.p_sSuccessDir = ConfigurationManager.AppSettings("SuccessDir")
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

    Public Function Get_CompanyList(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   Get_CompanyList()
        '   Purpose     :   This function will provide the company list 
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   October 2014
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty


        Try

            sFuncName = "Get_CompanyList()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            oCompany.UseTrusted = True
            oCompany.DbUserName = p_oCompDef.sDBUser
            oCompany.DbPassword = p_oCompDef.sDBPwd
            oCompany.Server = p_oCompDef.sServer
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            oRecordSet = oCompany.GetCompanyList

            'Dim s As String = oRecordSet.Fields.Item(0).Name
            'Dim s1 As String = oRecordSet.Fields.Item(1).Name

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Get_CompanyList = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Get_CompanyList = RTN_ERROR
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

    Public Function MoveAllItemsTo(ByVal fromPathInfo As DirectoryInfo, ByVal toPath As String _
                                   , ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "MoveAllItemsTo()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            ''Create the target directory if necessary
            Dim toPathInfo = New DirectoryInfo(toPath)
            If (Not toPathInfo.Exists) Then
                toPathInfo.Create()
            End If
            ''move all files
            For Each file As FileInfo In fromPathInfo.GetFiles()
                file.MoveTo(Path.Combine(toPath, file.Name))
            Next
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            MoveAllItemsTo = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            MoveAllItemsTo = RTN_ERROR
        End Try

    End Function

    Public Function Encryption(ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "FileMoveToArchive"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.p_sPaymentDir)
            Dim files() As System.IO.FileInfo

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting all the files from the Input folder", sFuncName)
            files = DirInfo.GetFiles("*.txt")

            For Each oFile As System.IO.FileInfo In files

                ''  Console.WriteLine("Encrypting the file... File Name : " & oFile.Name)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Encrypting the file... File Name : " & oFile.Name, sFuncName)

                Dim strFileName As String = "--encrypt " & oFile.FullName
                Process.Start("C:\SAP\Bank file - DBS Encryption\IDEALConnect\IDEALConnect.exe", strFileName)

                Thread.Sleep(15000)

                ' Console.WriteLine("Successfully encrypted the file. File Name : " & oFile.Name)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully encrypted the file. File Name : " & oFile.Name, sFuncName)

                ' Console.WriteLine("Moving the file to process folder")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToLocation()", sFuncName)

                If FileMoveToLocation(oFile.Directory.FullName & "\" & oFile.Name, p_oCompDef.p_sSuccessDir & "\" & oFile.Name, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                'Console.WriteLine("Successfully moved to prcess folder")

                ' Console.WriteLine("          ")

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            Encryption = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            Encryption = RTN_ERROR

        End Try

    End Function

    Public Function FileMoveToLocation(ByVal sFileToMove As String, ByVal sMoveToFile As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "FileMoveToArchive"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            File.Move(sFileToMove, sMoveToFile)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            FileMoveToLocation = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in renaming/copying/moving", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            FileMoveToLocation = RTN_ERROR
        End Try
    End Function

End Module
