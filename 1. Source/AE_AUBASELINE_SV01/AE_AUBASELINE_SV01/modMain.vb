Module modMain

#Region "Variables"

    ' Company Default Structure
    Public Structure CompanyDefault
        Public sServer As String
        Public sLicenceServer As String
        Public iServerLanguage As Integer
        Public iServerType As Integer
        Public sSAPUser As String
        Public sSAPPwd As String
        Public sSAPDBName As String
        Public sDBUser As String
        Public sDBPwd As String

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
    Public p_oCompany As SAPbobsCOM.Company

    Public p_oDtSuccess As DataTable
    Public p_oDtError As DataTable
    Public p_oDtReport As DataTable

    Public p_SyncDateTime As String

#End Region
#Region "Main"
    Sub Main()
        Dim sFuncName As String = "Main()"
        Dim sErrDesc As String = String.Empty
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            Console.WriteLine("Getting Initial Database connections")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("System Initialization", sFuncName)
            If GetCompanyInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sFuncName)

            Console.WriteLine("Starting Baseline interface")

            Start()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End
        End Try
    End Sub
#End Region

End Module
