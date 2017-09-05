
Module modMain

#Region "Variables"

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

        Public sCreditNoteGL As String
        Public sNonStockItem As String
        Public sFFSItemCode As String
        ''New Items codes
        Public sFFSItemCodeNonPanel As String
        Public s3FSItemCode As String
        Public s3FSItemCodeNonPanel As String
        '' 

        Public sCAPItemCode As String
        Public sTPAItemCode As String
        Public sFFSGLCode As String
        Public sCAPGLCode As String
        Public s3FSGLCode As String
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
        Public sGIROGLAccountAIA As String

        Public sGJ_CheckGLAccount As String
        Public sGJ_GIROGLAccount As String
        Public sGJ_CheckBankAccount As String
        Public sGJ_CheckBankCode As String
        Public sGJ_FFSGLCode As String
        Public sGJ_CAPGLCode As String

        Public p_sSMSUserName As String
        Public p_sSMSPassword As String
        Public p_sSMSFrom As String
        Public p_sGIROSMS As String
        Public p_sCheckSMS As String

        Public sDBS_CheckGLAccount As String
        Public sDBS_CheckBankAccount As String
        Public sDBS_CheckBankCode As String

        Public sDBS_AONCheckGLAccount As String
        Public sDBS_AONCheckBankAccount As String
        Public sDBS_AONCheckBankCode As String

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
    Public sGJDBName As String = String.Empty
    Public sGJCostCenter As String = String.Empty
    Public p_sPatientType As String = String.Empty
    Public p_oDtSMS As DataTable

#End Region

    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty

        Try
            sFuncName = "MBMS Interface Synchronization"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Console.Title = "MBMS Interface Synchronization"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            Start()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End
        End Try
    End Sub
  
End Module
