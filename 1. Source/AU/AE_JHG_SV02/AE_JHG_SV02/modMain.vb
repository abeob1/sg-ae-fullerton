
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

        Public sCashAccount As String
        Public sTransferAccount As String
        Public sCheckAccount As String
        Public sBankCode As String

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
    Public p_SyncDateTime As String
    Public p_oCompany As SAPbobsCOM.Company



#End Region

    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty

        Try
            sFuncName = "INFORCARE Interface Synchronization"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Console.Title = "INFORCARE Interface Synchronization"
            
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("==================== SATART =" & "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "] ===============================", sFuncName)

            Start()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("==================== END ====" & "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "] ================================", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End
        End Try
    End Sub
  
End Module
