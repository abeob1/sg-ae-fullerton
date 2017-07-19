Module Module1
    Public p_oApps As SAPbouiCOM.SboGuiApi
    Public sFuncName As String
    Public WithEvents p_oSBOApplication As SAPbouiCOM.Application
    Public p_iDebugMode As Int16
    Public p_oUICompany As SAPbouiCOM.Company
    Public p_oDICompany As SAPbobsCOM.Company
    Public p_iDeleteDebugLog As Int16
    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    Public sErrDesc As String
    Public p_oEventHandler As clsEventhandler
    Public Const ERR_DISPLAY_STATUS As Int16 = 1
    Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2

    Public Const DEBUG_ON As Int16 = 1
    Public p_iErrDispMethod As Int16
    Public sHanaPassword As String


    <STAThread()>
    Sub Main(ByVal args() As String)
        Dim sconn As String = String.Empty
        sFuncName = "Main()"

        Try
            p_iDebugMode = DEBUG_ON
            p_iErrDispMethod = ERR_DISPLAY_STATUS

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Addon startup function", sFuncName)
            p_oApps = New SAPbouiCOM.SboGuiApi
            p_oApps.Connect(args(0))

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing public SBO Application object", sFuncName)
            p_oSBOApplication = p_oApps.GetApplication

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO application company handle", sFuncName)
            p_oUICompany = p_oSBOApplication.Company

            p_oDICompany = New SAPbobsCOM.Company
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retrived SBO application company handle", sFuncName)
            p_oDICompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
            If Not p_oDICompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Event handler class", sFuncName)
            p_oEventHandler = New clsEventHandler(p_oSBOApplication, p_oDICompany)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddMenuItems()", sFuncName)
            p_oEventHandler.AddMenuItems()

            If Not String.IsNullOrEmpty(Configuration.ConfigurationManager.AppSettings("HANAPassword")) Then
                sHanaPassword = Configuration.ConfigurationManager.AppSettings("HANAPassword")
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating tables", sFuncName)
            CreateTable()

            AddUserQuery("CoCodeList")
            AssignFMS("CoCodeList", "10")
            AssignFMS("CoCodeList", "12")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            p_oSBOApplication.StatusBar.SetText("Addon Started Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try


    End Sub

    Private Sub CreateTable()
        Try
            'CREATE UDO TABLE FOR ENTITY LIST
            CreateUDOTable("AE_ENTITYLIST", "AE_ENTITYLIST", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            addField("@AE_ENTITYLIST", "GROUPCODE", "GROUP CODE", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("@AE_ENTITYLIST", "ENTITYCODE", "ENTITY CODE", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("@AE_ENTITYLIST", "ENTITYNAME", "ENTITY DESCRIPTION", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("@AE_ENTITYLIST", "ANAPLANCODE", "ANAPLANCODE", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub

End Module
