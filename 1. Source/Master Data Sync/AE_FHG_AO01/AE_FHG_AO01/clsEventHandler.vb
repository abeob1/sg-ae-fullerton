Option Explicit On
Imports System.Windows.Forms

Namespace AE_FHG_AO01
    Public Class clsEventHandler
        Dim WithEvents SBO_Application As SAPbouiCOM.Application ' holds connection with SBO
        Dim p_oDICompany As New SAPbobsCOM.Company

        Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)
            Dim sFuncName As String = String.Empty
            Try
                sFuncName = "Class_Initialize()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO Application handle", sFuncName)
                SBO_Application = oApplication
                p_oDICompany = oCompany

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Call WriteToLogFile(exc.Message, sFuncName)
            End Try
        End Sub

        Public Function SetApplication(ByRef sErrDesc As String) As Long
            ' **********************************************************************************
            '   Function   :    SetApplication()
            '   Purpose    :    This function will be calling to initialize the default settings
            '                   such as Retrieving the Company Default settings, Creating Menus, and
            '                   Initialize the Event Filters
            '               
            '   Parameters :    ByRef sErrDesc AS string
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "SetApplication()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetMenus()", sFuncName)
                If SetMenus(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetFilters()", sFuncName)
                If SetFilters(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                SetApplication = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(exc.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SetApplication = RTN_ERROR
            End Try
        End Function

        Private Function SetMenus(ByRef sErrDesc As String) As Long
            ' **********************************************************************************
            '   Function   :    SetMenus()
            '   Purpose    :    This function will be gathering to create the customized menu
            '               
            '   Parameters :    ByRef sErrDesc AS string
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty
            ' Dim oMenuItem As SAPbouiCOM.MenuItem
            Try
                sFuncName = "SetMenus()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                SetMenus = RTN_SUCCESS
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SetMenus = RTN_ERROR
            End Try
        End Function

        Private Function SetFilters(ByRef sErrDesc As String) As Long

            ' **********************************************************************************
            '   Function   :    SetFilters()
            '   Purpose    :    This function will be gathering to declare the event filter 
            '                   before starting the AddOn Application
            '               
            '   Parameters :    ByRef sErrDesc AS string
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            ' **********************************************************************************

            Dim oFilters As SAPbouiCOM.EventFilters
            Dim oFilter As SAPbouiCOM.EventFilter
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "SetFilters()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing EventFilters object", sFuncName)
                oFilters = New SAPbouiCOM.EventFilters



                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding filters", sFuncName)
                SBO_Application.SetFilter(oFilters)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                SetFilters = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SetFilters = RTN_ERROR
            End Try
        End Function

        Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_AppEvent()
            '   Purpose    :    This function will be handling the SAP Application Event
            '               
            '   Parameters :    ByVal EventType As SAPbouiCOM.BoAppEventTypes
            '                       EventType = set the SAP UI Application Eveny Object        
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty
            Dim sErrDesc As String = String.Empty
            Dim sMessage As String = String.Empty

            Try
                sFuncName = "SBO_Application_AppEvent()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                Select Case EventType
                    Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                        sMessage = String.Format("Please wait for a while to disconnect the AddOn {0} ....", System.Windows.Forms.Application.ProductName)
                        p_oSBOApplication.SetStatusBarMessage(sMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        End
                End Select

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                sErrDesc = ex.Message
                WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                ShowErr(sErrDesc)
            Finally
                GC.Collect()  'Forces garbage collection of all generations.
            End Try
        End Sub

        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_MenuEvent()
            '   Purpose    :    This function will be handling the SAP Menu Event
            '               
            '   Parameters :    ByRef pVal As SAPbouiCOM.MenuEvent
            '                       pVal = set the SAP UI MenuEvent Object
            '                   ByRef BubbleEvent As Boolean
            '                       BubbleEvent = set the True/False        
            ' **********************************************************************************
            ' Dim oForm As SAPbouiCOM.Form = Nothing
            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim oForm As SAPbouiCOM.Form = Nothing
            Try
                sFuncName = "SBO_Application_MenuEvent()"
                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If

                If pVal.BeforeAction = False Then
                    Select Case pVal.MenuUID

                        Case "MDS"
                            Try
                                LoadFromXML("MasterSync.srf", SBO_Application)
                                oForm = p_oSBOApplication.Forms.Item("MSD")
                                oForm.Items.Item("Item_3").Visible = False
                                oForm.Items.Item("Item_4").Visible = False
                                oForm.Visible = True
                                If EntityLoad(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                Exit Try
                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End Try
                            Exit Sub
                    End Select
                End If

                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                BubbleEvent = False
                ShowErr(exc.Message)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
            End Try
        End Sub

        Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
                ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_ItemEvent()
            '   Purpose    :    This function will be handling the SAP Menu Event
            '               
            '   Parameters :    ByVal FormUID As String
            '                       FormUID = set the FormUID
            '                   ByRef pVal As SAPbouiCOM.ItemEvent
            '                       pVal = set the SAP UI ItemEvent Object
            '                   ByRef BubbleEvent As Boolean
            '                       BubbleEvent = set the True/False        
            ' **********************************************************************************

            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim p_oDVJE As DataView = Nothing
            Dim oDTDistinct As DataTable = Nothing
            Dim oDTRowFilter As DataTable = Nothing
            Dim oMatrix As SAPbouiCOM.Matrix
            Try
                sFuncName = "SBO_Application_ItemEvent()"
                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not IsNothing(p_oDICompany) Then
                    If Not p_oDICompany.Connected Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                        If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If

                If pVal.BeforeAction = False Then

                    Select Case pVal.FormUID

                        Case "MSD"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.ItemUID = "Item_0" Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Clear_Matrix()", sFuncName)
                                If Clear_Matrix(oForm, "5", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                If oForm.Items.Item("Item_0").Specific.selected.description = "COA" Then
                                    If p_iCOACount = 0 Then
                                        If AddChooseFromList_COA(oForm, "1", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    If CFL_DataBindingUser(oForm, "AcctCode", "CFL1", "CFL2", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                ElseIf oForm.Items.Item("Item_0").Specific.selected.description = "OUSR" Then
                                    If p_iUserCount = 0 Then
                                        If AddChooseFromList_User(oForm, "12", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    If CFL_DataBindingUser(oForm, "USER_CODE", "CFL3", "CFL4", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvento = pVal
                                Dim sCFL_ID As String
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                                Try
                                    If oCFLEvento.BeforeAction = False Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects

                                        If oForm.Items.Item("Item_0").Specific.selected.description = "COA" Then
                                            If pVal.ItemUID = "txtCode" Then
                                                oForm.Items.Item("txtCode").Specific.string = oDataTable.GetValue("AcctCode", 0)
                                            ElseIf pVal.ItemUID = "Item_1" Then
                                                oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("AcctCode", 0)
                                            End If
                                        ElseIf oForm.Items.Item("Item_0").Specific.selected.description = "OUSR" Then
                                            If pVal.ItemUID = "txtCode" Then
                                                oForm.Items.Item("txtCode").Specific.string = oDataTable.GetValue("USER_CODE", 0)
                                            ElseIf pVal.ItemUID = "Item_1" Then
                                                oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("USER_CODE", 0)
                                            End If
                                        End If
                                    End If

                                Catch ex As Exception

                                End Try
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                                If pVal.ItemUID = "Item_0" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                                    oForm.Items.Item("txtCode").Specific.string = String.Empty
                                    oForm.Items.Item("Item_1").Specific.string = String.Empty

                                End If
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                Dim oCheckCol As SAPbouiCOM.CheckBox
                                Try
                                    oMatrix = oForm.Items.Item("5").Specific
                                    If pVal.ItemUID = "5" And pVal.ColUID = "V_1" Then
                                        oCheckCol = oMatrix.Columns.Item("V_1").Cells.Item(1).Specific

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
                                        Call DisplayStatus(oForm, "Please wait... while selcting/Deselcting the Entities", sErrDesc)

                                        oForm.Freeze(True)

                                        If oCheckCol.Checked = True Then

                                            For iRow As Integer = 1 To oMatrix.RowCount
                                                oCheckCol = oMatrix.Columns.Item("V_1").Cells.Item(iRow).Specific
                                                oCheckCol.Checked = False
                                            Next
                                        Else
                                            For iRow As Integer = 1 To oMatrix.RowCount
                                                oCheckCol = oMatrix.Columns.Item("V_1").Cells.Item(iRow).Specific
                                                oCheckCol.Checked = True
                                            Next
                                        End If

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Clear_Matrix()", sFuncName)
                                        If Clear_Matrix(oForm, "5", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                        oForm.Freeze(False)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                                        Call EndStatus(sErrDesc)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                                    Call EndStatus(sErrDesc)
                                End Try
                                
                            End If

                    End Select

                Else

                    Select Case pVal.FormUID

                        Case "MSD"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                p_iCOACount = 0
                                p_iUserCount = 0
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "btnGntFile" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                    Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim sSQL As String = String.Empty
                                    dtTable = New DataTable
                                    Dim sCheck As String = String.Empty
                                    Dim oDICompany() As SAPbobsCOM.Company = Nothing
                                    Dim sMasterDataType As String = String.Empty
                                    Dim sMasterDataCodeF As String = String.Empty
                                    Dim sMasterDataCodeT As String = String.Empty
                                    oMatrix = oForm.Items.Item("5").Specific
                                    Dim sCompanyDB As String = String.Empty

                                    Try

                                        SBO_Application.SetStatusBarMessage("Clearing the status messages...", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Clear_Matrix()", sFuncName)
                                        If Clear_Matrix(oForm, "5", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)

                                        SBO_Application.SetStatusBarMessage("Validation Process Started ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        If HeaderValidation(oForm, sErrDesc) = 0 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                        SBO_Application.SetStatusBarMessage("Validation Completed ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                        sMasterDataType = oForm.Items.Item("Item_0").Specific.selected.description.trim()
                                        sMasterDataCodeF = oForm.Items.Item("txtCode").Specific.String
                                        sMasterDataCodeT = oForm.Items.Item("Item_1").Specific.String

                                        ReDim oDICompany(oDT_Entities.Rows.Count)
                                        If oDT_Entities.Rows.Count > 0 Then

                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
                                            Call DisplayStatus(oForm, "Master data synching to selected databases.  Please wait....", sErrDesc)

                                            For imjs As Integer = 0 To oDT_Entities.Rows.Count - 1
                                                Dim irow As Integer = oDT_Entities.Rows(imjs).Item(0).ToString

                                                oDICompany(imjs) = New SAPbobsCOM.Company
                                              
                                                sCompanyDB = oDT_Entities.Rows(imjs).Item("Entity").ToString.Trim()
                                                Dim sCurrentComp As String = p_oDICompany.CompanyDB.ToString().ToUpper()
                                                If sCompanyDB.ToString().ToUpper() = sCurrentComp Then Continue For

                                                oMatrix.Columns.Item("Col_3").Cells.Item(irow).Specific.String = "Processing..."
                                                oMatrix.Columns.Item("Col_4").Cells.Item(irow).Specific.String = ""

                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                                                SBO_Application.SetStatusBarMessage("Connecting to the Target Company " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                If ConnectToTargetCompany(oDICompany(imjs), sCompanyDB, oDT_Entities.Rows(imjs).Item("UserName").ToString.Trim(), oDT_Entities.Rows(imjs).Item("Password").ToString.Trim(), sErrDesc) <> RTN_SUCCESS Then
                                                    Throw New ArgumentException(sErrDesc)
                                                End If

                                                SBO_Application.SetStatusBarMessage("Connecting to the target company Successfull " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDICompany(imjs).CompanyDB, sFuncName)

                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDICompany(imjs).CompanyDB, sFuncName)

                                                SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                If MasterDataSync(oForm, irow, p_oDICompany, oDICompany(imjs), sMasterDataType, sMasterDataCodeF, sMasterDataCodeT, sErrDesc) <> RTN_SUCCESS Then
                                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                                    oMatrix.Columns.Item("Col_3").Cells.Item(irow).Specific.String = "FAIL"
                                                    oMatrix.Columns.Item("Col_4").Cells.Item(irow).Specific.String = "Error occurred.. please refer to error log for more information."

                                                    SBO_Application.SetStatusBarMessage("Completed with ERROR " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                Else
                                                    oMatrix.Columns.Item("Col_3").Cells.Item(irow).Specific.String = "SUCCESS"
                                                    oMatrix.Columns.Item("Col_4").Cells.Item(irow).Specific.String = ""
                                                    SBO_Application.SetStatusBarMessage("Completed with SUCCESS " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If
                                            Next imjs

                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling WriteLogFiles()", sFuncName)
                                            Call WriteLogFiles(sErrDesc)

                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                                            Call EndStatus(sErrDesc)

                                        End If

                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                            If Not oDICompany(lCounter) Is Nothing Then
                                                If oDICompany(lCounter).Connected = True Then
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                    oDICompany(lCounter).Disconnect()
                                                    oDICompany(lCounter) = Nothing
                                                End If
                                            End If
                                        Next
                                        '' oMatrix.AutoResizeColumns()
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS .......", sFuncName)

                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling WriteLogFiles()", sFuncName)
                                        Call WriteLogFiles(sErrDesc)

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                                        Call EndStatus(sErrDesc)

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try

                                End If
                            End If
                    End Select
                End If

            Catch exc As Exception
                BubbleEvent = False

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling WriteLogFiles()", sFuncName)
                Call WriteLogFiles(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                Call EndStatus(sErrDesc)

                sErrDesc = exc.Message
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
                ShowErr(sErrDesc)
            End Try

        End Sub

        Sub AddMenuItems()
            Dim oMenus As SAPbouiCOM.Menus
            Dim oMenuItem As SAPbouiCOM.MenuItem
            oMenus = SBO_Application.Menus

            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = (SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))
            oMenuItem = SBO_Application.Menus.Item("43520") 'Modules

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "PWC"
            oCreationPackage.String = "Customization"
            oCreationPackage.Enabled = True
            oCreationPackage.Position = -1

            oCreationPackage.Image = System.Windows.Forms.Application.StartupPath & "\Logo.bmp"
            oMenus = oMenuItem.SubMenus

            Try
                'If the manu already exists this code will fail
                If Not p_oSBOApplication.Menus.Exists("PWC") Then
                    oMenus.AddEx(oCreationPackage)
                End If

            Catch
            End Try


            Try
                'Get the menu collection of the newly added pop-up item
                oMenuItem = SBO_Application.Menus.Item("PWC")
                oMenus = oMenuItem.SubMenus

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "MDS"
                oCreationPackage.String = "Master Data Synchronization"

                If Not p_oSBOApplication.Menus.Exists("MDS") Then
                    oMenus.AddEx(oCreationPackage)
                End If



            Catch
                'Menu already exists
                SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End Try
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent

            Select Case BusinessObjectInfo.FormTypeEx

                ''    Case "804"

                ''        If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then ' Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then


                ''            Dim ival As Integer
                ''            Dim IsError As Boolean
                ''            Dim iErr As Integer = 0
                ''            Dim sErr As String = String.Empty

                ''            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(804, p_FrmType)
                ''            Dim sAcctCode As String = String.Empty
                ''            sAcctCode = oForm.DataSources.DBDataSources.Item(0).GetValue("AcctCode", "0").ToString.Trim
                ''            Dim oChartofAccounts As SAPbobsCOM.ChartOfAccounts = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                ''            If oChartofAccounts.GetByKey(sAcctCode) Then
                ''                oChartofAccounts.FrozenFor = SAPbobsCOM.BoYesNoEnum.tYES
                ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Set InActive for Account Code " & sAcctCode, sFuncName)
                ''                ival = oChartofAccounts.Update()
                ''                If ival <> 0 Then
                ''                    IsError = True
                ''                    p_oDICompany.GetLastError(iErr, sErr)
                ''                    p_oSBOApplication.StatusBar.SetText(sErr, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                ''                    Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                ''                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                ''                    Exit Sub
                ''                End If
                ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sAcctCode, sFuncName)
                ''            End If

                ''        End If


            End Select

        End Sub

        Private Sub SBO_Application_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.LayoutKeyEvent

        End Sub
    End Class
End Namespace


