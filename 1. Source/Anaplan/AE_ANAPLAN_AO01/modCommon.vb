Imports System.Data
Imports System.Data.Common
Imports System.Data.OleDb

Module modCommon

    Public Function ConnectDICompSSO(ByRef objCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    ConnectDICompSSO()
        '   Purpose    :    Connect To DI Company Object
        '
        '   Parameters :    ByRef objCompany As SAPbobsCOM.Company
        '                       objCompany = set the SAP Company Object
        '                   ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Sri
        '   Date       :    29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sCookie As String = String.Empty
        Dim sConnStr As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim lRetval As Long
        Dim iErrCode As Int32
        Try
            sFuncName = "ConnectDICompSSO()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            objCompany = New SAPbobsCOM.Company

            sCookie = objCompany.GetContextCookie
            sConnStr = p_oUICompany.GetConnectionContext(sCookie)
            'sConnStr = p_oSBOApplication.Company.GetConnectionContext(sCookie)
            lRetval = objCompany.SetSboLoginContext(sConnStr)

            If Not lRetval = 0 Then
                Throw New ArgumentException("SetSboLoginContext of Single SignOn Failed.")
            End If
            p_oSBOApplication.StatusBar.SetText("Please Wait While Company Connecting... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            lRetval = objCompany.Connect
            If lRetval <> 0 Then
                objCompany.GetLastError(iErrCode, sErrDesc)
                Throw New ArgumentException("Connect of Single SignOn failed : " & sErrDesc)
            Else
                p_oSBOApplication.StatusBar.SetText("Company Connection Has Established with the " & objCompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
            End If
            ConnectDICompSSO = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectDICompSSO = RTN_ERROR
        End Try
    End Function

    Public Sub ShowErr(ByVal sErrMsg As String)
        ' ***********************************************************************************
        '   Function   :    ShowErr()
        '   Purpose    :    Show Error Message
        '   Parameters :  
        '                   ByVal sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Dev
        '   Date       :    23 Jan 2007
        '   Change     :
        ' ***********************************************************************************
        Try
            If sErrMsg <> "" Then
                If Not p_oSBOApplication Is Nothing Then
                    If p_iErrDispMethod = ERR_DISPLAY_STATUS Then

                        p_oSBOApplication.SetStatusBarMessage("Error : " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short)
                    ElseIf p_iErrDispMethod = ERR_DISPLAY_DIALOGUE Then
                        p_oSBOApplication.MessageBox("Error : " & sErrMsg)
                    End If
                End If
            End If
        Catch exc As Exception
            WriteToLogFile(exc.Message, "ShowErr()")
        End Try
    End Sub

    Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
        Try
            Dim oXmlDoc As New Xml.XmlDocument
            Dim sPath As String
            ''sPath = IO.Directory.GetParent(Application.StartupPath).ToString
            sPath = System.Windows.Forms.Application.StartupPath.ToString
            'oXmlDoc.Load(sPath & "\AE_FleetMangement\" & FileName)
            oXmlDoc.Load(sPath & "\" & FileName)
            ' MsgBox(Application.StartupPath)

            Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Public Function EntityLoad(ByRef oform As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

        Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("5").Specific
        Dim sSQL As String = String.Empty
        Dim oCombo As SAPbouiCOM.ComboBox = oform.Items.Item("Item_0").Specific
        Try
            sFuncName = "EntityLoad()"

            oform.Freeze(True)

            oMatrix.Columns.Item("Col_1").Visible = False
            oMatrix.Columns.Item("Col_2").Visible = False

            oCombo.ValidValues.Add("--Select--", "0")
            oCombo.ValidValues.Add("Chart of Account", "COA")
            oCombo.ValidValues.Add("Users", "OUSR")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            '' sSQL = "SELECT T0.[U_AB_COMCODE], T0.[U_AB_COMPANYNAME], T0.[U_AB_USERCODE], T0.[U_AB_PASSWORD]  FROM [dbo].[@AB_COMPANYDATA]  T0"

            sSQL = "SELECT T0.""Name"" ""U_AB_COMCODE"", T0.""U_DBNAME"" ""U_AB_COMPANYNAME"", T0.""U_SAPUSER"" ""U_AB_USERCODE"", " & _
                " T0.""U_SAPPASSWORD"" ""U_AB_PASSWORD"" FROM  ""@AI_TB01_COMPANYDATA""  T0"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sSQL, sFuncName)
            Try
                oform.DataSources.DataTables.Add("U_AB_COMPANYNAME")
            Catch ex As Exception
            End Try
            oform.DataSources.DataTables.Item("U_AB_COMPANYNAME").ExecuteQuery(sSQL)
            oMatrix.Clear()
            'oMatrix.Columns.Item("V_1").DataBind.Bind("U_AB_COMPANYNAME", "Choose")
            oMatrix.Columns.Item("Col_0").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_COMCODE")
            oMatrix.Columns.Item("V_0").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_COMPANYNAME")
            oMatrix.Columns.Item("Col_1").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_USERCODE")
            oMatrix.Columns.Item("Col_2").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_PASSWORD")
            oMatrix.LoadFromDataSource()

            For imjs As Integer = 1 To oMatrix.RowCount
                oMatrix.Columns.Item("V_-1").Cells.Item(imjs).Specific.String = imjs
            Next imjs

            'oMatrix.AutoResizeColumns()
            oMatrix.Columns.Item("V_0").Width = 200
            oMatrix.Columns.Item("Col_0").Width = 170
            oMatrix.Columns.Item("Col_3").Width = 100
            oMatrix.Columns.Item("Col_4").Width = 400

            ''oMatrix.AutoResizeColumns()

            oform.Freeze(False)
            EntityLoad = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            oform.Freeze(False)
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            EntityLoad = RTN_ERROR
        End Try
    End Function

    Public Function ExecuteQuery(ByVal oCompany As SAPbobsCOM.Company, ByVal sSql As String, ByVal sEntity As String) As DataSet
        Dim sFuncName As String = "ExecuteQuery"
        Dim sErrDesc As String = String.Empty

        Dim cmd As New Odbc.OdbcCommand
        Dim ods As New DataSet
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()

        Try
            ''--" & oCompany.DbPassword & "
            Con.ConnectionString = "DRIVER={HDBODBC32};UID=" & oCompany.DbUserName & ";PWD=" & sHanaPassword & ";SERVERNODE=" & oCompany.Server & ";CS=" & sEntity
            Con.Open()

            cmd.CommandType = CommandType.Text
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            cmd.CommandText = sSql
            cmd.Connection = Con
            cmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(cmd)
            da.Fill(ods)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try
        Return ods
    End Function

    Public Function ExecuteNonQuery(ByVal oCompany As SAPbobsCOM.Company, ByVal sSql As String, ByVal sEntity As String) As DataSet
        Dim sFuncName As String = "ExecuteNonQuery"
        Dim sErrDesc As String = String.Empty

        Dim cmd As New Odbc.OdbcCommand
        Dim ods As New DataSet
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()

        Try
            ''--" & oCompany.DbPassword & "
            Con.ConnectionString = "DRIVER={HDBODBC32};UID=" & oCompany.DbUserName & ";PWD=" & sHanaPassword & ";SERVERNODE=" & oCompany.Server & ";CS=" & sEntity
            Con.Open()

            cmd.CommandType = CommandType.Text
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            cmd.CommandText = sSql
            cmd.Connection = Con
            cmd.CommandTimeout = 0
            cmd.ExecuteNonQuery()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try
        Return ods
    End Function

    Public Function CreateUDOTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        Dim intRetCode As Integer
        Dim objUserTableMD As SAPbobsCOM.UserTablesMD
        objUserTableMD = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        Try
            If (Not objUserTableMD.GetByKey(TableName)) Then
                objUserTableMD.TableName = TableName
                objUserTableMD.TableDescription = TableDescription
                objUserTableMD.TableType = TableType
                intRetCode = objUserTableMD.Add()
                If (intRetCode = 0) Then
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
            Throw New ArgumentException(sErrDesc)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
            GC.Collect()
        End Try
    End Function

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            objUserFieldMD = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            If (Not isColumnExist(TableName, ColumnName)) Then
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    sErrDesc = p_oDICompany.GetLastErrorCode() & ":" & p_oDICompany.GetLastErrorDescription()
                    Throw New ArgumentException(sErrDesc)
                End If
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
            Throw New ArgumentException(sErrDesc)
        Finally

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()

        End Try


    End Sub

    Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
        Dim objRecordSet As SAPbobsCOM.Recordset
        objRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            objRecordSet.DoQuery("SELECT COUNT(*) FROM ""CUFD"" WHERE ""TableID"" = '" & TableName & "' AND ""AliasID"" = '" & ColumnName & "'")
            If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
            GC.Collect()
        End Try

    End Function

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    Public Sub AddUserQuery(ByVal sFMSName As String)
        Dim sSQL As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sSQL = "SELECT ""IntrnalKey"" FROM ""OUQR"" WHERE ""QName"" = '" & sFMSName & "'"
            oRecordSet.DoQuery(sSQL)
            If oRecordSet.RecordCount > 0 Then
            Else
                Dim oUserQry As SAPbobsCOM.UserQueries
                oUserQry = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)
                oUserQry.Query = "SELECT DISTINCT ""U_GROUPCODE"" FROM ""@AE_ENTITYLIST"" "
                oUserQry.QueryCategory = -1
                oUserQry.QueryDescription = sFMSName

                If oUserQry.Add() <> 0 Then
                    sErrDesc = p_oDICompany.GetLastErrorDescription
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while creating User Query", sFuncName)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)

        End Try
    End Sub

    Public Sub AssignFMS(ByVal sFMSName As String, ByVal sItemID As String)
        Dim sInternalKey As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim iQryId As Integer

        Try
            sFuncName = "AssignFMS"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT ""IntrnalKey"" FROM ""OUQR"" WHERE ""QName"" = '" & sFMSName & "'")
            If oRecordSet.RecordCount > 0 Then
                iQryId = oRecordSet.Fields.Item("IntrnalKey").Value

                Dim oFMS As SAPbobsCOM.FormattedSearches
                oFMS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                oFMS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                Dim oFMSKey As Integer
                Dim bUpdate As Boolean = False
                sSQL = "select ""FormID"",""ItemID"",""IndexID"" from ""CSHS"" WHERE ""FormID"" = 'GTB' AND ""ItemID"" = '" & sItemID & "'"
                oRecordSet.DoQuery(sSQL)
                If oRecordSet.RecordCount > 0 Then
                    oFMSKey = oRecordSet.Fields.Item("IndexID").Value
                    bUpdate = True
                End If
                oFMS.QueryID = iQryId
                oFMS.FormID = "GTB"
                oFMS.ItemID = sItemID
                oFMS.ColumnID = "-1"
                oFMS.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tNO
                oFMS.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                oFMS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                'oFMS.FieldID = ""
                If bUpdate = True Then
                    If oFMS.Update() <> 0 Then
                        sErrDesc = p_oDICompany.GetLastErrorDescription
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while creating FMS", sFuncName)
                    End If
                Else
                    If oFMS.Add() <> 0 Then
                        sErrDesc = p_oDICompany.GetLastErrorDescription
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while assigning FMS", sFuncName)
                    End If
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub

    Public Sub CreateProcedure(ByVal sProcName As String)
        Dim sFile As String = String.Empty
        Dim sTextLine As String = String.Empty
        Dim oRset As SAPbobsCOM.Recordset

        Try
            sFuncName = "CreateProcedure"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sFile = System.Windows.Forms.Application.StartupPath & "\" & sProcName & ".txt"

            If IO.File.Exists(sFile) Then
                Using objReader As New System.IO.StreamReader(sFile)
                    sTextLine = objReader.ReadToEnd()
                End Using

                If sTextLine <> String.Empty Then
                    Dim sQuery As String = String.Empty
                    Dim iCount As Integer = 0
                    sQuery = "SELECT ""PROCEDURE_NAME"" FROM ""PROCEDURES"" WHERE ""SCHEMA_NAME""  = '" & p_oDICompany.CompanyDB & "' AND UPPER(""PROCEDURE_NAME"") = '" + sProcName.ToUpper() + "'"
                    oRset.DoQuery(sQuery)
                    If oRset.RecordCount > 0 Then
                        sQuery = "DROP PROCEDURE " & sProcName & ""
                        oRset.DoQuery(sQuery)
                    End If

                    oRset.DoQuery(sTextLine)

                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRset)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

End Module
