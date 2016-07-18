Module modCoCodeList

    Private oGrid As SAPbouiCOM.Grid

    Public Function LoadCoCodeListForm(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "LoadCoCodeListForm"

        Try
            LoadFromXML("Co Code List.srf", p_oSBOApplication)
            objForm = p_oSBOApplication.Forms.Item("LCOD")
            objForm.DataSources.DataTables.Add("dtCodeList")
            LoadGrid(objForm)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            LoadCoCodeListForm = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            LoadCoCodeListForm = RTN_ERROR
        End Try
    End Function

    Public Function LoadGrid(ByVal objForm As SAPbouiCOM.Form) As Long
        Dim sFuncName As String = "LoadGrid"
        Dim Sql As String

        Try
            Sql = "SELECT  ""U_GROUPCODE"" AS ""Group Code"",""U_ENTITYCODE"" AS ""Entity Code"" ,""U_ENTITYNAME"" AS ""Entity Decription"" FROM ""@AE_ENTITYLIST"""
            oGrid = objForm.Items.Item("1").Specific
            objForm.DataSources.DataTables.Item("dtCodeList").Rows.Clear()
            objForm.DataSources.DataTables.Item("dtCodeList").ExecuteQuery(Sql)
            oGrid.DataTable = objForm.DataSources.DataTables.Item("dtCodeList")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            LoadGrid = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            LoadGrid = RTN_ERROR
        End Try

    End Function

End Module
