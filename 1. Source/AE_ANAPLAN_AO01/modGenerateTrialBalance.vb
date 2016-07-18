Imports System.Data
Imports System.Threading
Imports System.Windows.Forms
Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Module modGenerateTrialBalance

    Private oEdit As SAPbouiCOM.EditText
    Private oGrid As SAPbouiCOM.Grid
    Private oCheck As SAPbouiCOM.CheckBox
    Private sFolderPath As String

    Private Sub FolderBrowserDialog(ByRef objForm As SAPbouiCOM.Form)
        Dim myThread As New System.Threading.Thread(AddressOf OpenFolderBrowserDialog)
        myThread.SetApartmentState(Threading.ApartmentState.STA)
        myThread.Start()
        myThread.Join()
    End Sub

    Private Sub OpenFolderBrowserDialog(ByVal objForm As SAPbouiCOM.Form)
        Dim DummyForm As New frmBrowserDialog
        sFolderPath = String.Empty
        DummyForm.Show()
        DummyForm.Visible = False
        DummyForm.TopMost = True
        Dim result As DialogResult = DummyForm.FolderBrowserDialog1.ShowDialog()
        If result = DialogResult.OK Then
            sFolderPath = DummyForm.FolderBrowserDialog1.SelectedPath
        End If
        System.Threading.Thread.CurrentThread.Abort()

    End Sub

    Private Sub AddUserDatasources(ByVal objForm As SAPbouiCOM.Form)
        oEdit = objForm.Items.Item("4").Specific
        objForm.DataSources.UserDataSources.Add("uPath", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oEdit.DataBind.SetBound(True, "", "uPath")

        oEdit = objForm.Items.Item("6").Specific
        objForm.DataSources.UserDataSources.Add("uDateFrm", SAPbouiCOM.BoDataType.dt_DATE, 50)
        oEdit.DataBind.SetBound(True, "", "uDateFrm")

        oEdit = objForm.Items.Item("8").Specific
        objForm.DataSources.UserDataSources.Add("uDateTo", SAPbouiCOM.BoDataType.dt_DATE, 50)
        oEdit.DataBind.SetBound(True, "", "uDateTo")

        oEdit = objForm.Items.Item("10").Specific
        objForm.DataSources.UserDataSources.Add("CoCodeFrm", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oEdit.DataBind.SetBound(True, "", "CoCodeFrm")

        oEdit = objForm.Items.Item("12").Specific
        objForm.DataSources.UserDataSources.Add("CoCodeTo", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oEdit.DataBind.SetBound(True, "", "CoCodeTo")

        oCheck = objForm.Items.Item("13").Specific
        objForm.DataSources.UserDataSources.Add("uOpen", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
        oCheck.DataBind.SetBound(True, "", "uOpen")
    End Sub

    Private Sub LoadGrid(ByVal objForm As SAPbouiCOM.Form)
        Dim sCoCodeFrom, sCoCodeTo As String
        Dim Sql As String

        oEdit = objForm.Items.Item("10").Specific
        sCoCodeFrom = oEdit.Value
        oEdit = objForm.Items.Item("12").Specific
        sCoCodeTo = oEdit.Value

        p_oSBOApplication.StatusBar.SetText("Processing... Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        If sCoCodeFrom = "" Then
            p_oSBOApplication.StatusBar.SetText("Select Co Code From", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        ElseIf sCoCodeTo = "" Then
            p_oSBOApplication.StatusBar.SetText("Select Co To From", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        If sCoCodeFrom <> "" And sCoCodeTo <> "" Then
            Sql = "SELECT NULL ""Select"",""U_ENTITYCODE"" AS ""Entity Code"" ,""U_ENTITYNAME"" AS ""Entity Decription"",""U_ANAPLANCODE"" AS ""AnaPlanCode"" FROM ""@AE_ENTITYLIST"" " & _
                  " WHERE U_GROUPCODE >= '" & sCoCodeFrom & "' AND U_GROUPCODE <= '" & sCoCodeTo & "' "
            oGrid = objForm.Items.Item("14").Specific
            objForm.DataSources.DataTables.Item("dtEntityList").Rows.Clear()
            objForm.DataSources.DataTables.Item("dtEntityList").ExecuteQuery(Sql)
            oGrid.DataTable = objForm.DataSources.DataTables.Item("dtEntityList")
            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = False
        End If

        p_oSBOApplication.StatusBar.SetText("Operation Completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    End Sub

    Private Function CheckGridDatas(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String)
        Dim bCheck As Boolean
        bCheck = True

        oEdit = objForm.Items.Item("4").Specific
        If oEdit.Value = "" Then
            bCheck = False
            sErrDesc = "Select the folder path"
            Return bCheck
            Exit Function
        End If

        oEdit = objForm.Items.Item("6").Specific
        If oEdit.Value = "" Then
            bCheck = False
            sErrDesc = "Enter Date from"
            Return bCheck
            Exit Function
        End If

        oEdit = objForm.Items.Item("8").Specific
        If oEdit.Value = "" Then
            bCheck = False
            sErrDesc = "Enter Date To"
            Return bCheck
            Exit Function
        End If

        'oEdit = objForm.Items.Item("4").Specific
        'If oEdit.Value = "" Then
        '    bCheck = False
        '    sErrDesc = "Select the folder path"
        '    Return bCheck
        '    Exit Function
        'End If

        'oEdit = objForm.Items.Item("4").Specific
        'If oEdit.Value = "" Then
        '    bCheck = False
        '    sErrDesc = "Select the folder path"
        '    Return bCheck
        '    Exit Function
        'End If

        oGrid = objForm.Items.Item("14").Specific
        For i As Integer = 0 To oGrid.Rows.Count - 1
            If oGrid.DataTable.GetValue("Select", i) = "Y" Then
                bCheck = True
                Exit For
            Else
                bCheck = False
            End If
        Next

        If bCheck = False Then
            sErrDesc = "Select any one entity in Grid"
            Return bCheck
            Exit Function
        End If

        Return bCheck
    End Function

    Private Function GenerateCSVFile(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Boolean
        Dim sFuncName As String = "GenerateCSVFile"
        Dim oDs As New DataSet
        Dim oFinalDs As New DataSet
        Dim sSql As String = String.Empty
        Dim sFolderPath As String = String.Empty
        Dim sEntityCode As String = String.Empty
        Dim sAnaPlanCode As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim dtFromDate As Date
        Dim dtToDate As Date
        Dim k As Integer = 0

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oEdit = objForm.Items.Item("4").Specific
            sFolderPath = oEdit.Value

            oEdit = objForm.Items.Item("6").Specific
            dtFromDate = GetDateTimeValue(oEdit.String)

            oEdit = objForm.Items.Item("8").Specific
            dtToDate = GetDateTimeValue(oEdit.Value)

            oGrid = objForm.Items.Item("14").Specific
            For i As Integer = 0 To oGrid.Rows.Count - 1
                If oGrid.DataTable.GetValue("Select", i) = "Y" Then
                    k = 1
                    sEntityCode = oGrid.DataTable.GetValue("Entity Code", i)
                    sAnaPlanCode = oGrid.DataTable.GetValue("AnaPlanCode", i)

                    '*******CREATE PROCEDURE IN THE SELECTED ENTITY
                    'CreateProcedure("AE_SP001_AnaplanTrailBalExtraction_Addon", sEntityCode)
                    CreateProcedure_Entity(p_oDICompany, "AE_SP001_AnaplanTrailBalExtraction_Addon", sEntityCode)

                    sSql = "CALL " & sEntityCode & ".""AE_SP001_AnaplanTrailBalExtraction_Addon""('" & dtFromDate.ToString("yyyy-MM-dd") & "','" & dtToDate.ToString("yyyy-MM-dd") & "','" & sAnaPlanCode & "')"
                    oDs = ExecuteQuery(p_oDICompany, sSql, sEntityCode)
                    If oDs.Tables(0).Rows.Count > 0 Then
                        Dim sFilePath As String
                        'sFilePath = sFolderPath & "\" & sEntityCode & " TRIAL BALANCE " & DateTime.Today.ToString("MMMM").ToUpper() & "-" & Date.Now.Year & ".csv"
                        sFilePath = sFolderPath & "\" & sEntityCode & " TRIAL BALANCE " & dtToDate.ToString("MMMM").ToUpper() & "-" & dtToDate.Year & ".csv"

                        '************************GENERATION OF CSV FILE WORKING CODE STARTS
                        Dim bFirstRecord As Boolean = True
                        Dim myString As String = String.Empty

                        Dim myWriter As New System.IO.StreamWriter(sFilePath)
                        For Each dt In oDs.Tables
                            bFirstRecord = True
                            For Each column As DataColumn In dt.Columns
                                If Not bFirstRecord Then
                                    myString = myString & ","
                                End If
                                myString = myString & column.ColumnName
                                bFirstRecord = False
                            Next
                            myString = myString & Environment.NewLine
                            For Each dr As DataRow In dt.rows
                                bFirstRecord = True

                                ''myWriter.WriteLine(dr(0))

                                For Each field As Object In dr.ItemArray
                                    If Not bFirstRecord Then
                                        myString = myString & ","
                                    End If
                                    myString = myString & field
                                    bFirstRecord = False
                                Next
                                myString = myString & Environment.NewLine
                            Next
                        Next

                        myWriter.WriteLine(myString)
                        myWriter.Close()

                        oCheck = objForm.Items.Item("13").Specific
                        If oCheck.Checked = True Then
                            If OpenExcelDemo(sFilePath, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                        '************************GENERATION OF CSV FILE WORKING CODE ENDS
                    End If

                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GenerateCSVFile = True
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GenerateCSVFile = False
        End Try
    End Function

    Public Sub CreateProcedure_Entity(ByVal oCompany As SAPbobsCOM.Company, ByVal sProcedure As String, ByVal sEntity As String)
        Dim sQuery As String = String.Empty
        Dim iCount As Integer = 0
        Dim oDs As New DataSet
        Dim sFile As String = String.Empty
        Dim sTextLine As String = String.Empty

        Try
            sFuncName = "CreateProcedure_Entity"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            sFile = Application.StartupPath & "\" & sProcedure & ".txt"

            If IO.File.Exists(sFile) Then
                Using objReader As New System.IO.StreamReader(sFile)
                    sTextLine = objReader.ReadToEnd()
                End Using

                sQuery = "SELECT COUNT(""PROCEDURE_NAME"") AS ""MNO"" FROM ""PROCEDURES"" WHERE ""SCHEMA_NAME""  = '" & sEntity & "' AND ""PROCEDURE_NAME"" = '" + sProcedure + "'"
                oDs = ExecuteQuery(oCompany, sQuery, sEntity)
                If oDs.Tables(0).Rows.Count > 0 Then
                    iCount = oDs.Tables(0).Rows(0).Item(0).ToString()
                End If
                If iCount > 0 Then
                    sQuery = "DROP PROCEDURE " & sEntity & ".""" & sProcedure & """ "
                    ExecuteNonQuery(oCompany, sQuery, sEntity)
                End If
                ExecuteNonQuery(oCompany, sTextLine, sEntity)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

    Public Function OpenExcelDemo(ByVal sFileName As String, ByRef sErrDesc As String) As Long  ', ByVal SheetName As String
        Dim sFuncName As String = "OpenExcelDemo"
        Dim xlsApp As Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsApp = New Microsoft.Office.Interop.Excel.Application

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If IO.File.Exists(sFileName) Then
                xlsApp.Visible = True
                xlsWB = xlsApp.Workbooks.Open(sFileName)
            Else
                sErrDesc = "File Not found"
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            OpenExcelDemo = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            OpenExcelDemo = RTN_ERROR
        End Try

    End Function

    Public Sub GTB_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "GTB_SBO_ItemEvent"
        Dim sErrDesc As String = String.Empty

        Try
            If pval.Before_Action = True Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "3" Then
                            If CheckGridDatas(objForm, sErrDesc) = False Then
                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        ElseIf pval.ItemUID = "15" Then
                            FolderBrowserDialog(objForm)
                            oEdit = objForm.Items.Item("4").Specific
                            oEdit.Value = sFolderPath
                        End If
                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "3" Then
                            p_oSBOApplication.StatusBar.SetText("Processing... Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            If GenerateCSVFile(objForm, sErrDesc) = False Then
                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            p_oSBOApplication.StatusBar.SetText("Operation Completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        ElseIf pval.ItemUID = "16" Then
                            objForm.Freeze(True)
                            LoadGrid(objForm)
                            objForm.Freeze(False)
                        End If

                End Select
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub

    Public Sub GTB_SBO_MenuEvent(ByVal pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form
                If pVal.MenuUID = "GTB" Then
                    LoadFromXML("Generate Text File.srf", p_oSBOApplication)
                    objForm = p_oSBOApplication.Forms.Item("GTB")
                    AddUserDatasources(objForm)
                    objForm.DataSources.DataTables.Add("dtEntityList")
                End If
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub

End Module
