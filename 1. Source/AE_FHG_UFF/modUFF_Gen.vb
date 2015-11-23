Imports System.IO
Imports System.Text



Module modUFF_Gen


    Public Function UFF_Generation(ByVal sDate As Date, ByVal spath As String, ByVal sFilename As String, ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "UFF_Generation"
        Dim sSQL As String = String.Empty
        Dim oDT_UFFData As New DataTable
        Dim sBatchNo As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            For imjs As Integer = 0 To frmUFFGen.lstBatch.Items.Count - 1
                sBatchNo += frmUFFGen.lstBatch.Items(imjs) & ","
            Next

            sBatchNo = Left(sBatchNo, Len(sBatchNo) - 1)

            sSQL = "call ""AE_SP004_UFF_Generation"" ('" & sBatchNo & "', '" & sTemplateType & "')"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
            oDT_UFFData = ExecuteSQLQuery(sSQL)

            If oDT_UFFData.Rows.Count = 0 Then
                WriteToStatusScreen(False, "No Matches Records Found")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Matches Records Found ", sFuncName)
                WriteToStatusScreen(False, "------------------------- Process Completed ------------------------")
                MsgBox("Validation Error: No Matches Records ", MsgBoxStyle.OkOnly, "DBD UFF File Generation")
                UFF_Generation = RTN_SUCCESS
                Exit Try
            End If

            WriteToStatusScreen(False, "Calling Function FormatUFF()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function FormatUFF()", sFuncName)
            If FormatUFF(oDT_UFFData, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)


            WriteToStatusScreen(False, "Calling Function GenerateUFF()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function GenerateUFF()", sFuncName)
            If GenerateUFF(oDT_UFFData, spath, sFilename, sDate.ToString("ddMMyyyy"), sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

            '' WriteToStatusScreen(False, "------------------------- Process Completed ------------------------")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed successfully.", sFuncName)
            UFF_Generation = RTN_SUCCESS
        Catch ex As Exception
            sErrdesc = ex.Message
            UFF_Generation = RTN_ERROR
            Call WriteToLogFile(sErrdesc, sFuncName)
            ''WriteToStatusScreen(False, "ERROR::" & sErrdesc)
            MsgBox(sErrdesc, MsgBoxStyle.OkOnly, "DBS UFF File Generation")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        End Try
    End Function

    Public Sub WriteToStatusScreen(ByVal Clear As Boolean, ByVal msg As String)
        If Clear Then
            frmUFFGen.txtStatusMsg.Text = ""
        End If
        frmUFFGen.txtStatusMsg.HideSelection = True
        frmUFFGen.txtStatusMsg.Text &= msg & vbCrLf
        frmUFFGen.txtStatusMsg.SelectAll()
        frmUFFGen.txtStatusMsg.ScrollToCaret()
        frmUFFGen.txtStatusMsg.Refresh()
    End Sub

    Public Function GenerateUFF(ByVal oDT_FinalResult As DataTable, ByVal sPAth As String, ByVal sFileName As String, ByVal sDate As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty


        Dim sw As StreamWriter = Nothing
        Dim irow As Integer
        Dim dCount As Integer = 0
        Dim dAmount As Double = 0.0

        Try
            sFuncName = "GenerateUFF()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            WriteToStatusScreen(False, "Starting Function " & sFuncName)

            frmUFFGen.txtFileName.Text = Now.Date.ToString("yyyyMMdd") & "_" & DateTime.Now.ToString("HH.mm.ss") & p_oCompDef.sOrgID
            sFileName = Now.Date.ToString("yyyyMMdd") & "_" & DateTime.Now.ToString("HH.mm.ss") & p_oCompDef.sOrgID


            If File.Exists(sPAth & sFileName & ".txt") Then
                File.Delete(sPAth & sFileName & ".txt")
            End If

            WriteToStatusScreen(False, "Attempting to Generate the UFF file format")

            sw = New StreamWriter(sPAth & sFileName & ".txt", False, Encoding.UTF8)
            ' Add some text to the file.
            '----------------------- Header Portion -----------------------------------------------
            sw.WriteLine("HEADER," & sDate & "," & p_oCompDef.sOrgID & "," & p_oCompDef.sSenderName)
            '----------------------- Detail Portion -----------------------------------------------
            For Each dr As DataRow In oDT_FinalResult.Rows
                dCount += 1
                dAmount += dr("DocTotal").ToString.Trim
                sw.WriteLine(dr("Header").ToString.Trim & dr("D64").ToString & dr("Col1").ToString.Trim)
            Next
            '----------------------- Detail Portion -----------------------------------------------
            sw.WriteLine("TRAILER," & dCount & "," & dAmount)
            sw.Close()
            GenerateUFF = RTN_SUCCESS
            WriteToStatusScreen(False, "Successfully generated the UFF file ")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)
            WriteToStatusScreen(False, "Completed With SUCCESS " & sFuncName)

        Catch ex As Exception
            GenerateUFF = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToStatusScreen(False, "Completed with ERROR " & sErrDesc)
        Finally
            sw.Dispose()
            sw.Close()
        End Try

    End Function

    Public Function FormatUFF(ByRef oDT_Result As DataTable, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty


        Dim sw As StreamWriter = Nothing
        Dim irow As Integer
        Dim dCount As Decimal = 0
        Dim dAmount As Decimal = 0
        Dim oDV_UFF As DataView = oDT_Result.DefaultView
        Dim oDT_Distinct As New DataTable
        Dim oDT_Formatedresult As New DataTable
        Dim sHeader As String = String.Empty
        Dim sD64 As String = String.Empty
        Dim sCol1 As String = String.Empty
        Dim dDoctotal As String = String.Empty


        oDT_Formatedresult.Columns.Add("Header", GetType(String))
        oDT_Formatedresult.Columns.Add("D64", GetType(String))
        oDT_Formatedresult.Columns.Add("Col1", GetType(String))
        oDT_Formatedresult.Columns.Add("DocTotal", GetType(Decimal))
        ''oDT_Formatedresult.Columns.Add("DocNum", GetType(String))

        Try
            sFuncName = "Formatting The File()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            WriteToStatusScreen(False, "Starting Function " & sFuncName)
            oDT_Distinct = oDV_UFF.Table.DefaultView.ToTable(True, "DocNum")

            For Each dr As DataRow In oDT_Distinct.Rows
                oDV_UFF.RowFilter = "DocNum='" & dr("DocNum").ToString.Trim & "'"
                sHeader = oDV_UFF.Item(0)("Header")
                sCol1 = oDV_UFF.Item(0)("Col1")
                dDoctotal = oDV_UFF.Item(0)("Doctotal")
                For Each drv As DataRowView In oDV_UFF
                    sD64 += drv("D64").ToString & Space(70)
                Next
                oDT_Formatedresult.Rows.Add(sHeader, Left(sD64, sD64.ToString.Length - 70), sCol1, dDoctotal)
                sD64 = String.Empty
            Next
            oDT_Result = oDT_Formatedresult
            FormatUFF = RTN_SUCCESS
           
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)
            WriteToStatusScreen(False, "Completed With SUCCESS " & sFuncName)

        Catch ex As Exception
            FormatUFF = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToStatusScreen(False, "Completed with ERROR " & sErrDesc)
        Finally
           
        End Try

    End Function

End Module
