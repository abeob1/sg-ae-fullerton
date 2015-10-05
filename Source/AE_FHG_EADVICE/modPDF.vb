Module modPDF

    ''Public Function ExportReports(ByVal sErrdesc As String) As Long
    ''    Dim sFuncName As String = "ExportReports"
    ''    Dim sSQL As String = String.Empty
    ''    Dim oDs As New DataSet
    ''    Dim iDocNum As Integer
    ''    Try
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

    ''        If p_oDtReport.Rows.Count > 0 Then

    ''            For Each row As DataRow In p_oDtReport.Rows
    ''                If row.Item("DocEntry") > 0 Then
    ''                    If row.Item("Type") = "AR" Then
    ''                        sSQL = "SELECT ""DocNum"" from OINV WHERE ""DocEntry""=" & row.Item("DocEntry")
    ''                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
    ''                        oDs = ExecuteSQLQuery(sSQL)
    ''                        If oDs.Tables(0).Rows.Count > 0 Then
    ''                            iDocNum = oDs.Tables(0).Rows(0).Item(0)
    ''                        End If
    ''                        If row.Item("Owner") = "AIA" Then
    ''                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
    ''                            If ExportToPDF(row.Item("DocEntry"), iDocNum, row.Item("BPCode").ToString, "AR", "AI_RPT_TaxInvoice_AIA.rpt", sErrdesc) <> RTN_SUCCESS Then
    ''                                Throw New ArgumentException(sErrdesc)
    ''                            End If
    ''                        Else
    ''                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
    ''                            If ExportToPDF(row.Item("DocEntry"), iDocNum, row.Item("BPCode").ToString, "AR", "AI_RPT_TaxInvoice.rpt", sErrdesc) <> RTN_SUCCESS Then
    ''                                Throw New ArgumentException(sErrdesc)
    ''                            End If
    ''                        End If
    ''                    End If

    ''                    If row.Item("Type") = "AP" Then
    ''                        sSQL = "SELECT ""DocNum"" from OPCH WHERE ""DocEntry""=" & row.Item("DocEntry")
    ''                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
    ''                        oDs = ExecuteSQLQuery(sSQL)
    ''                        If oDs.Tables(0).Rows.Count > 0 Then
    ''                            iDocNum = oDs.Tables(0).Rows(0).Item(0)
    ''                        End If
    ''                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
    ''                        If ExportToPDF(row.Item("DocEntry"), iDocNum, row.Item("BPCode").ToString, "AP", "AI_RPT_AP_Invoice.rpt", sErrdesc) <> RTN_SUCCESS Then
    ''                            Throw New ArgumentException(sErrdesc)
    ''                        End If
    ''                    End If

    ''                    If row.Item("Type") = "PY" Then
    ''                        sSQL = "SELECT ""DocNum"" from OVPM WHERE ""DocEntry""=" & row.Item("DocEntry")
    ''                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
    ''                        oDs = ExecuteSQLQuery(sSQL)
    ''                        If oDs.Tables(0).Rows.Count > 0 Then
    ''                            iDocNum = oDs.Tables(0).Rows(0).Item(0)
    ''                        End If
    ''                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
    ''                        If ExportToPDF(row.Item("DocEntry"), iDocNum, row.Item("BPCode").ToString, "PY", "AI_RPT_PaymentAdvice.rpt", sErrdesc) <> RTN_SUCCESS Then
    ''                            Throw New ArgumentException(sErrdesc)
    ''                        End If
    ''                    End If

    ''                    If row.Item("Type") = "TPA" Then
    ''                        sSQL = "SELECT ""DocNum"" from ORPC WHERE ""DocEntry""=" & row.Item("DocEntry")
    ''                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
    ''                        oDs = ExecuteSQLQuery(sSQL)
    ''                        If oDs.Tables(0).Rows.Count > 0 Then
    ''                            iDocNum = oDs.Tables(0).Rows(0).Item(0)
    ''                        End If
    ''                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
    ''                        If ExportToPDF(row.Item("DocEntry"), iDocNum, row.Item("BPCode").ToString, "TPA", "AI_RPT_TaxInvoice_TPAFee.rpt", sErrdesc) <> RTN_SUCCESS Then
    ''                            Throw New ArgumentException(sErrdesc)
    ''                        End If
    ''                    End If

    ''                End If
    ''            Next
    ''        End If
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed successfully.", sFuncName)
    ''        ExportReports = RTN_SUCCESS
    ''    Catch ex As Exception
    ''        sErrdesc = ex.Message
    ''        ExportReports = RTN_ERROR
    ''        Call WriteToLogFile(sErrdesc, sFuncName)
    ''        Throw New ArgumentException(sErrdesc)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
    ''    End Try
    ''End Function

    'Public Function ExportToPDF(ByVal iDocNum As Integer, _
    '                            ByVal sTargetFileName As String, _
    '                            ByVal sReportFileName As String, _
    '                            ByRef sErrDesc As String) As Long

    '    ' *********************************************************************************************
    '    '   Function   :   ExportToPDF()
    '    '   Purpose    :   ExportToPDF
    '    '   Parameters :   ByVal sPath As Integer
    '    '                  sPath=Report Path
    '    '                  ByRef sErrDesc As String
    '    '                   sErrDesc=Error Description to be returned to calling function
    '    '   Return     :   0 - FAILURE
    '    '                  1 - SUCCESS
    '    '   Date       :   29/11/2013
    '    '   Change     :
    '    ' *********************************************************************************************

    '    Dim sFuncName As String = String.Empty
    '    Dim intCounter As Integer
    '    Dim intCounter1 As Integer
    '    Dim iCount As Integer
    '    Dim iSubRParaCount As Integer
    '    'Crystal Report's report document object

    '    Dim objReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    '    'object of table Log on info of Crystal report
    '    Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
    '    Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
    '    Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
    '    Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table = Nothing
    '    'Sub report object of crystal report.
    '    Dim mySubReportObject As CrystalDecisions.CrystalReports.Engine.SubreportObject
    '    'Sub report document of crystal report.
    '    Dim mySubRepDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    '    Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
    '    Dim CrDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions()
    '    Dim CrFormatTypeOptions As New CrystalDecisions.Shared.PdfRtfWordFormatOptions()
    '    Dim sTargetFile As String = sTargetFileName
    '    Dim sRptFileName As String = sReportFileName
    '    Dim sSQL As String = String.Empty
    '    Dim sCompanyName As String = String.Empty
    '    Dim sSVCID As String = String.Empty


    '    Try
    '        sFuncName = "ExportToPDF()"
    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Report File Name:" & sRptFileName, sFuncName)
    '        'Load the report
    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Loading Report", sFuncName)

    '        objReport.Load(sRptFileName, CrystalDecisions.[Shared].OpenReportMethod.OpenReportByTempCopy)
    '        'oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        'Set the connection information to crConnectionInfo object so that we can apply the 
    '        ' connection information on each table in the reporteport

    '        With crConnectionInfo
    '            .DatabaseName = p_oCompDef.sSAPDBName
    '            .Password = p_oCompDef.sDBPwd
    '            .ServerName = p_oCompDef.sReportDSN
    '            .UserID = p_oCompDef.sDBUser

    '        End With

    '        CrTables = objReport.Database.Tables
    '        For Each CrTable In CrTables
    '            crtableLogoninfo = CrTable.LogOnInfo
    '            crtableLogoninfo.ConnectionInfo = crConnectionInfo
    '            CrTable.ApplyLogOnInfo(crtableLogoninfo)
    '            CrTable.Location = p_oCompDef.sSAPDBName & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
    '        Next
    '        ' Loop through each section on the report then look  through each object in the section
    '        ' if the object is a subreport, then apply logon info on each table of that sub report
    '        For iCount = 0 To objReport.ReportDefinition.Sections.Count - 1
    '            For intCounter = 0 To objReport.ReportDefinition.Sections(iCount).ReportObjects.Count - 1
    '                With objReport.ReportDefinition.Sections(iCount)
    '                    If .ReportObjects(intCounter).Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then
    '                        mySubReportObject = CType(.ReportObjects(intCounter), CrystalDecisions.CrystalReports.Engine.SubreportObject)
    '                        mySubRepDoc = mySubReportObject.OpenSubreport(mySubReportObject.SubreportName)
    '                        'get the subreport parameter count to exclude passing data
    '                        iSubRParaCount += mySubRepDoc.DataDefinition.ParameterFields.Count
    '                        For intCounter1 = 0 To mySubRepDoc.Database.Tables.Count - 1
    '                            CrTable.ApplyLogOnInfo(crtableLogoninfo)
    '                            CrTable.Location = p_oCompDef.sSAPDBName & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
    '                        Next
    '                    End If
    '                End With
    '            Next
    '        Next
    '        'Check if there are parameters or not in report and exclude the subreport parameters.
    '        intCounter = objReport.DataDefinition.ParameterFields.Count - iSubRParaCount
    '        'As parameter fields collection also picks the selection formula which is not the parameter
    '        ' so if total parameter count is 1 then we check whether its a parameter or selection formula.
    '        If intCounter = 1 Then
    '            If InStr(objReport.DataDefinition.ParameterFields(0).ParameterFieldName, ".", CompareMethod.Text) > 0 Then
    '                intCounter = 0
    '            End If
    '        End If

    '        ' set the parameter to the report
    '        objReport.SetParameterValue(0, iDocNum)

    '        'Export to PDF

    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target File Name:" & sTargetFile, sFuncName)

    '        CrDiskFileDestinationOptions.DiskFileName = sTargetFile
    '        CrExportOptions = objReport.ExportOptions
    '        With CrExportOptions
    '            'Set the destination to a disk file 
    '            .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
    '            'Set the format to PDF 
    '            .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
    '            'Set the destination options to DiskFileDestinationOptions object 
    '            .DestinationOptions = CrDiskFileDestinationOptions
    '            .FormatOptions = CrFormatTypeOptions
    '        End With
    '        'Export the report 
    '        objReport.Export()

    '        ExportToPDF = RTN_SUCCESS
    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
    '    Catch ex As Exception
    '        sErrDesc = ex.Message
    '        ExportToPDF = RTN_ERROR
    '        Call WriteToLogFile(sErrDesc, sFuncName)
    '        Throw New ArgumentException(sErrDesc)
    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
    '    Finally
    '        objReport.Dispose()
    '        crConnectionInfo = Nothing
    '        mySubRepDoc = Nothing
    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
    '    End Try

    'End Function

    Public Function ExportToPDF(ByVal sCardCodeFrom As String, _
                                ByVal sCardCodeTo As String, _
                                ByVal _Date As Date, _
                                ByVal sTargetFileName As String, _
                                ByVal sReportFileName As String, _
                                ByRef sErrDesc As String) As Long

        ' *********************************************************************************************
        '   Function   :   ExportToPDF()
        '   Purpose    :   ExportToPDF
        '   Parameters :   ByVal sPath As Integer
        '                  sPath=Report Path
        '                  ByRef sErrDesc As String
        '                   sErrDesc=Error Description to be returned to calling function
        '   Return     :   0 - FAILURE
        '                  1 - SUCCESS
        '   Date       :   29/11/2013
        '   Change     :
        ' *********************************************************************************************

        Dim sFuncName As String = String.Empty
        Dim intCounter As Integer
        Dim intCounter1 As Integer
        Dim iCount As Integer
        Dim iSubRParaCount As Integer
        'Crystal Report's report document object

        Dim objReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        'object of table Log on info of Crystal report
        Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
        Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
        Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
        Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table = Nothing

        Dim CrTables1 As CrystalDecisions.CrystalReports.Engine.Tables
        Dim CrTable1 As CrystalDecisions.CrystalReports.Engine.Table = Nothing

        'Sub report object of crystal report.
        Dim mySubReportObject As CrystalDecisions.CrystalReports.Engine.SubreportObject
        'Sub report document of crystal report.
        Dim mySubRepDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
        Dim CrDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions()
        Dim CrFormatTypeOptions As New CrystalDecisions.Shared.PdfRtfWordFormatOptions()
        Dim sTargetFile As String = sTargetFileName
        Dim sRptFileName As String = sReportFileName
        Dim sSQL As String = String.Empty
        Dim sCompanyName As String = String.Empty
        Dim sSVCID As String = String.Empty


        Try
            sFuncName = "ExportToPDF()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Report File Name:" & sRptFileName, sFuncName)
            'Load the report
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Loading Report", sFuncName)

            objReport.Load(sRptFileName, CrystalDecisions.[Shared].OpenReportMethod.OpenReportByTempCopy)
            'oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Set the connection information to crConnectionInfo object so that we can apply the 
            ' connection information on each table in the reporteport

            '' p_oCompDef.sSAPDBName = frmSOA.SCompany.Text

            With crConnectionInfo
                .DatabaseName = p_oCompDef.sSAPDBName
                .Password = p_oCompDef.sDBPwd
                .ServerName = p_oCompDef.sReportDSN
                .UserID = p_oCompDef.sDBUser

            End With

            CrTables = objReport.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = p_oCompDef.sSAPDBName & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next

            ' Loop through each section on the report then look  through each object in the section
            ' if the object is a subreport, then apply logon info on each table of that sub report
            For iCount = 0 To objReport.ReportDefinition.Sections.Count - 1
                For intCounter = 0 To objReport.ReportDefinition.Sections(iCount).ReportObjects.Count - 1
                    With objReport.ReportDefinition.Sections(iCount)
                        If .ReportObjects(intCounter).Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then
                            mySubReportObject = CType(.ReportObjects(intCounter), CrystalDecisions.CrystalReports.Engine.SubreportObject)
                            mySubRepDoc = mySubReportObject.OpenSubreport(mySubReportObject.SubreportName)

                            'get the subreport parameter count to exclude passing data

                            iSubRParaCount += mySubRepDoc.DataDefinition.ParameterFields.Count
                            For intCounter1 = 0 To mySubRepDoc.Database.Tables.Count - 1
                                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                                CrTable.Location = p_oCompDef.sSAPDBName & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                            Next
                            iSubRParaCount += mySubRepDoc.DataDefinition.ParameterFields.Count
                        End If
                    End With
                Next
            Next

            'Check if there are parameters or not in report and exclude the subreport parameters.
            intCounter = objReport.DataDefinition.ParameterFields.Count - iSubRParaCount
            'As parameter fields collection also picks the selection formula which is not the parameter
            ' so if total parameter count is 1 then we check whether its a parameter or selection formula.
            If intCounter = 1 Then
                If InStr(objReport.DataDefinition.ParameterFields(0).ParameterFieldName, ".", CompareMethod.Text) > 0 Then
                    intCounter = 0
                End If

            End If

            ' set the parameter to the report
            objReport.SetParameterValue(0, sCardCodeFrom)
            objReport.SetParameterValue(1, sCardCodeTo)
            objReport.SetParameterValue(2, _Date)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target File Name:" & sTargetFile, sFuncName)

            CrDiskFileDestinationOptions.DiskFileName = sTargetFileName '"C:\Users\sri\Desktop\Abeo\MBMS\Reports\PDF\SOA.pdf"
            CrExportOptions = objReport.ExportOptions
            With CrExportOptions
                'Set the destination to a disk file 
                .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                'Set the format to PDF 
                .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                'Set the destination options to DiskFileDestinationOptions object 
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With
            'Export the report 
            objReport.Export()

            ExportToPDF = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            ExportToPDF = RTN_ERROR
            Call WriteToLogFile(sErrDesc, sFuncName)
            Throw New ArgumentException(sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        Finally
            objReport.Dispose()
            crConnectionInfo = Nothing
            mySubRepDoc = Nothing
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
        End Try

    End Function


    Public Function ExportToPDF_New(ByVal iDocEntry As Integer, ByVal sCompanyDB As String, _
                                 ByRef sTargetFileName As String, _
                                 ByVal sRptFileName As String, _
                                 ByRef sErrDesc As String) As Long

        ' *********************************************************************************************
        '   Function   :   ExportToPDF()
        '   Purpose    :   ExportToPDF
        '   Parameters :   ByVal sPath As Integer
        '                  sPath=Report Path
        '                  ByRef sErrDesc As String
        '                   sErrDesc=Error Description to be returned to calling function
        '   Return     :   0 - FAILURE
        '                  1 - SUCCESS
        '   Date       :   29/11/2013
        '   Change     :
        ' *********************************************************************************************

        Dim sFuncName As String = String.Empty
        Dim intCounter As Integer
        Dim intCounter1 As Integer
        Dim iCount As Integer
        Dim iSubRParaCount As Integer
        'Crystal Report's report document object

        Dim objReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        'object of table Log on info of Crystal report
        Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
        Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
        Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
        Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table = Nothing
        'Sub report object of crystal report.
        Dim mySubReportObject As CrystalDecisions.CrystalReports.Engine.SubreportObject
        'Sub report document of crystal report.
        Dim mySubRepDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
        Dim CrDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions()
        Dim CrFormatTypeOptions As New CrystalDecisions.Shared.PdfRtfWordFormatOptions()
        Dim sSQL As String = String.Empty
        Dim sCompanyName As String = String.Empty
        Dim sSVCID As String = String.Empty


        Try
            sFuncName = "ExportToPDF()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Report File Name:" & sRptFileName, sFuncName)
            'Load the report
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Loading Report", sFuncName)

            objReport.Load(sRptFileName, CrystalDecisions.[Shared].OpenReportMethod.OpenReportByTempCopy)

            'Set the connection information to crConnectionInfo object so that we can apply the 
            ' connection information on each table in the reporteport
            With crConnectionInfo
                .DatabaseName = sCompanyDB.ToString()
                .Password = p_oCompDef.sDBPwd
                .ServerName = p_oCompDef.sReportDSN
                .UserID = p_oCompDef.sDBUser

            End With


            CrTables = objReport.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = p_oCompDef.sSAPDBName & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next
            ' Loop through each section on the report then look  through each object in the section
            ' if the object is a subreport, then apply logon info on each table of that sub report
            For iCount = 0 To objReport.ReportDefinition.Sections.Count - 1
                For intCounter = 0 To objReport.ReportDefinition.Sections(iCount).ReportObjects.Count - 1
                    With objReport.ReportDefinition.Sections(iCount)
                        If .ReportObjects(intCounter).Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then
                            mySubReportObject = CType(.ReportObjects(intCounter), CrystalDecisions.CrystalReports.Engine.SubreportObject)
                            mySubRepDoc = mySubReportObject.OpenSubreport(mySubReportObject.SubreportName)
                            'get the subreport parameter count to exclude passing data
                            iSubRParaCount += mySubRepDoc.DataDefinition.ParameterFields.Count
                            For intCounter1 = 0 To mySubRepDoc.Database.Tables.Count - 1
                                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                                CrTable.Location = p_oCompDef.sSAPDBName & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                            Next
                        End If
                    End With
                Next
            Next

            'Check if there are parameters or not in report and exclude the subreport parameters.
            intCounter = objReport.DataDefinition.ParameterFields.Count - iSubRParaCount
            'As parameter fields collection also picks the selection formula which is not the parameter
            ' so if total parameter count is 1 then we check whether its a parameter or selection formula.
            If intCounter = 1 Then
                If InStr(objReport.DataDefinition.ParameterFields(0).ParameterFieldName, ".", CompareMethod.Text) > 0 Then
                    intCounter = 0
                End If
            End If

            ' set the parameter to the report
            objReport.SetParameterValue(0, iDocEntry)

            'Export to PDF

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target File Name:" & sTargetFileName, sFuncName)

            CrDiskFileDestinationOptions.DiskFileName = sTargetFileName
            CrExportOptions = objReport.ExportOptions
            With CrExportOptions
                'Set the destination to a disk file 
                .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                'Set the format to PDF 
                .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                'Set the destination options to DiskFileDestinationOptions object 
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With
            'Export the report 
            objReport.Export()

            ExportToPDF_New = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            ExportToPDF_New = RTN_ERROR
            Call WriteToLogFile(sErrDesc, sFuncName)
            Throw New ArgumentException(sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        Finally
            objReport.Dispose()
            crConnectionInfo = Nothing
            mySubRepDoc = Nothing
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
        End Try

    End Function

End Module
