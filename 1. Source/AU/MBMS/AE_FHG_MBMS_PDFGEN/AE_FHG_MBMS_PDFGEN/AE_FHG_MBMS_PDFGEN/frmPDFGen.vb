Public Class frmPDFGen


    Public Sub WriteToStatusScreen(ByVal Clear As Boolean, ByVal msg As String)
        If Clear Then
            txtStatusMsg.Text = ""
        End If
        txtStatusMsg.HideSelection = True
        txtStatusMsg.Text &= msg & vbCrLf
        txtStatusMsg.SelectAll()
        txtStatusMsg.ScrollToCaret()
        txtStatusMsg.Refresh()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If (Me.DateTimePicker2.Checked = False) Then
            MsgBox("Please Choose the Date ", MsgBoxStyle.Information, "MBMS  PDF Generation")
            Exit Sub
        End If

        If (Me.DateTimePicker1.Checked = False) Then
            MsgBox("Please Choose the Reimbursed Date ", MsgBoxStyle.Information, "MBMS  PDF Generation")
            Exit Sub
        End If
        'If Me.txtBatchNo.Text = String.Empty Then
        '    MsgBox("Please Enter Batch No", MsgBoxStyle.Information, "MBMS PDF Generation")
        '    Exit Sub
        'End If

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        WriteToStatusScreen(True, "Please wait PDF Genearation in progress ...")
        Me.Button1.Enabled = Not Me.Button1.Enabled
        Call btnPDFGen_Click(Me, New System.EventArgs)
        Me.Button1.Enabled = Not Me.Button1.Enabled
        System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

  

    Public Function ExportReports(ByVal sErrdesc As String) As Long
        Dim sFuncName As String = "ExportReports"
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet
        Dim sTargetFileName As String = String.Empty
        Dim sRptFileName As String = String.Empty
        Dim sBatchDir As String = String.Empty
        Dim sPAYADVDir As String = String.Empty
        Dim sPAYADVDirTMP As String = String.Empty
        Dim sTAXINVDir As String = String.Empty
        Dim sTAXINVDirtTMP As String = String.Empty
        Dim sTPAFEEDir As String = String.Empty
        Dim sTPAFEEDirTMP As String = String.Empty
        Dim sARTAXINVDir As String = String.Empty
        Dim sARTAXINVDirTMP As String = String.Empty
        Dim sPARENTDIR As String = String.Empty
        Dim sPAYSMDir As String = String.Empty
        Dim sPAYSMDirTMP As String = String.Empty
        Dim sComboSelect As String = String.Empty
        Dim dReimbrusedDate As Date

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Dim dDate As Date = DateTimePicker2.Text
            dReimbrusedDate = DateTimePicker1.Text

            WriteToStatusScreen(False, "=============================== Start Date :: " & dDate.ToString("yyyyMMdd") & "  ===========================")

            If Not System.IO.Directory.Exists(p_oCompDef.sReportPDFPath) Then
                System.IO.Directory.CreateDirectory(p_oCompDef.sReportPDFPath)
            End If



            ' sBatchDir = p_oCompDef.sReportPDFPath & "\" & dDate.ToString("yyyyMMdd")
            sBatchDir = p_oCompDef.sReportPDFPath & "\" & Now.Date.ToString("yyyyMMdd")

            WriteToStatusScreen(False, "Creating Root Folder  :: " & Now.Date.ToString("yyyyMMdd"))

            If Not System.IO.Directory.Exists(sBatchDir) Then
                System.IO.Directory.CreateDirectory(sBatchDir)
            End If

            sPAYADVDir = "PAYADV_" & dDate.ToString("yyyyMMdd")
            sTAXINVDir = "TAXINV_" & dDate.ToString("yyyyMMdd")
            sTPAFEEDir = "TPAFEE_" & dDate.ToString("yyyyMMdd")
            sARTAXINVDir = "ARTAXINV_" & dDate.ToString("yyyyMMdd")
            sPAYSMDir = "PAYSUM_" & dDate.ToString("yyyyMMdd")

            sComboSelect = "Payment Advice"


            Select Case sComboSelect.Trim()

                Case "AR Invoice"

                    'If Not System.IO.Directory.Exists(sARTAXINVDir) Then
                    '    System.IO.Directory.CreateDirectory(sARTAXINVDir)
                    'End If

                    'If Me.txtBPF.Text = String.Empty Or Me.txtBPT.Text = String.Empty Then
                    '    sSQL = "SELECT ""DocEntry"",""CardCode"",""DocNum"",""NumAtCard"",""CardName"" from OINV " & _
                    '      " WHERE ""ImportEnt""=" & Me.txtBatchNo.Text
                    'Else
                    '    sSQL = "SELECT ""DocEntry"",""CardCode"",""DocNum"",""NumAtCard"",""CardName"" from OINV " & _
                    '      " WHERE ""ImportEnt""=" & Me.txtBatchNo.Text & " and  ""DocNum"">=" & Me.txtDocNumFrom.Text & " and ""DocNum""<=" & Me.txtDocNumTo.Text
                    'End If

                    'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    'oDs = ExecuteSQLQuery(sSQL)
                    'If oDs.Tables(0).Rows.Count > 0 Then
                    '    For Each row As DataRow In oDs.Tables(0).Rows
                    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                    '        WriteToStatusScreen(False, "Generating PDF for DocNo::" & row.Item("DocNum").ToString)

                    '        sTargetFileName = "BATCH" & Me.txtBatchNo.Text & "_ARTI_" & row.Item("CardCode").ToString & ".pdf"
                    '        sTargetFileName = sTAXINVDir & "\" & sTargetFileName

                    '        If UCase(row.Item("CardName").ToString) = UCase("AIA Singapore Pte. Ltd.") Then
                    '            sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_TaxInvoice_AIA.rpt"
                    '        ElseIf UCase(row.Item("CardName").ToString) = UCase("Lenovo (Singapore) Pte Ltd") Then
                    '            sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_TaxInvoice_Lenovo.rpt"
                    '        Else
                    '            sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_TaxInvoice.rpt"
                    '        End If

                    '        If ExportToPDF(row.Item("DocEntry"), sTargetFileName, sRptFileName, sErrdesc) <> RTN_SUCCESS Then
                    '            Throw New ArgumentException(sErrdesc)
                    '        End If
                    '        WriteToStatusScreen(False, "Successfully  generated PDF ::" & sTargetFileName)
                    '    Next
                    'End If

                Case "AP Invoice"

                    If Me.txtBPF.Text = String.Empty Or Me.txtBPT.Text = String.Empty Then
                        sSQL = "SELECT T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""NumAtCard"",T0.""CardName"",T0.""CardCode"" , T0.""ImportEnt"" , " & _
                        "T1.""U_AI_FHN3Code"" from OPCH T0 JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" " & _
                        " WHERE T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "'  and T0.""CANCELED"" = 'N'"
                    Else
                        sSQL = "SELECT T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""NumAtCard"",T0.""CardName"",T0.""CardCode"" , T0.""ImportEnt"" , " & _
                        "T1.""U_AI_FHN3Code"" from OPCH T0 JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" " & _
                         " WHERE T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' and T0.""CardCode"">= '" & Me.txtBPF.Text & "' and T0.""CardCode""<= '" & Me.txtBPT.Text & "'  and T0.""CANCELED"" = 'N'"
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    oDs = ExecuteSQLQuery(sSQL)
                    If oDs.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In oDs.Tables(0).Rows
                            sPARENTDIR = sBatchDir & "\" & row.Item("U_AI_FHN3Code").ToString & "_" & row.Item("CardCode").ToString
                            If Not System.IO.Directory.Exists(sPARENTDIR) Then
                                System.IO.Directory.CreateDirectory(sPARENTDIR)
                            End If
                            sTAXINVDirtTMP = sPARENTDIR & "\" & sTAXINVDir
                            If Not System.IO.Directory.Exists(sTAXINVDirtTMP) Then
                                System.IO.Directory.CreateDirectory(sTAXINVDirtTMP)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)

                            WriteToStatusScreen(False, "Generating PDF for Provider::" & row.Item("CardCode").ToString)

                            sTargetFileName = "BATCH" & row.Item("ImportEnt").ToString & "_TI_" & row.Item("CardCode").ToString & ".pdf"
                            sTargetFileName = sTAXINVDirtTMP & "\" & sTargetFileName

                            sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_AP_Invoice.rpt"

                            If ExportToPDF(row.Item("DocEntry"), sTargetFileName, sRptFileName, sErrdesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrdesc)
                            End If
                            WriteToStatusScreen(False, "Successfully  generated PDF ::" & sTargetFileName)
                        Next
                    End If

                Case "Credit Memo - TPA Fee"

                    'If Not System.IO.Directory.Exists(sTPAFEEDir) Then
                    '    System.IO.Directory.CreateDirectory(sTPAFEEDir)
                    'End If

                    If Me.txtBPF.Text = String.Empty Or Me.txtBPT.Text = String.Empty Then

                        sSQL = "SELECT T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""NumAtCard"",T0.""CardName"", T0.""ImportEnt"" , " & _
                      "T1.""U_AI_FHN3Code"" from ORPC T0 JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" " & _
                      " WHERE T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' AND ""CANCELED""='N'"
                    Else
                        sSQL = "SELECT T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""NumAtCard"",T0.""CardName"",T0.""CardCode"" , T0.""ImportEnt"" , " & _
                       "T1.""U_AI_FHN3Code"" from ORPC T0 JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" " & _
                        " WHERE T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' AND ""CANCELED""='N' and T0.""CardCode"">='" & Me.txtBPF.Text & "' and T0.""CardCode""<= '" & Me.txtBPT.Text & "'"
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    oDs = ExecuteSQLQuery(sSQL)
                    If oDs.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In oDs.Tables(0).Rows
                            sPARENTDIR = sBatchDir & "\" & row.Item("U_AI_FHN3Code").ToString & "_" & row.Item("CardCode").ToString
                            If Not System.IO.Directory.Exists(sPARENTDIR) Then
                                System.IO.Directory.CreateDirectory(sPARENTDIR)
                            End If
                            sTPAFEEDirTMP = sPARENTDIR & "\" & sTPAFEEDir
                            If Not System.IO.Directory.Exists(sTPAFEEDirTMP) Then
                                System.IO.Directory.CreateDirectory(sTPAFEEDirTMP)
                            End If
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                            WriteToStatusScreen(False, "Generating PDF for Provider::" & row.Item("CardCode").ToString)

                            sTargetFileName = "BATCH" & row.Item("ImportEnt").ToString & "_TF_" & row.Item("CardCode").ToString & ".pdf"
                            sTargetFileName = sTPAFEEDirTMP & "\" & sTargetFileName

                            sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_TaxInvoice_TPAFee.rpt"
                            If ExportToPDF(row.Item("DocEntry"), sTargetFileName, sRptFileName, sErrdesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrdesc)
                            End If
                            WriteToStatusScreen(False, "Successfully  generated PDF ::" & sTargetFileName)
                        Next
                    End If

                Case "Payment Advice"

                    'If Not System.IO.Directory.Exists(sPAYADVDir) Then
                    '    System.IO.Directory.CreateDirectory(sPAYADVDir)
                    'End If
                    'll
                    'If Me.txtBPF.Text = String.Empty Or Me.txtBPT.Text = String.Empty Then
                    '    sSQL = " SELECT T0.""DocEntry"",T0.""DocNum"",T0.""CardName"",T0.""CardCode"" , T2.""ImportEnt"", T3.""U_AI_FHN3Code"" from OVPM T0 " & _
                    '           " INNER JOIN VPM2 T1 ON T0.""DocNum""=T1.""DocNum"" " & _
                    '           " LEFT JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry"" " & _
                    '            " LEFT JOIN OCRD T3 ON T3.""CardCode""=T0.""CardCode"" " & _
                    '           " WHERE T1.""InvType""=18 and T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' and T0.""Canceled"" = 'N'"
                    'Else
                    '    sSQL = " SELECT T0.""DocEntry"",T0.""DocNum"",T0.""CardName"",T0.""CardCode"" , T2.""ImportEnt"" , T3.""U_AI_FHN3Code""  from OVPM T0 " & _
                    '            " INNER JOIN VPM2 T1 ON T0.""DocNum""=T1.""DocNum"" " & _
                    '            " LEFT JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry"" " & _
                    '             " LEFT JOIN OCRD T3 ON T3.""CardCode""=T0.""CardCode"" " & _
                    '            " WHERE T1.""InvType""=18 and T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' " & _
                    '            " and T0.""CardCode"">='" & Me.txtBPF.Text & "' and T0.""CardCode""<= '" & Me.txtBPT.Text & "' and T0.""Canceled"" = 'N'"
                    'End If
                    If Me.txtBPF.Text = String.Empty Or Me.txtBPT.Text = String.Empty Then
                        sSQL = " SELECT T0.""DocEntry"",T0.""DocNum"",T0.""CardName"",T0.""CardCode"" ,   T3.""U_AI_FHN3Code"" from OVPM T0 " & _
                               " INNER JOIN VPM2 T1 ON T0.""DocNum""=T1.""DocNum"" " & _
                               " LEFT JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry"" " & _
                                " LEFT JOIN OCRD T3 ON T3.""CardCode""=T0.""CardCode"" " & _
                               " WHERE T1.""InvType""=18 and T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' and T0.""Canceled"" = 'N'" & _
                               " GROUP BY T0.""DocEntry"",T0.""DocNum"",T0.""CardName"",T0.""CardCode"" ,   T3.""U_AI_FHN3Code"" "
                    Else
                        sSQL = " SELECT T0.""DocEntry"",T0.""DocNum"",T0.""CardName"",T0.""CardCode"" ,   T3.""U_AI_FHN3Code""  from OVPM T0 " & _
                                " INNER JOIN VPM2 T1 ON T0.""DocNum""=T1.""DocNum"" " & _
                                " LEFT JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry"" " & _
                                 " LEFT JOIN OCRD T3 ON T3.""CardCode""=T0.""CardCode"" " & _
                                " WHERE T1.""InvType""=18 and T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' " & _
                                " and T0.""CardCode"">='" & Me.txtBPF.Text & "' and T0.""CardCode""<= '" & Me.txtBPT.Text & "' and T0.""Canceled"" = 'N' " & _
                                " GROUP BY T0.""DocEntry"",T0.""DocNum"",T0.""CardName"",T0.""CardCode"" ,   T3.""U_AI_FHN3Code"" "

                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    oDs = ExecuteSQLQuery(sSQL)
                    If oDs.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In oDs.Tables(0).Rows
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                            sPARENTDIR = sBatchDir & "\" & row.Item("U_AI_FHN3Code").ToString & "_" & row.Item("CardCode").ToString
                            If Not System.IO.Directory.Exists(sPARENTDIR) Then
                                System.IO.Directory.CreateDirectory(sPARENTDIR)
                            End If
                            sPAYADVDirTMP = sPARENTDIR & "\" & sPAYADVDir
                            If Not System.IO.Directory.Exists(sPAYADVDirTMP) Then
                                System.IO.Directory.CreateDirectory(sPAYADVDirTMP)
                            End If

                            WriteToStatusScreen(False, "Generating PDF for CardName::" & row.Item("CardCode").ToString)
                            sTargetFileName = "DOCNUM" & row.Item("DocNum").ToString & "_PA_" & row.Item("CardCode").ToString & ".pdf"
                            sTargetFileName = sPAYADVDirTMP & "\" & sTargetFileName

                            'sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_PaymentAdvice.rpt"
                            'If ExportToPDF(row.Item("DocEntry"), sTargetFileName, sRptFileName, sErrdesc) <> RTN_SUCCESS Then
                            '    Throw New ArgumentException(sErrdesc)
                            'End If

                            ''sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_PaymentAdviceByBatch.rpt"
                            sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_PaymentAdvice.rpt"
                            If ExportToPDF_PaymentAdvice(row.Item("DocEntry"), sTargetFileName, sRptFileName, dReimbrusedDate, sErrdesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrdesc)
                            End If


                            WriteToStatusScreen(False, "Successfully  generated PDF ::" & sTargetFileName)
                        Next
                    End If
                Case "Payment Summary"

                    'If Not System.IO.Directory.Exists(sPAYADVDir) Then
                    '    System.IO.Directory.CreateDirectory(sPAYADVDir)
                    'End If
                    'll
                    If Me.txtBPF.Text = String.Empty Or Me.txtBPT.Text = String.Empty Then
                        'sSQL = " SELECT T0.""DocEntry"",T0.""DocNum"",T0.""CardName"",T0.""CardCode"" , T2.""ImportEnt"", T3.""U_AI_FHN3Code"" from OVPM T0 " & _
                        '       " INNER JOIN VPM2 T1 ON T0.""DocNum""=T1.""DocNum"" " & _
                        '       " LEFT JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry"" " & _
                        '        " LEFT JOIN OCRD T3 ON T3.""CardCode""=T0.""CardCode"" " & _
                        '       " WHERE T1.""InvType""=18 and T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "'"
                        sSQL = " SELECT T0.""DocEntry"",T0.""DocNum"" ""ImportEnt"",T0.""CardName"",T0.""CardCode"" , T3.""U_AI_FHN3Code"" from OVPM T0 " & _
                               " INNER JOIN VPM2 T1 ON T0.""DocNum""=T1.""DocNum"" " & _
                               " LEFT JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry"" " & _
                                " LEFT JOIN OCRD T3 ON T3.""CardCode""=T0.""CardCode"" " & _
                               " WHERE T1.""InvType""=18 and T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' and T0.""Canceled"" = 'N' GROUP BY T0.""DocEntry"",T0.""DocNum"" ,T0.""CardName"",T0.""CardCode"" , T3.""U_AI_FHN3Code"" "
                    Else
                        'sSQL = " SELECT T0.""DocEntry"",T0.""DocNum""  ,T0.""CardName"",T0.""CardCode"" , T2.""ImportEnt"" ,  T3.""U_AI_FHN3Code""  from OVPM T0 " & _
                        '        " INNER JOIN VPM2 T1 ON T0.""DocNum""=T1.""DocNum"" " & _
                        '        " LEFT JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry"" " & _
                        '         " LEFT JOIN OCRD T3 ON T3.""CardCode""=T0.""CardCode"" " & _
                        '        " WHERE T1.""InvType""=18 and T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' " & _
                        '        " and T0.""CardCode"">='" & Me.txtBPF.Text & "' and T0.""CardCode""<= '" & Me.txtBPT.Text & "'"

                        sSQL = " SELECT T0.""DocEntry"",T0.""DocNum"" ""ImportEnt"",T0.""CardName"",T0.""CardCode""  , T3.""U_AI_FHN3Code""  from OVPM T0 " & _
                               " INNER JOIN VPM2 T1 ON T0.""DocNum""=T1.""DocNum"" " & _
                               " LEFT JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry"" " & _
                                " LEFT JOIN OCRD T3 ON T3.""CardCode""=T0.""CardCode"" " & _
                               " WHERE T1.""InvType""=18 and T0.""DocDate""='" & dDate.ToString("yyyyMMdd") & "' " & _
                               " and T0.""CardCode"">='" & Me.txtBPF.Text & "' and T0.""CardCode""<= '" & Me.txtBPT.Text & "' and T0.""Canceled"" = 'N' GROUP BY T0.""DocEntry"",T0.""DocNum"" ,T0.""CardName"",T0.""CardCode"" , T3.""U_AI_FHN3Code"" "
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    oDs = ExecuteSQLQuery(sSQL)
                    If oDs.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In oDs.Tables(0).Rows
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                            sPARENTDIR = sBatchDir & "\" & row.Item("U_AI_FHN3Code").ToString & "_" & row.Item("CardCode").ToString
                            If Not System.IO.Directory.Exists(sPARENTDIR) Then
                                System.IO.Directory.CreateDirectory(sPARENTDIR)
                            End If
                            sPAYSMDirTMP = sPARENTDIR & "\" & sPAYSMDir
                            If Not System.IO.Directory.Exists(sPAYSMDirTMP) Then
                                System.IO.Directory.CreateDirectory(sPAYSMDirTMP)
                            End If

                            WriteToStatusScreen(False, "Generating PDF for CardName::" & row.Item("CardCode").ToString)
                            sTargetFileName = row.Item("ImportEnt").ToString & "_PS_" & row.Item("CardCode").ToString & ".pdf"
                            sTargetFileName = sPAYSMDirTMP & "\" & sTargetFileName

                            'sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_PaymentAdvice.rpt"
                            'If ExportToPDF(row.Item("DocEntry"), sTargetFileName, sRptFileName, sErrdesc) <> RTN_SUCCESS Then
                            '    Throw New ArgumentException(sErrdesc)
                            'End If

                            sRptFileName = p_oCompDef.sReportsPath & "\Payment Advice Form.rpt"
                            If ExportToPDF_PaymentSummary(row.Item("DocEntry"), sTargetFileName, sRptFileName, sErrdesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrdesc)
                            End If
                            WriteToStatusScreen(False, "Successfully  generated PDF ::" & sTargetFileName)
                        Next
                    End If
            End Select

            WriteToStatusScreen(False, "================================ Completed Date :: " & dDate.ToString("yyyyMMdd") & "  ===========================")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed successfully.", sFuncName)
            ExportReports = RTN_SUCCESS
        Catch ex As Exception
            sErrdesc = ex.Message
            ExportReports = RTN_ERROR
            Call WriteToLogFile(sErrdesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        End Try
    End Function

    Public Function ExportToPDF(ByVal iDocEntry As Integer, _
                                ByVal sTargetFileName As String, _
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
            'oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Set the connection information to crConnectionInfo object so that we can apply the 
            ' connection information on each table in the reporteport

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

    Public Function ExportToPDF_PaymentAdvice(ByVal iDocEntry As Integer, _
                                ByVal sTargetFileName As String, _
                                ByVal sRptFileName As String, ByVal ddate As Date, _
                                ByRef sErrDesc As String) As Long

        ' *********************************************************************************************
        '   Function   :   ExportToPDF_PaymentAdvice()
        '   Purpose    :   ExportToPDF_PaymentAdvice
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
            'oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Set the connection information to crConnectionInfo object so that we can apply the 
            ' connection information on each table in the reporteport

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
            objReport.SetParameterValue(1, ddate)

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

            ExportToPDF_PaymentAdvice = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            ExportToPDF_PaymentAdvice = RTN_ERROR
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

    Public Function ExportToPDF_PaymentSummary(ByVal iDocEntry As Integer, _
                              ByVal sTargetFileName As String, _
                              ByVal sRptFileName As String, _
                              ByRef sErrDesc As String) As Long

        ' *********************************************************************************************
        '   Function   :   ExportToPDF_PaymentSummary()
        '   Purpose    :   ExportToPDF_PaymentSummary
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
            'oRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Set the connection information to crConnectionInfo object so that we can apply the 
            ' connection information on each table in the reporteport

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
            '' objReport.SetParameterValue(1, iBatchNo)

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

            ExportToPDF_PaymentSummary = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            ExportToPDF_PaymentSummary = RTN_ERROR
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


    Private Sub btnPDFGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPDFGen.Click
        Dim sErrdesc As String = String.Empty
        If ExportReports(sErrdesc) <> RTN_SUCCESS Then
            MsgBox(sErrdesc)
            Exit Sub
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.txtStatusMsg.Clear()
        Me.txtBPF.Clear()
        Me.txtBPT.Clear()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    
    Private Sub frmPDFGen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sErrdesc As String = String.Empty
        If GetSystemIntializeInfo(p_oCompDef, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
    End Sub

    Private Sub Bt_CFL1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bt_CFL1.Click
        Dim sErrdesc As String = String.Empty
       
        Dim ssql As String = "SELECT T0.""CardCode"", T0.""CardName"" FROM OCRD T0 WHERE T0.""CardType"" = 'S' and T0.""CardCode"" <> '' ORDER BY T0.""CardCode"""
        Dim oCFL = New CFL(ssql)
        oCFL.ShowDialog()
        txtBPF.Text = sCFL

    End Sub

    Private Sub Bt_CFL2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bt_CFL2.Click
        Dim sErrdesc As String = String.Empty

        Dim ssql As String = "SELECT T0.""CardCode"", T0.""CardName"" FROM OCRD T0 WHERE T0.""CardType"" = 'S' and T0.""CardCode"" <> '' ORDER BY T0.""CardCode"""
        Dim oCFL = New CFL(ssql)
        oCFL.ShowDialog()
        txtBPT.Text = sCFL
    End Sub

   
End Class
