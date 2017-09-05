
Imports System.IO
Imports Microsoft.VisualBasic

Public Class frmPaymentadvice

    Dim sErrdesc As String = String.Empty

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

  

    Private Function ExportReports(ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "ExportReports"
        Dim sSQL As String = String.Empty
        Dim oDS As New DataSet
        Dim oDS1 As New DataSet
        Dim sTargetFileName As String = String.Empty
        Dim sRptFileName As String = String.Empty
        Dim sBatchDir As String = String.Empty
        Dim sPAYADVDir As String = String.Empty
        Dim sTAXINVDir As String = String.Empty
        Dim sTPAFEEDir As String = String.Empty
        Dim sARTAXINVDir As String = String.Empty

        Dim sDocF As String = String.Empty
        Dim sDocT As String = String.Empty
        Dim dDateF, dDateT As Date
        Dim sBPCode As String = String.Empty
        Dim sPayCode As String = String.Empty
        Dim sBpcodeTo As String = String.Empty
        Dim sBPcode_Array(100, 2) As String
        Dim iCount_Array As Integer = 0
        Dim sBody As String = String.Empty
        Dim sSubject As String = String.Empty
        Dim sEPaymentEmail, sESOAEmail As String
        Dim sEmail, sDocentry, sInvoice As String

        Dim sSOARptName As String = String.Empty
        Dim sPytRptName As String = String.Empty
        Dim sPytDtlRptName As String = String.Empty


        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            WriteToStatusScreen(False, "=============================== Start  ===========================")

            If Not System.IO.Directory.Exists(p_oCompDef.sReportsPath) Then
                System.IO.Directory.CreateDirectory(p_oCompDef.sReportsPath)
                WriteToStatusScreen(False, "Created Folder for Report :: " & p_oCompDef.sReportsPath)
            End If

            'Get E-Advice contanct email address
            sSQL = "select ""AliasName"" as ePaymentEmail,""E_Mail"" as eSOAEmail from " & """" & Me.SCompany.Text & """" & ".OADM T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
            oDS = ExecuteSQLQuery(sSQL)
            If oDS.Tables(0).Rows.Count > 0 Then
                sEPaymentEmail = oDS.Tables(0).Rows(0).Item(0).ToString '"sahayar@abeo-electra.com" '
                sESOAEmail = oDS.Tables(0).Rows(0).Item(1).ToString
            End If


            sSQL = "Select * from " & """" & Me.SCompany.Text & """" & ".""@EADVISE"""
            p_oCompDef.sSAPDBName = Me.SCompany.Text

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
            oDS = ExecuteSQLQuery(sSQL)
            If oDS.Tables(0).Rows.Count > 0 Then
                sSOARptName = oDS.Tables(0).Rows(0).Item("U_SOARptName").ToString
                sPytRptName = oDS.Tables(0).Rows(0).Item("U_PARptName").ToString
                sPytDtlRptName = oDS.Tables(0).Rows(0).Item("U_PARptDet").ToString
            End If


            Select Case DocType.Text

                Case "Payment Advice"

                    WriteToStatusScreen(False, "Attempting to Payment Advise Portion :: ")
                    iCount_Array = 0
                    ReDim sBPcode_Array(gvtaxinvoice.Rows.Count, 2)

                    For imjs As Integer = 0 To gvtaxinvoice.Rows.Count - 1
                        If Convert.ToBoolean(gvtaxinvoice.Rows(imjs).Cells(0).Value) = True Then
                            sEmail = Convert.ToString(gvtaxinvoice.Rows(imjs).Cells("clmnemail").Value)
                            sDocentry = Convert.ToString(gvtaxinvoice.Rows(imjs).Cells("clmdocentry").Value)
                            sInvoice = Convert.ToString(gvtaxinvoice.Rows(imjs).Cells("clmnInvoiceno").Value)
                            lblMsg.Text = "Processing the Invoice " & sInvoice
                            Me.Refresh()


                            sSQL = "SELECT T1.""E_Mail"",T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""CardName"", T0.""DocDate"",T1.""U_AI_EMAILSOA"" " & _
                " FROM " & """" & Me.SCompany.Text & """" & ".OVPM T0  LEFT JOIN " & """" & Me.SCompany.Text & """" & ".OCRD T1" & _
                " ON T0.""CardCode""= T1.""CardCode"" LEFT JOIN " & """" & Me.SCompany.Text & """" & ".OOND T2 ON T1.""IndustryC""= T2.""IndCode"" " & _
                "WHERE T1.""U_AI_EMAILSOA""='YES' and T0.""Canceled"" <>'Y'  and T0.""DocEntry"" = '" & sDocentry & "'"

                            sSQL += "GROUP BY T1.""E_Mail"",T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""CardName"", T0.""DocDate"", T1.""U_AI_EMAILSOA"" ORDER BY T0.""DocNum"""

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                            oDS = ExecuteSQLQuery(sSQL)

                            If oDS.Tables(0).Rows.Count > 0 Then
                                For Each row As DataRow In oDS.Tables(0).Rows
                                    If row.Item("E_Mail").ToString <> String.Empty Then
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                                        WriteToStatusScreen(False, "Generating PDF for DocNo::" & row.Item("DocNum").ToString)
                                        Dim dDateCon As Date = row.Item("DocDate").ToString

                                        sTargetFileName = "PA_" & row.Item("DocNum").ToString & "_" & row.Item("CardCode").ToString & "_" & dDateCon.ToString("ddMMyyyy") & ".pdf"
                                        sTargetFileName = p_oCompDef.sReportPDFPath & "\" & sTargetFileName
                                        sRptFileName = p_oCompDef.sReportsPath & "\" & sPytRptName
                                        ' ''If Microsoft.VisualBasic.Left(row.Item("CardCode").ToString, 2) = "MV" Then
                                        ' ''    'sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_PaymentAdvice_Dtls.rpt"
                                        ' ''    sRptFileName = p_oCompDef.sReportsPath & "\" & sPytDtlRptName
                                        ' ''Else
                                        ' ''    'sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_PaymentAdvice.rpt"
                                        ' ''    sRptFileName = p_oCompDef.sReportsPath & "\" & sPytRptName
                                        ' ''End If

                                        If ExportToPDF_New(row.Item("DocEntry"), SCompany.Text.ToString, sTargetFileName, sRptFileName, sErrdesc) <> RTN_SUCCESS Then
                                            Throw New ArgumentException(sErrdesc)
                                        End If

                                        ' WriteToStatusScreen(False, "Successfully  generated PDF ::" & sTargetFileName)
                                        WriteToStatusScreen(False, "Successfully  generated the PDF")
                                        WriteToStatusScreen(False, "Sending Email for Payment No ::" & row.Item("DocNum").ToString)

                                        sBody = "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                        sBody = sBody & " Dear Sir/Madam,<br /><br /> Account Name: " & row.Item("CardCode").ToString.Trim & "<br /><br /><br />"
                                        ''    sBody = sBody & " This payment will be credited to your account on " & Me.DateTimePicker2.Value.Date.ToString("dd/MM/yyyy") & ".<br /><br/><br/>"
                                        sBody = sBody & " Please find attached your remittance.  This payment will be credited to your account on " & Format(row.Item("DocDate"), "dd-MMM-yyyy") & " <br /><br />"

                                        Dim sCompanyName As String = String.Empty
                                        oCompList.TryGetValue(Me.SCompany.Text, sCompanyName)

                                        sSubject = "Accounts payable remittances " ''"Payment advice from " & sCompanyName
                                        Dim sFromEmail As String = String.Empty
                                        sFromEmail = cmbSenderEmail.Text.Trim()

                                        ''row.Item("E_Mail").ToString
                                        '' If SendEmailNotification(sTargetFileName, sFromEmail, "sahayar@abeo-electra.com", sBody, sSubject, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                                        If SendEmailNotification(sTargetFileName, sFromEmail, row.Item("E_Mail").ToString, sBody, sSubject, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                                        WriteToStatusScreen(False, "Successfully sent email Payment No ::" & row.Item("DocNum").ToString)
                                        WriteToStatusScreen(False, "===========================================================")
                                        gvtaxinvoice.Rows(imjs).Cells(5).Value = "Successfully sent Email "
                                    Else
                                        iCount_Array += 1
                                        sBPcode_Array(iCount_Array, 0) = row.Item("CardCode").ToString
                                        sBPcode_Array(iCount_Array, 1) = row.Item("CardName").ToString
                                        sBPcode_Array(iCount_Array, 2) = row.Item("DocNum").ToString
                                        gvtaxinvoice.Rows(imjs).Cells(5).Value = "Email address is blank for the Payment No. :: " & sInvoice
                                        gvtaxinvoice.Rows(imjs).Cells(5).Style.ForeColor = Color.DarkRed
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email address is blank for the Customer  - Invoice No." & row.Item("DocNum").ToString, sFuncName)
                                        WriteToStatusScreen(False, "Email address is blank for the Customer in this Invoice No. :: " & row.Item("DocNum").ToString)
                                    End If
                                Next
                            Else
                                WriteToStatusScreen(False, "No records Found for the above selection, Please try again with different Selection.")
                                gvtaxinvoice.Rows(imjs).Cells(5).Value = "No outstanding Balance "
                            End If
                        End If
                    Next imjs
                    '-----------------
                    lblMsg.Text = "Completed Successfully "
                    Me.Refresh()
                    If iCount_Array > 0 Then
                        Write_TextFile_BPList(sBPcode_Array, "0")
                    End If
                    '---------------

                    ' ''Case "SOA"

                    ' ''    WriteToStatusScreen(False, "Attempting to Satement of Accounts Portion :: ")

                    ' ''    sSQL = "SELECT T0.""E_Mail"",T0.""CardCode"", T0.""CardName"" from " & """" & Me.SCompany.Text & """" & ".OCRD T0 " & _
                    ' ''              " WHERE T0.""frozenFor"" ='N' AND T0.""U_AI_EMAILSOA""='YES' AND T0.""CardCode""  between  '" & sBPCode & "' and '" & sBpcodeTo & "'"

                    ' ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Satement of Accounts Portion", sFuncName)
                    ' ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    ' ''    oDS = ExecuteSQLQuery(sSQL)

                    ' ''    iCount_Array = 0
                    ' ''    ReDim sBPcode_Array(oDS.Tables(0).Rows.Count, 2)
                    ' ''    If oDS.Tables(0).Rows.Count > 0 Then
                    ' ''        For Each row As DataRow In oDS.Tables(0).Rows
                    ' ''            If row.Item("E_Mail").ToString <> String.Empty Then
                    ' ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                    ' ''                WriteToStatusScreen(False, "Generating PDF for BP::" & row.Item("CardName").ToString)
                    ' ''                ''  Dim dDateCon As Date = DateTimePicker1.Text

                    ' ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking BP Balance ..", sFuncName)
                    ' ''                sSQL = "CALL " & """" & Me.SCompany.Text & """" & ".""AE_SP002_SOA"" ('" & row.Item("CardCode").ToString & "','" & row.Item("CardCode").ToString & "','" & dDateCon.ToString("yyyyMMdd") & "')"
                    ' ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                    ' ''                oDS1 = ExecuteSQLQuery(sSQL)

                    ' ''                If oDS1.Tables(0).Rows.Count > 0 Then

                    ' ''                    ''  sTargetFileName = "SOA_" & row.Item(1).ToString & "_" & dDateCon.ToString("ddMMyyyy") & ".pdf"
                    ' ''                    sTargetFileName = p_oCompDef.sReportPDFPath & "\" & sTargetFileName
                    ' ''                    'sRptFileName = p_oCompDef.sReportsPath & "\AI_RPT_SOA.rpt"

                    ' ''                    sRptFileName = p_oCompDef.sReportsPath & "\" & sSOARptName

                    ' ''                    If ExportToPDF(row.Item("CardCode").ToString, row.Item("CardCode").ToString, dDateCon, sTargetFileName, sRptFileName, sErrdesc) <> RTN_SUCCESS Then
                    ' ''                        Throw New ArgumentException(sErrdesc)
                    ' ''                    End If

                    ' ''                    WriteToStatusScreen(False, "Successfully  generated PDF ::" & sTargetFileName)

                    ' ''                    sBody = "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                    ' ''                    sBody = sBody & " Dear Sir/Madam,<br /><br /> Please find the latest Statement of Account. We appreciate your settlement of the outstanding amount to keep to credit terms.<br /><br />"
                    ' ''                    sBody = sBody & " Please check your statement and contact us within 7 days if there is any discrepancy. <br /><br/><br/>"
                    ' ''                    sBody = sBody & " Please do not reply to this email. For clarification, please contact " & sESOAEmail & "<br /><br />"

                    ' ''                    Dim sCompanyName As String = String.Empty
                    ' ''                    oCompList.TryGetValue(Me.SCompany.Text, sCompanyName)

                    ' ''                    ''  sSubject = "Important - Statement of Account as at " & Me.DateTimePicker1.Value.ToString("dd MM yyyy") & " from " & sCompanyName

                    ' ''                    WriteToStatusScreen(False, "Sending Email for BP ::" & row.Item("CardName").ToString)
                    ' ''                    If SendEmailNotification(sTargetFileName, row.Item("E_Mail").ToString, sBody, sSubject, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

                    ' ''                    ''    SOAEmailLog(row.Item("CardCode").ToString, row.Item("CardName").ToString, Me.DateTimePicker1.Value, row.Item("E_Mail").ToString)

                    ' ''                    WriteToStatusScreen(False, "Successfully sent email BP ::" & row.Item("CardName").ToString)
                    ' ''                    WriteToStatusScreen(False, "===========================================================")
                    ' ''                Else
                    ' ''                    WriteToStatusScreen(False, "No outstanding Balance for BP::" & row.Item("CardName").ToString)
                    ' ''                End If

                    ' ''            Else
                    ' ''                iCount_Array += 1
                    ' ''                sBPcode_Array(iCount_Array, 0) = row.Item("CardCode").ToString
                    ' ''                sBPcode_Array(iCount_Array, 1) = row.Item("CardName").ToString

                    ' ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email address is blank for the Customer " & row.Item("CardName").ToString, sFuncName)
                    ' ''                WriteToStatusScreen(False, "Email address is blank for the Customer :: " & row.Item("CardName").ToString)
                    ' ''            End If

                    ' ''        Next

                    ' ''        If iCount_Array > 0 Then
                    ' ''            Write_TextFile_BPList(sBPcode_Array, "1")
                    ' ''        End If
                    ' ''    Else
                    ' ''        WriteToStatusScreen(False, "No records Found for the above selection, Please try again with different Selection.")
                    ' ''    End If

                    ' ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Satement of Accounts Completed Successfully :: ", sFuncName)


                    ' ''    WriteToStatusScreen(False, "Satement of Accounts Completed Successfully :: ")

            End Select

            WriteToStatusScreen(False, "================================ Completed ===========================")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed successfully.", sFuncName)
            ExportReports = RTN_SUCCESS
        Catch ex As Exception
            sErrdesc = ex.Message
            ExportReports = RTN_ERROR
            Call WriteToLogFile(sErrdesc, sFuncName)
            WriteToStatusScreen(False, "ERROR::" & sErrdesc)
            MsgBox(sErrdesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        End Try
    End Function

    Private Sub btnPDFGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPDFGen.Click
        Dim sErrdesc As String = String.Empty
        If ExportReports(sErrdesc) <> RTN_SUCCESS Then
            'Throw New ArgumentException(sErrdesc)
            'MsgBox(sErrdesc)

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.txtStatusMsg.Clear()
        Me.txtDocNumFrom.Clear()
        Me.txtDocNumTo.Clear()

        Me.txtBpCode.Clear()

    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub frmPDFGen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        oCompany = New SAPbobsCOM.Company
        ''------------  Selecting the Document Type and set Select as default
        DocType.SelectedIndex = 0
        ''-------------------------------------------------------------------

        If GetSystemIntializeInfo(p_oCompDef, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

        If Get_CompanyList(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)


    End Sub

    Private Sub DocType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DocType.SelectedIndexChanged

        If DocType.SelectedItem <> "---Select---" Then
            If DocType.SelectedItem <> "SOA" Then
                DocNo.Visible = True
            Else
                DocNo.Visible = False
            End If
        End If



    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim sDocF As String = String.Empty
        Dim sDocT As String = String.Empty
        Dim dDateF, dDateT As Date
        Dim sBPCode As String = String.Empty
        Dim sPayCode As String = String.Empty
        Dim sBpcodeTo As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDS As New DataSet
        Dim sFuncName As String = String.Empty


        If Me.SCompany.Text = "--Select Company-- " Then Exit Sub
        If Me.DocType.Text = "---Select---" Then Exit Sub
        If Me.cmbSenderEmail.Text = String.Empty Then Exit Sub
        sFuncName = "Show()"

        If Not Me.DocType.Text = "SOA" Then
            If Me.txtDocNumFrom.Text = String.Empty And Me.txtDocNumTo.Text = String.Empty And Me.txtBpCode.Text = String.Empty _
                And Me.DateTimePicker4.Checked = False And Me.DateTimePicker5.Checked = False Then
                MsgBox("Enter either DocNum range OR DocDate range OR BP Code", MsgBoxStyle.Information, "E-Advice")
                Exit Sub
            End If

            If (Me.txtDocNumFrom.Text = String.Empty And Me.txtDocNumTo.Text <> String.Empty) Or (Me.txtDocNumFrom.Text <> String.Empty And Me.txtDocNumTo.Text = String.Empty) Then
                MsgBox("Please enter the DocNum Range From and To", MsgBoxStyle.Information, "E-Advice")
                Exit Sub
            End If

            If (Me.DateTimePicker4.Checked = True And Me.DateTimePicker5.Checked = False) Or (Me.DateTimePicker4.Checked = False And Me.DateTimePicker5.Checked = True) Then
                MsgBox("Please enter the DocDate Range From and To", MsgBoxStyle.Information, "E-Advice")
                Exit Sub
            End If
        End If

        If Me.DocType.Text = "Payment Advice" Then
            If Me.DateTimePicker2.Checked = False Then
                MsgBox("Please select value date", MsgBoxStyle.Information, "E-Advice")
                Exit Sub
            End If
        End If

        If (Me.txtDocNumFrom.Text <> String.Empty And Me.txtDocNumTo.Text <> String.Empty) And _
                       (Me.DateTimePicker4.Checked = False And Me.DateTimePicker5.Checked = False) And _
                       Me.txtBpCode.Text = String.Empty Then

            sDocF = Me.txtDocNumFrom.Text
            sDocT = Me.txtDocNumTo.Text
            dDateF = "1/1/1"
            dDateT = "1/1/1"
            sBPCode = "%"
        ElseIf (Me.txtDocNumFrom.Text <> String.Empty And Me.txtDocNumTo.Text <> String.Empty) And _
           (Me.DateTimePicker4.Checked = True And Me.DateTimePicker5.Checked = True) And _
           Me.txtBpCode.Text = String.Empty Then

            sDocF = Me.txtDocNumFrom.Text
            sDocT = Me.txtDocNumTo.Text
            dDateF = Me.DateTimePicker4.Value
            dDateT = Me.DateTimePicker5.Value
            sBPCode = "%"

        ElseIf (Me.txtDocNumFrom.Text <> String.Empty And Me.txtDocNumTo.Text <> String.Empty) And _
           (Me.DateTimePicker4.Checked = True And Me.DateTimePicker5.Checked = True) And _
           Me.txtBpCode.Text <> String.Empty Then

            sDocF = Me.txtDocNumFrom.Text
            sDocT = Me.txtDocNumTo.Text
            dDateF = Me.DateTimePicker4.Value
            dDateT = Me.DateTimePicker5.Value
            sBPCode = Me.txtBpCode.Text

        ElseIf (Me.txtDocNumFrom.Text = String.Empty And Me.txtDocNumTo.Text = String.Empty) And _
               (Me.DateTimePicker4.Checked = True And Me.DateTimePicker5.Checked = True) And _
               Me.txtBpCode.Text <> String.Empty Then

            sDocF = "%"
            sDocT = "%"
            dDateF = Me.DateTimePicker4.Value
            dDateT = Me.DateTimePicker5.Value
            sBPCode = Me.txtBpCode.Text

        ElseIf (Me.txtDocNumFrom.Text = String.Empty And Me.txtDocNumTo.Text = String.Empty) And _
           (Me.DateTimePicker4.Checked = True And Me.DateTimePicker5.Checked = True) And _
           Me.txtBpCode.Text = String.Empty Then

            sDocF = "%"
            sDocT = "%"
            dDateF = Me.DateTimePicker4.Value
            dDateT = Me.DateTimePicker5.Value
            sBPCode = "%"

        ElseIf (Me.txtDocNumFrom.Text <> String.Empty And Me.txtDocNumTo.Text <> String.Empty) And _
       (Me.DateTimePicker4.Checked = False And Me.DateTimePicker5.Checked = False) And _
       Me.txtBpCode.Text <> String.Empty Then

            sDocF = Me.txtDocNumFrom.Text
            sDocT = Me.txtDocNumTo.Text
            dDateF = "1/1/1"
            dDateT = "1/1/1"
            sBPCode = Me.txtBpCode.Text

        ElseIf (Me.txtDocNumFrom.Text = String.Empty And Me.txtDocNumTo.Text = String.Empty) And _
  (Me.DateTimePicker4.Checked = False And Me.DateTimePicker5.Checked = False) And _
  Me.txtBpCode.Text <> String.Empty Then

            sDocF = "%"
            sDocT = "%"
            dDateF = "1/1/1"
            dDateT = "1/1/1"
            sBPCode = Me.txtBpCode.Text
        End If


        If Not String.IsNullOrEmpty(Me.txtpaymentcode.Text) Then
            sPayCode = Me.txtpaymentcode.Text
        Else
            sPayCode = "%"
        End If

        sSQL = "SELECT T1.""E_Mail"",T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""CardName"", T0.""DocDate"",T1.""U_AI_EMAILSOA"" from " & """" & Me.SCompany.Text & """" & ".OVPM T0 " & _
                         " LEFT OUTER JOIN " & """" & Me.SCompany.Text & """" & ".OCRD T1 ON T0.""CardCode""= T1.""CardCode"" LEFT OUTER JOIN " & """" & Me.SCompany.Text & """" & ".OPEX T2 " & _
                         "ON T0.""DocEntry""= T2.""PaymDocNum"""

        If sDocF = "%" And dDateT = "1/1/1" Then
            sSQL += " WHERE ifnull(""TrsfrAcct"",'') <>'' AND ((LEFT(T0.""CardCode"",2)<>'MV') OR  (LEFT(T0.""CardCode"",2)='MV' AND T1.""U_AI_EMAILSOA""='YES'))  AND T0.""Canceled"" <>'Y' AND T0.""CardCode""  like '" & sBPCode & "'"
        ElseIf sDocF <> "%" And dDateF = "1/1/1" Then
            sSQL += " WHERE ifnull(""TrsfrAcct"",'') <>'' AND ""DocNum"">= " & sDocF & " and ""DocNum""<= " & sDocT &
              " AND ((LEFT(T0.""CardCode"",2)<>'MV') OR  (LEFT(T0.""CardCode"",2)='MV' AND T1.""U_AI_EMAILSOA""='YES')) AND T0.""Canceled"" <>'Y' and T0.""CardCode"" like '" & sBPCode & "'"
        ElseIf sDocF = "%" And dDateT <> "1/1/1" Then
            sSQL += " WHERE ifnull(""TrsfrAcct"",'') <>'' AND ""DocDate"" between '" & dDateF.ToString("yyyy-MM-dd") & "' and '" & dDateT.ToString("yyyy-MM-dd") & "'" & _
             " AND ((LEFT(T0.""CardCode"",2)<>'MV') OR  (LEFT(T0.""CardCode"",2)='MV' AND T1.""U_AI_EMAILSOA""='YES')) AND T0.""Canceled"" <>'Y' and T0.""CardCode"" like '" & sBPCode & "' "
        Else
            sSQL += " WHERE ifnull(""TrsfrAcct"",'') <>'' AND ""DocNum"">= " & sDocF & " and ""DocNum""<= " & sDocT & " and ""DocDate"" between '" & dDateF.ToString("yyyy-MM-dd") & "' and '" & dDateT.ToString("yyyy-MM-dd") & "'" & _
            " AND ((LEFT(T0.""CardCode"",2)<>'MV') OR  (LEFT(T0.""CardCode"",2)='MV' AND T1.""U_AI_EMAILSOA""='YES'))  AND T0.""Canceled"" <>'Y' and T0.""CardCode"" like '" & sBPCode & "'"
        End If

        sSQL += "AND T2.""PaymWizCod"" LIKE '" & sPayCode & "'"
        sSQL += "GROUP BY T1.""E_Mail"",T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""CardName"", T0.""DocDate"", T1.""U_AI_EMAILSOA"" ORDER BY T0.""DocNum"""

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Payment Advise Portion :: ", sFuncName)

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
        oDS = ExecuteSQLQuery(sSQL)

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Load_PaymentAdvise()" & sSQL, sFuncName)
        If Load_PaymentAdvise(oDS, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
        Me.btnPDFGen.Enabled = True

        'System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        'WriteToStatusScreen(True, "Please wait PDF Genearation in progress ...")
        'Me.Button1.Enabled = Not Me.Button1.Enabled
        'Call btnPDFGen_Click(Me, New System.EventArgs)
        'Me.Button1.Enabled = Not Me.Button1.Enabled
        'System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Bt_CFL1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bt_CFL1.Click
        Dim sErrdesc As String = String.Empty
        If Not Me.SCompany.Text = "--Select Company-- " Then
            ''If Customer_CFL("3", Me.SCompany.Text, sErrdesc) <> RTN_SUCCESS Then
            ''    Throw New ArgumentException(sErrdesc)
            ''End If
            ''txtBpCode.Text = sCFL
            Dim ssql As String = "SELECT T0.""CardCode"", T0.""CardName"" FROM " & """" & sCompanyName.ToString().Trim() & """" & ".OCRD T0 WHERE T0.""CardType"" = 'S' and T0.""CardCode"" <> '' ORDER BY T0.""CardName"""
            Dim oCFL = New CFL(ssql, "1")
            oCFL.ShowDialog()
            txtBpCode.Text = sCFL
        End If
    End Sub

    ''Private Sub Bt_CFL2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ''    Dim sErrdesc As String = String.Empty
    ''    If Not Me.SCompany.Text = "--Select Company-- " Then
    ''        If Customer_CFL("1", Me.SCompany.Text, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
    ''    End If

    ''End Sub

    ''Private Sub Bt_CFL3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ''    Dim sErrdesc As String = String.Empty
    ''    If Not Me.SCompany.Text = "--Select Company-- " Then
    ''        If Customer_CFL("2", Me.SCompany.Text, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
    ''    End If
    ''End Sub

    Private Sub SOAEmailLog(ByVal sCardCode As String, _
                            ByVal sCardName As String, _
                            ByVal dStmtDate As Date, _
                            ByVal sEmail As String)
        Dim sSQL As String = String.Empty
        Dim iCode As Integer
        Dim sErrdesc As String = String.Empty
        Dim sFuncName As String = String.Empty
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling function..", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Geting Max code..", sFuncName)
            GetMaxCode(iCode)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Max code ::" & iCode, sFuncName)

            sSQL = "insert into ""@SOAEMAILLOG"" values(" & iCode & "," & iCode & "," & "'" & sCardCode & "','" & sCardName & "','" & dStmtDate.Date.ToString("yyyyMMdd") & "','" & Now.Date.ToString("yyyyMMdd") & "','" & sEmail & "')"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing sSQL", sFuncName)

            ExecuteSQLNonQuery(sSQL)

        Catch ex As Exception
            sErrdesc = ex.Message
            Call WriteToLogFile(sErrdesc, sFuncName)
            Throw New ArgumentException(sErrdesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        End Try
    End Sub

    Private Function Get_CompanyList(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   Get_CompanyList()
        '   Purpose     :   This function will provide the company list 
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   October 2014
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sQuery As String = String.Empty
        Dim iCount As Integer = 0
        Dim oDTCompanyList As DataTable = Nothing

        Try

            sFuncName = "Get_CompanyList()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            '' ''oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            '' ''oCompany.UseTrusted = True
            '' ''oCompany.DbUserName = p_oCompDef.sDBUser
            '' ''oCompany.DbPassword = p_oCompDef.sDBPwd
            '' ''oCompany.Server = p_oCompDef.sServer
            '' ''oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            '' ''oRecordSet = oCompany.GetCompanyList

            ' '' ''Dim s As String = oRecordSet.Fields.Item(0).Name
            ' '' ''Dim s1 As String = oRecordSet.Fields.Item(1).Name

            ' '' ''PaymentAdvice_F = New frmPaymentadvice
            '' ''Me.SCompany.Items.Clear()
            '' ''Me.SCompany.Items.Add("--Select Company-- ")

            '' ''oCompList.Clear()
            '' ''Do Until oRecordSet.EoF = True
            '' ''    Me.SCompany.Items.Add(oRecordSet.Fields.Item(0).Value)
            '' ''    oCompList.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
            '' ''    oRecordSet.MoveNext()
            '' ''Loop
            Me.SCompany.Items.Clear()
            Me.SCompany.Items.Add("--Select Company-- ")
            oCompListSOA.Clear()

            sQuery = "SELECT T0.""Name"", T0.""U_DBNAME"" FROM ""@AI_TB01_COMPANYDATA""  T0"
            oDTCompanyList = ExecuteSQLQuery_DataTable(sQuery, p_oCompDef.sSAPDBName, sErrDesc)

            Do Until iCount = oDTCompanyList.Rows.Count
                Me.SCompany.Items.Add(oDTCompanyList.Rows(iCount).Item(0).ToString().Trim())
                oCompListSOA.Add(oDTCompanyList.Rows(iCount).Item(0).ToString().Trim(), oDTCompanyList.Rows(iCount).Item(1).ToString().Trim())
                iCount += 1
            Loop

            Me.SCompany.SelectedIndex = 0

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Get_CompanyList = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Get_CompanyList = RTN_ERROR
        End Try
    End Function

    Private Sub SCompany_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles SCompany.SelectedIndexChanged

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SCompany_SelectedIndexChanged()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            sCompanyName = SCompany.SelectedItem.ToString.Trim()
            If UCase(sCompanyName.ToString().Trim()) = UCase("--Select Company--") Then Exit Sub

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Load_CompanyList() ", sFuncName)
            If Load_CompanyList(sCompanyName, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

    Private Function Load_CompanyList(ByVal sCompanyDB As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oDTEmailList As DataTable = New DataTable
        Dim sQuery As String = String.Empty


        Try
            sFuncName = "Load_CompanyList()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Me.cmbSenderEmail.Text = String.Empty
            Me.cmbSenderEmail.Items.Clear()

            sQuery = "SELECT T0.""U_AI_EMAIL"" FROM ""@AE_EMAILSENDERLIST"" T0"

            oDTEmailList = ExecuteSQLQuery_DataTable(sQuery, sCompanyName.ToString().Trim(), "")

            For Each drRow As DataRow In oDTEmailList.Rows
                Me.cmbSenderEmail.Items.Add(drRow.Item(0).ToString.Trim())
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Load_CompanyList = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message().ToString()
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Load_CompanyList = RTN_ERROR

        End Try
    End Function

    Private Sub btnPaymentCFL_Click(sender As System.Object, e As System.EventArgs)
        Dim sErrdesc As String = String.Empty
        If Not Me.SCompany.Text = "--Select Company-- " Then
            ''If Customer_CFL("3", Me.SCompany.Text, sErrdesc) <> RTN_SUCCESS Then
            ''    Throw New ArgumentException(sErrdesc)
            ''End If
            ''txtBpCode.Text = sCFL
            Dim ssql As String = "select T1.""WizardName"",T0.""PaymWizCod"" from OPEX T0  INNER JOIN OPWZ T1 ON T0.""PaymWizCod"" = T1.""IdNumber"" " & _
               "group by T1.""WizardName"",T0.""PaymWizCod"""
            Dim oCFL = New CFL(ssql, "4")
            oCFL.ShowDialog()
            txtPayment.Text = sCFL
            txtpaymentcode.Text = sCFL1
        End If
    End Sub

    Private Function Load_PaymentAdvise(ByVal oDS As DataSet, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   Load_PaymentAdvise()
        '   Purpose     :   This function will provide the BP Group list 
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   July 2015
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim ddate As Date

        Try

            sFuncName = "Load_PaymentAdvise()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Me.gvtaxinvoice.Rows.Clear()
            For imjs As Integer = 0 To oDS.Tables(0).Rows.Count - 1
                ddate = oDS.Tables(0).Rows(imjs)("DocDate").ToString
                Me.gvtaxinvoice.Rows.Add(1)
                '' MsgBox(oDsBPList.Tables(0).Rows(imjs)("CardCode").ToString)
                Me.gvtaxinvoice.Rows.Item(imjs).Cells.Item("clmnInvoiceno").Value = oDS.Tables(0).Rows(imjs)("DocNum").ToString
                Me.gvtaxinvoice.Rows.Item(imjs).Cells.Item("ClmnDate").Value = Format(ddate, "dd/MM/yyyy") '' Microsoft.VisualBasic.Left(oDS.Tables(0).Rows(imjs)("DocDate").ToString, 10)
                Me.gvtaxinvoice.Rows.Item(imjs).Cells.Item("clmnemail").Value = oDS.Tables(0).Rows(imjs)("E_Mail").ToString
                Me.gvtaxinvoice.Rows.Item(imjs).Cells.Item("clmnsend").Value = oDS.Tables(0).Rows(imjs)("U_AI_EMAILSOA").ToString
                Me.gvtaxinvoice.Rows.Item(imjs).Cells.Item("clmdocentry").Value = oDS.Tables(0).Rows(imjs)("DocEntry").ToString
            Next
            Me.gvtaxinvoice.Rows(0).Cells(2).Selected = True

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Load_PaymentAdvise = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Load_PaymentAdvise = RTN_ERROR
        End Try
    End Function

    Private Sub gvtaxinvoice_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gvtaxinvoice.CellContentClick
        If gvtaxinvoice.Columns.Item(e.ColumnIndex).Name = "Choose" And e.RowIndex = -1 Then
            Dim bflag As Boolean = False
            If Convert.ToBoolean(gvtaxinvoice.Rows(0).Cells(0).Value) = True Then
                bflag = False
            Else
                bflag = True
            End If

            Me.gvtaxinvoice.Rows(0).Cells(2).Selected = True
            For imjs As Integer = 0 To Me.gvtaxinvoice.Rows.Count - 2 'oDsBPList.Tables(0).Rows.Count - 1
                Me.gvtaxinvoice.Rows.Item(imjs).Cells.Item("Choose").Value = bflag  'oDsBPList.Tables(0).Rows(imjs)("Check").ToString
            Next
            Me.gvtaxinvoice.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            txtBpCode.Focus()


        End If
    End Sub

    Private Sub DocNo_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles DocNo.Paint

    End Sub
End Class
