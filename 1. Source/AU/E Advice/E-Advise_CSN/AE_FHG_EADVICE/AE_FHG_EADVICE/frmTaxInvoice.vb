
Imports System.IO
Imports Microsoft.VisualBasic

Public Class frmTaxInvoice
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If Me.DocType.Text = String.Empty Then Exit Sub

        If Me.txtBpCode.Text = String.Empty Then
            MsgBox("Please Enter Batch No", MsgBoxStyle.Information, "MBMS PDF Generation")
            Exit Sub
        End If

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        WriteToStatusScreen(True, "Please wait PDF Genearation in progress ...")
        ' Me.Button1.Enabled = Not Me.Button1.Enabled
        Call btnPDFGen_Click(Me, New System.EventArgs)
        ' Me.Button1.Enabled = Not Me.Button1.Enabled
        System.Windows.Forms.Cursor.Current = Cursors.Default
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
        Dim sBpcodeTo As String = String.Empty
        Dim sBPcode_Array(100, 2) As String
        Dim iCount_Array As Integer = 0
        Dim sBody As String = String.Empty
        Dim sSubject As String = String.Empty
        Dim sEPaymentEmail, sESOAEmail As String

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

            If Not DocType.Text = "SOA" Then
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
            End If


            WriteToStatusScreen(False, "Attempting to Tax Invoice Portion :: ")


            '            sSQL = "SELECT T1.""E_Mail"",T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""CardName"", T0.""DocDate"" " & _
            '                  " FROM " & """" & Me.SCompany.Text & """" & ".OINV T0  LEFT JOIN " & """" & Me.SCompany.Text & """" & ".OCRD T1" & _
            '                  "LEFT JOIN ""OITM"" T on X1.""ItemCode"" = T.""ItemCode"" LEFT JOIN ""OCRD"" A on A.""CardCode"" = X.""CardCode"" " & _
            '                " LEFT JOIN ""OCRY"" B on B.""Code"" = A.""Country"" LEFT JOIN ""OCPR"" C on C.""CntctCode"" = X.""CntctCode""" & _
            '               " left join ""OCTG"" K on K.""GroupNum"" = X.""GroupNum"",""ADM1"" Z, ""OADM"" J ,""OPDT"" O"
            'where X."DocEntry" = :DocKey and O."TextCode"='Bank_DBS'"


            '             " ON T0.""CardCode""= T1.""CardCode"" WHERE T0.""CANCELED"" <>'Y' "


            sSQL = "SELECT T1.""E_Mail"",T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""CardName"", T0.""DocDate"" " & _
                  " FROM " & """" & Me.SCompany.Text & """" & ".OINV T0  LEFT JOIN " & """" & Me.SCompany.Text & """" & ".OCRD T1" & _
                  " ON T0.""CardCode""= T1.""CardCode"" WHERE T0.""CANCELED"" <>'Y' "

            If sDocF <> "%" Then
                sSQL += " AND T0.""DocNum"">='" & sDocF & "'"
            End If

            If sDocT <> "%" Then
                sSQL += " AND T0.""DocNum""<='" & dDateT & "'"
            End If

            If dDateF <> "1/1/1" Then
                sSQL += " AND T0.""DocDate"">='" & dDateF.ToString("yyyy-MM-dd") & "'"
            End If

            If dDateT <> "1/1/1" Then
                sSQL += " AND T0.""DocDate""<='" & dDateT.ToString("yyyy-MM-dd") & "'"
            End If

            If sBPCode <> "%" Then
                sSQL += " AND T0.""CardCode""<='" & sBPCode & "'"
            End If

            sSQL += "GROUP BY T1.""E_Mail"",T0.""DocEntry"",T0.""CardCode"",T0.""DocNum"",T0.""CardName"", T0.""DocDate"""

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Payment Advise Portion :: ", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
            oDS = ExecuteSQLQuery(sSQL)
            iCount_Array = 0

            ReDim sBPcode_Array(oDS.Tables(0).Rows.Count, 2)
            '-----------------
            If oDS.Tables(0).Rows.Count > 0 Then

                Select Case DocType.Text.ToUpper()
                    Case "INJURY(ITEMTYPE)"
                        sRptFileName = p_oCompDef.sReportsPath & "\Tax Invoice_Injury_ItemType.rpt"
                    Case "PREEMPLOYMENT(ITEMTYPE)"
                        sRptFileName = p_oCompDef.sReportsPath & "\Tax Invoice_PreEmployment_ItemType.rpt"
                    Case "PREEMPLOYMENT(SERVICETYPE)"
                        sRptFileName = p_oCompDef.sReportsPath & "\Tax Invoice_PreEmployment_ServiceType.rpt"
                End Select

                For Each row As DataRow In oDS.Tables(0).Rows
                    If row.Item("E_Mail").ToString <> String.Empty Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                        WriteToStatusScreen(False, "Generating PDF for DocNo::" & row.Item("DocNum").ToString)
                        Dim dDateCon As Date = row.Item("DocDate").ToString

                        sTargetFileName = "TAXINV_" & row.Item("DocNum").ToString & "_" & row.Item("CardCode").ToString & "_" & dDateCon.ToString("ddMMyyyy") & ".pdf"
                        sTargetFileName = p_oCompDef.sReportPDFPath & "\" & sTargetFileName

                        If ExportToPDF_New(row.Item("DocEntry"), SCompany.Text.ToString().Trim(), sTargetFileName, sRptFileName, sErrdesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrdesc)
                        End If
                        ' WriteToStatusScreen(False, "Successfully  generated PDF ::" & sTargetFileName)
                        WriteToStatusScreen(False, "Successfully  generated the PDF")
                        WriteToStatusScreen(False, "Sending Email for Payment No ::" & row.Item("DocNum").ToString)

                        sBody = "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                        sBody = sBody & " Dear Sir/Madam,<br /><br /> Please find Tax invoice for your record.<br /><br />"
                        sBody = sBody & " This payment will be credited to your account on " & Me.DateTimePicker2.Value.Date.ToString("dd/MM/yyyy") & ".<br /><br/><br/>"
                        sBody = sBody & " Please do not reply to this email. For clarification, please contact " & sEPaymentEmail & "<br /><br />"

                        Dim sCompanyName As String = String.Empty
                        oCompList.TryGetValue(Me.SCompany.Text, sCompanyName)

                        sSubject = "Tax invoice from " & sCompanyName
                        Dim sFromEmail As String = String.Empty
                        sFromEmail = cmbSenderEmail.Text.Trim()
                        ''row.Item("E_Mail").ToString
                        If SCompany.Text.ToString().ToUpper().Trim().Contains("CSN") Then
                            ''call the send email funtion for CSN company
                            '' If SendEmailNotificationForCSN(sTargetFileName, sFromEmail, row.Item("E_Mail").ToString, sBody, sSubject, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                            If SendEmailNotification(sTargetFileName, sFromEmail, row.Item("E_Mail").ToString, sBody, sSubject, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                            '' If SendEmailNotificationForCSN(sTargetFileName, sFromEmail, "vivekrm60@gmail.com;srinivab@abeo-electra.com;chongxiang.weng@fullertonhealthcare.com", sBody, sSubject, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                        Else
                            If SendEmailNotification(sTargetFileName, sFromEmail, row.Item("E_Mail").ToString, sBody, sSubject, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                        End If
                        WriteToStatusScreen(False, "Successfully sent email Payment No ::" & row.Item("DocNum").ToString)
                        WriteToStatusScreen(False, "===========================================================")
                    Else
                        iCount_Array += 1
                        sBPcode_Array(iCount_Array, 0) = row.Item("CardCode").ToString
                        sBPcode_Array(iCount_Array, 1) = row.Item("CardName").ToString
                        sBPcode_Array(iCount_Array, 2) = row.Item("DocNum").ToString
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email address is blank for the Customer  - Invoice No." & row.Item("DocNum").ToString, sFuncName)
                        WriteToStatusScreen(False, "Email address is blank for the Customer in this Invoice No. :: " & row.Item("DocNum").ToString)
                    End If
                Next
            Else
                WriteToStatusScreen(False, "No records Found for the above selection, Please try again with different Selection.")
            End If

            If iCount_Array > 0 Then
                Write_TextFile_BPList(sBPcode_Array, "0")
            End If

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
        cmbSenderEmail.Text = String.Empty
        cmbSenderEmail.Items.Clear()
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


        If Me.SCompany.Text = "--Select Company-- " Then Exit Sub
        If Me.DocType.Text = "---Select---" Then Exit Sub
        If Me.cmbSenderEmail.Text = String.Empty Then Exit Sub

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

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        WriteToStatusScreen(True, "Please wait PDF Genearation in progress ...")
        Me.Button1.Enabled = Not Me.Button1.Enabled
        Call btnPDFGen_Click(Me, New System.EventArgs)
        Me.Button1.Enabled = Not Me.Button1.Enabled
        System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Bt_CFL1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bt_CFL1.Click
        Dim sErrdesc As String = String.Empty
        If Not Me.SCompany.Text = "--Select Company-- " Then
            ''If Customer_CFL("3", Me.SCompany.Text, sErrdesc) <> RTN_SUCCESS Then
            ''    Throw New ArgumentException(sErrdesc)
            ''End If
            ''txtBpCode.Text = sCFL
            Dim ssql As String = "SELECT T0.""CardCode"", T0.""CardName"" FROM " & """" & sCompanyName & """" & ".OCRD T0 WHERE T0.""CardType"" = 'C' and T0.""CardCode"" <> '' ORDER BY T0.""CardName"""
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

        Try

            sFuncName = "Get_CompanyList()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            oCompany.UseTrusted = True
            oCompany.DbUserName = p_oCompDef.sDBUser
            oCompany.DbPassword = p_oCompDef.sDBPwd
            oCompany.Server = p_oCompDef.sServer
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            oRecordSet = oCompany.GetCompanyList

            'Dim s As String = oRecordSet.Fields.Item(0).Name
            'Dim s1 As String = oRecordSet.Fields.Item(1).Name

            'PaymentAdvice_F = New frmPaymentadvice
            Me.SCompany.Items.Clear()
            Me.SCompany.Items.Add("--Select Company-- ")

            oCompList.Clear()
            Do Until oRecordSet.EoF = True
                Me.SCompany.Items.Add(oRecordSet.Fields.Item(0).Value)
                oCompList.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
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

            sQuery = "SELECT T0.""Name"" FROM ""@AE_EMAILSENDERLIST""  T0"

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

End Class
