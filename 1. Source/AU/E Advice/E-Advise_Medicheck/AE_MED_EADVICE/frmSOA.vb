
Imports System.IO
Imports Microsoft.VisualBasic

Public Class frmSOA

    Dim sErrdesc As String = String.Empty

    Public Sub WriteToStatusScreen(ByVal Clear As Boolean, ByVal msg As String)
        ''If Clear Then
        ''    txtStatusMsg.Text = ""
        ''End If
        ''txtStatusMsg.HideSelection = True
        ''txtStatusMsg.Text &= msg & vbCrLf
        ''txtStatusMsg.SelectAll()
        ''txtStatusMsg.ScrollToCaret()
        ''txtStatusMsg.Refresh()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If Me.DocType.Text = String.Empty Then Exit Sub
        If Me.cmbSenderEmail.Text = String.Empty Then Exit Sub


        ''If Me.txtBpCode.Text = String.Empty Then
        ''    MsgBox("Please Enter Batch No", MsgBoxStyle.Information, "MBMS PDF Generation")
        ''    Exit Sub
        ''End If

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

            ' WriteToStatusScreen(False, "=============================== Start  ===========================")

            If Not System.IO.Directory.Exists(p_oCompDef.sReportsPath) Then
                System.IO.Directory.CreateDirectory(p_oCompDef.sReportsPath)
                '  WriteToStatusScreen(False, "Created Folder for Report :: " & p_oCompDef.sReportsPath)
            End If


            'Exit Function

            'Get E-Advice contanct email address
            'sSQL = "select ""AliasName"" as ePaymentEmail,""E_Mail"" as eSOAEmail from " & """" & Me.SCompany.Text & """" & ".OADM T0"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
            'oDS = ExecuteSQLQuery(sSQL)
            'If oDS.Tables(0).Rows.Count > 0 Then
            '    sEPaymentEmail = oDS.Tables(0).Rows(0).Item(0).ToString
            '    sESOAEmail = oDS.Tables(0).Rows(0).Item(1).ToString ' '
            'End If

            If Me.cmbSenderEmail.Text.Trim <> String.Empty Then
                sEPaymentEmail = Me.cmbSenderEmail.Text.Trim
                sESOAEmail = Me.cmbSenderEmail.Text.Trim
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

                Case "SOA"

                    Dim sEmail As String = String.Empty
                    Dim sCardcode As String = String.Empty
                    Dim sCardName As String = String.Empty

                    For imjs As Integer = 0 To Dgv_BPList.Rows.Count - 1

                        If Convert.ToBoolean(Dgv_BPList.Rows(imjs).Cells(0).Value) = True Then

                            '' sEmail = "srinivasanm@abeo-electra.com"

                            sEmail = Convert.ToString(Dgv_BPList.Rows(imjs).Cells(3).Value)
                            sCardcode = Convert.ToString(Dgv_BPList.Rows(imjs).Cells(1).Value)
                            sCardName = Convert.ToString(Dgv_BPList.Rows(imjs).Cells(2).Value)
                            lblMsg.Text = "Processing for CardName " & sCardName
                            Me.Refresh()
                            If sEmail <> String.Empty Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF()", sFuncName)
                                Dim dDateCon As Date = DateTimePicker1.Text

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking BP Balance ..", sFuncName)

                                sSQL = "CALL " & """" & Me.SCompany.Text & """" & ".""AE_SP002_SOA"" ('" & sCardcode & "','" & sCardcode & "','" & dDateCon.ToString("yyyyMMdd") & "')"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
                                oDS1 = ExecuteSQLQuery(sSQL)

                                If oDS1.Tables(0).Rows.Count > 0 Then

                                    sTargetFileName = "SOA_" & sCardcode & "_" & dDateCon.ToString("ddMMyyyy") & ".pdf"
                                    sTargetFileName = p_oCompDef.sReportPDFPath & "\" & sTargetFileName
                                    sRptFileName = p_oCompDef.sReportsPath & "\" & sSOARptName

                                    If ExportToPDF_SOA(sCardcode, sCardcode, dDateCon, sTargetFileName, sRptFileName, SCompany.Text().ToString().Trim(), sErrdesc) <> RTN_SUCCESS Then
                                        Throw New ArgumentException(sErrdesc)
                                    End If
                                    WriteToStatusScreen(False, "Successfully  generated PDF ::" & sTargetFileName)

                                    ''sBody = "<div align=left style='font-size:11.0pt;font-family:Calibri'>"
                                    ''sBody = sBody & "<br />Dear Customer, <br/>Please contact us immediately if you are unable to detach or download your statement. <br /><br /> "
                                    ''sBody = sBody & " Many thanks & kind regards, <br />AR Accounts <br/> "
                                    ''sBody = sBody & " Corporate Services Network <br/> "
                                    ''sBody = sBody & " Level 2,280 George St, Sydney NSW 2000 <br/> <br/>"
                                    ''sBody = sBody & " <B>T</B> +61 2 8256 1704 | <B>F</B> +61 2 82561785 | <B>W</B> fullertonhealthcare.com.au. <br/> <br/><br/>"

                                    sBody = "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                    sBody = sBody & " Dear Sir/Madam,<br /><br /> Account Code: " & oDS1.Tables(0).Rows(0).Item("CardCode").ToString & "<br />"
                                    sBody = sBody & "Statement of Account as on " & Me.DateTimePicker1.Value.ToString("dd-MMM-yyyy") & ".<br /><br/>"
                                    sBody = sBody & " Please find attached your invoice/statement.  If you have any queries please contact us on (07) 4046 8600 or email accounts@cairnshealth.com.au <br /><br />"
                                    sBody = sBody & "<br/><br/>"
                                    sBody = sBody & "<br/><br/>"
                                    sBody = sBody & "<br/><br/>"
                                    sBody = sBody & "Gayle Baxter |Accounts| Central Plaza Doctors/Smithfield Central Doctors | 60 McLeod St Cairns QLD 4870 | <br/>"
                                    sBody = sBody & "Ph: (07) 4046 8600| Fax: (07) 4046 8699| accounts@cairnshealth.com.au <br/>"
                                    sBody = sBody & "www.cairnshealth.com.au"


                                    Dim sCompanyName As String = String.Empty
                                    oCompListSOA.TryGetValue(Me.SCompany.Text, sCompanyName)

                                    ''  sSubject = "Important - Statement of Account as at " & Me.DateTimePicker1.Value.ToString("dd MM yyyy") & " from " & sCompanyName
                                    ''  sSubject = "Invoices / Statements (Statement of Accounts)"
                                    sSubject = "Statement of Accounts from Medicheck (" & Me.SCompany.Text & ")"

                                    WriteToStatusScreen(False, "Sending Email for BP ::" & sCardName)
                                    lblMsg.Text = sCardName & " - Sending Email ::"
                                    Dim sFromEmail As String = String.Empty
                                    sFromEmail = cmbSenderEmail.Text.Trim()
                                    Me.Refresh()
                                    '' If SendEmailNotification(sTargetFileName, sFromEmail, "sahayar@abeo-electra.com", sBody, sSubject, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                                    If SendEmailNotification(sTargetFileName, sFromEmail, sEmail, sBody, sSubject, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                                    lblMsg.Text = sCardName & " - Sending Email Successfully ::"
                                    Me.Refresh()
                                    SOAEmailLog(sCardcode, sCardName, Me.DateTimePicker1.Value, sEmail)
                                    Dgv_BPList.Rows(imjs).Cells(5).Value = "Successfully sent email BP ::" & sCardName

                                Else
                                    Dgv_BPList.Rows(imjs).Cells(5).Value = "No outstanding Balance for BP::" & sCardName
                                    Dgv_BPList.Rows(imjs).Cells(5).Style.ForeColor = Color.DarkRed
                                End If

                            Else
                                iCount_Array += 1
                                sBPcode_Array(iCount_Array, 0) = sCardcode
                                sBPcode_Array(iCount_Array, 1) = sCardName

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email address is blank for the Customer " & sCardName, sFuncName)
                                Dgv_BPList.Rows(imjs).Cells(5).Value = "Email address is blank for the Customer :: " & sCardName
                                Dgv_BPList.Rows(imjs).Cells(5).Style.ForeColor = Color.DarkRed
                            End If
                        End If

                    Next imjs
            End Select
            lblMsg.Text = "Completed Successfully"
            Me.Refresh()
            lblMsg1.Text = ""
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

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub frmPDFGen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        oCompany = New SAPbobsCOM.Company
        '------------  Selecting the Document Type and set Select as default
        DocType.SelectedIndex = 0
        '-------------------------------------------------------------------

        If GetSystemIntializeInfo(p_oCompDef, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
        If Get_CompanyList(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
        '' If Get_BPGroupList(sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim bCheck As Boolean = False
        If Me.SCompany.Text = "--Select Company-- " Then Exit Sub
        If Me.DocType.Text = "---Select---" Then Exit Sub
        If Me.CheckBox1.Checked = True Then
            If (Me.DateTimePicker1.Checked = False) Then
                MsgBox("Enter the DocDate ", MsgBoxStyle.Information, "E-Advice")
                Exit Sub
            End If
        Else
            If (Me.txtGroup.Text = String.Empty) Or (Me.DateTimePicker1.Checked = False) Then
                MsgBox("Enter the BP Group and DocDate ", MsgBoxStyle.Information, "E-Advice")
                Exit Sub
            End If
        End If

        For imjs As Integer = 0 To Dgv_BPList.Rows.Count - 1
            If Convert.ToBoolean(Dgv_BPList.Rows(imjs).Cells(0).Value) = True Then
                bCheck = True
            End If
        Next imjs

        If bCheck = False Then
            MsgBox("Choose the Customer ....")
        End If

        lblMsg1.Text = "Please Wait ... "
        Me.Refresh()

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        WriteToStatusScreen(True, "Please wait PDF Genearation in progress ...")
        Me.Button1.Enabled = Not Me.Button1.Enabled
        Call btnPDFGen_Click(Me, New System.EventArgs)
        Me.Button1.Enabled = Not Me.Button1.Enabled
        System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    ''Private Sub Bt_CFL1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ''    Dim sErrdesc As String = String.Empty
    ''    If Not Me.SCompany.Text = "--Select Company-- " Then
    ''        If Customer_CFL("3", Me.SCompany.Text, sErrdesc) <> RTN_SUCCESS Then
    ''            Throw New ArgumentException(sErrdesc)
    ''        End If
    ''    End If
    ''End Sub

    ''Private Sub Bt_CFL2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ''    Dim sErrdesc As String = String.Empty
    ''    If Not Me.SCompany.Text = "--Select Company-- " Then
    ''        If Customer_CFL("1", Me.SCompany.Text, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
    ''    End If

    ''End Sub

    '' ''Private Sub Bt_CFL3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '' ''    Dim sErrdesc As String = String.Empty
    '' ''    If Not Me.SCompany.Text = "--Select Company-- " Then
    '' ''        If Customer_CFL("2", Me.SCompany.Text, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
    '' ''    End If
    '' ''End Sub

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

            sSQL = "insert into """ & sCompanyName & """ .""@SOAEMAILLOG"" values(" & iCode & "," & iCode & "," & "'" & sCardCode & "','" & Replace(sCardName, "'", "''") & "','" & dStmtDate.Date.ToString("yyyyMMdd") & "','" & Now.Date.ToString("yyyyMMdd") & "','" & sEmail & "')"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing sSQL " & sSQL, sFuncName)

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
        Dim oDTCompanyList As DataTable = Nothing
        Dim sQuery As String = String.Empty
        Dim iCount As Integer = 0
        Try

            sFuncName = "Get_CompanyList()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

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

    Private Function Get_CompanyList_Backup(ByRef sErrDesc As String) As Long

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
            oCompListSOA.Clear()

            Do Until oRecordSet.EoF = True
                Me.SCompany.Items.Add(oRecordSet.Fields.Item(0).Value)
                oCompListSOA.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Loop

            Me.SCompany.SelectedIndex = 0

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Get_CompanyList_Backup = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Get_CompanyList_Backup = RTN_ERROR
        End Try
    End Function

    Private Function Get_BPGroupList(ByRef sErrDesc As String, ByVal sCompanyName As String) As Long

        ' **********************************************************************************
        '   Function    :   Get_BPGroupList()
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
        Dim sConnection As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDsBPList As DataSet = Nothing
        Try

            sFuncName = "Get_BPGroupList()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            ''sSQL = "SELECT  ""---- Select ----"" ""GroupCode"", ""---- Select ----"" ""GroupName"" "
            ''sSQL += "Union all"
            sSQL = "SELECT T0.""GroupCode"",T0.""GroupName"" FROM " & """" & sCompanyName & """" & " .OCRG T0 WHERE T0.""GroupType"" = 'C'"
            ''sSQL += "Union all"
            ''sSQL += "SELECT  ""ALL"" ""GroupCode"", ""ALL"" ""GroupName"" "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
            oDsBPList = ExecuteSQLQuery(sSQL)

            Me.CboxBPlist.Items.Clear()
            Me.CboxBPlist.DataSource = oDsBPList.Tables(0)
            Me.CboxBPlist.ValueMember = "GroupCode"
            Me.CboxBPlist.DisplayMember = "GroupName"
            ''Me.CboxBPlist.Items.Add("--- BP Group ---")
            ''Me.CboxBPlist.Items.Add("ALL")

            ' ''Do Until oRecordSet.EoF = True
            ' ''    Me.SCompany.Items.Add(oRecordSet.Fields.Item(0).Value)
            ' ''    oCompList.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
            ' ''    oRecordSet.MoveNext()
            ' ''Loop

            Me.CboxBPlist.SelectedIndex = Me.CboxBPlist.Items.Count - 1

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Get_BPGroupList = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Get_BPGroupList = RTN_ERROR
        End Try
    End Function

    ''Private Sub SCompany_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles SCompany.SelectedIndexChanged
    ''    Dim sErrdesc As String = String.Empty
    ''    Try
    ''        If Me.SCompany.Text <> "--Select Company-- " Then
    ''            If Get_BPGroupList(sErrdesc, Me.SCompany.Text) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
    ''        End If

    ''    Catch ex As Exception
    ''        Eadvice_F.ToolStripStatusLabel1.Text = sErrdesc
    ''    End Try

    ''End Sub

    Private Sub btnGroup_Click(sender As System.Object, e As System.EventArgs) Handles btnGroup.Click
        Dim sErrdesc As String = String.Empty
        If Not Me.SCompany.Text = "--Select Company-- " Then
            '' If CustomerGroup_CFL("1", Me.SCompany.Text, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
            Dim ssql As String = "SELECT T0.""GroupCode"",T0.""GroupName"" FROM " & """" & sCompanyName & """" & " .OCRG T0 WHERE T0.""GroupType"" = 'C'"
            Dim oCFL = New CFL(ssql, "1")
            oCFL.ShowDialog()
            txtGroup.Text = sCFL
            txtGroupNo.Text = sCFL1
        End If
    End Sub

    Private Sub btnShow_Click(sender As System.Object, e As System.EventArgs) Handles btnShow.Click

        Dim sBPGroup As String = String.Empty

        If Me.CheckBox1.Checked = True Then
            sBPGroup = "%"
        Else
            If (Me.txtGroup.Text = String.Empty) Then
                MsgBox("Enter the BP Group  ", MsgBoxStyle.Information, "E-Advice")
                Exit Sub
            Else
                sBPGroup = Me.txtGroup.Text
            End If
        End If

        Dim sErrDesc As String = String.Empty
        Load_BPGroupList(sErrDesc, Me.SCompany.Text, sBPGroup)
      
    End Sub


    Private Function Load_BPGroupList(ByRef sErrDesc As String, ByVal sCompanyName As String, ByVal sGroupCode As String) As Long

        ' **********************************************************************************
        '   Function    :   Load_BPGroupList()
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
        Dim sConnection As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDsBPList As DataSet = Nothing
        Try

            sFuncName = "Load_BPGroupList()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSQL = "SELECT T0.""CardCode"", T0.""CardName"",T0.""E_Mail"", T0.""U_AI_EMAILSOA"" from " & """" & Me.SCompany.Text & """" & ".OCRD T0 " & _
                          " WHERE T0.""frozenFor"" ='N' AND T0.""GroupCode"" like '" & sGroupCode & "' AND T0.""CardType"" = 'C' Order by  T0.""CardCode"""

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
            oDsBPList = ExecuteSQLQuery(sSQL)
            Me.Dgv_BPList.Rows.Clear()
            For imjs As Integer = 0 To oDsBPList.Tables(0).Rows.Count - 1
                Me.Dgv_BPList.Rows.Add(1)
                '' MsgBox(oDsBPList.Tables(0).Rows(imjs)("CardCode").ToString)
                Me.Dgv_BPList.Rows.Item(imjs).Cells.Item("CardCode").Value = oDsBPList.Tables(0).Rows(imjs)("CardCode").ToString
                Me.Dgv_BPList.Rows.Item(imjs).Cells.Item("CardName").Value = oDsBPList.Tables(0).Rows(imjs)("CardName").ToString
                Me.Dgv_BPList.Rows.Item(imjs).Cells.Item("Email").Value = oDsBPList.Tables(0).Rows(imjs)("E_Mail").ToString
                Me.Dgv_BPList.Rows.Item(imjs).Cells.Item("SendEmail").Value = oDsBPList.Tables(0).Rows(imjs)("U_AI_EMAILSOA").ToString

            Next
            Me.Dgv_BPList.Rows(0).Cells(2).Selected = True

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Load_BPGroupList = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Load_BPGroupList = RTN_ERROR
        End Try
    End Function

    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            Me.txtGroup.Clear()
            Me.txtGroup.Enabled = False
            Me.btnGroup.Enabled = False
        Else
            Me.txtGroup.Enabled = True
            Me.btnGroup.Enabled = True
        End If
    End Sub

    Private Sub Dgv_BPList_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Dgv_BPList.CellContentClick

        If Dgv_BPList.Columns.Item(e.ColumnIndex).Name = "Choose" And e.RowIndex = -1 Then
            Dim bflag As Boolean = False
            If Convert.ToBoolean(Dgv_BPList.Rows(0).Cells(0).Value) = True Then
                bflag = False
            Else
                bflag = True
            End If

            Me.Dgv_BPList.Rows(0).Cells(2).Selected = True
            For imjs As Integer = 0 To Me.Dgv_BPList.Rows.Count - 2 'oDsBPList.Tables(0).Rows.Count - 1
                Me.Dgv_BPList.Rows.Item(imjs).Cells.Item("Choose").Value = bflag  'oDsBPList.Tables(0).Rows(imjs)("Check").ToString
            Next
            Me.Dgv_BPList.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            CheckBox1.Focus()

            ''Me.Dgv_BPList.Rows(1).Cells(0).Selected = True
            'Me.Dgv_BPList.Rows(1).Cells(0). = True
            '' Me.Dgv_BPList.Rows.Add(1)
            '' MsgBox(oDsBPList.Tables(0).Rows(imjs)("CardCode").ToString)
            ''Dim sSQL As String = String.Empty
            ''Dim oDsBPList As DataSet = Nothing
            ''Dim sGroupCode As String = String.Empty

            ''If Me.CheckBox1.Checked = True Then
            ''    sGroupCode = "%"
            ''Else
            ''    sGroupCode = Me.txtGroup.Text
            ''End If

            ''If Convert.ToBoolean(Dgv_BPList.Rows(0).Cells(0).Value) = True Then
            ''    sSQL = "SELECT 'False' ""Check"", T0.""CardCode"", T0.""CardName"",T0.""E_Mail"", T0.""U_AI_EMAILSOA"" from " & """" & Me.SCompany.Text & """" & ".OCRD T0 " & _
            ''              " WHERE T0.""frozenFor"" ='N' AND T0.""GroupCode"" like '" & sGroupCode & "'"
            ''Else
            ''    sSQL = "SELECT 'True' ""Check"",T0.""CardCode"", T0.""CardName"",T0.""E_Mail"", T0.""U_AI_EMAILSOA"" from " & """" & Me.SCompany.Text & """" & ".OCRD T0 " & _
            ''              " WHERE T0.""frozenFor"" ='N' AND T0.""GroupCode"" like '" & sGroupCode & "'"
            ''End If

            ''Me.Dgv_BPList.Rows.Item(imjs).Cells.Item("CardCode").Value = oDsBPList.Tables(0).Rows(imjs)("CardCode").ToString
            ''Me.Dgv_BPList.Rows.Item(imjs).Cells.Item("CardName").Value = oDsBPList.Tables(0).Rows(imjs)("CardName").ToString
            ''Me.Dgv_BPList.Rows.Item(imjs).Cells.Item("Email").Value = oDsBPList.Tables(0).Rows(imjs)("E_Mail").ToString
            ''Me.Dgv_BPList.Rows.Item(imjs).Cells.Item("SendEmail").Value = oDsBPList.Tables(0).Rows(imjs)("U_AI_EMAILSOA").ToString

            ''oDsBPList = ExecuteSQLQuery(sSQL)
            '' Me.Dgv_BPList.Rows.Clear()


        End If
    End Sub


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

            cmbSenderEmail.SelectedIndex = 0


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
