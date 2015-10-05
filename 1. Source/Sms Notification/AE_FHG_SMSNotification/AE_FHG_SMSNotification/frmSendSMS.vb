Option Explicit On
Imports System.Threading

Public Class frmSendSMS

    Dim sErrDesc As String = String.Empty
    Private dtSMSHeader As DataTable
    Dim dtBatch As New DataTable
    Dim iTotRowCount As Integer
    Dim iSuccessCount As Integer
    Dim iErrorCount As Integer
    Dim sPaymentType As String = String.Empty

    Private Sub btnSendSMS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendSMS.Click

        Dim sFuncName As String = String.Empty
        Dim sMessage As String = String.Empty


        Try
            sFuncName = "btnSendSMS_Click()"
           If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sPaymentType = String.Empty

            sPaymentType = cmbPaymentType.Text.Trim()

            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

            WriteToStatusScreen(True, "Please wait... ", sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Header_Validation()", sFuncName)
            If Header_Validation(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessSMS()", sFuncName)
            If ProcessSMS(txtBatchNo.Text.Trim(), sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            'WriteToStatusScreen(False, "Batch No. validation completed successfully.", sFuncName, sErrDesc)

            If dtSMSHeader.Rows.Count > 0 Then

                iSuccessCount = 0
                iErrorCount = 0

                ' WriteToStatusScreen(False, "Please wait... while validating phone numbers.", sFuncName, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Validate_PhoneNumbers()", sFuncName)
                If Validate_PhoneNumbers(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                ' WriteToStatusScreen(False, "Phone Number validation completed successfully.", sFuncName, sErrDesc)


                If iTotRowCount <> iSuccessCount Then
                    ' sMessage = String.Format("Records Successed " & iSuccessCount & " Out Of " & iTotRowCount & "")
                    sMessage = "Some of the documents doesn't have mobil no."
                    sMessage += Environment.NewLine & Environment.NewLine & Environment.NewLine
                    sMessage += "Do you want to continue.."
                    Dim result As Integer = MessageBox.Show(sMessage, sTitle, MessageBoxButtons.YesNoCancel)

                    If result = DialogResult.Yes Then
                        ' WriteToStatusScreen(False, "Please Wait...", sFuncName, sErrDesc)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendSMS()", sFuncName)
                        If SendSMS(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                    End If
                Else
                    'WriteToStatusScreen(False, "Please Wait...", sFuncName, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendSMS()", sFuncName)
                    If SendSMS(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If
            Else
                sErrDesc = "No records found in SAP"
                ''Throw New ArgumentException(sErrDesc)
                MessageBox.Show(sErrDesc, sTitle, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Default

            WriteToStatusScreen(False, "============== COMPLETED ==============", sFuncName, sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            System.Windows.Forms.Cursor.Current = Cursors.Default
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            MessageBox.Show(sErrDesc, sTitle, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        End
    End Sub

    Function SendSMS_BackUp(ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oWebClient As New System.Net.WebClient
        Dim sUserName As String = String.Empty
        Dim sPassword As String = String.Empty
        Dim sMobileNo As String = String.Empty
        Dim sFrom As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sDocNum As String = String.Empty
        Dim sAmount As String = String.Empty
        Dim sQuery As String = String.Empty

        Try

            sFuncName = "SendSMS()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sUserName = p_oCompDef.p_sSMSUserName
            sPassword = p_oCompDef.p_sSMSPassword
            sFrom = p_oCompDef.p_sSMSFrom

            For iRow As Integer = 0 To dtSMSHeader.Rows.Count - 1

                sDocNum = dtSMSHeader.Rows(iRow)(0).ToString().Trim()
                sMobileNo = dtSMSHeader.Rows(iRow)(1).ToString().Trim()
                sAmount = dtSMSHeader.Rows(iRow)(2).ToString().Trim()

                If sMobileNo.ToString() = String.Empty Then
                    Call WriteToLogFile("Check Document Number : " & sDocNum & ". Mobile Number is Blank!", sFuncName)
                    WriteToStatusScreen(False, "Check Document Number : " & sDocNum & ". Mobile Number is Blank!", sFuncName, sErrDesc)
                    Continue For

                End If

                Dim sUrl As String = "http://mx.fortdigital.net/http/send-message?username={0}&password={1}&to=%2B91{2}&from={3}&message={4}"

                sMessage = String.Format(p_oCompDef.p_sCheckSMS, sAmount)

                Dim QStr As String = String.Empty
                QStr = String.Format(sUrl, sUserName, sPassword, sMobileNo, sFrom, sMessage)
                oWebClient.Encoding = System.Text.Encoding.ASCII
                oWebClient.UseDefaultCredentials = False

                WriteToStatusScreen(False, "Sending SMS for Mobile No." & sMobileNo, sFuncName, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sending SMS for Mobile NO." & sMobileNo, sFuncName)

                Dim ReturnMsg As String = oWebClient.DownloadString(QStr)

                ''GETTING THE STATUS FOR SEND SMS
                sUrl = "http://mx.fortdigital.net/http/request-status-update?username={0}&password={1}&message-id={2}"
                QStr = String.Empty
                QStr = String.Format(sUrl, sUserName, sPassword, Mid(ReturnMsg, 5, ReturnMsg.Length - 8))
                oWebClient.Encoding = System.Text.Encoding.ASCII
                oWebClient.UseDefaultCredentials = False

                Dim sStatus As String = String.Empty

                sStatus = oWebClient.DownloadString(QStr)

                WriteToStatusScreen(False, "Message ID :" & Mid(ReturnMsg, 5, ReturnMsg.Length - 8), sFuncName, sErrDesc)

                If Mid(sStatus, 1, 7).ToString().ToUpper() = "SUCCESS" Then
                    WriteToStatusScreen(False, "Successfully Sent SMS to mobile No." & sMobileNo, sFuncName, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Sent SMS to mobile No." & sMobileNo, sFuncName)

                    WriteToStatusScreen(False, "Updating the Flag for Mobile No." & sMobileNo, sFuncName, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery for Update the Flag. Mobile NO." & sMobileNo, sFuncName)
                    sQuery = "UPDATE OVPM SET ""U_AI_SentSMS""='Y' WHERE ""DocNum""='" & sDocNum & "'"
                    ExecuteSQLNonQuery(sQuery)
                Else
                    WriteToStatusScreen(False, sStatus & ".Mobile No." & sMobileNo, sFuncName, sErrDesc)
                    WriteToStatusScreen(False, Mid(sStatus, 8, sStatus.Length - 1) & ".Mobile No." & sMobileNo, sFuncName, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Sent SMS to mobile No." & sMobileNo, sFuncName)
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SendSMS_BackUp = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            WriteToStatusScreen(False, "Completed with ERROR. Error : " & sErrDesc, sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SendSMS_BackUp = RTN_ERROR

        End Try

    End Function

    Function SendSMS(ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oWebClient As New System.Net.WebClient
        Dim sUserName As String = String.Empty
        Dim sPassword As String = String.Empty
        Dim sMobileNo As String = String.Empty
        Dim sFrom As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sDocNum As String = String.Empty
        Dim sAmount As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim oDTSMS As DataTable = New DataTable
        Dim sReturnMsg As String = String.Empty
        Dim sStatusUrl As String = String.Empty
        Dim sStatusQStr As String = String.Empty
        Dim sMessageID As String = String.Empty
        Dim k As Integer
        Dim sDecimalAmount As String = String.Empty

        Try

            sFuncName = "SendSMS()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sUserName = p_oCompDef.p_sSMSUserName
            sPassword = p_oCompDef.p_sSMSPassword
            sFrom = p_oCompDef.p_sSMSFrom

            oDTSMS = CreateDataTable("DocNum", "MobileNo", "MessgaeID")

            For iRow As Integer = 0 To dtSMSHeader.Rows.Count - 1

                sDocNum = dtSMSHeader.Rows(iRow)(0).ToString().Trim()
                sMobileNo = dtSMSHeader.Rows(iRow)(1).ToString().Trim()
                sAmount = dtSMSHeader.Rows(iRow)(2).ToString().Trim()

                k = InStrRev(sAmount, ".")
                If k = 0 Then
                    sAmount += ".00"
                Else
                    sDecimalAmount = Microsoft.VisualBasic.Right(sAmount, Len(sAmount) - k).Trim + "0"
                    sAmount = Microsoft.VisualBasic.Left(sAmount, k) + Microsoft.VisualBasic.Left(sDecimalAmount, 2)
                End If

                If sMobileNo.ToString() = String.Empty Then
                    Call WriteToLogFile("Check Document Number : " & sDocNum & ". Mobile Number is Blank!", sFuncName)
                    WriteToStatusScreen(False, "Check Document Number : " & sDocNum & ". Mobile Number is Blank!", sFuncName, sErrDesc)
                    Continue For

                End If

                Dim sUrl As String = "http://mx.fortdigital.net/http/send-message?username={0}&password={1}&to=%2B65{2}&from={3}&message={4}"

                If sPaymentType.ToString().ToUpper = "CHEQUE" Then
                    sMessage = String.Format(p_oCompDef.p_sCheckSMS, sAmount)
                ElseIf sPaymentType.ToString().ToUpper = "GIRO" Then
                    sMessage = String.Format(p_oCompDef.p_sGIROSMS, sAmount)
                End If

                Dim QStr As String = String.Empty
                QStr = String.Format(sUrl, sUserName, sPassword, sMobileNo, sFrom, sMessage)
                oWebClient.Encoding = System.Text.Encoding.ASCII
                oWebClient.UseDefaultCredentials = False

                WriteToStatusScreen(False, "Sending SMS to Mobile No." & sMobileNo, sFuncName, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sending SMS for Mobile NO." & sMobileNo, sFuncName)

                sReturnMsg = oWebClient.DownloadString(QStr)

                AddDataToTable(oDTSMS, sDocNum, sMobileNo, Mid(sReturnMsg, 5, sReturnMsg.Length - 8))

            Next

            Thread.Sleep(30000)

            For iRow As Integer = 0 To oDTSMS.Rows.Count - 1

                ''GETTING THE STATUS FOR SEND SMS
                sDocNum = oDTSMS.Rows(iRow)(0).ToString().Trim()
                sMobileNo = oDTSMS.Rows(iRow)(1).ToString().Trim()
                sMessageID = oDTSMS.Rows(iRow)(2).ToString().Trim()

                sStatusUrl = "http://mx.fortdigital.net/http/request-status-update?username={0}&password={1}&message-id={2}"

                sStatusQStr = String.Format(sStatusUrl, sUserName, sPassword, sMessageID)
                oWebClient.Encoding = System.Text.Encoding.ASCII
                oWebClient.UseDefaultCredentials = False

                ' System.Diagnostics.EventLog.WriteEntry("AE_FHG_SMSNotification", "SMS Status")

                Dim sStatus As String = String.Empty

                sStatus = oWebClient.DownloadString(sStatusQStr)

                '' WriteToStatusScreen(False, "Message ID : " & sMessageID, sFuncName, sErrDesc)

                If Mid(sStatus, 1, 7).ToString().ToUpper() = "SUCCESS" Then
                    WriteToStatusScreen(False, "Successfully Sent SMS to mobile No." & sMobileNo, sFuncName, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Sent SMS to mobile No." & sMobileNo, sFuncName)

                    ''WriteToStatusScreen(False, "Updating the Flag to Mobile No." & sMobileNo, sFuncName, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery for Update the Flag. Mobile NO." & sMobileNo, sFuncName)
                    sQuery = "UPDATE OVPM SET ""U_AI_SentSMS""='Y' WHERE ""DocNum""='" & sDocNum & "'"
                    ExecuteSQLNonQuery(sQuery)
                Else
                    WriteToStatusScreen(False, sStatus & ".Mobile No." & sMobileNo, sFuncName, sErrDesc)
                    '' WriteToStatusScreen(False, Mid(sStatus, 8, sStatus.Length - 1) & ".Mobile No." & sMobileNo, sFuncName, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Sent SMS to mobile No." & sMobileNo, sFuncName)
                End If

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SendSMS = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            WriteToStatusScreen(False, "Completed with ERROR. Error : " & sErrDesc, sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SendSMS = RTN_ERROR

        End Try

    End Function

    Private Sub frmSendSMS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim sFuncName As String = String.Empty

        Try

            sFuncName = "frmSendSMS_Load()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)


            Dim img As New System.Drawing.Icon(Application.StartupPath & "\SMSNotification.ico")
            Me.Icon = img

            With cmbType
                .SelectedIndex = 1
                .DropDownStyle = ComboBoxStyle.DropDown
            End With

            With cmbPaymentType
                .SelectedIndex = 0
                .DropDownStyle = ComboBoxStyle.DropDown
            End With

            'Getting the Parameter Values from App Cofig File

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            p_sSAPConnString = String.Empty

            p_sSAPConnString = "Data Source=" & p_oCompDef.p_sServerName & ";Initial Catalog=" & p_oCompDef.p_sDataBaseName & ";User ID=" & p_oCompDef.p_sDBUserName & "; Password=" & p_oCompDef.p_sDBPassword

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            Console.WriteLine("Completed with ERROR : " & sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

    Function WriteToStatusScreen(ByVal bIsClear As Boolean, ByVal strErrText As String, _
                                 ByVal strSourceName As String, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   WriteToStatusScreen()
        '   Purpose     :   This Subroutine will Display the status message in the Message text box.
        '   Author      :   SRINIVASANM
        '   Date        :   JULY 2015
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim strTempString As String = String.Empty

        Try

            sFuncName = "WriteToStatusScreen()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            'strTempString = Space(IIf(Len(strSourceName) > 30, 0, 30 - Len(strSourceName)))
            'strSourceName = strTempString & strSourceName
            'strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText

            If bIsClear Then
                rtxtMessage.Text = ""
            End If

            With rtxtMessage
                .HideSelection = True
                .Text &= strErrText & vbCrLf
                .SelectAll()
                .ScrollToCaret()
                .Refresh()
            End With

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            WriteToStatusScreen = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToStatusScreen = RTN_ERROR
        End Try

    End Function

    Function Header_Validation(ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "Header_Validation()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If String.Format(txtBatchNo.Text) = String.Empty Then
                txtBatchNo.Focus()
                sErrDesc = "Batch Number Should not be Blank!"
                Call WriteToLogFile(sErrDesc, sFuncName)
                WriteToStatusScreen(False, sErrDesc, sFuncName, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Return RTN_ERROR
            ElseIf String.Format(cmbType.SelectedItem) = String.Empty Then
                cmbType.Focus()
                sErrDesc = "Type Should not be Blank!"
                Call WriteToLogFile(sErrDesc, sFuncName)
                WriteToStatusScreen(False, sErrDesc, sFuncName, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Return RTN_ERROR
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Header_Validation = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Header_Validation = RTN_ERROR
        End Try
    End Function

    Function ProcessSMS(ByVal sBatchNo As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sQuery As String = String.Empty

        Try
            sFuncName = "ProcessSMS()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            ''FETCHING ALL THE EXISTING BATCH NUMBERS IN SAPAND STORE IT IN DT
            sQuery = "SELECT T0.""U_AI_BatchNo"" FROM OVPM T0 WHERE T0.""U_AI_BatchNo"" IS NOT NULL AND T0.""U_AI_BatchNo""='" & sBatchNo & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQueryForDT() for getting All the BatchNumbers.", sFuncName)
            dtBatch = ExecuteSQLQueryForDT(sQuery, p_oCompDef.p_sDataBaseName)

            dtBatch.DefaultView.RowFilter = "U_AI_BatchNo = '" & sBatchNo & "'"
            If dtBatch.DefaultView.Count = 0 Then
                sErrDesc = "Batch NO :: " & sBatchNo & " provided does not exist in SAP."
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If sPaymentType.ToString().ToUpper = "CHEQUE" Then

                sQuery = "SELECT T0.""DocNum"", T0.""U_AI_MobileNo"", T0.""DocTotal"" FROM OVPM T0  INNER JOIN VPM1 T1 ON T0.""DocEntry"" " & _
               " = T1.""DocNum"" WHERE T0.""U_AI_BatchNo"" ='" & sBatchNo & "' AND IFNULL( T0.""U_AI_SentSMS"" ,'N')='N' " & _
               " GROUP BY T0.""DocNum"", T0.""U_AI_MobileNo"", T0.""DocTotal"""

            ElseIf sPaymentType.ToString().ToUpper = "GIRO" Then

                sQuery = "SELECT T0.""DocNum"", T0.""U_AI_MobileNo"",T0.""TrsfrSum"" FROM OVPM T0 WHERE T0.""TrsfrSum"" >0 " & _
               "  AND T0.""U_AI_BatchNo"" ='" & sBatchNo & "' AND IFNULL( T0.""U_AI_SentSMS"" ,'N')='N' GROUP BY T0.""DocNum"", T0.""U_AI_MobileNo"",T0.""TrsfrSum"" "
            End If

            ''QUERY FOR FETCHING ALL TH E CHECK DETAILS AND STORE IT IN DATATABLE

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQueryForDT() for getting All Check payment Details.", sFuncName)
            dtSMSHeader = ExecuteSQLQueryForDT(sQuery, p_oCompDef.p_sDataBaseName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessSMS = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessSMS = RTN_ERROR
        End Try

    End Function

    Function Validate_PhoneNumbers(ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sMobileNo As String = String.Empty
        Dim sDocNum As String = String.Empty

        Try
            sFuncName = "Validate_PhoneNumbers()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            iTotRowCount = dtSMSHeader.Rows.Count

            For iRow As Integer = 0 To dtSMSHeader.Rows.Count - 1

                sDocNum = dtSMSHeader.Rows(iRow)(0).ToString().Trim()
                sMobileNo = dtSMSHeader.Rows(iRow)(1).ToString().Trim()

                If sMobileNo.ToString().Length = 0 Then
                    Call WriteToLogFile("Please Check Document Number : " & sDocNum & ". Mobile Number is Blank!", sFuncName)
                    iErrorCount = iErrorCount + 1
                Else
                    iSuccessCount = iSuccessCount + 1
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Validate_PhoneNumbers = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Validate_PhoneNumbers = RTN_ERROR
        End Try

    End Function

End Class
