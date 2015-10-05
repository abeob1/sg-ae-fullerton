
Module modProcess

    Public Sub Start()
        Dim sFuncname As String = "Start"
        Uploadfiles()

        'Send Error Email if Datable has rows.
        If p_oDtError.Rows.Count > 0 Then
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Error()", sFuncName)
            EmailTemplate_Error()
        End If
        p_oDtError.Rows.Clear()

    End Sub

    Private Sub Uploadfiles()

        Dim sFuncName As String = "Uploadfiles()"
        Dim sErrDesc As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function..", sFuncName)
            Console.WriteLine("Starting Upload Files.....")
            p_oDtError = CreateDataTable("FileName", "Status", "ErrDesc")

            '================================ MED AND FAH  =======================================================================

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function UploadMED_FAH()", sFuncName)
            If UploadFAH(sErrDesc) <> RTN_SUCCESS Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while Uploading the AR csv files", sFuncName)
                WriteToLogFile(sErrDesc, sFuncName)
            End If

            Console.WriteLine("Completed Upload the Files.")
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in upload setup", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

End Module
