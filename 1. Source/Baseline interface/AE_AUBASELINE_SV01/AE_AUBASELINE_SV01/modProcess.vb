Module modProcess

#Region "Start"
    Public Sub Start()
        Dim sFuncName As String = "Start()"
        Dim sErrDesc As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("calling ReadExcel()", sFuncName)

            Console.WriteLine("Reading CSV values")

            UploadFiles(sErrDesc)

            'Send Error Email if Datable has rows.
            If p_oDtError.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Error()", sFuncName)
                EmailTemplate_Error()
            End If
            p_oDtError.Rows.Clear()

            'Send Success Email if Datable has rows..
            If p_oDtSuccess.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Success()", sFuncName)
                EmailTemplate_Success()
            End If
            p_oDtSuccess.Rows.Clear()


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End
        End Try

    End Sub
#End Region

#Region "Read Excel Files"
    Private Function UploadFiles(ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "UploadFiles"
        Dim bIsFileExists As Boolean = False
        Dim oDVData As DataView = New DataView

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting funciton", sFuncName)

            p_oDtSuccess = CreateDataTable("FileName", "Status")
            p_oDtError = CreateDataTable("FileName", "Status", "ErrDesc")
            p_oDtReport = CreateDataTable("Type", "DocEntry", "BPCode", "Owner")

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo

            files = DirInfo.GetFiles("*.csv")

            For Each file As System.IO.FileInfo In files
                sErrDesc = String.Empty
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is: " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File: " & file.Name.ToUpper)

                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)

                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & sFileName
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                'If sFileName.Contains("_Baseline Group_") Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel(.csv) file into Dataview", sFuncName)
                oDVData = GetDataViewFromCSV(p_oCompDef.sInboxDir & "\" & file.Name)

                If Not oDVData Is Nothing Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessSalesDocFile()", sFuncName)
                    If ProcessSalesDocFile(file, oDVData, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                    Continue For
                End If
                'End If
            Next

        Catch ex As Exception

        End Try
    End Function
#End Region

End Module
