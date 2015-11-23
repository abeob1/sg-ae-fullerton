Imports Microsoft.VisualBasic.Devices
Imports System.IO
Imports System.Threading


Module modDBSEncryption

    Public Sub Start()

        Dim sFuncName As String = String.Empty
        Dim permanent As Boolean = True
        Dim Command As String = String.Empty
        Dim Argument As String = String.Empty
        Dim sFileName As String = String.Empty
        Dim bFileExist As Boolean = False

        Try
            sFuncName = "Start()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)


            Dim oFromDirInfo = New DirectoryInfo(p_oCompDef.sInboxDir)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling MoveAllItemsTo()", sFuncName)
            If MoveAllItemsTo(oFromDirInfo, p_oCompDef.sPaymentFileDir, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sPaymentFileDir)
            Dim files() As System.IO.FileInfo

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting all the files from the Input folder", sFuncName)
            files = DirInfo.GetFiles("*.csv")

            For Each oFile As System.IO.FileInfo In files

                bFileExist = True

                sFileName = Replace(oFile.Name, ".csv", "") & "_" & Now.Year & Now.Month & Now.Day & "_" & Now.Hour & Now.Minute & Now.Second

                Console.WriteLine("Rename the file")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rename the file", sFuncName)
                Rename(oFile.FullName, oFile.Directory.FullName & "\" & sFileName & oFile.Extension)

                Console.WriteLine("Encrypting the file... File Name : " & sFileName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Encrypting the file... File Name : " & sFileName, sFuncName)

                Dim strFileName As String = "--encrypt " & oFile.Directory.FullName & "\" & sFileName & oFile.Extension
                Process.Start("C:\SAP\Bank file - DBS Encryption\IDEALConnect\IDEALConnect.exe", strFileName)

                Thread.Sleep(15000)

                Console.WriteLine("Successfully encrypted the file. File Name : " & sFileName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully encrypted the file. File Name : " & sFileName, sFuncName)

                Console.WriteLine("Moving the file to process folder")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToLocation()", sFuncName)

                If FileMoveToLocation(oFile.Directory.FullName & "\" & sFileName & oFile.Extension, p_oCompDef.sSuccessDir & "\" & sFileName & oFile.Extension, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Console.WriteLine("Successfully moved to prcess folder")

                Console.WriteLine("          ")

            Next

            If bFileExist = False Then
                Console.WriteLine("No files found to encrypt.")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No files found in input folder", sFuncName)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message().ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
        End Try
    End Sub

End Module
