Imports System.Threading

Public Class frmDBSEncryption

    Private Sub frmDBSEncryption_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "frmDBSEncryption_Load()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            txtSelectFileFolder.Text = p_oCompDef.sInboxDir
            txtSelectProcessFolder.Text = p_oCompDef.sSuccessDir

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            WriteToStatusScreen(False, "Completed with ERROR. Error : " & sErrDesc, sFuncName, sErrDesc)
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

    Function FolderSelection(ByRef sErrDesc As String) As String

        Dim sFuncName As String = String.Empty
        Dim folderDlg As New FolderBrowserDialog
        Dim sFolderPath As String = String.Empty

        Try

            sFuncName = "FolderSelection()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            folderDlg.ShowNewFolderButton = True
            If (folderDlg.ShowDialog() = DialogResult.OK) Then
                sFolderPath = folderDlg.SelectedPath
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            FolderSelection = sFolderPath

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            WriteToStatusScreen(False, "Completed with ERROR. Error : " & sErrDesc, sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            FolderSelection = String.Empty
        End Try

    End Function

    Private Sub btnSelectFileFolder_Click(sender As System.Object, e As System.EventArgs) Handles btnSelectFileFolder.Click

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "btnSelectFileFolder_Click()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FolderSelection()", sFuncName)
            txtSelectFileFolder.Text = FolderSelection(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            WriteToStatusScreen(False, "Completed with ERROR. Error : " & sErrDesc, sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)

        End Try
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        End
    End Sub

    Private Sub btnSelectProcessFolder_Click(sender As System.Object, e As System.EventArgs) Handles btnSelectProcessFolder.Click
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "btnSelectProcessFolder_Click()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FolderSelection()", sFuncName)
            txtSelectProcessFolder.Text = FolderSelection(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            WriteToStatusScreen(False, "Completed with ERROR. Error : " & sErrDesc, sFuncName, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)

        End Try
    End Sub

    Private Sub btnEncrypt_Click(sender As System.Object, e As System.EventArgs) Handles btnEncrypt.Click
        Dim sFuncName As String = String.Empty
        Dim permanent As Boolean = True
        Dim Command As String = String.Empty
        Dim Argument As String = String.Empty
        Dim sFileName As String = String.Empty
        Dim bFileExist As Boolean = False

        Try
            sFuncName = "Start()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Validation()", sFuncName)
            If Validation(sErrDesc) <> RTN_SUCCESS Then Exit Sub

            WriteToStatusScreen(True, "Please wait... ", sFuncName, sErrDesc)

            Cursor.Current = Cursors.WaitCursor

            Dim DirInfo As New System.IO.DirectoryInfo(txtSelectFileFolder.Text.Trim())
            Dim files() As System.IO.FileInfo

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting all the files from the Input folder", sFuncName)
            files = DirInfo.GetFiles("*.csv")

            For Each oFile As System.IO.FileInfo In files

                bFileExist = True

                sFileName = Replace(oFile.Name, ".csv", "") & "_" & Now.Year & Now.Month & Now.Day & "_" & Now.Hour & Now.Minute & Now.Second

                'Console.WriteLine("Rename the file")
                WriteToStatusScreen(False, "Rename the file", sFuncName, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rename the file", sFuncName)
                Rename(oFile.FullName, oFile.Directory.FullName & "\" & sFileName & oFile.Extension)

                'Console.WriteLine("Encrypting the file... File Name : " & sFileName)
                WriteToStatusScreen(False, "Encrypting the file... File Name : " & sFileName, sFuncName, sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Encrypting the file... File Name : " & sFileName, sFuncName)

                Dim strFileName As String = "--encrypt " & oFile.Directory.FullName & "\" & sFileName & oFile.Extension
                Process.Start("C:\SAP\Bank file - DBS Encryption\IDEALConnect\IDEALConnect.exe", strFileName)

                Thread.Sleep(20000)
                WriteToStatusScreen(False, "Successfully encrypted the file. File Name : " & sFileName, sFuncName, sErrDesc)
                'Console.WriteLine("Successfully encrypted the file. File Name : " & sFileName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully encrypted the file. File Name : " & sFileName, sFuncName)

                'Console.WriteLine("Moving the file to process folder")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToLocation()", sFuncName)

                FileMoveToLocation(oFile.Directory.FullName & "\" & sFileName & oFile.Extension, txtSelectProcessFolder.Text.Trim() & "\" & sFileName & oFile.Extension, sErrDesc)

                'Console.WriteLine("Successfully moved to prcess folder")

                'Console.WriteLine("          ")

            Next

            If bFileExist = False Then
                WriteToStatusScreen(False, "No files found to encrypt.", sFuncName, sErrDesc)
                'Console.WriteLine("No files found to encrypt.")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No files found in input folder", sFuncName)
            End If

            WriteToStatusScreen(False, "============================ COMPLETED ============================", sFuncName, sErrDesc)
            Cursor.Current = Cursors.Default

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message().ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            Cursor.Current = Cursors.Default
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
        End Try
    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click

        Dim sFuncName As String = String.Empty

        Try

            sFuncName = "btnClear_Click()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            rtxtMessage.Clear()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message().ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
        End Try
    End Sub

    Private Function Validation(sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "Validation()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If txtSelectFileFolder.Text.Trim() = String.Empty Then
                MessageBox.Show("Select File Folder cannot be null!", "DBSFile Encryption", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

                Return RTN_ERROR
            ElseIf txtSelectProcessFolder.Text.Trim() = String.Empty Then
                MessageBox.Show("Select Process Folder cannot be null!", "DBSFile Encryption", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

                Return RTN_ERROR
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            Validation = RTN_SUCCESS
        Catch ex As Exception
            Validation = RTN_ERROR
            sErrDesc = ex.Message().ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
        End Try
    End Function

End Class
