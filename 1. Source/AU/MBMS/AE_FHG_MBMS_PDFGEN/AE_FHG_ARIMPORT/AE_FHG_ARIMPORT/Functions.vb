Imports System.IO

Public Class Functions
    Public Shared Sub LoadConfigFile()
        Dim IniFile As String
        IniFile = Application.StartupPath + "\ConfigFile.ini"
        If File.Exists(IniFile) Then
            Dim objIniFile As New INIClass(IniFile)
            PublicVariable.SMTPServer = objIniFile.GetString("SMTP", "SMTP_Server", "")
            PublicVariable.SMTPPort = objIniFile.GetString("SMTP", "SMTP_Port", "")
            PublicVariable.SMTPEmail = objIniFile.GetString("SMTP", "SMTP_Email", "")
            PublicVariable.SMTPPwd = objIniFile.GetString("SMTP", "SMTP_Pwd", "")
            PublicVariable.EmailList = objIniFile.GetString("SMTP", "EMAILLIST", "")

            PublicVariable.InputPath = objIniFile.GetString("FILE", "FILE_INPUTPATH", "")
            PublicVariable.LogFilePath = objIniFile.GetString("FILE", "FILE_LOGFILEPATH", "")

            PublicVariable.SAPUser = objIniFile.GetString("SAP", "SAPUser", "")
            PublicVariable.SAPPwd = objIniFile.GetString("SAP", "SAPPwd", "")
            PublicVariable.SAPLicenseServer = objIniFile.GetString("SAP", "SAPLicenseServer", "")

            PublicVariable.SQLServer = objIniFile.GetString("SQL", "SQLServer", "")
            PublicVariable.SQLUser = objIniFile.GetString("SQL", "SQLUser", "")
            PublicVariable.SQLPwd = objIniFile.GetString("SQL", "SQLPwd", "")
            PublicVariable.SQLDB = objIniFile.GetString("SQL", "SQLDB", "")
            PublicVariable.SQLType = objIniFile.GetString("SQL", "SQLType", "")

        End If
    End Sub
    Public Shared Sub WriteLog(ByVal Str As String)
        Dim oWrite As IO.StreamWriter
        Dim FilePath As String
        FilePath = Application.StartupPath + "\ErrorLog.txt"

        If IO.File.Exists(FilePath) Then
            oWrite = IO.File.AppendText(FilePath)
        Else
            oWrite = IO.File.CreateText(FilePath)
        End If
        oWrite.Write(Now.ToString() + ":" + Str + vbCrLf)
        oWrite.Close()
    End Sub
End Class
