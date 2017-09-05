Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Data.SqlTypes
Public Class Connection
    Public Shared sConn As SqlConnection
    Public Shared sConnSAP As SqlConnection
    Public Shared bConnect As Boolean

    Public Sub setDB()
        Try
            Dim strConnect As String = ""
            Dim sCon As String = ""
            Dim SQLType As String = ""
            Dim sErrMsg As String = ""
            strConnect = "SAPConnect"
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString.trim()
         


            Dim objIniFile As New INIClass(sPath.Replace("bin", "") & "\" & "ConfigFile.ini")
            PublicVariable.SQLDB = objIniFile.GetString("SQL", "SQLDB", "")
            PublicVariable.SAPUser = objIniFile.GetString("SAP", "SAPUser", "")
            PublicVariable.SAPPwd = objIniFile.GetString("SAP", "SAPPwd", "")
            PublicVariable.SQLServer = objIniFile.GetString("SQL", "SQLServer", "")
            PublicVariable.SQLUser = objIniFile.GetString("SQL", "SQLUser", "")
            PublicVariable.SQLPwd = objIniFile.GetString("SQL", "SQLPwd", "")
            PublicVariable.SAPLicenseServer = objIniFile.GetString("SAP", "SAPLicenseServer", "")
            PublicVariable.SQLType = objIniFile.GetString("SQL", "SQLType", "")

            ' sCon = System.Configuration.ConfigurationSettings.AppSettings.Get(strConnect)
            '  MyArr = sCon.Split(";")
            If IsNothing(PublicVariable.oCompany) Then
                PublicVariable.oCompany = New SAPbobsCOM.Company
            End If
            PublicVariable.oCompany.CompanyDB = PublicVariable.SQLDB ' MyArr(0).ToString.Trim()
            PublicVariable.oCompany.UserName = PublicVariable.SAPUser ' MyArr(1).ToString.trim()
            PublicVariable.oCompany.Password = PublicVariable.SAPPwd ' MyArr(2).ToString.trim()
            PublicVariable.oCompany.Server = PublicVariable.SQLServer ' MyArr(3).ToString.trim()
            PublicVariable.oCompany.DbUserName = PublicVariable.SQLUser ' MyArr(4).ToString.trim()
            PublicVariable.oCompany.DbPassword = PublicVariable.SQLPwd ' MyArr(5).ToString.Trim()
            'PublicVariable.oCompany.LicenseServer = PublicVariable.SAPLicenseServer ' MyArr(6)
            SQLType = PublicVariable.SQLType ' MyArr(7)
            PublicVariable.oCompany.UseTrusted = False
            If SQLType = 2008 Then
                PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            Else
                PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            End If

            'sCon = "server= " + MyArr(3).ToString.Trim() + ";database=" + MyArr(0).ToString.Trim() + " ;uid=" + MyArr(4).ToString.Trim() + "; pwd=" + MyArr(5).ToString.Trim() + ";"
            ' sConnSAP = New SqlConnection(sCon)
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
    Public Function connectDB() As Boolean
        Try
            Dim sErrMsg As String = ""
            Dim connectOk As Integer = 0
            If PublicVariable.oCompany.Connect() <> 0 Then
                PublicVariable.oCompany.GetLastError(connectOk, sErrMsg)
                bConnect = False
                Functions.WriteLog("Error Msg:" & sErrMsg)
                Return False
            Else
                bConnect = True
                Return True
            End If
        Catch ex As Exception
            'Dim file As System.IO.StreamWriter = New System.IO.StreamWriter("C:\\connectDB.txt", True)
            'file.WriteLine(ex)
            'file.Close()
            Return False
            Functions.WriteLog(ex.ToString)
        End Try
    End Function
#Region "ADO SAP"
    Private Function GetConnectionString_SAP() As SqlConnection
        If sConnSAP.State = ConnectionState.Open Then
            sConnSAP.Close()
        End If
        Try
            sConnSAP.Open()
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
        Return sConnSAP
    End Function
    Public Function ObjectGetAll_Query_SAP(ByVal QueryString As String) As DataSet
        Try

            Using myConn = GetConnectionString_SAP()
                Dim MyCommand As SqlCommand = New SqlCommand(QueryString, myConn)
                MyCommand.CommandType = CommandType.Text
                Dim da As SqlDataAdapter = New SqlDataAdapter()
                Dim mytable As DataSet = New DataSet()
                da.SelectCommand = MyCommand
                da.Fill(mytable)
                myConn.Close()
                Return mytable
            End Using
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
            Return Nothing
        End Try
    End Function
#End Region
    
End Class
