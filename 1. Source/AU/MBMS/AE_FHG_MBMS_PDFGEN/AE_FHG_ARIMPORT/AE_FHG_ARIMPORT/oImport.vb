Imports System.Net.Mail
Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient


Public Class oImport
    Dim TDT As String = Format(Now.Date, "yyyyMMdd")
    Dim TDT1 As String = Format(Now.Date, "ddMMyyyy")
    Dim FIH As String = ""
    Dim FIR As String = ""
    Dim CMaster As Boolean = True
    Dim CName As Boolean = True
    Dim CAddr As Boolean = True
    Dim CSales As Boolean = True
    Dim Crel As Boolean = True
    Dim SOH As Boolean = True
    Dim IH As Boolean = True
    Dim oFunctions As New Functions
    'Function RemoveAccessDatabase( _
    'ByVal FileName As String, _
    'ByVal WaitTime As Integer, _
    'ByVal Loops As Integer) As Boolean

    '    Dim Success As Boolean = False

    '    Dim LockFile As String = IO.Path.ChangeExtension(FileName, "ldb")

    '    For Counter As Integer = 0 To Loops
    '        If IO.File.Exists(LockFile) Then
    '            System.Threading.Thread.Sleep(WaitTime)
    '            IO.File.Delete(FileName)
    '        Else
    '            Success = True
    '            Exit For
    '        End If
    '    Next

    '    Return Success

    'End Function
    Public Sub Import()
        Dim connect As New Connection()

        If Connection.bConnect = False Then
            connect.setDB()
            If Not connect.connectDB() Then
                Functions.WriteLog("SAP Connection Failed")
            Else


                AR_INVOICE_FULL("Fullerton Indonesia AR Invoice Header Upload_G.csv", "Fullerton Indonesia AR Invoice Lines Upload_G.csv", "C:\TEST\")
                MsgBox("Complected")
                'Import1()
                ' Connection.bConnect = False
                '  Me.Close()
            End If
        Else
            AR_INVOICE_FULL("Fullerton Indonesia AR Invoice Header Upload_G.csv", "Fullerton Indonesia AR Invoice Lines Upload_G.csv", "C:\TEST\")
            MsgBox("Complected")
            'Import1()
            ' Connection.bConnect = False
            '  Me.Close()
        End If
    End Sub
    Public Sub Import1()
        Try
            Dim Custaddr As String()
            Dim FCustaddr As String = ""
            Dim dt As String = Format(Now.Date.AddDays(-1), "yyyyMMdd")
            Dim str As String = "CUSTADDR_" & dt
            Dim sPath As String = IO.Directory.GetParent(Application.StartupPath).ToString.Trim()
            Dim objIniFile As New INIClass(sPath.Replace("bin", "") & "\" & "ConfigFile.ini")
            PublicVariable.InputPath = objIniFile.GetString("FILE", "FILE_INPUTPATH", "")
            PublicVariable.RollBack = objIniFile.GetString("SAP", "RollBack", "")
            PublicVariable.LogFilePath = objIniFile.GetString("FILE", "FILE_LOGFILEPATH", "")
            PublicVariable.SuccessFolder = objIniFile.GetString("FILE", "FILE_SUCCESS", "")
            PublicVariable.ErrorFolder = objIniFile.GetString("FILE", "FILE_ERROR", "")
            Custaddr = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If Custaddr.Length >= 1 Then
                FCustaddr = Custaddr(0)
            End If
            Dim Custname As String()
            Dim FCustname As String = ""
            str = "CUSTNAME_" & dt
            Custname = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If Custname.Length >= 1 Then
                FCustname = Custname(0)
            End If
            Dim CustMaster As String()
            Dim FCustMaster As String = ""
            str = "CUSTOMER_" & dt
            CustMaster = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If CustMaster.Length >= 1 Then
                FCustMaster = CustMaster(0)
            End If
            Dim CustRelH As String()
            Dim FCustRel_H As String = ""
            str = "RELATION_" & dt
            CustRelH = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If CustRelH.Length >= 1 Then
                FCustRel_H = CustRelH(0)
            End If
            Dim CustRelR As String()
            Dim FCustRel_R As String = ""
            str = "CUSTRELT_" & dt
            CustRelR = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If CustRelR.Length >= 1 Then
                FCustRel_R = CustRelR(0)
            End If
            Dim CustSales As String()
            Dim FCustSales As String = ""
            str = "SPST_" & dt
            CustSales = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If CustSales.Length >= 1 Then
                FCustSales = CustSales(0)
            End If
            Dim IH As String()

            str = "SOH_" & dt
            IH = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If IH.Length >= 1 Then
                FIH = IH(0)
            End If

            Dim IR As String()

            str = "SOR_" & dt
            IR = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If IR.Length >= 1 Then
                FIR = IR(0)
            End If
            Dim IH1 As String()
            Dim FIH1 As String = ""
            str = "IH_" & dt
            IH1 = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If IH1.Length >= 1 Then
                FIH1 = IH1(0)
            End If
            Dim IR1 As String()
            Dim FIR1 As String = ""
            str = "IR_" & dt
            IR1 = Directory.GetFiles(PublicVariable.InputPath, String.Format("{0}*.csv", str))
            If IR1.Length >= 1 Then
                FIR1 = IR1(0)
            End If
            Try
                Dim Filenum As Integer = FreeFile()
                FileOpen(Filenum, "" & PublicVariable.LogFilePath & "ErrorLog.txt", OpenMode.Output)
                FileClose()
                FileOpen(Filenum, "" & PublicVariable.LogFilePath & "ErrorLog_Master" & TDT & ".txt", OpenMode.Output)
                FileClose()
                FileOpen(Filenum, "" & PublicVariable.LogFilePath & "ErrorLog_Name" & TDT & ".txt", OpenMode.Output)
                FileClose()
                FileOpen(Filenum, "" & PublicVariable.LogFilePath & "ErrorLog_Addr" & TDT & ".txt", OpenMode.Output)
                FileClose()
                FileOpen(Filenum, "" & PublicVariable.LogFilePath & "ErrorLog_Sales" & TDT & ".txt", OpenMode.Output)
                FileClose()
                FileOpen(Filenum, "" & PublicVariable.LogFilePath & "ErrorLog_ARInvoice_IH" & TDT & ".txt", OpenMode.Output)
                FileClose()
                FileOpen(Filenum, "" & PublicVariable.LogFilePath & "ErrorLog_ARInvoice_SOH" & TDT & ".txt", OpenMode.Output)
                FileClose()
                FileOpen(Filenum, "" & PublicVariable.LogFilePath & "ErrorLog_CustRelation" & TDT & ".txt", OpenMode.Output)
                FileClose()
                'ErrorLog_ARInvoice

            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try
            Dim file As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog.txt", True)
            file.Close()
            Dim csvFileFolder As String = PublicVariable.InputPath
            Dim RollBack As String = PublicVariable.RollBack
            Try
                If FCustMaster <> "" Then
                    Try
                        Del_schema(csvFileFolder)
                        file.Close()
                    Catch ex As Exception

                    End Try
                    'Dim oFunction As Functions
                    'oFunction = New Functions
                    Functions.WriteLog("Customer Master Start:" & DateTime.Now)
                    CustomerMain(FCustMaster, csvFileFolder)
                    Functions.WriteLog("Customer Master End:" & DateTime.Now)
                End If
                If FCustname <> "" Then
                    Try
                        Del_schema(csvFileFolder)
                        file.Close()
                    Catch ex As Exception

                    End Try
                    'CustomerName(FCustname, csvFileFolder)
                End If
                If FCustaddr <> "" Then
                    Try
                        Del_schema(csvFileFolder)
                        file.Close()
                    Catch ex As Exception

                    End Try
                    ' If Now.Date.DayOfWeek = DayOfWeek.Saturday Then
                    Functions.WriteLog("Customer Address Start:" & DateTime.Now)
                    FCustaddress(FCustaddr, csvFileFolder)
                    Functions.WriteLog("Customer Address End:" & DateTime.Now)
                    'End If

                End If
                If FCustSales <> "" Then
                    Try
                        Del_schema(csvFileFolder)
                        file.Close()
                    Catch ex As Exception
                    End Try
                    'If Now.Date.DayOfWeek = DayOfWeek.Saturday Then
                    Functions.WriteLog("Customer Sales Person Start:" & DateTime.Now)
                    FCustSalesPer(FCustSales, csvFileFolder)
                    Functions.WriteLog("Customer Sales Person End:" & DateTime.Now)
                    'End If
                End If
                If FCustRel_H <> "" And FCustRel_R <> "" Then
                    Try
                        Del_schema(csvFileFolder)
                        file.Close()
                    Catch ex As Exception
                    End Try
                    'If Now.Date.DayOfWeek = DayOfWeek.Saturday Then
                    Functions.WriteLog("Customer Relation Start:" & DateTime.Now)
                    Customer_Relation(FCustRel_H, FCustRel_R, csvFileFolder)
                    Functions.WriteLog("Customer Relation End:" & DateTime.Now)
                    'End If
                End If



                If FIH <> "" And FIR <> "" Then
                    Try
                        Del_schema(csvFileFolder)
                        file.Close()
                    Catch ex As Exception
                    End Try
                    If RollBack = "Yes" Then
                        Functions.WriteLog("Invoice Start:" & DateTime.Now)
                        Dim bol As Boolean = Invoice21(FIH, FIR, csvFileFolder)
                        Functions.WriteLog("Invoice End:" & DateTime.Now)
                    Else
                        Functions.WriteLog("Invoice Start:" & DateTime.Now)
                        Invoice2(FIH, FIR, csvFileFolder)
                        Functions.WriteLog("Invoice End:" & DateTime.Now)
                    End If
                End If
                If FIH1 <> "" And FIR1 <> "" Then
                    Try
                        Del_schema(csvFileFolder)
                        file.Close()
                    Catch ex As Exception
                    End Try
                    If RollBack = "Yes" Then
                        Functions.WriteLog("Invoice_SH Start:" & DateTime.Now)
                        Dim bol As Boolean = Invoice11(FIH1, FIR1, csvFileFolder)
                        Functions.WriteLog("Invoice_SH End:" & DateTime.Now)
                        If bol = False Then
                            PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        Else
                            PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If
                    Else
                        Functions.WriteLog("Invoice_SH Start:" & DateTime.Now)
                        Invoice1(FIH1, FIR1, csvFileFolder)
                        Functions.WriteLog("Invoice_SH End:" & DateTime.Now)
                    End If
                End If
            Catch ex As Exception
                Functions.WriteLog(ex.ToString)
            End Try
            'MsgBox("Complected Successfully")

            EMail(FCustMaster, FCustname, FCustaddr, FCustSales, FCustRel_H, FIH, FIH1, FCustRel_R, FIR, FIR1, csvFileFolder)
            PublicVariable.oCompany.Disconnect()
            'PublicVariable.oCompany.Dispose()
            Try
                Del_schema(csvFileFolder)
                file.Close()
            Catch ex As Exception
            End Try
            GC.Collect()
            Environment.Exit(0)
            'MsgBox("Completed")
          
            'Dim oForm As New Form1
            'oForm.Close()
            'Call PublicVariable.oCompany.Disconnect()
            'PublicVariable.oCompany.Disconnect()

            'Connection.bConnect = False
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
            'MsgBox(ex.Message)
            'Call PublicVariable.oCompany.Disconnect()
            'PublicVariable.oCompany = Nothing
            'Exit Sub
        End Try
    End Sub
    Private Function AR_INVOICE_FULL(ByVal FIH As String, ByVal FIR As String, ByVal csvFileFolder As String) As Boolean
        Try
            Dim bo As Boolean = True
            ' PublicVariable.oCompany.StartTransaction()

            Dim csvFileName As String = FIH.Replace(csvFileFolder, "")
            Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            Dim s1, s2, s3, s4, s5 As String
            s1 = "[" & csvFileName & "]"
            s2 = "ColNameHeader=True"
            s3 = "Format=CSVDelimited"
            s4 = "MaxScanRows=0"
            s5 = "CharacterSet=OEM"
            srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf + s5.ToString())
            srOutput.Close()
            fsOutput.Close()

            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            Dim conn As New Odbc.OdbcConnection(connString)
            Dim da As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIH As System.Data.DataSet
            DtSet_FIH = New System.Data.DataSet
            da.Fill(DtSet_FIH)
            Dim rowcount_FIH As Integer = DtSet_FIH.Tables(0).Rows.Count
            '**********************************************************
            csvFileName = FIR.Replace(csvFileFolder, "")
            'connString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
            '            & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            'Dim conn1 As New Odbc.OdbcConnection(connString)
            Dim da1 As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIR As System.Data.DataSet
            DtSet_FIR = New System.Data.DataSet
            da1.Fill(DtSet_FIR)
            Dim rowcount_FIR As Integer = DtSet_FIR.Tables(0).Rows.Count
            Dim i, j As Integer
            Dim oINV As SAPbobsCOM.Documents
            Dim return_value As Integer
            Dim SerrorMsg As String = ""
            'PublicVariable.oCompany.StartTransaction()
            Dim oRecordSet6 As SAPbobsCOM.Recordset
            Try
                For i = 0 To rowcount_FIH - 1

                    oRecordSet6 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim str As String = "SELECT T0.U_AB_PERIODE FROM OINV T0 WHERE ifnull( T0.U_AB_PERIODE ,'') <> '' and  T0.U_AB_PERIODE ='" & DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() & "'"
                    oRecordSet6.DoQuery(str)
                    If oRecordSet6.RecordCount = 0 Then
                        oINV = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                        Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                        Try
                            If DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() <> "" Then
                                dt1 = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim()
                                dt1 = dt1.Insert(4, "-")
                                dt1 = dt1.Insert(7, "-")
                            End If
                        Catch ex As Exception

                        End Try
                        oINV.DocDate = dt1
                        dt1 = Format(Now.Date, "yyyy-MM-dd")
                        Try
                            If DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.Trim() <> "" Then
                                dt1 = DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.Trim()
                                dt1 = dt1.Insert(4, "-")
                                dt1 = dt1.Insert(7, "-")
                            End If
                        Catch ex As Exception
                        End Try
                        oINV.DocDueDate = dt1
                        oINV.CardCode = DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim()
                        oINV.NumAtCard = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                        oINV.DocCurrency = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString.Trim()
                        oINV.Comments = DtSet_FIH.Tables(0).Rows(i).Item(8).ToString.Trim()
                        oINV.JournalMemo = DtSet_FIH.Tables(0).Rows(i).Item(9).ToString.Trim()
                        oINV.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        oINV.UserFields.Fields.Item("U_AB_PERIODE").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim()
                        Dim Bol As Boolean = False
                        'If i = 233 Then
                        '    MsgBox("Hi")
                        'End If
                        Dim HEad As Integer = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim()

                        For j = 0 To rowcount_FIR - 1
                            'If j = 286 Then
                            '    MsgBox("Hi11")
                            'End If
                            Dim Row As Integer = DtSet_FIR.Tables(0).Rows(j).Item(0).ToString.Trim() '
                            If Row = HEad Then
                                Bol = True
                                Dim Len As Integer = DtSet_FIR.Tables(0).Rows(j).Item(3).ToString.Trim().Length
                                If Len > 100 Then
                                    oINV.Lines.ItemDescription = DtSet_FIR.Tables(0).Rows(j).Item(3).ToString.Trim().Substring(0, 99)
                                Else
                                    oINV.Lines.ItemDescription = DtSet_FIR.Tables(0).Rows(j).Item(3).ToString.Trim()
                                End If

                                oINV.Lines.Quantity = 1
                                oINV.Lines.AccountCode = DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim()
                                oINV.Lines.LineTotal = DtSet_FIR.Tables(0).Rows(j).Item(4).ToString.Trim()
                                oINV.Lines.VatGroup = DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim()
                                oINV.Lines.CostingCode = DtSet_FIR.Tables(0).Rows(j).Item(6).ToString.Trim()
                                oINV.Lines.CostingCode2 = DtSet_FIR.Tables(0).Rows(j).Item(7).ToString.Trim()
                                oINV.Lines.CostingCode3 = DtSet_FIR.Tables(0).Rows(j).Item(8).ToString.Trim()
                                oINV.Lines.Add()
                            End If
                        Next
                        return_value = oINV.Add()
                        If return_value = 0 Then
                        Else
                            PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                            Functions.WriteLog("Record No: " & DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() & " Err Msg:" & SerrorMsg)
                            bo = False
                            ' Exit Function
                        End If
                    End If
                Next
                If bo = False Then
                    '       PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                Else
                    '      PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
            Catch ex As Exception
                Functions.WriteLog(ex.Message)
                ' PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End Try

        Catch ex As Exception
            Functions.WriteLog(ex.Message)
            'PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try


    End Function

#Region "Import"
    Public Sub Del_schema(ByVal csvFileFolder As String)
        Try
            Dim FileToDelete As String
            FileToDelete = csvFileFolder & "\\schema.ini"
            If System.IO.File.Exists(FileToDelete) = True Then
                System.IO.File.Delete(FileToDelete)
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Customer_Relation(ByVal FCustRel_H As String, ByVal FCustRel_R As String, ByVal csvFileFolder As String)
        Dim file_CustRelation As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_CustRelation" & TDT & ".txt", True)
        ' Dim file_Sales As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_Sales" & TDT & ".txt", True)
        Try

            Dim CRBool As Boolean = False
            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                      & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""

            Dim conn As New Odbc.OdbcConnection(connString)
            Dim csvFileName As String = FCustRel_H.Replace(csvFileFolder, "")
            Dim da3 As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_Rel_H As System.Data.DataSet
            DtSet_Rel_H = New System.Data.DataSet
            da3.Fill(DtSet_Rel_H)
            Dim rowcount_Rel_H As Integer = DtSet_Rel_H.Tables(0).Rows.Count
            Dim oRecordSet2 As SAPbobsCOM.Recordset
            Dim oRecordSet3 As SAPbobsCOM.Recordset
            Dim oRecordSet4 As SAPbobsCOM.Recordset

            csvFileName = FCustRel_R.Replace(csvFileFolder, "")
            Dim da2 As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_Rel_R As System.Data.DataSet
            DtSet_Rel_R = New System.Data.DataSet
            da2.Fill(DtSet_Rel_R)
            Dim rowcount_Rel_R As Integer = DtSet_Rel_R.Tables(0).Rows.Count
            Dim i, j As Integer
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oSons As SAPbobsCOM.GeneralDataCollection
            Dim oSon As SAPbobsCOM.GeneralData
            Dim sCmp As SAPbobsCOM.CompanyService
            Dim OldRelCode As String = ""
            Dim NewRelCode As String = ""
            Dim oform As New Form1

            '****************************************************
            Dim oRecordSetBP11 As SAPbobsCOM.Recordset
            oRecordSetBP11 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSetBP11.DoQuery("Delete from ACUST")
            Dim connectionString As String = "Server=" & PublicVariable.SQLServer & ";Database=" & PublicVariable.SQLDB & ";User Id=" & PublicVariable.SQLUser & ";Password=" & PublicVariable.SQLPwd & ""
            Dim Cn = New SqlConnection(connectionString)
            Dim sql As String = "select Distinct(CARDCODE) 'CardCode' from OQUT union select Distinct(CARDCODE) 'CardCode' from ORDR union select Distinct(CARDCODE) 'CardCode' from OINV Union select Distinct(CARDCODE) 'CardCode' from ORIN"
            Dim connection As New SqlConnection(connectionString)
            Dim dataadapter As New SqlDataAdapter(sql, connection)
            Dim ds As New DataSet()
            connection.Open()
            dataadapter.Fill(ds, "Authors_table")
            connection.Close()
            Dim rd1 As DataTableReader = ds.Tables("Authors_table").CreateDataReader()

            Using destinationConnection As SqlConnection = _
               New SqlConnection(connectionString)
                destinationConnection.Open()
                Using copy As New SqlBulkCopy(destinationConnection)
                    copy.DestinationTableName = "ACUST"
                    copy.WriteToServer(rd1)
                End Using
                destinationConnection.Close()
            End Using
            '****************************************************
            For i = 0 To rowcount_Rel_H - 1
                Dim oRecordSetBP2 As SAPbobsCOM.Recordset
                oRecordSetBP2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSetBP2.DoQuery("Select CardCode from ACUST where CardCode='" & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                If oRecordSetBP2.RecordCount > 0 Then
                    CRBool = False
                    'oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oRecordSet2.DoQuery("SELECT CardCode  FROM OCRD  where CardCode='" & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                    'If oRecordSet2.RecordCount > 0 Then
                    'oform.Label1.Text = "Customer Relation " & rowcount_Rel_H & "-of " & i
                    'oform.Label1.Show()
                    Try
                        'If i = 3665 Then
                        '    MsgBox(i)
                        'End If
                        NewRelCode = DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim()
                        If NewRelCode <> OldRelCode And NewRelCode <> "" Then
                            OldRelCode = NewRelCode
                            oRecordSet4 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet4.DoQuery("SELECT T0.[U_AI_RelationCode] FROM [dbo].[@CUSTRELT]  T0 WHERE T0.[U_AI_RelationCode] ='" & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                            If oRecordSet4.RecordCount = 0 Then
                                CRBool = False
                                sCmp = PublicVariable.oCompany.GetCompanyService
                                oGeneralService = sCmp.GetGeneralService("CustRelation")
                                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                                oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet3.DoQuery("SELECT T0.[U_AI_RelationCode] FROM [dbo].[@RELATION]  T0")
                                oGeneralData.SetProperty("Code", (oRecordSet3.RecordCount + 1).ToString.Trim())
                                oGeneralData.SetProperty("Name", (oRecordSet3.RecordCount + 1).ToString.Trim())
                                oGeneralData.SetProperty("U_AI_RelationCode", DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim())
                                oGeneralData.SetProperty("U_AI_RelationName", DtSet_Rel_H.Tables(0).Rows(i).Item(1).ToString.Trim())
                                oGeneralData.SetProperty("U_AI_RelationType", DtSet_Rel_H.Tables(0).Rows(i).Item(2).ToString.Trim())
                                Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                                Try
                                    If DtSet_Rel_H.Tables(0).Rows(i).Item(3).ToString.Trim() <> "" Then
                                        dt1 = DtSet_Rel_H.Tables(0).Rows(i).Item(3).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try

                                oGeneralData.SetProperty("U_AI_DateAppointed", dt1)
                                oGeneralData.SetProperty("U_AI_Status", DtSet_Rel_H.Tables(0).Rows(i).Item(4).ToString.Trim())
                                '---------------
                                For j = 0 To rowcount_Rel_R - 1
                                    If DtSet_Rel_R.Tables(0).Rows(j).Item(0).ToString.Trim() = DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() Then

                                        Dim oRecordSetBPC As SAPbobsCOM.Recordset
                                        oRecordSetBPC = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSetBPC.DoQuery("SELECT CardCode  FROM OCRD  where CardCode='" & DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim() & "'")
                                        If oRecordSetBPC.RecordCount > 0 Then
                                            CRBool = True
                                            dt1 = Format(Now.Date, "yyyy-MM-dd")
                                            Try
                                                If DtSet_Rel_H.Tables(0).Rows(i).Item(4).ToString.Trim() <> "" Then
                                                    dt1 = DtSet_Rel_R.Tables(0).Rows(j).Item(4).ToString.Trim()
                                                    dt1 = dt1.Insert(4, "-")
                                                    dt1 = dt1.Insert(7, "-")
                                                End If
                                            Catch ex As Exception

                                            End Try

                                            Dim oRecordSetBP As SAPbobsCOM.Recordset
                                            oRecordSetBP = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            'Dim s As String = "UPDATE OCRD set [U_AI_DateAppointed]='" & dt1 & "', [U_AI_RelationCode]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(0).ToString.Trim().Replace("'", " ") & "', [U_AI_RelationName]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(1).ToString.Trim() & "', [U_AI_RelationType]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(3).ToString.Trim() & "' WHERE [CardCode] ='" & DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim() & "'"
                                            oRecordSetBP.DoQuery("UPDATE OCRD set [U_AI_DateAppointed]='" & dt1 & "', [U_AI_RelationCode]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(0).ToString.Trim() & "', [U_AI_RelationName]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(1).ToString.Trim().Replace("'", " ") & "', [U_AI_RelationType]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(3).ToString.Trim() & "' WHERE [CardCode] ='" & DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim() & "'")


                                            oSons = oGeneralData.Child("CUSTRELT")
                                            oSon = oSons.Add
                                            oSon.SetProperty("U_AI_RelationCode", DtSet_Rel_R.Tables(0).Rows(j).Item(0).ToString.Trim())
                                            oSon.SetProperty("U_AI_RelationName", DtSet_Rel_R.Tables(0).Rows(j).Item(1).ToString.Trim())
                                            Dim dt As String = DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim()
                                            oSon.SetProperty("U_AI_EntityCode", DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim())
                                            oSon.SetProperty("U_AI_RelationType", DtSet_Rel_R.Tables(0).Rows(j).Item(3).ToString.Trim())
                                            Dim dt2 As String = Format(Now.Date, "yyyy-MM-dd")
                                            Try
                                                If DtSet_Rel_R.Tables(0).Rows(j).Item(4).ToString.Trim() <> "" Then
                                                    dt2 = DtSet_Rel_R.Tables(0).Rows(j).Item(4).ToString.Trim()
                                                    dt2 = dt1.Insert(4, "-")
                                                    dt2 = dt1.Insert(7, "-")
                                                End If
                                                ''MsgBox(dt2)
                                                Dim dt3 As String = dt2.Remove(8, 1)
                                                oSon.SetProperty("U_AI_DateAppointed", dt3)
                                            Catch ex As Exception

                                            End Try

                                            oSon.SetProperty("U_AI_Alternate", DtSet_Rel_R.Tables(0).Rows(j).Item(5).ToString.Trim())
                                            oSon.SetProperty("U_AI_Industry", DtSet_Rel_R.Tables(0).Rows(j).Item(6).ToString.Trim())
                                            oSon.SetProperty("U_AI_Country", DtSet_Rel_R.Tables(0).Rows(j).Item(7).ToString.Trim())
                                        Else
                                            file_CustRelation.WriteLine("Header " & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() & "-Relation Code Failed; Line No - " & i + 1 & ":Row " & DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim() & " BP- " & j & "- Line No Error Message: Customer code not defined in business partner master; Date Time : " & DateTime.Now & "")
                                            Crel = False
                                        End If

                                    End If
                                Next
                                If CRBool = True Then
                                    oGeneralService.Add(oGeneralData)
                                End If

                            Else
                                '===============================

                                CRBool = False
                                Dim oGeneralService1 As SAPbobsCOM.GeneralService
                                Dim oHeaderParams1 As SAPbobsCOM.GeneralDataParams
                                Dim oHeadTableRow1 As SAPbobsCOM.GeneralData
                                Dim sCmp1 As SAPbobsCOM.CompanyService = PublicVariable.oCompany.GetCompanyService
                                oGeneralService1 = sCmp1.GetGeneralService("CustRelation")
                                sCmp1 = PublicVariable.oCompany.GetCompanyService
                                oHeadTableRow1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                                ' Set the params for receiving a specific record
                                oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet3.DoQuery("SELECT T0.[U_AI_RelationCode] FROM [dbo].[@RELATION]  T0")
                                oHeaderParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                                'YOUR RECORD CODE

                                Dim oRecordSet_EC1 As SAPbobsCOM.Recordset

                                oRecordSet_EC1 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet_EC1.DoQuery("SELECT T0.[Code] FROM [dbo].[@RELATION]  T0 WHERE T0.[U_AI_RelationCode] ='" & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")


                                oHeaderParams1.SetProperty("Code", oRecordSet_EC1.Fields.Item(0).Value.ToString.Trim()) '


                                oHeadTableRow1.SetProperty("Code", (oRecordSet3.RecordCount + 1).ToString.Trim())
                                oHeadTableRow1.SetProperty("Name", oRecordSet_EC1.Fields.Item(0).Value.ToString.Trim())
                                oHeadTableRow1.SetProperty("U_AI_RelationCode", DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim())
                                oHeadTableRow1.SetProperty("U_AI_RelationName", DtSet_Rel_H.Tables(0).Rows(i).Item(1).ToString.Trim())
                                oHeadTableRow1.SetProperty("U_AI_RelationType", DtSet_Rel_H.Tables(0).Rows(i).Item(2).ToString.Trim())
                                Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                                Try
                                    If DtSet_Rel_H.Tables(0).Rows(i).Item(3).ToString.Trim() <> "" Then
                                        dt1 = DtSet_Rel_H.Tables(0).Rows(i).Item(3).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try

                                oHeadTableRow1.SetProperty("U_AI_DateAppointed", dt1)
                                oHeadTableRow1.SetProperty("U_AI_Status", DtSet_Rel_H.Tables(0).Rows(i).Item(4).ToString.Trim())
                                '---------------

                                Dim ST1 As String = "UPDATE [dbo].[@RELATION] SET [U_AI_RelationCode]='" & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() & "', [U_AI_RelationName]='" & DtSet_Rel_H.Tables(0).Rows(i).Item(1).ToString.Trim().Replace("'", " ") & "', [U_AI_RelationType]='" & DtSet_Rel_H.Tables(0).Rows(i).Item(2).ToString.Trim() & "', [U_AI_DateAppointed]='" & dt1 & "', [U_AI_Status]='" & DtSet_Rel_H.Tables(0).Rows(i).Item(4).ToString.Trim() & "' WHERE [Code] ='" & oRecordSet_EC1.Fields.Item(0).Value.ToString.Trim() & "'"
                                Dim oRecordSet_ECUP1 As SAPbobsCOM.Recordset
                                oRecordSet_ECUP1 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet_ECUP1.DoQuery(ST1)

                                oHeadTableRow1 = oGeneralService1.GetByParams(oHeaderParams1)
                                For j = 0 To rowcount_Rel_R - 1

                                    If DtSet_Rel_R.Tables(0).Rows(j).Item(0).ToString.Trim() = DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() Then


                                        Dim oRecordSetBPC As SAPbobsCOM.Recordset
                                        oRecordSetBPC = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSetBPC.DoQuery("SELECT CardCode  FROM OCRD  where CardCode='" & DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim() & "'")
                                        If oRecordSetBPC.RecordCount > 0 Then
                                            CRBool = True
                                            Dim oRecordSetBP1 As SAPbobsCOM.Recordset
                                            oRecordSetBP1 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRecordSetBP1.DoQuery("UPDATE OCRD set [U_AI_DateAppointed]='" & dt1 & "', [U_AI_RelationCode]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(0).ToString.Trim() & "', [U_AI_RelationName]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(1).ToString.Trim().Replace("'", " ") & "', [U_AI_RelationType]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(3).ToString.Trim() & "' WHERE [CardCode] ='" & DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim() & "'")



                                            Dim oRecordSet_EC As SAPbobsCOM.Recordset

                                            oRecordSet_EC = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRecordSet_EC.DoQuery("SELECT T1.[U_AI_EntityCode],T1.[LineId] FROM [dbo].[@RELATION]  T0 , [dbo].[@CUSTRELT]  T1 WHERE T0.[Code] = T1.[Code] and  T1.[U_AI_EntityCode] ='" & DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim() & "'")

                                            If oRecordSet_EC.RecordCount = 0 Then
                                                Dim oChildTableRows1 As SAPbobsCOM.GeneralDataCollection = oHeadTableRow1.Child("CUSTRELT")
                                                Dim oChildTableRow1 As SAPbobsCOM.GeneralData = oChildTableRows1.Add()
                                                oChildTableRow1.SetProperty("U_AI_RelationCode", DtSet_Rel_R.Tables(0).Rows(j).Item(0).ToString.Trim())
                                                oChildTableRow1.SetProperty("U_AI_RelationName", DtSet_Rel_R.Tables(0).Rows(j).Item(1).ToString.Trim())
                                                Dim dt As String = DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim()
                                                oChildTableRow1.SetProperty("U_AI_EntityCode", DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim())
                                                oChildTableRow1.SetProperty("U_AI_RelationType", DtSet_Rel_R.Tables(0).Rows(j).Item(3).ToString.Trim())
                                                Dim dt2 As String = Format(Now.Date, "yyyy-MM-dd")
                                                Try
                                                    If DtSet_Rel_R.Tables(0).Rows(j).Item(4).ToString.Trim() <> "" Then
                                                        dt2 = DtSet_Rel_R.Tables(0).Rows(j).Item(4).ToString.Trim()
                                                        dt2 = dt1.Insert(4, "-")
                                                        dt2 = dt1.Insert(7, "-")
                                                    End If
                                                    Dim dt3 As String = dt2.Remove(8, 1)
                                                    oChildTableRow1.SetProperty("U_AI_DateAppointed", dt3)
                                                Catch ex As Exception

                                                End Try

                                                ''MsgBox(dt2)

                                                oChildTableRow1.SetProperty("U_AI_Alternate", DtSet_Rel_R.Tables(0).Rows(j).Item(5).ToString.Trim())
                                                oChildTableRow1.SetProperty("U_AI_Industry", DtSet_Rel_R.Tables(0).Rows(j).Item(6).ToString.Trim())
                                                oChildTableRow1.SetProperty("U_AI_Country", DtSet_Rel_R.Tables(0).Rows(j).Item(7).ToString.Trim())
                                            Else

                                                Dim LineNum As Integer = oRecordSet_EC.Fields.Item(1).Value - 1
                                                Try
                                                    Dim oChildTableRows1 As SAPbobsCOM.GeneralDataCollection = oHeadTableRow1.Child("CUSTRELT")

                                                    Dim oChildTableRow1 As SAPbobsCOM.GeneralData = oChildTableRows1.Item(LineNum)

                                                    oChildTableRow1.SetProperty("U_AI_RelationCode", DtSet_Rel_R.Tables(0).Rows(j).Item(0).ToString.Trim())
                                                    oChildTableRow1.SetProperty("U_AI_RelationName", DtSet_Rel_R.Tables(0).Rows(j).Item(1).ToString.Trim())
                                                    Dim dt11 As String = DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim()
                                                    oChildTableRow1.SetProperty("U_AI_EntityCode", DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim())
                                                    oChildTableRow1.SetProperty("U_AI_RelationType", DtSet_Rel_R.Tables(0).Rows(j).Item(3).ToString.Trim())
                                                    Dim dt211 As String = Format(Now.Date, "yyyy-MM-dd")
                                                    Try


                                                        If DtSet_Rel_R.Tables(0).Rows(j).Item(4).ToString.Trim() <> "" Then
                                                            dt211 = DtSet_Rel_R.Tables(0).Rows(j).Item(4).ToString.Trim()
                                                            dt211 = dt1.Insert(4, "-")
                                                            dt211 = dt1.Insert(7, "-")
                                                        End If
                                                        ''MsgBox(dt2)
                                                        Dim dt311 As String = dt211.Remove(8, 1)
                                                        oChildTableRow1.SetProperty("U_AI_DateAppointed", dt311)
                                                    Catch ex As Exception

                                                    End Try
                                                    oChildTableRow1.SetProperty("U_AI_Alternate", DtSet_Rel_R.Tables(0).Rows(j).Item(5).ToString.Trim())
                                                    oChildTableRow1.SetProperty("U_AI_Industry", DtSet_Rel_R.Tables(0).Rows(j).Item(6).ToString.Trim())
                                                    oChildTableRow1.SetProperty("U_AI_Country", DtSet_Rel_R.Tables(0).Rows(j).Item(7).ToString.Trim())
                                                Catch ex As Exception

                                                End Try

                                                'Dim dt As String = DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim()
                                                'Dim dt2 As String = Format(Now.Date, "yyyy-MM-dd")
                                                'If DtSet_Rel_R.Tables(0).Rows(j).Item(4).ToString.Trim() <> "" Then
                                                '    dt2 = DtSet_Rel_R.Tables(0).Rows(j).Item(4).ToString.Trim()
                                                '    dt2 = dt1.Insert(4, "-")
                                                '    dt2 = dt1.Insert(7, "-")
                                                'End If
                                                ' ''MsgBox(dt2)
                                                'Dim dt3 As String = dt2.Remove(8, 1)
                                                'Dim ST As String = "UPDATE  [dbo].[@CUSTRELT] SET [U_AI_RelationCode]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(0).ToString.Trim() & "', [U_AI_RelationName]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(1).ToString.Trim() & "', [U_AI_EntityCode]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim() & "',[U_AI_RelationType]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(3).ToString.Trim() & "',[U_AI_DateAppointed]='" & dt3 & "', [U_AI_Alternate]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(5).ToString.Trim() & "',[U_AI_Industry]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(6).ToString.Trim() & "', [U_AI_Country]='" & DtSet_Rel_R.Tables(0).Rows(j).Item(7).ToString.Trim() & "' WHERE [Code] ='" & oRecordSet_EC1.Fields.Item(0).Value.ToString.Trim() & "' and  [LineId] =" & LineNum & ""
                                                'Dim oRecordSet_ECUP As SAPbobsCOM.Recordset
                                                'oRecordSet_ECUP = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                ''  oRecordSet_ECUP.DoQuery(ST)
                                                'Dim s As String = oRecordSet_EC1.Fields.Item(0).Value.ToString.Trim()
                                                '

                                                '=========
                                            End If
                                        Else
                                            file_CustRelation.WriteLine("Header " & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() & "-Relation Code Failed; Line No - " & i + 1 & ":Row " & DtSet_Rel_R.Tables(0).Rows(j).Item(2).ToString.Trim() & " BP- " & j & " - Line No Error Message: Customer code not defined in business partner master; Date Time : " & DateTime.Now & "")
                                            Crel = False
                                        End If
                                        'Else
                                        '    '======

                                        '    file_CustRelation.WriteLine(" " & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.trim() & "-Relation Code Failed; Line No - " & i + 1 & " Error Message: Entiry code alrady in the Relation Table; Date Time : " & DateTime.Now & "")
                                    End If
                                Next
                                If CRBool = True Then
                                    oGeneralService1.Update(oHeadTableRow1)
                                End If

                                '================================

                                'file_CustRelation.WriteLine(" " & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.trim() & "-Relation Code Failed; Line No - " & i + 1 & " Error Message:This Relation Code already Exists; Date Time : " & DateTime.Now & "")
                            End If
                        End If
                    Catch ex As Exception
                        file_CustRelation.WriteLine(" " & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() & "-Relation Code Failed; Line No - " & i + 1 & " Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
                        Crel = False
                    End Try

                    'Else
                    '    file_CustRelation.WriteLine(" " & DtSet_Rel_H.Tables(0).Rows(i).Item(0).ToString.Trim() & "-Relation Code Failed; Line No - " & i + 1 & " Error Message: Customer code not defined in business partner master; Date Time : " & DateTime.Now & "")
                    'End If

                End If

            Next

            Try
                file_CustRelation.Close()
            Catch ex As Exception

            End Try
        Catch ex As Exception
            file_CustRelation.WriteLine("Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
            Crel = False
            file_CustRelation.Close()
        End Try
    End Sub
#Region "Invoice"
    Private Sub Invoice1(ByVal FIH As String, ByVal FIR As String, ByVal csvFileFolder As String)
        Dim file_ARInv As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_ARInvoice_IH" & TDT & ".txt", True)
        Try


            Dim csvFileName As String = FIH.Replace(csvFileFolder, "")
            Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            Dim s1, s2, s3, s4, s5 As String
            s1 = "[" & csvFileName & "]"
            s2 = "ColNameHeader=True"
            s3 = "Format=CSVDelimited"
            s4 = "MaxScanRows=0"
            s5 = "CharacterSet=OEM"
            srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf + s5.ToString())
            srOutput.Close()
            fsOutput.Close()

            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""

            '  Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & csvFileFolder & ";Extended  Properties=Text;"

            Dim conn As New Odbc.OdbcConnection(connString)
            Dim da As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIH As System.Data.DataSet
            DtSet_FIH = New System.Data.DataSet
            da.Fill(DtSet_FIH)
            Dim rowcount_FIH As Integer = DtSet_FIH.Tables(0).Rows.Count
            '**********************************************************
            csvFileName = FIR.Replace(csvFileFolder, "")
            'connString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
            '            & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            'Dim conn1 As New Odbc.OdbcConnection(connString)
            Dim da1 As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIR As System.Data.DataSet
            DtSet_FIR = New System.Data.DataSet
            da1.Fill(DtSet_FIR)
            Dim rowcount_FIR As Integer = DtSet_FIR.Tables(0).Rows.Count


            Dim i, j As Integer
            Dim oINV As SAPbobsCOM.Documents

            Dim oRIN As SAPbobsCOM.Documents

            Dim oPRJ As SAPbobsCOM.Project

            Dim return_value As Integer
            Dim SerrorMsg As String = ""


            'file_ARInv.Close()
            Dim oRecordSet2 As SAPbobsCOM.Recordset
            Dim oRecordSet3 As SAPbobsCOM.Recordset
            Dim oRecordSet4 As SAPbobsCOM.Recordset
            Dim oRecordSet5 As SAPbobsCOM.Recordset
            Dim oRecordSet6 As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject
            For i = 0 To rowcount_FIH - 1
                'If i = 772 Then
                '    MsgBox("Hi..")
                'End If
                Try
                    oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet2.DoQuery("SELECT CardCode  FROM OCRD  where CardCode='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                    If oRecordSet2.RecordCount > 0 Then
                        oRecordSet6 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet6.DoQuery("SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[PrjCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                        Try
                            If oRecordSet6.RecordCount = 0 Then
                                Try
                                    oCmpSrv = PublicVariable.oCompany.GetCompanyService
                                    projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                                    project = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject)
                                    project.Code = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                    project.Name = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                    projectService.AddProject(project)
                                Catch ex As Exception
                                End Try
                            End If
                        Catch ex As Exception
                        End Try

                        If DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "A_CS" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "DR_N" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "A_GE" Then
                            oINV = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            oRecordSet5 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet5.DoQuery("SELECT T0.[NumAtCard] FROM OINV T0 WHERE T0.[NumAtCard] ='" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "'")
                            If oRecordSet5.RecordCount = 0 Then
                                oINV.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim()
                                Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                                Try
                                    If DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try

                                oINV.TaxDate = dt1
                                dt1 = Format(Now.Date, "yyyy-MM-dd")
                                Try
                                    If DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try

                                oINV.DocDate = dt1
                                oINV.CardCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oINV.NumAtCard = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim()
                                oINV.Project = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oINV.DocCurrency = DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.Trim()
                                Dim DocRt As Double = 0.0
                                Try
                                    DocRt = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString.Trim
                                Catch ex As Exception
                                End Try
                                Dim st As String = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString
                                oINV.DocRate = DocRt
                                oINV.Comments = DtSet_FIH.Tables(0).Rows(i).Item(7).ToString.Trim()
                                oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet3.DoQuery("SELECT T0.[U_AI_SalesType],T0.[U_AI_BVIP1Or2] FROM OCRD T0 WHERE T0.[CardCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                                If oRecordSet3.RecordCount > 0 Then
                                    Dim stest As String = oRecordSet3.Fields.Item(0).Value
                                    If stest = "N" Or stest = "E" Then
                                        oINV.UserFields.Fields.Item("U_AI_TypeofSales").Value = oRecordSet3.Fields.Item(0).Value
                                    End If
                                    oINV.UserFields.Fields.Item("U_AI_Half").Value = oRecordSet3.Fields.Item(1).Value ',T0.[U_AI_BVIP1Or2]
                                End If

                                'oINV.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim()


                                Dim Bol As Boolean = False
                                For j = 0 To rowcount_FIR - 1
                                    If DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() Then
                                        oRecordSet4 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet4.DoQuery("SELECT T0.[ItemCode] FROM OITM T0 WHERE T0.[ItemCode] ='" & DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim() & "'")
                                        If oRecordSet4.RecordCount > 0 Then
                                            Bol = True
                                            oINV.Lines.ItemCode = DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim()
                                            ' oINV.Lines.ItemDescription = DtSet_FIR.Tables(0).Rows(j).Item(6).ToString.Trim()
                                            'AI_ItemName
                                            Dim ItemName As String = DtSet_FIR.Tables(0).Rows(j).Item(6).ToString.Trim()
                                            If ItemName.Length > 100 Then
                                                oINV.Lines.ItemDescription = ItemName.Substring(0, 99)
                                                oINV.Lines.UserFields.Fields.Item("U_AI_ItemName").Value = ItemName
                                            Else
                                                oINV.Lines.ItemDescription = ItemName
                                                oINV.Lines.UserFields.Fields.Item("U_AI_ItemName").Value = ItemName
                                            End If
                                            oINV.Lines.Quantity = DtSet_FIR.Tables(0).Rows(j).Item(7).ToString.Trim()
                                            '  oINV.Lines.UnitPrice = DtSet_FIR.Tables(0).Rows(j).Item(9).ToString.trim().Replace("$", "")
                                            oINV.Lines.UserFields.Fields.Item("U_AI_LineNo").Value = DtSet_FIR.Tables(0).Rows(j).Item(11).ToString.Trim()
                                            oINV.Lines.LineTotal = DtSet_FIR.Tables(0).Rows(j).Item(12).ToString.Trim().Replace("$", "")
                                            oINV.Lines.ProjectCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                            oINV.Lines.Add()
                                        Else
                                            file_ARInv.WriteLine("" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & " - Invoice Header Line No " & i + 1 & " - " & DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() & "-DraftNo Failed; Invoice Row Line No - " & j + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                            IH = False
                                        End If
                                    End If
                                Next
                                'If Bol = False Then
                                '    'file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Invoic Header Line No - " & i + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                'Else
                                return_value = oINV.Add()
                                If return_value = 0 Then
                                Else
                                    PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                                    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "; Date Time : " & DateTime.Now & "")
                                    IH = False
                                End If
                                ' End If
                            Else
                                file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: This Invoice already Entered in SAP; Date Time : " & DateTime.Now & "")
                                IH = False
                            End If
                            '************************************
                        ElseIf DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "1CR_C" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "1CR_N" Then
                            oRIN = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            oRecordSet5 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet5.DoQuery("SELECT T0.[NumAtCard] FROM ORIN T0 WHERE T0.[NumAtCard] ='" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "'")
                            If oRecordSet5.RecordCount = 0 Then
                                oRIN.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim()
                                Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                                Try
                                    If DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try

                                oRIN.TaxDate = dt1
                                dt1 = Format(Now.Date, "yyyy-MM-dd")
                                If DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                    dt1 = DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim()
                                    dt1 = dt1.Insert(4, "-")
                                    dt1 = dt1.Insert(7, "-")
                                End If
                                oRIN.DocDate = dt1
                                oRIN.CardCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oRIN.NumAtCard = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim()

                                oRIN.DocCurrency = DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.Trim()

                                oRIN.DocRate = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString.Trim()
                                'oRIN.Comments = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.trim()
                                oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet3.DoQuery("SELECT T0.[U_AI_SalesType],T0.[U_AI_BVIP1Or2] FROM OCRD T0 WHERE T0.[CardCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                                If oRecordSet3.RecordCount > 0 Then
                                    Dim stest As String = oRecordSet3.Fields.Item(0).Value
                                    If stest = "N" Or stest = "E" Then
                                        oRIN.UserFields.Fields.Item("U_AI_TypeofSales").Value = oRecordSet3.Fields.Item(0).Value
                                    End If
                                    oRIN.UserFields.Fields.Item("U_AI_Half").Value = oRecordSet3.Fields.Item(1).Value ',T0.[U_AI_BVIP1Or2]
                                End If


                                Dim Bol As Boolean = False
                                For j = 0 To rowcount_FIR - 1
                                    If DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() Then
                                        oRecordSet4 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet4.DoQuery("SELECT T0.[ItemCode] FROM OITM T0 WHERE T0.[ItemCode] ='" & DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim() & "'")
                                        If oRecordSet4.RecordCount > 0 Then
                                            Bol = True
                                            oRIN.Lines.ItemCode = DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim()
                                            oRIN.Lines.Quantity = DtSet_FIR.Tables(0).Rows(j).Item(7).ToString.Trim()
                                            ' oRIN.Lines.UnitPrice = DtSet_FIR.Tables(0).Rows(j).Item(9).ToString.trim().Replace("$", "")
                                            oRIN.Lines.UserFields.Fields.Item("U_AI_LineNo").Value = DtSet_FIR.Tables(0).Rows(j).Item(11).ToString.Trim()
                                            oRIN.Lines.LineTotal = DtSet_FIR.Tables(0).Rows(j).Item(12).ToString.Trim().Replace("$", "")
                                            oRIN.Lines.ProjectCode = ""
                                            oRIN.Lines.Add()
                                        Else
                                            file_ARInv.WriteLine("" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & " - Invoice Header Line No " & i + 1 & " -  " & DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() & "-DraftNo Failed; Invoice Row Line No - " & j + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                            IH = False
                                        End If
                                    End If
                                Next
                                If Bol = False Then
                                    '  file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Invoic Header Line No - " & i + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                Else
                                    'return_value = oRIN.Add()
                                    If return_value = 0 Then
                                    Else
                                        PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                                        'file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "; Date Time : " & DateTime.Now & "")
                                    End If
                                End If
                            Else
                                file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: This Credit Memo already Entered in SAP; Date Time : " & DateTime.Now & "")
                                IH = False
                            End If

                            '**************************
                        Else
                            file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: Invoice Type Not Definied; Date Time : " & DateTime.Now & "")
                            IH = False
                        End If
                    Else
                        file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: Customer code not defined in business partner master; Date Time : " & DateTime.Now & "")
                        IH = False
                    End If
                Catch ex As Exception
                    ' MsgBox(ex.Message)
                    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
                    IH = False
                End Try

            Next
            Try
                file_ARInv.Close()
            Catch ex As Exception

            End Try


        Catch ex As Exception
            file_ARInv.WriteLine("Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
            IH = False
            file_ARInv.Close()
        End Try


    End Sub
    Private Function Invoice11(ByVal FIH As String, ByVal FIR As String, ByVal csvFileFolder As String) As Boolean
        Dim file_ARInv As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_ARInvoice_IH" & TDT & ".txt", True)
        Try
            Dim bo As Boolean = True
            ' PublicVariable.oCompany.StartTransaction()

            Dim csvFileName As String = FIH.Replace(csvFileFolder, "")
            Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            Dim s1, s2, s3, s4, s5 As String
            s1 = "[" & csvFileName & "]"
            s2 = "ColNameHeader=True"
            s3 = "Format=CSVDelimited"
            s4 = "MaxScanRows=0"
            s5 = "CharacterSet=OEM"
            srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf + s5.ToString())
            srOutput.Close()
            fsOutput.Close()

            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            Dim conn As New Odbc.OdbcConnection(connString)
            Dim da As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIH As System.Data.DataSet
            DtSet_FIH = New System.Data.DataSet
            da.Fill(DtSet_FIH)
            Dim rowcount_FIH As Integer = DtSet_FIH.Tables(0).Rows.Count
            '**********************************************************
            csvFileName = FIR.Replace(csvFileFolder, "")
            'connString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
            '            & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            'Dim conn1 As New Odbc.OdbcConnection(connString)
            Dim da1 As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIR As System.Data.DataSet
            DtSet_FIR = New System.Data.DataSet
            da1.Fill(DtSet_FIR)
            Dim rowcount_FIR As Integer = DtSet_FIR.Tables(0).Rows.Count


            Dim i, j As Integer
            Dim oINV As SAPbobsCOM.Documents

            Dim oRIN As SAPbobsCOM.Documents

            Dim oPRJ As SAPbobsCOM.Project

            Dim return_value As Integer
            Dim SerrorMsg As String = ""


            'file_ARInv.Close()
            Dim oRecordSet2 As SAPbobsCOM.Recordset
            Dim oRecordSet3 As SAPbobsCOM.Recordset
            Dim oRecordSet4 As SAPbobsCOM.Recordset
            Dim oRecordSet5 As SAPbobsCOM.Recordset
            Dim oRecordSet6 As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject
            For i = 0 To rowcount_FIH - 1
                oRecordSet6 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet6.DoQuery("SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[PrjCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                Try
                    If oRecordSet6.RecordCount = 0 Then
                        Try
                            oCmpSrv = PublicVariable.oCompany.GetCompanyService
                            projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                            project = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject)
                            project.Code = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                            project.Name = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                            projectService.AddProject(project)
                        Catch ex As Exception
                        End Try
                    End If
                Catch ex As Exception
                End Try

            Next
            PublicVariable.oCompany.StartTransaction()
            For i = 0 To rowcount_FIH - 1
                'If i = 772 Then
                '    MsgBox("Hi..")
                'End If
                Try
                    oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet2.DoQuery("SELECT CardCode  FROM OCRD  where CardCode='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                    If oRecordSet2.RecordCount > 0 Then
                        If DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "A_CS" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "DR_N" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "A_GE" Then
                            oINV = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            oRecordSet5 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet5.DoQuery("SELECT T0.[NumAtCard] FROM OINV T0 WHERE T0.[NumAtCard] ='" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "'")
                            If oRecordSet5.RecordCount = 0 Then
                                oINV.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim()
                                Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                                Try


                                    If DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try
                                oINV.TaxDate = dt1
                                dt1 = Format(Now.Date, "yyyy-MM-dd")
                                Try

                                    If DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If

                                Catch ex As Exception

                                End Try
                                oINV.DocDate = dt1
                                oINV.CardCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oINV.NumAtCard = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim()
                                oINV.Project = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oINV.DocCurrency = DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.Trim()
                                Dim DocRt As Double = 0.0
                                Try
                                    DocRt = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString.Trim
                                Catch ex As Exception
                                End Try

                                oINV.DocRate = DocRt
                                'oINV.Comments = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.trim()
                                oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet3.DoQuery("SELECT T0.[U_AI_SalesType],T0.[U_AI_BVIP1Or2] FROM OCRD T0 WHERE T0.[CardCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                                If oRecordSet3.RecordCount > 0 Then
                                    Dim stest As String = oRecordSet3.Fields.Item(0).Value
                                    If stest = "N" Or stest = "E" Then
                                        oINV.UserFields.Fields.Item("U_AI_TypeofSales").Value = oRecordSet3.Fields.Item(0).Value
                                    End If
                                    oINV.UserFields.Fields.Item("U_AI_Half").Value = oRecordSet3.Fields.Item(1).Value ',T0.[U_AI_BVIP1Or2]
                                End If

                                'oINV.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim()


                                Dim Bol As Boolean = False
                                For j = 0 To rowcount_FIR - 1
                                    If DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() Then
                                        oRecordSet4 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet4.DoQuery("SELECT T0.[ItemCode] FROM OITM T0 WHERE T0.[ItemCode] ='" & DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim() & "'")
                                        If oRecordSet4.RecordCount > 0 Then
                                            Bol = True
                                            oINV.Lines.ItemCode = DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim()
                                            'oINV.Lines.ItemDescription = DtSet_FIR.Tables(0).Rows(j).Item(6).ToString.Trim()
                                            Dim ItemName As String = DtSet_FIR.Tables(0).Rows(j).Item(6).ToString.Trim()
                                            If ItemName.Length > 100 Then
                                                oINV.Lines.ItemDescription = ItemName.Substring(0, 99)
                                                oINV.Lines.UserFields.Fields.Item("U_AI_ItemName").Value = ItemName
                                            Else
                                                oINV.Lines.ItemDescription = ItemName
                                                oINV.Lines.UserFields.Fields.Item("U_AI_ItemName").Value = ItemName
                                            End If
                                            oINV.Lines.Quantity = DtSet_FIR.Tables(0).Rows(j).Item(7).ToString.Trim()
                                            '  oINV.Lines.UnitPrice = DtSet_FIR.Tables(0).Rows(j).Item(9).ToString.trim().Replace("$", "")
                                            oINV.Lines.UserFields.Fields.Item("U_AI_LineNo").Value = DtSet_FIR.Tables(0).Rows(j).Item(11).ToString.Trim()
                                            oINV.Lines.LineTotal = DtSet_FIR.Tables(0).Rows(j).Item(12).ToString.Trim().Replace("$", "")
                                            oINV.Lines.ProjectCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                            oINV.Lines.Add()
                                        Else
                                            file_ARInv.WriteLine("" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & " - Invoice Header Line No " & i + 1 & " - " & DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() & "-DraftNo Failed; Invoic Row Line No - " & j + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                            IH = False
                                            bo = False
                                        End If
                                    End If
                                Next
                                'If Bol = False Then
                                '    'file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Invoic Header Line No - " & i + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                'Else
                                return_value = oINV.Add()
                                If return_value = 0 Then
                                Else
                                    PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                                    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "; Date Time : " & DateTime.Now & "")
                                    IH = False
                                    bo = False
                                    ' PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                                ' End If
                            Else

                                file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: This Invoice already Entered in SAP; Date Time : " & DateTime.Now & "")
                                IH = False
                                bo = False
                            End If
                            '************************************
                        ElseIf DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "1CR_C" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "1CR_N" Then
                            oRIN = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            oRecordSet5 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet5.DoQuery("SELECT T0.[NumAtCard] FROM ORIN T0 WHERE T0.[NumAtCard] ='" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "'")
                            If oRecordSet5.RecordCount = 0 Then
                                oRIN.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim()
                                Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                                Try



                                    If DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try
                                oRIN.TaxDate = dt1
                                dt1 = Format(Now.Date, "yyyy-MM-dd")
                                Try
                                    If DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try

                                oRIN.DocDate = dt1
                                oRIN.CardCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oRIN.NumAtCard = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim()

                                oRIN.DocCurrency = DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.Trim()
                                oRIN.DocRate = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString.Trim()
                                'oRIN.Comments = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.trim()
                                oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet3.DoQuery("SELECT T0.[U_AI_SalesType],T0.[U_AI_BVIP1Or2] FROM OCRD T0 WHERE T0.[CardCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                                If oRecordSet3.RecordCount > 0 Then
                                    Dim stest As String = oRecordSet3.Fields.Item(0).Value
                                    If stest = "N" Or stest = "E" Then
                                        oRIN.UserFields.Fields.Item("U_AI_TypeofSales").Value = oRecordSet3.Fields.Item(0).Value
                                    End If
                                    oRIN.UserFields.Fields.Item("U_AI_Half").Value = oRecordSet3.Fields.Item(1).Value ',T0.[U_AI_BVIP1Or2]
                                End If


                                Dim Bol As Boolean = False
                                For j = 0 To rowcount_FIR - 1
                                    If DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() Then
                                        oRecordSet4 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet4.DoQuery("SELECT T0.[ItemCode] FROM OITM T0 WHERE T0.[ItemCode] ='" & DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim() & "'")
                                        If oRecordSet4.RecordCount > 0 Then
                                            Bol = True
                                            oRIN.Lines.ItemCode = DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim()
                                            oRIN.Lines.Quantity = DtSet_FIR.Tables(0).Rows(j).Item(7).ToString.Trim()
                                            ' oRIN.Lines.UnitPrice = DtSet_FIR.Tables(0).Rows(j).Item(9).ToString.trim().Replace("$", "")
                                            oRIN.Lines.UserFields.Fields.Item("U_AI_LineNo").Value = DtSet_FIR.Tables(0).Rows(j).Item(11).ToString.Trim()
                                            oRIN.Lines.LineTotal = DtSet_FIR.Tables(0).Rows(j).Item(12).ToString.Trim().Replace("$", "")
                                            oRIN.Lines.ProjectCode = ""
                                            oRIN.Lines.Add()
                                        Else
                                            file_ARInv.WriteLine("" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & " - Invoice Header Line No " & i + 1 & " -  " & DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() & "-DraftNo Failed; Invoic Row Line No - " & j + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                            IH = False
                                        End If
                                    End If
                                Next
                                If Bol = False Then
                                    '  file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Invoic Header Line No - " & i + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                Else
                                    'return_value = oRIN.Add()
                                    If return_value = 0 Then
                                    Else
                                        PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                                        'file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "; Date Time : " & DateTime.Now & "")
                                    End If
                                End If
                            Else
                                file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: This Credit Memo already Entered in SAP; Date Time : " & DateTime.Now & "")
                                IH = False
                                bo = False
                            End If

                            '**************************
                        Else
                            file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: Invoice Type Not Definied; Date Time : " & DateTime.Now & "")
                            IH = False
                            bo = False
                        End If

                    Else
                        file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: Customer code not defined in business partner master; Date Time : " & DateTime.Now & "")
                        IH = False
                        bo = False
                    End If
                Catch ex As Exception
                    ' MsgBox(ex.Message)
                    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
                    IH = False
                    bo = False
                End Try

            Next
            Try
                file_ARInv.Close()
            Catch ex As Exception

            End Try
            If bo = False Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Else
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If


        Catch ex As Exception
            file_ARInv.WriteLine("Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
            IH = False
            file_ARInv.Close()
        End Try


    End Function
    Public Sub Import_CustomerCode(ByVal FIH As String, ByVal FIR As String, ByVal csvFileFolder As String)
        Dim csvFileName As String = FIH.Replace(csvFileFolder, "")
        Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
               & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
        Dim conn As New Odbc.OdbcConnection(connString)
        Dim da As New Odbc.OdbcDataAdapter("SELECT [Master File] FROM [" & csvFileName & "]", conn)
        Dim DtSet_FIH As System.Data.DataSet
        DtSet_FIH = New System.Data.DataSet
        da.Fill(DtSet_FIH)
        Dim rowcount_FIH As Integer = DtSet_FIH.Tables(0).Rows.Count
        '**********************************************************
        Dim rd As DataTableReader = DtSet_FIH.Tables(0).CreateDataReader()
        Dim connectionString As String = "Server=" & PublicVariable.SQLServer & ";Database=" & PublicVariable.SQLDB & ";User Id=" & PublicVariable.SQLUser & ";Password=" & PublicVariable.SQLPwd & ""
        'Dim Cn = New SqlConnection(connectionString)
        

        Using destinationConnection As SqlConnection = _
           New SqlConnection(connectionString)
            destinationConnection.Open()
            ' dataadapter.Fill(ds, "Authors_table")
            ''  destinationConnection.Open()
            Using copy As New SqlBulkCopy(destinationConnection)
                ' copy.ColumnMappings.Add(0, 0)
                'copy.ColumnMappings.Add(1, 1)
                'copy.ColumnMappings.Add(2, 2)
                'copy.ColumnMappings.Add(3, 3)
                'copy.ColumnMappings.Add(4, 4)
                copy.DestinationTableName = "ACUST"
                copy.WriteToServer(rd)
            End Using
            destinationConnection.Close()
        End Using
    End Sub
    'TETST
    Private Function Invoice21(ByVal FIH As String, ByVal FIR As String, ByVal csvFileFolder As String) As Boolean
        Dim file_ARInv As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_ARInvoice_SOH" & TDT & ".txt", True)
        Try
            Dim b1 As Boolean = True
            Dim csvFileName As String = FIH.Replace(csvFileFolder, "")
            Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            Dim s1, s2, s3, s4, s5 As String
            s1 = "[" & csvFileName & "]"
            s2 = "ColNameHeader=True"
            s3 = "Format=CSVDelimited"
            s4 = "MaxScanRows=0"
            s5 = "CharacterSet=OEM"
            srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf + s5.ToString())
            srOutput.Close()
            fsOutput.Close()

            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            Dim conn As New Odbc.OdbcConnection(connString)
            Dim da As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIH As System.Data.DataSet
            DtSet_FIH = New System.Data.DataSet
            da.Fill(DtSet_FIH)
            Dim rowcount_FIH As Integer = DtSet_FIH.Tables(0).Rows.Count
            '**********************************************************
            csvFileName = FIR.Replace(csvFileFolder, "")
            'connString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
            '            & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            'Dim conn1 As New Odbc.OdbcConnection(connString)
            Dim da1 As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIR As System.Data.DataSet
            DtSet_FIR = New System.Data.DataSet
            da1.Fill(DtSet_FIR)
            Dim rowcount_FIR As Integer = DtSet_FIR.Tables(0).Rows.Count


            Dim i, j As Integer
            Dim oINV As SAPbobsCOM.Documents

            Dim oRIN As SAPbobsCOM.Documents

            Dim oPRJ As SAPbobsCOM.Project

            Dim return_value As Integer
            Dim SerrorMsg As String = ""


            'file_ARInv.Close()
            Dim oRecordSet2 As SAPbobsCOM.Recordset
            Dim oRecordSet3 As SAPbobsCOM.Recordset
            Dim oRecordSet4 As SAPbobsCOM.Recordset
            Dim oRecordSet5 As SAPbobsCOM.Recordset
            Dim oRecordSet6 As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject

            For i = 0 To rowcount_FIH - 1
                oRecordSet6 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet6.DoQuery("SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[PrjCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                Try
                    If oRecordSet6.RecordCount = 0 Then
                        Try
                            oCmpSrv = PublicVariable.oCompany.GetCompanyService
                            projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                            project = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject)
                            project.Code = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                            project.Name = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                            projectService.AddProject(project)
                        Catch ex As Exception
                        End Try
                    End If
                Catch ex As Exception
                End Try
            Next
            PublicVariable.oCompany.StartTransaction()
            For i = 0 To rowcount_FIH - 1
                'If i = 772 Then
                '    MsgBox("Hi..")
                'End If
                Try

                    oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet2.DoQuery("SELECT CardCode  FROM OCRD  where CardCode='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                    If oRecordSet2.RecordCount > 0 Then

                        If DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "A_CS" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "DR_N" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "A_GE" Then
                            oINV = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            oRecordSet5 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet5.DoQuery("SELECT T0.[NumAtCard] FROM OINV T0 WHERE T0.[NumAtCard] ='" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "'")
                            If oRecordSet5.RecordCount = 0 Then
                                oINV.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim()
                                Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                                Try
                                    If DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try

                                oINV.TaxDate = dt1
                                dt1 = Format(Now.Date, "yyyy-MM-dd")
                                Try


                                    If DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try
                                oINV.DocDate = dt1
                                oINV.CardCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oINV.NumAtCard = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim()
                                oINV.Project = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oINV.DocCurrency = DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.Trim()
                                Dim DocRt As Double = 0.0
                                Try
                                    DocRt = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString.Trim
                                Catch ex As Exception
                                End Try
                                oINV.DocRate = DocRt
                                oINV.Comments = DtSet_FIH.Tables(0).Rows(i).Item(7).ToString.Trim()
                                oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet3.DoQuery("SELECT T0.[U_AI_SalesType],T0.[U_AI_BVIP1Or2] FROM OCRD T0 WHERE T0.[CardCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                                If oRecordSet3.RecordCount > 0 Then
                                    Dim stest As String = oRecordSet3.Fields.Item(0).Value
                                    If stest = "N" Or stest = "E" Then
                                        oINV.UserFields.Fields.Item("U_AI_TypeofSales").Value = oRecordSet3.Fields.Item(0).Value
                                    End If
                                    oINV.UserFields.Fields.Item("U_AI_Half").Value = oRecordSet3.Fields.Item(1).Value ',T0.[U_AI_BVIP1Or2]
                                End If

                                'oINV.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim()


                                Dim Bol As Boolean = False
                                For j = 0 To rowcount_FIR - 1
                                    If DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() Then
                                        oRecordSet4 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet4.DoQuery("SELECT T0.[ItemCode] FROM OITM T0 WHERE T0.[ItemCode] ='" & DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim() & "'")
                                        If oRecordSet4.RecordCount > 0 Then
                                            Bol = True
                                            oINV.Lines.ItemCode = DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim()
                                            'oINV.Lines.ItemDescription = DtSet_FIR.Tables(0).Rows(j).Item(6).ToString.Trim()
                                            Dim ItemName As String = DtSet_FIR.Tables(0).Rows(j).Item(6).ToString.Trim()
                                            If ItemName.Length > 100 Then

                                                oINV.Lines.ItemDescription = ItemName.Substring(0, 99)
                                                oINV.Lines.UserFields.Fields.Item("U_AI_ItemName").Value = ItemName
                                            Else
                                                oINV.Lines.ItemDescription = ItemName
                                                oINV.Lines.UserFields.Fields.Item("U_AI_ItemName").Value = ItemName
                                            End If
                                            oINV.Lines.Quantity = DtSet_FIR.Tables(0).Rows(j).Item(7).ToString.Trim()
                                            'oINV.Lines.UnitPrice = DtSet_FIR.Tables(0).Rows(j).Item(11).ToString.trim().Replace("$", "")
                                            oINV.Lines.UserFields.Fields.Item("U_AI_LineNo").Value = DtSet_FIR.Tables(0).Rows(j).Item(10).ToString.Trim()
                                            oINV.Lines.LineTotal = DtSet_FIR.Tables(0).Rows(j).Item(12).ToString.Trim().Replace("$", "")
                                            oINV.Lines.ProjectCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                            oINV.Lines.Add()
                                        Else
                                            file_ARInv.WriteLine("" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & " - Invoice Header Line No " & i + 1 & " - " & DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() & "-DraftNo Failed; Invoic Row Line No - " & j + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                            b1 = False
                                            SOH = False
                                        End If
                                    End If
                                Next
                                'If Bol = False Then
                                '    'file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Invoic Header Line No - " & i + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                'Else
                                return_value = oINV.Add()
                                If return_value = 0 Then
                                Else
                                    PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                                    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "; Date Time : " & DateTime.Now & "")
                                    SOH = False
                                    b1 = False
                                End If
                                ' End If
                            Else
                                file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: This Invoice already Entered in SAP; Date Time : " & DateTime.Now & "")
                                SOH = False
                                b1 = False
                            End If
                            '************************************
                        ElseIf DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "1CR_C" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "1CR_N" Then

                            'oRIN = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            'oRecordSet5 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRecordSet5.DoQuery("SELECT T0.[NumAtCard] FROM ORIN T0 WHERE T0.[NumAtCard] ='" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "'")
                            'If oRecordSet5.RecordCount = 0 Then
                            '    oRIN.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim()
                            '    Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                            '    If DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim() <> "" Then
                            '        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.trim()
                            '        dt1 = dt1.Insert(4, "-")
                            '        dt1 = dt1.Insert(7, "-")
                            '    End If
                            '    oRIN.TaxDate = dt1
                            '    dt1 = Format(Now.Date, "yyyy-MM-dd")
                            '    If DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim() <> "" Then
                            '        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.trim()
                            '        dt1 = dt1.Insert(4, "-")
                            '        dt1 = dt1.Insert(7, "-")
                            '    End If
                            '    oRIN.DocDate = dt1
                            '    oRIN.CardCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.trim()
                            '    oRIN.NumAtCard = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim()

                            '    oRIN.DocCurrency = DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.trim()
                            '    oRIN.DocRate = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString.trim()
                            '    'oRIN.Comments = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.trim()
                            '    oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRecordSet3.DoQuery("SELECT T0.[U_AI_SalesType],T0.[U_AI_BVIP1Or2] FROM OCRD T0 WHERE T0.[CardCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.trim() & "'")
                            '    If oRecordSet3.RecordCount > 0 Then
                            '        Dim stest As String = oRecordSet3.Fields.Item(0).Value
                            '        If stest = "N" Or stest = "E" Then
                            '            oRIN.UserFields.Fields.Item("U_AI_TypeofSales").Value = oRecordSet3.Fields.Item(0).Value
                            '        End If
                            '        oRIN.UserFields.Fields.Item("U_AI_Half").Value = oRecordSet3.Fields.Item(1).Value ',T0.[U_AI_BVIP1Or2]
                            '    End If


                            '    Dim Bol As Boolean = False
                            '    For j = 0 To rowcount_FIR - 1
                            '        If DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.trim() = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() Then
                            '            oRecordSet4 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '            oRecordSet4.DoQuery("SELECT T0.[ItemCode] FROM OITM T0 WHERE T0.[ItemCode] ='" & DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.trim() & "'")
                            '            If oRecordSet4.RecordCount > 0 Then
                            '                Bol = True
                            '                oRIN.Lines.ItemCode = DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.trim()
                            '                oRIN.Lines.Quantity = DtSet_FIR.Tables(0).Rows(j).Item(7).ToString.trim()
                            '                oRIN.Lines.UnitPrice = DtSet_FIR.Tables(0).Rows(j).Item(9).ToString.trim().Replace("$", "")
                            '                oRIN.Lines.UserFields.Fields.Item("U_AI_LineNo").Value = DtSet_FIR.Tables(0).Rows(j).Item(11).ToString.trim()
                            '                oRIN.Lines.LineTotal = DtSet_FIR.Tables(0).Rows(j).Item(12).ToString.trim().Replace("$", "")
                            '                oRIN.Lines.Add()
                            '            Else
                            '                file_ARInv.WriteLine("" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & " - Invoice Header Line No " & j + 1 & " -  " & DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.trim() & "-DraftNo Failed; Invoic Row Line No - " & j + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                            '            End If
                            '        End If
                            '    Next
                            '    If Bol = False Then
                            '        '  file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Invoic Header Line No - " & i + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                            '    Else
                            '        return_value = oRIN.Add()
                            '        If return_value = 0 Then
                            '        Else
                            '            PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                            '            file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "; Date Time : " & DateTime.Now & "")
                            '        End If
                            '    End If
                            'Else
                            '    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: This Credit Memo already Entered in SAP; Date Time : " & DateTime.Now & "")
                            'End If

                            '**************************
                        Else
                            file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: Invoice Type Not Definied; Date Time : " & DateTime.Now & "")
                            b1 = False
                            SOH = False
                        End If
                    Else
                        file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: Customer code not defined in business partner master; Date Time : " & DateTime.Now & "")
                        b1 = False
                        SOH = False
                    End If
                Catch ex As Exception
                    ' MsgBox(ex.Message)
                    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
                    b1 = False
                    SOH = False
                End Try

            Next
            Try
                file_ARInv.Close()
            Catch ex As Exception

            End Try
            If b1 = False Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If

        Catch ex As Exception
            file_ARInv.WriteLine("Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
            file_ARInv.Close()
            SOH = False

        End Try


    End Function
    Private Sub Invoice2(ByVal FIH As String, ByVal FIR As String, ByVal csvFileFolder As String)
        Dim file_ARInv As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_ARInvoice_SOH" & TDT & ".txt", True)
        Try
            'Dim conn As System.Data.OleDb.OleDbConnection
            'Dim ExcelConnectionStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + csvFileFolder + ";Extended Properties=""text;HDR=Yes;FMT=Delimited"""
            Dim csvFileName As String = FIH.Replace(csvFileFolder, "")
            Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            Dim s1, s2, s3, s4, s5 As String
            s1 = "[" & csvFileName & "]"
            s2 = "ColNameHeader=True"
            s3 = "Format=CSVDelimited"
            s4 = "MaxScanRows=0"
            s5 = "CharacterSet=OEM"
            srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf + s5.ToString())
            srOutput.Close()
            fsOutput.Close()

            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            Dim conn As New Odbc.OdbcConnection(connString)
            Dim da As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIH As System.Data.DataSet
            DtSet_FIH = New System.Data.DataSet
            da.Fill(DtSet_FIH)
            Dim rowcount_FIH As Integer = DtSet_FIH.Tables(0).Rows.Count
            '**********************************************************
            csvFileName = FIR.Replace(csvFileFolder, "")
            'connString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
            '            & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            'Dim conn1 As New Odbc.OdbcConnection(connString)
            Dim da1 As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_FIR As System.Data.DataSet
            DtSet_FIR = New System.Data.DataSet
            da1.Fill(DtSet_FIR)
            Dim rowcount_FIR As Integer = DtSet_FIR.Tables(0).Rows.Count


            Dim i, j As Integer
            Dim oINV As SAPbobsCOM.Documents

            Dim oRIN As SAPbobsCOM.Documents

            Dim oPRJ As SAPbobsCOM.Project

            Dim return_value As Integer
            Dim SerrorMsg As String = ""


            'file_ARInv.Close()
            Dim oRecordSet2 As SAPbobsCOM.Recordset
            Dim oRecordSet3 As SAPbobsCOM.Recordset
            Dim oRecordSet4 As SAPbobsCOM.Recordset
            Dim oRecordSet5 As SAPbobsCOM.Recordset
            Dim oRecordSet6 As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject
            For i = 0 To rowcount_FIH - 1
                'If i = 772 Then
                '    MsgBox("Hi..")
                'End If
                Try
                    oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet2.DoQuery("SELECT CardCode  FROM OCRD  where CardCode='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                    If oRecordSet2.RecordCount > 0 Then
                        oRecordSet6 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet6.DoQuery("SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[PrjCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                        Try
                            If oRecordSet6.RecordCount = 0 Then
                                Try
                                    oCmpSrv = PublicVariable.oCompany.GetCompanyService
                                    projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                                    project = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject)
                                    project.Code = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                    project.Name = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                    projectService.AddProject(project)
                                Catch ex As Exception
                                End Try
                            End If
                        Catch ex As Exception
                        End Try

                        If DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "A_CS" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "DR_N" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "A_GE" Then
                            oINV = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            oRecordSet5 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet5.DoQuery("SELECT T0.[NumAtCard] FROM OINV T0 WHERE T0.[NumAtCard] ='" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "'")
                            If oRecordSet5.RecordCount = 0 Then
                                oINV.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim()
                                Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                                Try


                                    If DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try
                                oINV.TaxDate = dt1
                                dt1 = Format(Now.Date, "yyyy-MM-dd")
                                Try


                                    If DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.Trim()
                                        dt1 = dt1.Insert(4, "-")
                                        dt1 = dt1.Insert(7, "-")
                                    End If
                                Catch ex As Exception

                                End Try
                                oINV.DocDate = dt1
                                oINV.CardCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oINV.NumAtCard = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim()
                                oINV.Project = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                oINV.DocCurrency = DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.Trim()
                                Dim DocRt As Double = 0.0
                                Try
                                    DocRt = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString.Trim
                                Catch ex As Exception
                                End Try
                                ' MsgBox(DtSet_FIH.Tables(0).Rows(i).Item(6).ToString)
                                oINV.DocRate = DocRt
                                oINV.Comments = DtSet_FIH.Tables(0).Rows(i).Item(7).ToString.Trim()
                                oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet3.DoQuery("SELECT T0.[U_AI_SalesType],T0.[U_AI_BVIP1Or2] FROM OCRD T0 WHERE T0.[CardCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim() & "'")
                                If oRecordSet3.RecordCount > 0 Then
                                    Dim stest As String = oRecordSet3.Fields.Item(0).Value
                                    If stest = "N" Or stest = "E" Then
                                        oINV.UserFields.Fields.Item("U_AI_TypeofSales").Value = oRecordSet3.Fields.Item(0).Value
                                    End If
                                    oINV.UserFields.Fields.Item("U_AI_Half").Value = oRecordSet3.Fields.Item(1).Value ',T0.[U_AI_BVIP1Or2]
                                End If

                                'oINV.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim()


                                Dim Bol As Boolean = False
                                For j = 0 To rowcount_FIR - 1
                                    If DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() Then
                                        oRecordSet4 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet4.DoQuery("SELECT T0.[ItemCode] FROM OITM T0 WHERE T0.[ItemCode] ='" & DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim() & "'")
                                        If oRecordSet4.RecordCount > 0 Then
                                            Bol = True
                                            oINV.Lines.ItemCode = DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.Trim()
                                            ' oINV.Lines.ItemDescription = DtSet_FIR.Tables(0).Rows(j).Item(6).ToString.Trim()
                                            Dim ItemName As String = DtSet_FIR.Tables(0).Rows(j).Item(6).ToString.Trim()


                                            'ItemName = dt2
                                            If ItemName.Length > 100 Then
                                                ' MsgBox(DtSet_FIR.Tables(0).Rows(j).Item(6).ToString)
                                                oINV.Lines.ItemDescription = ItemName.Substring(0, 99)
                                                oINV.Lines.UserFields.Fields.Item("U_AI_ItemName").Value = ItemName
                                                'oINV.Lines.UserFields.Fields.Item("U_AI_ItemName1").Value = ItemName
                                            Else
                                                oINV.Lines.ItemDescription = ItemName
                                                oINV.Lines.UserFields.Fields.Item("U_AI_ItemName").Value = ItemName
                                                ' oINV.Lines.UserFields.Fields.Item("U_AI_ItemName1").Value = ItemName
                                            End If
                                            oINV.Lines.Quantity = DtSet_FIR.Tables(0).Rows(j).Item(7).ToString.Trim()
                                            'oINV.Lines.UnitPrice = DtSet_FIR.Tables(0).Rows(j).Item(11).ToString.trim().Replace("$", "")
                                            oINV.Lines.UserFields.Fields.Item("U_AI_LineNo").Value = DtSet_FIR.Tables(0).Rows(j).Item(10).ToString.Trim()
                                            oINV.Lines.LineTotal = DtSet_FIR.Tables(0).Rows(j).Item(12).ToString.Trim().Replace("$", "")
                                            oINV.Lines.ProjectCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.Trim()
                                            oINV.Lines.Add()
                                        Else
                                            file_ARInv.WriteLine("" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & " - Invoice Header Line No " & i + 1 & " - " & DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.Trim() & "-DraftNo Failed; Invoic Row Line No - " & j + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                            SOH = False
                                        End If
                                    End If
                                Next
                                'If Bol = False Then
                                '    'file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Invoic Header Line No - " & i + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                                'Else
                                return_value = oINV.Add()
                                If return_value = 0 Then
                                Else
                                    PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                                    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "; Date Time : " & DateTime.Now & "")
                                    SOH = False
                                End If
                                ' End If
                            Else
                                file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: This Invoice already Entered in SAP; Date Time : " & DateTime.Now & "")
                                SOH = False
                            End If
                            '************************************
                        ElseIf DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "1CR_C" Or DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.Trim() = "1CR_N" Then

                            'oRIN = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            'oRecordSet5 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRecordSet5.DoQuery("SELECT T0.[NumAtCard] FROM ORIN T0 WHERE T0.[NumAtCard] ='" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "'")
                            'If oRecordSet5.RecordCount = 0 Then
                            '    oRIN.UserFields.Fields.Item("U_AI_Type").Value = DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim()
                            '    Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                            '    If DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim() <> "" Then
                            '        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(1).ToString.trim()
                            '        dt1 = dt1.Insert(4, "-")
                            '        dt1 = dt1.Insert(7, "-")
                            '    End If
                            '    oRIN.TaxDate = dt1
                            '    dt1 = Format(Now.Date, "yyyy-MM-dd")
                            '    If DtSet_FIH.Tables(0).Rows(i).Item(0).ToString.trim() <> "" Then
                            '        dt1 = DtSet_FIH.Tables(0).Rows(i).Item(2).ToString.trim()
                            '        dt1 = dt1.Insert(4, "-")
                            '        dt1 = dt1.Insert(7, "-")
                            '    End If
                            '    oRIN.DocDate = dt1
                            '    oRIN.CardCode = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.trim()
                            '    oRIN.NumAtCard = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim()

                            '    oRIN.DocCurrency = DtSet_FIH.Tables(0).Rows(i).Item(5).ToString.trim()
                            '    oRIN.DocRate = DtSet_FIH.Tables(0).Rows(i).Item(6).ToString.trim()
                            '    'oRIN.Comments = DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.trim()
                            '    oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRecordSet3.DoQuery("SELECT T0.[U_AI_SalesType],T0.[U_AI_BVIP1Or2] FROM OCRD T0 WHERE T0.[CardCode] ='" & DtSet_FIH.Tables(0).Rows(i).Item(3).ToString.trim() & "'")
                            '    If oRecordSet3.RecordCount > 0 Then
                            '        Dim stest As String = oRecordSet3.Fields.Item(0).Value
                            '        If stest = "N" Or stest = "E" Then
                            '            oRIN.UserFields.Fields.Item("U_AI_TypeofSales").Value = oRecordSet3.Fields.Item(0).Value
                            '        End If
                            '        oRIN.UserFields.Fields.Item("U_AI_Half").Value = oRecordSet3.Fields.Item(1).Value ',T0.[U_AI_BVIP1Or2]
                            '    End If


                            '    Dim Bol As Boolean = False
                            '    For j = 0 To rowcount_FIR - 1
                            '        If DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.trim() = DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() Then
                            '            oRecordSet4 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '            oRecordSet4.DoQuery("SELECT T0.[ItemCode] FROM OITM T0 WHERE T0.[ItemCode] ='" & DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.trim() & "'")
                            '            If oRecordSet4.RecordCount > 0 Then
                            '                Bol = True
                            '                oRIN.Lines.ItemCode = DtSet_FIR.Tables(0).Rows(j).Item(5).ToString.trim()
                            '                oRIN.Lines.Quantity = DtSet_FIR.Tables(0).Rows(j).Item(7).ToString.trim()
                            '                oRIN.Lines.UnitPrice = DtSet_FIR.Tables(0).Rows(j).Item(9).ToString.trim().Replace("$", "")
                            '                oRIN.Lines.UserFields.Fields.Item("U_AI_LineNo").Value = DtSet_FIR.Tables(0).Rows(j).Item(11).ToString.trim()
                            '                oRIN.Lines.LineTotal = DtSet_FIR.Tables(0).Rows(j).Item(12).ToString.trim().Replace("$", "")
                            '                oRIN.Lines.Add()
                            '            Else
                            '                file_ARInv.WriteLine("" & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & " - Invoice Header Line No " & j + 1 & " -  " & DtSet_FIR.Tables(0).Rows(j).Item(2).ToString.trim() & "-DraftNo Failed; Invoic Row Line No - " & j + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                            '            End If
                            '        End If
                            '    Next
                            '    If Bol = False Then
                            '        '  file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Invoic Header Line No - " & i + 1 & " Error Message: Item Code Not Find in SAP; Date Time : " & DateTime.Now & "")
                            '    Else
                            '        return_value = oRIN.Add()
                            '        If return_value = 0 Then
                            '        Else
                            '            PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                            '            file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "; Date Time : " & DateTime.Now & "")
                            '        End If
                            '    End If
                            'Else
                            '    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: This Credit Memo already Entered in SAP; Date Time : " & DateTime.Now & "")
                            'End If

                            '**************************
                        Else
                            file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: Invoice Type Not Definied; Date Time : " & DateTime.Now & "")
                            SOH = False
                        End If
                    Else
                        file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: Customer code not defined in business partner master; Date Time : " & DateTime.Now & "")
                        SOH = False
                    End If
                Catch ex As Exception
                    ' MsgBox(ex.Message)
                    file_ARInv.WriteLine(" " & DtSet_FIH.Tables(0).Rows(i).Item(4).ToString.Trim() & "-DraftNo Failed; Line No - " & i + 1 & " Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
                    SOH = False
                End Try

            Next
            Try
                file_ARInv.Close()
            Catch ex As Exception

            End Try


        Catch ex As Exception
            file_ARInv.WriteLine("Error Message: " & ex.Message & "; Date Time : " & DateTime.Now & "")
            file_ARInv.Close()
            SOH = False
        End Try


    End Sub
#End Region
#Region "Customer"
    Private Sub FCustSalesPer(ByVal FCustSaels As String, ByVal csvFileFolder As String)
        'Dim file As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath  & "ErrorLog.txt", True)
        Dim file_Sales As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_Sales" & TDT & ".txt", True)
        Try

            '   file_Sales.Close()
        Catch ex As Exception

        End Try

        Try


            Dim csvFileName As String = FCustSaels.Replace(csvFileFolder, "")
            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""

            Dim conn As New Odbc.OdbcConnection(connString)
            Dim da As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_Sales1 As System.Data.DataSet
            DtSet_Sales1 = New System.Data.DataSet
            da.Fill(DtSet_Sales1)
            Dim rd As DataTableReader = DtSet_Sales1.Tables(0).CreateDataReader()
            '**********************
            Dim oRecordSetBP1 As SAPbobsCOM.Recordset
            oRecordSetBP1 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSetBP1.DoQuery("Delete from TEMP2_SPST")
            Dim connectionString As String = "Server=" & PublicVariable.SQLServer & ";Database=" & PublicVariable.SQLDB & ";User Id=" & PublicVariable.SQLUser & ";Password=" & PublicVariable.SQLPwd & ""
            Using destinationConnection As SqlConnection = _
               New SqlConnection(connectionString)
                destinationConnection.Open()
                Using copy As New SqlBulkCopy(destinationConnection)
                    copy.DestinationTableName = "TEMP2_SPST"
                    copy.WriteToServer(rd)
                End Using
                destinationConnection.Close()
            End Using
            Dim DtSet_Sales As System.Data.DataSet
            DtSet_Sales = New System.Data.DataSet
            Dim connection As New SqlConnection(connectionString)
            Dim sql As String = "execute [dbo].[AB_CustomerMaster_SALES]"
            'Dim dataadapter As New SqlDataAdapter(sql, connection)
            'connection.Open()
            'dataadapter.Fill(DtSet_Sales)
            'connection.Close()
            Dim SQLCMD As SqlCommand
            SQLCMD = New SqlCommand(sql, connection)
            SQLCMD.CommandTimeout = 600
            Dim reader As SqlDataReader
            Try
                connection.Open()
                reader = SQLCMD.ExecuteReader
                reader.Read()
                DtSet_Sales.Load(reader, LoadOption.OverwriteChanges, "TAB1")
                connection.Close()
            Catch ex As Exception
                file_Sales.WriteLine(" Error Message: " & ex.Message & "Date Time : " & DateTime.Now & "")
                Exit Sub
            End Try

            '**************************
            Dim rowcount_Sales As Integer = DtSet_Sales.Tables(0).Rows.Count
            If rowcount_Sales = 0 Then

                ' file.WriteLine("No Data in Customer Address CSV File- Date Time : " & DateTime.Now & "")
            End If

            Dim i As Integer
            Dim oCRD As SAPbobsCOM.BusinessPartners
            oCRD = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            Dim return_value As Integer
            Dim SerrorMsg As String = ""

            For i = 0 To rowcount_Sales - 1
                Try

                        Dim oRecordSet2 As SAPbobsCOM.Recordset
                        Dim oRecordSet3 As SAPbobsCOM.Recordset
                        oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet2.DoQuery("SELECT CardCode  FROM OCRD  where CardCode = '" & DtSet_Sales.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            'oCRD.Addresses.Add()

                            oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet3.DoQuery("SELECT T0.[SlpCode], T0.[SlpName] FROM OSLP T0 WHERE T0.[SlpName] like '" & DtSet_Sales.Tables(0).Rows(i).Item(1).ToString.Trim() & "%'")

                            If oRecordSet3.RecordCount > 0 Then
                                oCRD.GetByKey(DtSet_Sales.Tables(0).Rows(i).Item(0).ToString.Trim())
                                If oRecordSet3.RecordCount > 0 Then
                                    oCRD.SalesPersonCode = oRecordSet3.Fields.Item(0).Value
                                Else
                                    oCRD.SalesPersonCode = -1

                                End If
                                Dim st As String = DtSet_Sales.Tables(0).Rows(i).Item(3).ToString.Trim()
                                If DtSet_Sales.Tables(0).Rows(i).Item(3).ToString.Trim() <> "" Then
                                    oCRD.UserFields.Fields.Item("U_AI_SalesType").Value = DtSet_Sales.Tables(0).Rows(i).Item(3).ToString.Trim()
                                End If

                                return_value = oCRD.Update()

                                If return_value = 0 Then
                                    'file_addr.WriteLine(" " & DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.trim() & "-BP Update Success; Line No - " & i + 1 & " Date Time : " & DateTime.Now & "")
                                Else
                                    PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                                    file_Sales.WriteLine(" " & DtSet_Sales.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & " Date Time : " & DateTime.Now & "")
                                    CSales = False
                                End If
                            Else
                                file_Sales.WriteLine(" " & DtSet_Sales.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: Sales Employee Not defined in SAP Date Time : " & DateTime.Now & "")
                                CSales = False
                            End If
                        Else
                            file_Sales.WriteLine(" " & DtSet_Sales.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: EntityCode not exist in SAP Customer Master Date Time : " & DateTime.Now & "")
                            CSales = False
                        End If


                Catch ex As Exception
                    '  MsgBox(ex.Message)
                    file_Sales.WriteLine(" " & DtSet_Sales.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: " & ex.Message & " Date Time : " & DateTime.Now & "")
                    CSales = False
                End Try
            Next
            file_Sales.Close()
        Catch ex As Exception
            file_Sales.WriteLine("Error Message: " & ex.Message & " Date Time : " & DateTime.Now & "")
            file_Sales.Close()
            CSales = False
        End Try
    End Sub
    Private Sub FCustaddress(ByVal FCustaddr As String, ByVal csvFileFolder As String)
        Dim file_addr As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_Addr" & TDT & ".txt", True)
        Try


            'Dim file As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath  & "ErrorLog.txt", True)
            Dim csvFileName As String = FCustaddr.Replace(csvFileFolder, "")
            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""

            Dim conn As New Odbc.OdbcConnection(connString)
            Dim da As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_Adrr1 As System.Data.DataSet
            DtSet_Adrr1 = New System.Data.DataSet
            da.Fill(DtSet_Adrr1)
            Dim rd As DataTableReader = DtSet_Adrr1.Tables(0).CreateDataReader()


           


            '*******************
            Dim oRecordSetBP1 As SAPbobsCOM.Recordset
            oRecordSetBP1 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSetBP1.DoQuery("Delete from TEMP2_ADDRESS")
            Dim connectionString As String = "Server=" & PublicVariable.SQLServer & ";Database=" & PublicVariable.SQLDB & ";User Id=" & PublicVariable.SQLUser & ";Password=" & PublicVariable.SQLPwd & ""
            Using destinationConnection As SqlConnection = _
               New SqlConnection(connectionString)
                destinationConnection.Open()
                Using copy As New SqlBulkCopy(destinationConnection)
                    copy.DestinationTableName = "TEMP2_ADDRESS"
                    copy.WriteToServer(rd)
                End Using
                destinationConnection.Close()
            End Using
            Dim DtSet_Adrr As System.Data.DataSet
            DtSet_Adrr = New System.Data.DataSet
            Dim connection As New SqlConnection(connectionString)
            Dim sql As String = "execute [dbo].[AB_CustomerMaster_ADDR]"
            'Dim dataadapter As New SqlDataAdapter(sql, connection)
            'connection.Open()
            'dataadapter.Fill(DtSet_Adrr)
            'connection.Close()
            Dim SQLCMD As SqlCommand
            SQLCMD = New SqlCommand(sql, connection)
            SQLCMD.CommandTimeout = 600
            Dim reader As SqlDataReader
            Try
                connection.Open()
                reader = SQLCMD.ExecuteReader
                reader.Read()

                DtSet_Adrr.Load(reader, LoadOption.OverwriteChanges, "TAB1")
                connection.Close()
                'connection.Open()


                'dataadapter.Fill(DtSet_Master)
                'connection.Close()
            Catch ex As Exception
                file_addr.WriteLine(" Error Message: " & ex.Message & "Date Time : " & DateTime.Now & "")
                Exit Sub
            End Try
            '**********************

            Dim rowcount_addr As Integer = DtSet_Adrr.Tables(0).Rows.Count
            If rowcount_addr = 0 Then
                ' file.WriteLine("No Data in Customer Address CSV File- Date Time : " & DateTime.Now & "")
            End If

            Dim i As Integer
            Dim oCRD As SAPbobsCOM.BusinessPartners
            oCRD = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            Dim return_value As Integer
            Dim SerrorMsg As String = ""

            For i = 0 To rowcount_addr - 1
                Try

                  
                        Dim oRecordSet3 As SAPbobsCOM.Recordset
                        Dim oRecordSet2 As SAPbobsCOM.Recordset
                        oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet2.DoQuery("SELECT CardCode  FROM OCRD  where CardCode='" & DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oCRD.GetByKey(DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.Trim())

                            'oCRD.Addresses.Add()
                            oRecordSet3 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet3.DoQuery("SELECT count(*) as count1 FROM CRD1 T0 WHERE T0.[AdresType]='B' and  T0.[CardCode] ='" & DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")

                            oCRD.Addresses.Add()
                            oCRD.Addresses.SetCurrentLine(oRecordSet3.Fields.Item(0).Value)
                            If DtSet_Adrr.Tables(0).Rows(i).Item(1).ToString.Trim() <> "" Then
                                oCRD.Addresses.UserFields.Fields.Item("U_AI_Type").Value = DtSet_Adrr.Tables(0).Rows(i).Item(1).ToString.Trim()
                            End If
                            oCRD.Addresses.AddressName = DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.Trim() & "-" & (oRecordSet3.Fields.Item(0).Value + 1)
                            oCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
                            If DtSet_Adrr.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                oCRD.Addresses.Street = DtSet_Adrr.Tables(0).Rows(i).Item(2).ToString.Trim()
                            End If
                            If DtSet_Adrr.Tables(0).Rows(i).Item(3).ToString.Trim() <> "" Then
                                oCRD.Addresses.Block = DtSet_Adrr.Tables(0).Rows(i).Item(3).ToString.Trim()
                            End If
                            If DtSet_Adrr.Tables(0).Rows(i).Item(4).ToString.Trim() <> "" Then
                                oCRD.Addresses.County = DtSet_Adrr.Tables(0).Rows(i).Item(4).ToString.Trim()
                            End If
                            If DtSet_Adrr.Tables(0).Rows(i).Item(5).ToString.Trim() <> "" Then
                                oCRD.Addresses.StreetNo = DtSet_Adrr.Tables(0).Rows(i).Item(5).ToString.Trim()
                            End If
                            If DtSet_Adrr.Tables(0).Rows(i).Item(6).ToString.Trim() <> "" Then
                                oCRD.Addresses.BuildingFloorRoom = DtSet_Adrr.Tables(0).Rows(i).Item(6).ToString.Trim()
                            End If
                            If DtSet_Adrr.Tables(0).Rows(i).Item(7).ToString.Trim() <> "" Then
                                oCRD.Addresses.City = DtSet_Adrr.Tables(0).Rows(i).Item(7).ToString.Trim()
                            End If
                            oCRD.Addresses.UserFields.Fields.Item("U_AI_CityNo").Value = DtSet_Adrr.Tables(0).Rows(i).Item(8).ToString.Trim()
                            If DtSet_Adrr.Tables(0).Rows(i).Item(9).ToString.Trim() <> "" Then
                                'puni
                                Dim oRecordSetBP As SAPbobsCOM.Recordset
                                oRecordSetBP = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSetBP.DoQuery("SELECT T0.[Code], T0.[Name] FROM OCRY T0 WHERE T0.[Code]='" & DtSet_Adrr.Tables(0).Rows(i).Item(9).ToString.Trim() & "'")
                                If oRecordSetBP.RecordCount > 0 Then
                                    oCRD.Addresses.Country = DtSet_Adrr.Tables(0).Rows(i).Item(9).ToString.Trim()
                                Else
                                    file_addr.WriteLine(" " & DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: Country code Not exist in SAP  Date Time : " & DateTime.Now & "")
                                    CAddr = False
                                    Exit Try
                                End If

                            End If
                            'If DtSet_Adrr.Tables(0).Rows(i).Item(11).ToString.trim() <> "" Then
                            '    Dim oRecordSetST As SAPbobsCOM.Recordset
                            '    oRecordSetST = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRecordSetST.DoQuery("SELECT T0.[Name] FROM OCST T0 WHERE T0.[Name] ='" & DtSet_Adrr.Tables(0).Rows(i).Item(11).ToString.Trim() & "' and  T0.[Country] ='" & DtSet_Adrr.Tables(0).Rows(i).Item(9).ToString.Trim() & "'")
                            '    If oRecordSetST.RecordCount > 0 Then
                            '        oCRD.Addresses.State = DtSet_Adrr.Tables(0).Rows(i).Item(11).ToString.trim()
                            '    Else
                            '        file_addr.WriteLine(" " & DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: State Name Not exist in SAP  Date Time : " & DateTime.Now & "")
                            '        Exit Try
                            '    End If



                            'End If
                            oCRD.Addresses.UserFields.Fields.Item("U_AI_StateName").Value = DtSet_Adrr.Tables(0).Rows(i).Item(11).ToString.Trim()
                            If DtSet_Adrr.Tables(0).Rows(i).Item(13).ToString.Trim() <> "" Then
                                oCRD.Addresses.UserFields.Fields.Item("U_AI_AddrNo").Value = DtSet_Adrr.Tables(0).Rows(i).Item(13).ToString.Trim()
                            End If
                            If DtSet_Adrr.Tables(0).Rows(i).Item(14).ToString.Trim() <> "" Then
                                If DtSet_Adrr.Tables(0).Rows(i).Item(14).ToString.Trim().Length = 5 And DtSet_Adrr.Tables(0).Rows(i).Item(9).ToString.Trim() = "SG" Then
                                    oCRD.Addresses.ZipCode = "0" & DtSet_Adrr.Tables(0).Rows(i).Item(14).ToString.Trim()
                                Else
                                    oCRD.Addresses.ZipCode = DtSet_Adrr.Tables(0).Rows(i).Item(14).ToString.Trim()
                                End If

                            End If
                            oCRD.BilltoDefault = DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.Trim() & "-" & (oRecordSet3.Fields.Item(0).Value + 1)
                            oCRD.Addresses.UserFields.Fields.Item("U_AI_Desc").Value = DtSet_Adrr.Tables(0).Rows(i).Item(15).ToString.Trim()

                            return_value = oCRD.Update()
                            If return_value = 0 Then
                                'file_addr.WriteLine(" " & DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.trim() & "-BP Update Success; Line No - " & i + 1 & " Date Time : " & DateTime.Now & "")
                            Else
                                PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                                file_addr.WriteLine(" " & DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & " Date Time : " & DateTime.Now & "")
                                CAddr = False
                            End If
                        Else
                            '     file_addr.WriteLine(" " & DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: EntityCode not exist in SAP Customer Master Date Time : " & DateTime.Now & "")
                        End If


                Catch ex As Exception
                    'MsgBox(ex.Message)
                    file_addr.WriteLine(" " & DtSet_Adrr.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: " & ex.Message & " Date Time : " & DateTime.Now & "")
                    CAddr = False
                End Try
            Next
            file_addr.Close()
        Catch ex As Exception
            file_addr.WriteLine("Error Message: " & ex.Message & " Date Time : " & DateTime.Now & "")
            file_addr.Close()
            CAddr = False
        End Try
    End Sub
    Private Sub CustomerName(ByVal FCustname As String, ByVal csvFileFolder As String)
        'Dim file As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog.txt", True)
        Dim file_Name As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_Name" & TDT & ".txt", True)
        Try


            Dim csvFileName As String = FCustname.Replace(csvFileFolder, "")
            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                   & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            Dim conn As New Odbc.OdbcConnection(connString)
            Dim da1 As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            Dim DtSet_Name As System.Data.DataSet
            DtSet_Name = New System.Data.DataSet
            da1.Fill(DtSet_Name)
            Dim rowcount_Name As Integer = DtSet_Name.Tables(0).Rows.Count
            If rowcount_Name = 0 Then
                ' file.WriteLine("No Data in Customer Name CSV File- Date Time : " & DateTime.Now & "")
            End If
            Dim i As Integer
            Dim oCRD As SAPbobsCOM.BusinessPartners
            oCRD = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            Dim return_value As Integer
            Dim SerrorMsg As String = ""

            For i = 0 To rowcount_Name - 1
                Try
                    Dim oRecordSet2 As SAPbobsCOM.Recordset
                    Dim oRecordSetBP1 As SAPbobsCOM.Recordset
                    oRecordSetBP1 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSetBP1.DoQuery("Select CardCode from ACUST where CardCode='" & DtSet_Name.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                    If oRecordSetBP1.RecordCount > 0 Then
                        oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet2.DoQuery("SELECT CardCode  FROM OCRD  where CardCode='" & DtSet_Name.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oCRD.GetByKey(DtSet_Name.Tables(0).Rows(i).Item(0).ToString.Trim())
                            oCRD.UserFields.Fields.Item("U_AI_CustNameStatus").Value = DtSet_Name.Tables(0).Rows(i).Item(1).ToString.Trim()
                            Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                            Try


                                If DtSet_Name.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                    dt1 = DtSet_Name.Tables(0).Rows(i).Item(2).ToString.Trim()
                                    dt1 = dt1.Insert(4, "-")
                                    dt1 = dt1.Insert(7, "-")
                                End If
                            Catch ex As Exception

                            End Try
                            If DtSet_Name.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                oCRD.UserFields.Fields.Item("U_AI_DateApplied").Value = dt1
                            End If

                            Try
                                If DtSet_Name.Tables(0).Rows(i).Item(0).ToString.Trim() <> "" Then
                                    dt1 = DtSet_Name.Tables(0).Rows(i).Item(3).ToString.Trim()
                                    dt1 = dt1.Insert(4, "-")
                                    dt1 = dt1.Insert(7, "-")
                                End If
                            Catch ex As Exception

                            End Try
                            If DtSet_Name.Tables(0).Rows(i).Item(3).ToString.Trim() <> "" Then
                                oCRD.UserFields.Fields.Item("U_AI_EffectiveDate").Value = dt1
                            End If


                            If DtSet_Name.Tables(0).Rows(i).Item(5).ToString.Trim() <> "" Then
                                Dim Cardname As String = DtSet_Name.Tables(0).Rows(i).Item(5).ToString.Trim()
                                If Cardname.Length > 100 Then
                                    oCRD.CardName = Cardname.Substring(0, 99)
                                    oCRD.UserFields.Fields.Item("U_AI_CardName").Value = Cardname
                                Else
                                    oCRD.CardName = Cardname
                                    oCRD.UserFields.Fields.Item("U_AI_CardName").Value = Cardname
                                End If

                            End If
                            return_value = oCRD.Update()
                            If return_value = 0 Then
                                ' file_Name.WriteLine(" " & DtSet_Name.Tables(0).Rows(i).Item(0).ToString.trim() & "-BP Update Success; Line No - " & i + 1 & " Date Time : " & DateTime.Now & "")
                            Else
                                PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                                file_Name.WriteLine(" " & DtSet_Name.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & " Date Time : " & DateTime.Now & "")
                                CName = False
                            End If
                        Else
                            file_Name.WriteLine(" " & DtSet_Name.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: EntityCode not exist in SAP Customer Master Date Time : " & DateTime.Now & "")
                            CName = False
                        End If


                    End If
                Catch ex As Exception
                    ' MsgBox(ex.Message)
                    file_Name.WriteLine(" " & DtSet_Name.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: " & ex.Message & " Date Time : " & DateTime.Now & "")
                    CName = False
                End Try
            Next
            file_Name.Close()
        Catch ex As Exception
            file_Name.WriteLine("Error Message: " & ex.Message & " Date Time : " & DateTime.Now & "")
            file_Name.Close()
            CName = False
        End Try
    End Sub

    Private Sub CustomerMain(ByVal FCustMaster As String, ByVal csvFileFolder As String)

        Dim file_Master As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath & "ErrorLog_Master" & TDT & ".txt", True)


        Try
            '--------------
            Dim csvFileName As String = FCustMaster.Replace(csvFileFolder, "")
            'Dim strConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & csvFileFolder & ";Extended Properties=Text;"
            'Dim conn As New OleDbConnection(strConnString)

            ''Try
            'conn.Open()
            'Dim cmd As New OleDbCommand("SELECT * FROM [" & csvFileName & "]", conn)
            'Dim da As New OleDbDataAdapter()

            'da.SelectCommand = cmd

            'Dim ds As New DataSet()
            'Dim DtSet_Master As New System.Data.DataSet

            'da.Fill(DtSet_Master)
            'da.Dispose()
            'Dim rowcount_Master As Integer = DtSet_Master.Tables(0).Rows.Count
            '   ds.Tables(0)
            'Catch
            '    '  Return Nothing
            'Finally
            '    conn.Close()
            'End Try
            '------------------
            '**********
            Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            Dim s1, s2, s3, s4, s5 As String
            s1 = "[" & csvFileName & "]"
            s2 = "ColNameHeader=True"
            s3 = "Format=CSVDelimited"
            s4 = "MaxScanRows=0"
            s5 = "CharacterSet=OEM"
            srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf + s5.ToString())
            srOutput.Close()
            fsOutput.Close()

            '        	FileStream fsOutput = new FileStream(txtCSVFolderPath.Text+"\\schema.ini",FileMode.Create, FileAccess.Write);
            'StreamWriter srOutput = new StreamWriter (fsOutput);
            'string s1,s2,s3,s4,s5;
            's1="["+strCSVFile+"]";
            's2="ColNameHeader="+bolColName.ToString();
            's3="Format="+strFormat;
            's4="MaxScanRows=25";
            's5="CharacterSet=OEM";
            'srOutput.WriteLine(s1.ToString()+'\n'+s2.ToString()+'\n'+s3.ToString()+'\n'+s4.ToString()+'\n'+s5.ToString());
            'srOutput.Close ();
            'fsOutput.Close ();	

            '****************

            'Dim file As System.IO.StreamWriter = New System.IO.StreamWriter("" & PublicVariable.LogFilePath  & "ErrorLog.txt", True)
            ' Dim csvFileName As String = FCustMaster.Replace(csvFileFolder, "")
            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                   & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""

            Dim strConnString As String = "Provider=Microsoft.Jet.OLEDB.14.0;Data Source=" & csvFileFolder & ";Extended Properties=Text;"
            Dim conn As New Odbc.OdbcConnection(connString)
            Dim da2 As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            'Dim da2 As New Odbc.OdbcDataAdapter("SELECT EntityCode,EntityName,DateEntered,ContactName,BillingContactCode,BillingContactTitle,BillingFirstName,BillingName,Administrator,PlaceofIncorporation,Manger,ShortName,PrincipalCode,PrincipalTitle,PrincipleFirstname,PrincipalName,DateofIncorporation,IncorporationNumber,Jurisdiction,Status,Notes,BVIP1or2 FROM [" & csvFileName & "]", conn)
            Dim DtSet_Master1 As System.Data.DataSet
            DtSet_Master1 = New System.Data.DataSet

            da2.Fill(DtSet_Master1)
            'DtSet_Master.Load(myReader, LoadOption.OverwriteChanges, New String(""))
            Dim rd As DataTableReader = DtSet_Master1.Tables(0).CreateDataReader()
            'objCommand.Connection.Close()
            '*****************
            Dim oRecordSetBP1 As SAPbobsCOM.Recordset
            oRecordSetBP1 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSetBP1.DoQuery("Delete from TEMP2")
            Dim connectionString As String = "Server=" & PublicVariable.SQLServer & ";Database=" & PublicVariable.SQLDB & ";User Id=" & PublicVariable.SQLUser & ";Password=" & PublicVariable.SQLPwd & ""
            Using destinationConnection As SqlConnection = _
               New SqlConnection(connectionString)
                destinationConnection.Open()
                Using copy As New SqlBulkCopy(destinationConnection)
                    copy.DestinationTableName = "TEMP2"
                    copy.WriteToServer(rd)
                End Using
                destinationConnection.Close()
            End Using
            Dim DtSet_Master As System.Data.DataSet
            DtSet_Master = New System.Data.DataSet
            Dim connection As New SqlConnection(connectionString)
            Dim sql As String = "execute [dbo].[AB_CustomerMaster]"
            ' Dim dataadapter As New SqlDataAdapter(sql, connection)
            Dim SQLCMD As SqlCommand
            SQLCMD = New SqlCommand(sql, connection)
            SQLCMD.CommandTimeout = 600
            Dim reader As SqlDataReader
            Try
                connection.Open()
                reader = SQLCMD.ExecuteReader
                reader.Read()

                DtSet_Master.Load(reader, LoadOption.OverwriteChanges, "TAB1")
                connection.Close()
                'connection.Open()


                'dataadapter.Fill(DtSet_Master)
                'connection.Close()
            Catch ex As Exception
                file_Master.WriteLine(" Error Message: " & ex.Message & "Date Time : " & DateTime.Now & "")
                Exit Sub
            End Try
            '********************

            Dim rowcount_Master As Integer = DtSet_Master.Tables(0).Rows.Count
            If rowcount_Master = 0 Then
                'file.WriteLine("No Data in Customer Master CSV File- Date Time : " & DateTime.Now & "")
            End If

            '------------
            Dim i As Integer
            Dim F As Integer
            Dim oCRD As SAPbobsCOM.BusinessPartners
            Dim oCRG As SAPbobsCOM.BusinessPartnerGroups
            oCRD = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oCRG = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups)
            Dim return_value As Integer
            Dim SerrorMsg As String = ""
            Dim oRecordSetSP_D As SAPbobsCOM.Recordset
            For i = 0 To rowcount_Master - 1
                Dim oRecordSetBP As SAPbobsCOM.Recordset
                oRecordSetBP = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSetBP.DoQuery("select cardcode from OCRD where CardCode='" & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                If oRecordSetBP.RecordCount = 0 Then
                    oCRD.CardCode = DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim()
                    Dim Cardname As String = DtSet_Master.Tables(0).Rows(i).Item(1).ToString.Trim()
                    If Cardname.Length > 100 Then
                        oCRD.CardName = Cardname.Substring(0, 99)
                        oCRD.UserFields.Fields.Item("U_AI_CardName").Value = Cardname
                    Else
                        oCRD.CardName = Cardname
                        oCRD.UserFields.Fields.Item("U_AI_CardName").Value = Cardname
                    End If
                    oCRD.Add()
                End If
            Next
            oCRD = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            For i = 0 To rowcount_Master - 1
                Try

                    Dim oRecordSetBP As SAPbobsCOM.Recordset
                    oRecordSetBP = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSetBP.DoQuery("select cardcode from OCRD where CardCode='" & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                    If oRecordSetBP.RecordCount = 0 Then
                        '----BP Add
                        oCRD.CardCode = DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim()
                        Dim Cardname As String = DtSet_Master.Tables(0).Rows(i).Item(1).ToString.Trim()
                        If Cardname.Length > 100 Then
                            oCRD.CardName = Cardname.Substring(0, 99)
                            oCRD.UserFields.Fields.Item("U_AI_CardName").Value = Cardname
                        Else
                            oCRD.CardName = Cardname
                            oCRD.UserFields.Fields.Item("U_AI_CardName").Value = Cardname
                        End If


                        Dim oRecordSet2 As SAPbobsCOM.Recordset
                        oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet2.DoQuery("SELECT groupcode  FROM OCRG  where GroupName='" & DtSet_Master.Tables(0).Rows(i).Item(19).ToString.Trim() & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oCRD.GroupCode = oRecordSet2.Fields.Item(0).Value
                        Else
                            oCRG.Name = DtSet_Master.Tables(0).Rows(i).Item(19).ToString.Trim()
                            oCRG.Add()
                            oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet2.DoQuery("SELECT groupcode  FROM OCRG  where GroupName='" & DtSet_Master.Tables(0).Rows(i).Item(19).ToString.Trim() & "'")
                            If oRecordSet2.RecordCount > 0 Then
                                oCRD.GroupCode = oRecordSet2.Fields.Item(0).Value
                            Else
                                oCRD.GroupCode = 100
                            End If

                        End If
                        Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                        Try


                            If DtSet_Master.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                dt1 = DtSet_Master.Tables(0).Rows(i).Item(2).ToString.Trim()
                                dt1 = dt1.Insert(4, "-")
                                dt1 = dt1.Insert(7, "-")
                            End If
                        Catch ex As Exception

                        End Try
                        If DtSet_Master.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                            oCRD.UserFields.Fields.Item("U_AI_DateEntered").Value = dt1
                        Else
                            oCRD.UserFields.Fields.Item("U_AI_DateEntered").Value = "1900-01-01"
                        End If
                        If DtSet_Master.Tables(0).Rows(i).Item(3).ToString.Trim() = "TRUE" Then
                            oCRD.UserFields.Fields.Item("U_AI_OnceOff").Value = "1"
                        Else
                            oCRD.UserFields.Fields.Item("U_AI_OnceOff").Value = "0"
                        End If
                        oCRD.CardType = SAPbobsCOM.BoCardTypes.cCustomer

                        Call oCRD.ContactEmployees.Add()
                        oCRD.ContactEmployees.SetCurrentLine(0)
                        oCRD.ContactEmployees.Title = DtSet_Master.Tables(0).Rows(i).Item(6).ToString.Trim()
                        oCRD.ContactEmployees.FirstName = DtSet_Master.Tables(0).Rows(i).Item(5).ToString.Trim()
                        oCRD.ContactEmployees.LastName = DtSet_Master.Tables(0).Rows(i).Item(8).ToString.Trim()
                        If DtSet_Master.Tables(0).Rows(i).Item(7).ToString.Trim() = "" Then
                            oCRD.ContactEmployees.Name = "-"
                            oCRD.ContactPerson = "-"
                        Else
                            oCRD.ContactEmployees.Name = DtSet_Master.Tables(0).Rows(i).Item(7).ToString.Trim()
                            oCRD.ContactPerson = DtSet_Master.Tables(0).Rows(i).Item(7).ToString.Trim()
                        End If
                        oCRD.UserFields.Fields.Item("U_AI_InvContactCode").Value = DtSet_Master.Tables(0).Rows(i).Item(4).ToString.Trim()
                        'oCRD.UserFields.Fields.Item("U_AI_BillContract").Value = DtSet_Master.Tables(0).Rows(i).Item(5).ToString.trim()
                        oCRD.UserFields.Fields.Item("U_AI_PlaceOfIncor").Value = DtSet_Master.Tables(0).Rows(i).Item(10).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_Administrator").Value = DtSet_Master.Tables(0).Rows(i).Item(9).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_Manger").Value = DtSet_Master.Tables(0).Rows(i).Item(11).ToString.Trim()
                        oCRD.CardForeignName = DtSet_Master.Tables(0).Rows(i).Item(12).ToString.Trim()
                        '-----
                        oCRD.UserFields.Fields.Item("U_AI_Principal").Value = DtSet_Master.Tables(0).Rows(i).Item(13).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_PrincipalTitle").Value = DtSet_Master.Tables(0).Rows(i).Item(14).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_PrincipalName").Value = DtSet_Master.Tables(0).Rows(i).Item(16).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_PrincipalFirst").Value = DtSet_Master.Tables(0).Rows(i).Item(15).ToString.Trim()
                        Dim dt2 As String = Format(Now.Date, "yyyy-MM-dd")
                        Try


                            If DtSet_Master.Tables(0).Rows(i).Item(17).ToString.Trim() <> "" Then
                                dt2 = DtSet_Master.Tables(0).Rows(i).Item(17).ToString.Trim()
                                dt2 = dt2.Insert(4, "-")
                                dt2 = dt2.Insert(7, "-")
                            End If
                        Catch ex As Exception

                        End Try
                        If DtSet_Master.Tables(0).Rows(i).Item(17).ToString.Trim() <> "" Then
                            oCRD.UserFields.Fields.Item("U_AI_DOI").Value = dt2
                        Else
                            oCRD.UserFields.Fields.Item("U_AI_DOI").Value = "1900-01-01"
                        End If
                        oCRD.UserFields.Fields.Item("U_AI_CustStatus").Value = DtSet_Master.Tables(0).Rows(i).Item(20).ToString.Trim()
                        If DtSet_Master.Tables(0).Rows(i).Item(20).ToString.Trim() = "8" Then
                            oCRD.Frozen = SAPbobsCOM.BoYesNoEnum.tYES
                            oCRD.Valid = SAPbobsCOM.BoYesNoEnum.tNO

                        Else
                            oCRD.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                            oCRD.Valid = SAPbobsCOM.BoYesNoEnum.tYES

                        End If
                        Try
                            ' fileFields = fileRows(i + 1).Split(ControlChars.Tab)
                            oCRD.UserFields.Fields.Item("U_AI_IncorNo").Value = DtSet_Master.Tables(0).Rows(i).Item(18).ToString.Trim() 'fileFields(18).Trim ' DtSet_Master.Tables(0).Rows(i).Item(18).ToString.Trim()
                        Catch ex As Exception
                            oCRD.UserFields.Fields.Item("U_AI_IncorNo").Value = DtSet_Master.Tables(0).Rows(i).Item(18).ToString.Trim()
                        End Try

                        oCRD.UserFields.Fields.Item("U_AI_BVIP1Or2").Value = DtSet_Master.Tables(0).Rows(i).Item(22).ToString.Trim()
                        oCRD.FreeText = DtSet_Master.Tables(0).Rows(i).Item(21).ToString.Trim()
                        return_value = oCRD.Add()
                        If return_value = 0 Then
                            'file_Master.WriteLine(" " & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.trim() & "-BP Created Success; Line No - " & i + 1 & " Date Time : " & DateTime.Now & "")
                        Else
                            PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                            file_Master.WriteLine(" " & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Created Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "Date Time : " & DateTime.Now & "")
                            CMaster = False
                        End If
                    Else
                        '---BP Update
                        '    GoTo 75
                        Try
                            oRecordSetSP_D = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim cname As String = DtSet_Master.Tables(0).Rows(i).Item(7).ToString.Trim()
                            If cname = "" Then
                                cname = "-"
                            End If
                            Dim stt As String = "DELETE FROM OCPR  WHERE [CardCode]='" & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim() & "' and Name= '" & cname & "'"
                            oRecordSetSP_D.DoQuery(stt)
                            oRecordSetSP_D = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            stt = "update ocrd set [CntctPrsn]='' where [CardCode] ='" & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim() & "'"
                            oRecordSetSP_D.DoQuery(stt)

                        Catch ex As Exception
                        End Try
                        oCRD.GetByKey(DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim())
                        Dim Cardname As String = DtSet_Master.Tables(0).Rows(i).Item(1).ToString.Trim()
                        If Cardname.Length > 100 Then
                            oCRD.CardName = Cardname.Substring(0, 99)
                            oCRD.UserFields.Fields.Item("U_AI_CardName").Value = Cardname
                        Else
                            oCRD.CardName = Cardname
                            oCRD.UserFields.Fields.Item("U_AI_CardName").Value = Cardname
                        End If
                        'oCRD.GroupCode = 100

                        Dim oRecordSet2 As SAPbobsCOM.Recordset
                        oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet2.DoQuery("SELECT groupcode  FROM OCRG  where GroupName='" & DtSet_Master.Tables(0).Rows(i).Item(19).ToString.Trim() & "'")
                        If oRecordSet2.RecordCount > 0 Then
                            oCRD.GroupCode = oRecordSet2.Fields.Item(0).Value
                        Else
                            oCRG.Name = DtSet_Master.Tables(0).Rows(i).Item(19).ToString.Trim()
                            oCRG.Add()
                            oRecordSet2 = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet2.DoQuery("SELECT groupcode  FROM OCRG  where GroupName='" & DtSet_Master.Tables(0).Rows(i).Item(19).ToString.Trim() & "'")
                            If oRecordSet2.RecordCount > 0 Then
                                oCRD.GroupCode = oRecordSet2.Fields.Item(0).Value
                            Else
                                oCRD.GroupCode = 100
                            End If

                        End If
                        Dim dt1 As String = Format(Now.Date, "yyyy-MM-dd")
                        Try


                            If DtSet_Master.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                                dt1 = DtSet_Master.Tables(0).Rows(i).Item(2).ToString.Trim()
                                dt1 = dt1.Insert(4, "-")
                                dt1 = dt1.Insert(7, "-")
                            End If
                        Catch ex As Exception

                        End Try
                        If DtSet_Master.Tables(0).Rows(i).Item(2).ToString.Trim() <> "" Then
                            oCRD.UserFields.Fields.Item("U_AI_DateEntered").Value = dt1
                        Else
                            oCRD.UserFields.Fields.Item("U_AI_DateEntered").Value = "1900-01-01"
                        End If

                        If DtSet_Master.Tables(0).Rows(i).Item(3).ToString.Trim() = "TRUE" Then
                            oCRD.UserFields.Fields.Item("U_AI_OnceOff").Value = "1"
                        Else
                            oCRD.UserFields.Fields.Item("U_AI_OnceOff").Value = "0"
                        End If
                        oCRD.CardType = SAPbobsCOM.BoCardTypes.cCustomer
                        Dim oRecordSetSP As SAPbobsCOM.Recordset

                        Call oCRD.ContactEmployees.Add()
                        ' Dim oRecordSetSP As SAPbobsCOM.Recordset
                        oRecordSetSP = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSetSP.DoQuery("SELECT T0.[Name] FROM OCPR T0 WHERE T0.[CardCode]='" & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim() & "'")
                        oCRD.ContactEmployees.SetCurrentLine(oRecordSetSP.RecordCount)
                        oCRD.ContactEmployees.Title = DtSet_Master.Tables(0).Rows(i).Item(6).ToString.Trim()
                        oCRD.ContactEmployees.FirstName = DtSet_Master.Tables(0).Rows(i).Item(5).ToString.Trim()
                        oCRD.ContactEmployees.LastName = DtSet_Master.Tables(0).Rows(i).Item(8).ToString.Trim()
                        If DtSet_Master.Tables(0).Rows(i).Item(7).ToString.Trim() = "" Then
                            oCRD.ContactEmployees.Name = "-"
                            oCRD.ContactPerson = "-"
                        Else
                            oCRD.ContactEmployees.Name = DtSet_Master.Tables(0).Rows(i).Item(7).ToString.Trim()
                            oCRD.ContactPerson = DtSet_Master.Tables(0).Rows(i).Item(7).ToString.Trim()
                        End If
                        '-----
                        oCRD.UserFields.Fields.Item("U_AI_InvContactCode").Value = DtSet_Master.Tables(0).Rows(i).Item(4).ToString.Trim()

                        oCRD.UserFields.Fields.Item("U_AI_PlaceOfIncor").Value = DtSet_Master.Tables(0).Rows(i).Item(10).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_Administrator").Value = DtSet_Master.Tables(0).Rows(i).Item(9).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_Manger").Value = DtSet_Master.Tables(0).Rows(i).Item(11).ToString.Trim()
                        oCRD.CardForeignName = DtSet_Master.Tables(0).Rows(i).Item(12).ToString.Trim()
                        '-----
                        oCRD.UserFields.Fields.Item("U_AI_Principal").Value = DtSet_Master.Tables(0).Rows(i).Item(13).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_PrincipalTitle").Value = DtSet_Master.Tables(0).Rows(i).Item(14).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_PrincipalName").Value = DtSet_Master.Tables(0).Rows(i).Item(16).ToString.Trim()
                        oCRD.UserFields.Fields.Item("U_AI_PrincipalFirst").Value = DtSet_Master.Tables(0).Rows(i).Item(15).ToString.Trim()
                        Dim dt2 As String = Format(Now.Date, "yyyy-MM-dd")
                        Try
                            If DtSet_Master.Tables(0).Rows(i).Item(17).ToString.Trim() <> "" Then
                                dt2 = DtSet_Master.Tables(0).Rows(i).Item(17).ToString.Trim()
                                dt2 = dt2.Insert(4, "-")
                                dt2 = dt2.Insert(7, "-")
                            End If
                        Catch ex As Exception

                        End Try
                        If DtSet_Master.Tables(0).Rows(i).Item(17).ToString.Trim() <> "" Then
                            oCRD.UserFields.Fields.Item("U_AI_DOI").Value = dt2
                        Else
                            oCRD.UserFields.Fields.Item("U_AI_DOI").Value = "1900-01-01"
                        End If
                        oCRD.UserFields.Fields.Item("U_AI_CustStatus").Value = DtSet_Master.Tables(0).Rows(i).Item(20).ToString.Trim()
                        If DtSet_Master.Tables(0).Rows(i).Item(20).ToString.Trim() = "8" Then
                            oCRD.Frozen = SAPbobsCOM.BoYesNoEnum.tYES
                            oCRD.Valid = SAPbobsCOM.BoYesNoEnum.tNO

                        Else
                            oCRD.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                            oCRD.Valid = SAPbobsCOM.BoYesNoEnum.tYES

                        End If
                        Try
                            'fileFields = fileRows(i + 1).Split(ControlChars.Tab)

                            Dim tset As String = DtSet_Master.Tables(0).Rows(i).Item(18).ToString.Trim()
                            oCRD.UserFields.Fields.Item("U_AI_IncorNo").Value = DtSet_Master.Tables(0).Rows(i).Item(18).ToString.Trim() 'fileFields(18).Trim()
                        Catch ex As Exception
                            oCRD.UserFields.Fields.Item("U_AI_IncorNo").Value = DtSet_Master.Tables(0).Rows(i).Item(18).ToString.Trim()
                        End Try

                        oCRD.UserFields.Fields.Item("U_AI_BVIP1Or2").Value = DtSet_Master.Tables(0).Rows(i).Item(22).ToString.Trim()

                        oCRD.FreeText = DtSet_Master.Tables(0).Rows(i).Item(21).ToString.Trim()
                        return_value = oCRD.Update()
                        If return_value = 0 Then
                            'file_Master.WriteLine(" " & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.trim() & "-BP Created Success; Line No - " & i + 1 & " Date Time : " & DateTime.Now & "")
                        Else
                            PublicVariable.oCompany.GetLastError(return_value, SerrorMsg)
                            file_Master.WriteLine(" " & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Update Failed; Line No - " & i + 1 & " Error Message: " & SerrorMsg & "Date Time : " & DateTime.Now & "")
                            CMaster = False
                        End If
                    End If

                Catch ex As Exception
                    ' MsgBox(ex.Message)
                    file_Master.WriteLine(" " & DtSet_Master.Tables(0).Rows(i).Item(0).ToString.Trim() & "-BP Created Failed; Line No - " & i + 1 & " Error Message: " & ex.Message & "Date Time : " & DateTime.Now & "")
                    CMaster = False
                End Try
75:             Dim stest1 As String = ""

            Next
            file_Master.Close()
        Catch ex As Exception
            file_Master.WriteLine("Error Message: " & ex.Message & "Date Time : " & DateTime.Now & "")
            CMaster = False
            file_Master.Close()
        End Try
    End Sub
#End Region
#End Region
    Private Sub EMail(ByVal FCustMaster As String, ByVal FCustname As String, ByVal FCustaddr As String, ByVal FCustSales As String, ByVal FCustRel_H As String, ByVal FIH As String, ByVal FIH1 As String, ByVal FCustRel_R As String, ByVal FIR As String, ByVal FIR1 As String, ByVal csvFileFolder As String)
        Try
            'Email
            If FCustMaster = "" And FCustname = "" And FCustaddr = "" And FCustSales = "" And FCustRel_H = "" And FIH = "" Then
                Exit Sub
            End If
            Dim ToAddress As String = ""
            Try


                Dim sPath As String = IO.Directory.GetParent(Application.StartupPath).ToString.Trim()
                Dim objIniFile As New INIClass(sPath.Replace("bin", "") & "\" & "ConfigFile.ini")
                PublicVariable.SMTPServer = objIniFile.GetString("SMTP", "SMTP_Server", "")
                PublicVariable.SMTPPort = objIniFile.GetString("SMTP", "SMTP_Port", "")
                PublicVariable.SMTPPwd = objIniFile.GetString("SMTP", "SMTP_Pwd", "")
                PublicVariable.SMTPEmail = objIniFile.GetString("SMTP", "SMTP_Email", "")
                PublicVariable.EmailList = objIniFile.GetString("SMTP", "EMAILLIST", "")
                '  PublicVariable.SMTPHost = objIniFile.GetString("SMTP", "SMTP_Host", "")
                'EmailList
                'SMTPHost
                ToAddress = PublicVariable.EmailList
            Catch ex As Exception
                '   ToAddress = "gokikopi@gmail.com"
            End Try
            Dim SmtpServer As New SmtpClient()
            SmtpServer.Credentials = New Net.NetworkCredential(PublicVariable.SMTPEmail, PublicVariable.SMTPPwd)
            SmtpServer.Port = PublicVariable.SMTPPort
            SmtpServer.Host = PublicVariable.SMTPServer
            SmtpServer.EnableSsl = True
            Dim mail As New MailMessage()
            mail = New MailMessage()
            Dim AttachmentFiles As ArrayList = Nothing
            Dim addr As String = ToAddress
            ' Dim j, iCnt As Integer

            Try
                mail.From = New MailAddress(PublicVariable.SMTPEmail, "HFS", System.Text.Encoding.UTF8)

                Dim i As Byte
                'For i = 0 To addr.Length - 1
                mail.To.Add(addr)
                'Next
                mail.Subject = "Viewpoint Integration Summary Log " & TDT1 & "  "
                mail.IsBodyHtml = True
                Dim Body1 As String = ""
                Dim Body2 As String = ""
                Dim Body3 As String = ""
                Dim Body4 As String = ""
                Dim Body5 As String = ""
                Dim Body6 As String = ""
                Dim Body7 As String = ""

                Dim Succ As String = ""
                If CMaster = True Then
                    Body1 = "" & FCustMaster & "-File Name which has been successfuly uploaded into SAP"
                End If
                If CName = True Then
                    Body2 = "" & FCustname & "-File Name which has been successfuly uploaded into SAP"
                End If
                If CAddr = True Then
                    Body3 = "" & FCustaddr & "-File Name which has been successfuly uploaded into SAP"
                End If
                If CSales = True Then
                    Body4 = "" & FCustSales & "-File Name which has been successfuly uploaded into SAP"
                End If
                If Crel = True Then
                    Body5 = "" & FCustRel_H & "-File Name which has been successfuly uploaded into SAP"
                End If
                If SOH = True Then
                    Body6 = "" & FIH & "-File Name which has been successfuly uploaded into SAP"
                End If
                If IH = True Then
                    Body7 = "" & FIH1 & "-File Name which has been successfuly uploaded into SAP"
                End If
                If Body1 <> "" And FCustMaster <> "" Then
                    Succ = Succ & Body1 & "<br />"
                End If
                If Body2 <> "" And FCustname <> "" Then
                    Succ = Succ & Body2 & "<br />"
                End If
                If Body3 <> "" And FCustaddr <> "" Then
                    Succ = Succ & Body3 & "<br />"
                End If
                If Body4 <> "" And FCustSales <> "" Then
                    Succ = Succ & Body4 & "<br />"
                End If
                If Body5 <> "" And FCustRel_H <> "" Then
                    Succ = Succ & Body5 & "<br />"
                End If
                If Body6 <> "" And FIH <> "" Then
                    Succ = Succ & Body6 & "<br />"
                End If
                If Body7 <> "" And FIH1 <> "" Then
                    Succ = Succ & Body7 & "<br />"
                End If
                Dim Body11 As String = ""
                Dim Body21 As String = ""
                Dim Body31 As String = ""
                Dim Body41 As String = ""
                Dim Body51 As String = ""
                Dim Body61 As String = ""
                Dim Body71 As String = ""

                Dim Fail As String = ""
                If CMaster = False Then
                    Body11 = "" & FCustMaster & "-File name which has failed to upload and ErrorLog_Master" & TDT & ".txt of the file containing error information."
                End If
                If CName = False Then
                    Body21 = "" & FCustname & "-File name which has failed to upload and ErrorLog_Name" & TDT & ".txt of the file containing error information."
                End If
                If CAddr = False Then
                    Body31 = "" & FCustaddr & "-File name which has failed to upload and ErrorLog_Addr" & TDT & ".txt of the file containing error information."
                End If
                If CSales = False Then
                    Body41 = "" & FCustSales & "-File name which has failed to upload and ErrorLog_Sales" & TDT & ".txt of the file containing error information."
                End If
                If Crel = False Then
                    Body51 = "" & FCustRel_H & "-File name which has failed to upload and ErrorLog_CustRelation" & TDT & ".txt of the file containing error information."
                End If
                If SOH = False Then
                    Body61 = "" & FIH & "-File name which has failed to upload and ErrorLog_ARInvoice_SOH" & TDT & ".txt of the file containing error information."
                End If
                If IH = False Then
                    Body71 = "" & FIH1 & "-File name which has failed to upload and ErrorLog_ARInvoice_IH" & TDT & ".txt of the file containing error information."
                End If
                If Body11 <> "" And FCustMaster <> "" Then
                    Fail = Fail & Body11 & "<br />"
                End If
                If Body21 <> "" And FCustname <> "" Then
                    Fail = Fail & Body21 & "<br />"
                End If
                If Body31 <> "" And FCustaddr <> "" Then
                    Fail = Fail & Body31 & "<br />"
                End If
                If Body41 <> "" And FCustSales <> "" Then
                    Fail = Fail & Body41 & "<br />"
                End If
                If Body51 <> "" And FCustRel_H <> "" Then
                    Fail = Fail & Body51 & "<br />"
                End If
                If Body61 <> "" And FIH <> "" Then
                    Fail = Fail & Body61 & "<br />"
                End If
                If Body71 <> "" And FIH1 <> "" Then
                    Fail = Fail & Body71 & "<br />"
                End If

                mail.Body = Succ & "<br />" & Fail
                '    <ol>
                '        <li>Test</li>
                '        <li>Test</li>
                '        <li>Test</li>
                '    </ol>
                '</div>"

                ' AttachmentFiles = "" & PublicVariable.LogFilePath & "ErrorLog_Master20130103.txt"
                'If Not AttachmentFiles Is Nothing Then
                '    iCnt = AttachmentFiles.Count - 1
                '    For i = 0 To iCnt
                '        If FileExists(AttachmentFiles(j)) Then _
                '          mail.Attachments.Add(AttachmentFiles(j))
                '    Next

                'End If

                Dim attachment As System.Net.Mail.Attachment
                attachment = New System.Net.Mail.Attachment("" & PublicVariable.LogFilePath & "\ErrorLog_Master" & TDT & ".txt")
                mail.Attachments.Add(attachment)
                '                Dim attachment As System.Net.Mail.Attachment
                attachment = New System.Net.Mail.Attachment("" & PublicVariable.LogFilePath & "\ErrorLog_Name" & TDT & ".txt")
                mail.Attachments.Add(attachment)
                attachment = New System.Net.Mail.Attachment("" & PublicVariable.LogFilePath & "\ErrorLog_Addr" & TDT & ".txt")
                mail.Attachments.Add(attachment)
                attachment = New System.Net.Mail.Attachment("" & PublicVariable.LogFilePath & "\ErrorLog_Sales" & TDT & ".txt")
                mail.Attachments.Add(attachment)
                attachment = New System.Net.Mail.Attachment("" & PublicVariable.LogFilePath & "\ErrorLog_CustRelation" & TDT & ".txt")
                mail.Attachments.Add(attachment)
                attachment = New System.Net.Mail.Attachment("" & PublicVariable.LogFilePath & "\ErrorLog_ARInvoice_SOH" & TDT & ".txt")
                mail.Attachments.Add(attachment)
                attachment = New System.Net.Mail.Attachment("" & PublicVariable.LogFilePath & "\ErrorLog_ARInvoice_IH" & TDT & ".txt")
                mail.Attachments.Add(attachment)
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
                mail.ReplyTo = New MailAddress(PublicVariable.SMTPEmail)
                SmtpServer.Send(mail)
                mail.Dispose()
            Catch ex As Exception
                mail.Dispose()
                Functions.WriteLog(ex.ToString)
            End Try
            '
            Try
                If FCustMaster <> "" And CMaster = True Then
                    FileMove(FCustMaster, PublicVariable.SuccessFolder & FCustMaster.Replace(csvFileFolder, ""))
                ElseIf FCustMaster <> "" And CMaster = False Then
                    FileMove(FCustMaster, PublicVariable.ErrorFolder & FCustMaster.Replace(csvFileFolder, ""))
                End If
                If FCustname <> "" And CName = True Then
                    FileMove(FCustname, PublicVariable.SuccessFolder & FCustname.Replace(csvFileFolder, ""))
                ElseIf FCustname <> "" And CName = False Then
                    FileMove(FCustname, PublicVariable.ErrorFolder & FCustname.Replace(csvFileFolder, ""))
                End If
                If FCustaddr <> "" And CAddr = True Then
                    FileMove(FCustaddr, PublicVariable.SuccessFolder & FCustaddr.Replace(csvFileFolder, ""))
                ElseIf FCustaddr <> "" And CAddr = False Then
                    FileMove(FCustaddr, PublicVariable.ErrorFolder & FCustaddr.Replace(csvFileFolder, ""))
                End If
                If FCustSales <> "" And CSales = True Then
                    FileMove(FCustSales, PublicVariable.SuccessFolder & FCustSales.Replace(csvFileFolder, ""))
                ElseIf FCustSales <> "" And CSales = False Then
                    FileMove(FCustSales, PublicVariable.ErrorFolder & FCustSales.Replace(csvFileFolder, ""))
                End If
                If FCustRel_H <> "" And Crel = True Then
                    FileMove(FCustRel_H, PublicVariable.SuccessFolder & FCustRel_H.Replace(csvFileFolder, ""))
                ElseIf FCustRel_H <> "" And Crel = False Then
                    FileMove(FCustRel_H, PublicVariable.ErrorFolder & FCustRel_H.Replace(csvFileFolder, ""))
                End If
                If FCustRel_R <> "" And Crel = True Then
                    FileMove(FCustRel_R, PublicVariable.SuccessFolder & FCustRel_R.Replace(csvFileFolder, ""))
                ElseIf FCustRel_R <> "" And Crel = False Then
                    FileMove(FCustRel_R, PublicVariable.ErrorFolder & FCustRel_R.Replace(csvFileFolder, ""))
                End If
                If FIH <> "" And SOH = True Then
                    FileMove(FIH, PublicVariable.SuccessFolder & FIH.Replace(csvFileFolder, ""))
                ElseIf FIH <> "" And SOH = False Then
                    FileMove(FIH, PublicVariable.ErrorFolder & FIH.Replace(csvFileFolder, ""))
                End If
                If FIR <> "" And SOH = True Then
                    FileMove(FIR, PublicVariable.SuccessFolder & FIR.Replace(csvFileFolder, ""))
                ElseIf FIR <> "" And SOH = False Then
                    FileMove(FIR, PublicVariable.ErrorFolder & FIR.Replace(csvFileFolder, ""))
                End If
                If FIH1 <> "" And IH = True Then
                    FileMove(FIH1, PublicVariable.SuccessFolder & FIH1.Replace(csvFileFolder, ""))
                ElseIf FIH1 <> "" And IH = False Then
                    FileMove(FIH1, PublicVariable.ErrorFolder & FIH1.Replace(csvFileFolder, ""))
                End If
            Catch ex As Exception
                Functions.WriteLog(ex.ToString)
            End Try
        Catch ex As Exception
        End Try
    End Sub

    Public Sub FileMove(ByVal sourcePath As String, ByVal dest As String)
        Try
            File.Move(sourcePath, dest)
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
End Class
