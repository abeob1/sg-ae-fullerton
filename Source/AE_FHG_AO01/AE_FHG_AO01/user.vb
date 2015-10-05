
Imports System.IO


Namespace AE_FHG_AO01
    Module user

        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing

        Public Function UsersSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                    ByRef sErrDesc As String) As Long

            'Function   :   UsersSync()
            'Purpose    :   
            'Parameters :   ByVal oForm As SAPbouiCOM.Form
            '                   oForm=Form Type
            '               ByRef sErrDesc As String
            '                   sErrDesc=Error Description to be returned to calling function
            '               
            '                   =
            'Return     :   0 - FAILURE
            '               1 - SUCCESS
            'Author     :   SRINIVASAN
            'Date       :   09/09/2015
            'Change     :

            Dim sFuncName As String = String.Empty
            sFuncName = "UsersSync()"
            Dim iErrCode As Double
            Try
                ''INITIALIZE THE OBJECT FOR RECORDSET
                Dim oRset As SAPbobsCOM.Recordset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oRset_T As SAPbobsCOM.Recordset = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                ''Initialize the SAP object for user
                Dim oUser_Holding As SAPbobsCOM.Users = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
                Dim oUser_Target As SAPbobsCOM.Users = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)

                Dim sSQL As String = "SELECT ""USERID"" FROM ""OUSR"" T0 WHERE T0.""USER_CODE"" = '" & sMasterdatacode & "'"

                Dim sFileName As String = System.Windows.Forms.Application.StartupPath & "\Users.xml"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sSQL, sFuncName)
                ''Execute the query and get the userid for the particular USER_CODE
                oRset.DoQuery(sSQL)

                ''Store the userid into the variables
                Dim iUsercode As Integer = oRset.Fields.Item("userid").Value
                oRset_T.DoQuery(sSQL)
                Dim iUsercode_T As Integer = oRset_T.Fields.Item("userid").Value

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)

                ''Check whether the userid is exist in Holding DB
                If oUser_Holding.GetByKey(iUsercode) Then

                    If File.Exists(sFileName) Then '' Check the file already exist with the same name
                        File.Delete(sFileName) '' if exist, delete the file from the startup path 
                    End If
                    oUser_Holding.SaveXML(sFileName) '' Save the data as XML file in the startup path
                    If oUser_Target.GetByKey(iUsercode_T) Then '' check the user exist or not in target company
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_Assignment()", sFuncName)
                      

                        oUser_Target.Browser.ReadXml(sFileName, 0) ''Read the data's from the xml file

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Update the User " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oUser_Target.Update() '' Updte the user datas in target company
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            UsersSync = RTN_ERROR
                            Exit Function
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item_Assignment()", sFuncName)
                        
                        oUser_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0) ''Read the values from the XML file 
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the User" & oTragetCompany.CompanyDB, sFuncName)
                        ival = oUser_Target.Add() '' Add the new user in target company
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            UsersSync = RTN_ERROR
                            Exit Function
                        End If
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    UsersSync = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErr, sFuncName)
                UsersSync = RTN_SUCCESS
            Catch ex As Exception
                UsersSync = RTN_ERROR
                iErrCode = Err.Number
                sErrDesc = Err.Description
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try


        End Function

    End Module
End Namespace

