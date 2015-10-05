
Imports System.Xml
Imports System.IO

Namespace AE_FHG_AO01

    Module modMasterData

        Dim sPath As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim xDoc As New XmlDocument

        Dim oMatrix As SAPbouiCOM.Matrix = Nothing

        Public Function ItemMaterSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                       ByRef sErrDesc As String) As Long

            'Function   :   ItemMaterSync()
            'Purpose    :   
            'Parameters :   ByVal oForm As SAPbouiCOM.Form
            '                   oForm=Form Type
            '               ByRef sErrDesc As String
            '                   sErrDesc=Error Description to be returned to calling function
            '               
            '                   =
            'Return     :   0 - FAILURE
            '               1 - SUCCESS
            'Author     :   SRI
            'Date       :   30/12/2007
            'Change     :

            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "ItemMaterSync()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)

                Dim oItemMaster As SAPbobsCOM.Items = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                Dim oItemMaster_Target As SAPbobsCOM.Items = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Item Master Sync Function ", sFuncName)

                If oItemMaster.GetByKey(sMasterdatacode) Then
                    If oItemMaster_Target.GetByKey(sMasterdatacode) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item_Assignment()", sFuncName)
                        'oItemMaster.SaveXML("C:\item006.xml")
                        ' oItemMaster_Target = oTragetCompany.GetBusinessObjectFromXML("C:\item006.xml", 0)
                        Item_Assignment(oItemMaster, oItemMaster_Target)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oItemMaster_Target.Update()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            ItemMaterSync = RTN_ERROR
                            Exit Function
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item_Assignment()", sFuncName)
                        oItemMaster_Target.ItemCode = oItemMaster.ItemCode
                        Item_Assignment(oItemMaster, oItemMaster_Target)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oItemMaster_Target.Add()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            ItemMaterSync = RTN_ERROR
                            Exit Function
                        End If
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else

                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    ItemMaterSync = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function

                    ' 
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                ItemMaterSync = RTN_SUCCESS
            Catch ex As Exception
                ItemMaterSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
            End Try

        End Function

        Public Function BPMaterSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                      ByRef sErrDesc As String) As Long

            'Function   :   BPMaterSync()
            'Purpose    :   
            'Parameters :   ByVal oForm As SAPbouiCOM.Form
            '                   oForm=Form Type
            '               ByRef sErrDesc As String
            '                   sErrDesc=Error Description to be returned to calling function
            '               
            '                   =
            'Return     :   0 - FAILURE
            '               1 - SUCCESS
            'Author     :   SRI
            'Date       :   30/12/2007
            'Change     :

            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "BPMaterSync()"
                Dim sBPPaymentMethods As String = String.Empty
                Dim sSQLString As String = String.Empty

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                Dim oRset As SAPbobsCOM.Recordset = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oDlfPaymenthod As SAPbobsCOM.Recordset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oBP_Holding As SAPbobsCOM.BusinessPartners = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                ' Dim oContact_Holding As SAPbobsCOM.PaymentRunExport = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentRunExport)
                Dim oBP_Target As SAPbobsCOM.BusinessPartners = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                '   Dim oContact_Target As SAPbobsCOM.Contacts = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts)

                Dim sFileName As String = System.Windows.Forms.Application.StartupPath & "\ BP.xml"
                If oBP_Holding.GetByKey(sMasterdatacode) Then
                    sSQLString = "SELECT T0.[CardCode] FROM OCRD T0 WHERE T0.[U_AB_SYNCCODE] = '" & sMasterdatacode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting the respective BP " & sSQLString, sFuncName)
                    oRset.DoQuery(sSQLString)
                    sMasterdatacode = oRset.Fields.Item("CardCode").Value

                    If oBP_Target.GetByKey(sMasterdatacode) Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_Assignment()", sFuncName)
                        '  oBP_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0)
                        BP_Assignment(oBP_Holding, oBP_Target)
                        'oBP_Target.SaveXML(sFileName)
                        For imjs As Integer = 1 To oBP_Target.BPPaymentMethods.Count
                            oBP_Target.BPPaymentMethods.SetCurrentLine(imjs - 1)
                            If String.IsNullOrEmpty(oBP_Target.BPPaymentMethods.PaymentMethodCode) Then
                                oBP_Target.BPPaymentMethods.Delete()
                            End If
                        Next imjs

                        For imjs As Integer = 0 To oBP_Target.BPPaymentMethods.Count - 1
                            oBP_Target.BPPaymentMethods.SetCurrentLine(imjs)
                            If Not String.IsNullOrEmpty(oBP_Target.BPPaymentMethods.PaymentMethodCode) Then
                                sBPPaymentMethods += "'" & oBP_Target.BPPaymentMethods.PaymentMethodCode & "',"
                            End If
                        Next
                        sBPPaymentMethods = Left(sBPPaymentMethods, Len(sBPPaymentMethods) - 1)
                        ''sSQLString = "SELECT T0.[DfltVendPM] FROM OADM T0"
                        sSQLString = "SELECT T0.[PayMethCod], T0.[Descript] FROM OPYM T0 where T0.[PayMethCod] not  in ( " & sBPPaymentMethods & ")"
                        oRset.DoQuery(sSQLString)
                        For imjs As Integer = 1 To oRset.RecordCount
                            oBP_Target.BPPaymentMethods.Add()
                            oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("PayMethCod").Value
                            oRset.MoveNext()
                        Next imjs

                        'sSQLString = "SELECT T0.[DfltVendPM] FROM OADM T0"
                        'oDlfPaymenthod.DoQuery(sSQLString)
                        'oBP_Target.PeymentMethodCode = oDlfPaymenthod.Fields.Item("DfltVendPM").Value


                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the BP Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oBP_Target.Update()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BPMaterSync = RTN_ERROR
                            Exit Function
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_Assignment()", sFuncName)
                        oBP_Target.CardCode = oBP_Holding.CardCode
                        oBP_Target.Series = oBP_Holding.Series
                        BP_Assignment(oBP_Holding, oBP_Target)
                        oBP_Target.UserFields.Fields.Item("U_AB_SYNCCODE").Value = oBP_Holding.CardCode
                        sSQLString = "SELECT T0.[PayMethCod], T0.[Descript] FROM OPYM T0 "
                        oRset.DoQuery(sSQLString)
                        For imjs As Integer = 0 To oRset.RecordCount - 1
                            oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("PayMethCod").Value
                            oBP_Target.BPPaymentMethods.Add()
                            oRset.MoveNext()
                        Next
                        ''sSQLString = "SELECT T0.[DfltVendPM] FROM OADM T0"
                        ''oDlfPaymenthod.DoQuery(sSQLString)
                        ''oBP_Target.PeymentMethodCode = oDlfPaymenthod.Fields.Item("DfltVendPM").Value


                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oBP_Target.Add()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BPMaterSync = RTN_ERROR
                            Exit Function
                        End If
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else

                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    BPMaterSync = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                    ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & oTragetCompany.CompanyDB, sFuncName)
                End If
                BPMaterSync = RTN_SUCCESS
            Catch ex As Exception
                BPMaterSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally



            End Try

        End Function

        Public Function BPMaterSyncOld(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                     ByRef sErrDesc As String) As Long

            'Function   :   BPMaterSync()
            'Purpose    :   
            'Parameters :   ByVal oForm As SAPbouiCOM.Form
            '                   oForm=Form Type
            '               ByRef sErrDesc As String
            '                   sErrDesc=Error Description to be returned to calling function
            '               
            '                   =
            'Return     :   0 - FAILURE
            '               1 - SUCCESS
            'Author     :   SRI
            'Date       :   30/12/2007
            'Change     :

            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "BPMaterSync()"
                Dim sBPPaymentMethods As String = String.Empty
                Dim sSQLString As String = String.Empty

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                Dim oRset As SAPbobsCOM.Recordset = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oBP_Holding As SAPbobsCOM.BusinessPartners = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                ' Dim oContact_Holding As SAPbobsCOM.PaymentRunExport = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentRunExport)
                Dim oBP_Target As SAPbobsCOM.BusinessPartners = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                '   Dim oContact_Target As SAPbobsCOM.Contacts = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts)

                oRset.DoQuery("SELECT T0.[DfltVendPM] FROM OADM T0")


                Dim sFileName As String = System.Windows.Forms.Application.StartupPath & "\ BP.xml"
                If oBP_Holding.GetByKey(sMasterdatacode) Then

                    ''If File.Exists(sFileName) Then
                    ''    File.Delete(sFileName)
                    ''End If
                    If oBP_Target.GetByKey(sMasterdatacode) Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_Assignment()", sFuncName)
                        '  oBP_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0)
                        BP_Assignment(oBP_Holding, oBP_Target)

                        For imjs As Integer = 1 To oBP_Holding.BPPaymentMethods.Count
                            oBP_Holding.BPPaymentMethods.SetCurrentLine(imjs - 1)
                            If Not String.IsNullOrEmpty(oBP_Holding.BPPaymentMethods.PaymentMethodCode) Then
                                If imjs <= oBP_Target.BPPaymentMethods.Count Then
                                    oBP_Target.BPPaymentMethods.SetCurrentLine(imjs - 1)
                                    oBP_Target.BPPaymentMethods.PaymentMethodCode = oBP_Holding.BPPaymentMethods.PaymentMethodCode
                                Else
                                    oBP_Target.BPPaymentMethods.PaymentMethodCode = oBP_Holding.BPPaymentMethods.PaymentMethodCode
                                    oBP_Target.BPPaymentMethods.Add()
                                End If
                            End If
                        Next imjs

                        oBP_Target.SaveXML(sFileName)

                        For imjs As Integer = 1 To oBP_Target.BPPaymentMethods.Count
                            oBP_Target.BPPaymentMethods.SetCurrentLine(imjs - 1)
                            If String.IsNullOrEmpty(oBP_Target.BPPaymentMethods.PaymentMethodCode) Then
                                oBP_Target.BPPaymentMethods.Delete()
                            End If
                        Next imjs

                        oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("DfltVendPM").Value
                        oBP_Target.BPPaymentMethods.Add()

                        oBP_Target.SaveXML(sFileName)

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the BP Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oBP_Target.Update()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BPMaterSyncOld = RTN_ERROR
                            Exit Function
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item_Assignment()", sFuncName)
                        oBP_Target.CardCode = oBP_Holding.CardCode
                        oBP_Target.Series = oBP_Holding.Series
                        BP_Assignment(oBP_Holding, oBP_Target)
                        ''oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("DfltVendPM").Value
                        ''oBP_Target.BPPaymentMethods.Add()

                        ''For imjs As Integer = 0 To oRset.RecordCount - 1
                        ''    '******RECORDSET OUTPUT COLUMN VALUE CHANGED BY JEEVA ON 07/07/2015 11:56 ISD******
                        ''    'oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("PayMethCod").Value
                        ''    oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("DfltVendPM").Value
                        ''    oBP_Target.BPPaymentMethods.Add()
                        ''    oRset.MoveNext()
                        ''Next
                        '' oBP_Target.SaveXML(sFileName)
                        ' oBP_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oBP_Target.Add()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BPMaterSyncOld = RTN_ERROR
                            Exit Function
                        End If
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else

                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    BPMaterSyncOld = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                    ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & oTragetCompany.CompanyDB, sFuncName)
                End If
                BPMaterSyncOld = RTN_SUCCESS
            Catch ex As Exception
                BPMaterSyncOld = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally



            End Try

        End Function

        Public Sub Item_Assignment(ByRef oItemMaster As SAPbobsCOM.Items, ByRef oItemMaster_Target As SAPbobsCOM.Items)

            sFuncName = "Item_Assignment()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oItemMaster.SaveXML("C:\item.xml")
            oItemMaster_Target.ItemName = oItemMaster.ItemName
            oItemMaster_Target.ItemType = oItemMaster.ItemType
            oItemMaster_Target.ItemsGroupCode = oItemMaster.ItemsGroupCode
            oItemMaster_Target.InventoryItem = oItemMaster.InventoryItem
            oItemMaster_Target.SalesItem = oItemMaster.SalesItem
            oItemMaster_Target.PurchaseItem = oItemMaster.PurchaseItem
            oItemMaster_Target.InventoryUOM = oItemMaster.InventoryUOM
            oItemMaster_Target.PurchaseVATGroup = oItemMaster.PurchaseVATGroup
            oItemMaster_Target.GLMethod = oItemMaster.GLMethod
            oItemMaster_Target.WTLiable = oItemMaster.WTLiable
            oItemMaster_Target.PurchaseUnit = oItemMaster.PurchaseUnit

            '  MsgBox(oItemMaster.WhsInfo.ExpensesAccount & "  - " & oItemMaster.WhsInfo.ForeignExpensAcc)
            For imjs As Integer = 0 To oItemMaster.WhsInfo.Count - 1
                oItemMaster.WhsInfo.SetCurrentLine(imjs)
                oItemMaster_Target.WhsInfo.WarehouseCode = oItemMaster.WhsInfo.WarehouseCode
                oItemMaster_Target.WhsInfo.ExpensesAccount = oItemMaster.WhsInfo.ExpensesAccount
                oItemMaster_Target.WhsInfo.ForeignExpensAcc = oItemMaster.WhsInfo.ForeignExpensAcc
                oItemMaster_Target.WhsInfo.PurchaseCreditAcc = oItemMaster.WhsInfo.PurchaseCreditAcc
                oItemMaster_Target.WhsInfo.ForeignPurchaseCreditAcc = oItemMaster.WhsInfo.ForeignPurchaseCreditAcc
                oItemMaster_Target.WhsInfo.Add()
            Next

            oItemMaster_Target.Employee = oItemMaster.Employee
            oItemMaster_Target.Properties(1) = oItemMaster.Properties(1)
            oItemMaster_Target.Properties(2) = oItemMaster.Properties(2)
            oItemMaster_Target.Properties(3) = oItemMaster.Properties(3)
            oItemMaster_Target.Properties(4) = oItemMaster.Properties(4)
            oItemMaster_Target.Properties(5) = oItemMaster.Properties(5)
            oItemMaster_Target.Properties(6) = oItemMaster.Properties(6)
            oItemMaster_Target.Properties(7) = oItemMaster.Properties(7)
            oItemMaster_Target.Properties(8) = oItemMaster.Properties(8)
            oItemMaster_Target.Properties(9) = oItemMaster.Properties(9)
            oItemMaster_Target.Properties(10) = oItemMaster.Properties(10)
            oItemMaster_Target.Properties(11) = oItemMaster.Properties(11)
            oItemMaster_Target.Properties(12) = oItemMaster.Properties(12)

            oItemMaster_Target.User_Text = oItemMaster.User_Text
            oItemMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
            oItemMaster_Target.Valid = SAPbobsCOM.BoYesNoEnum.tYES

            'If oItemMaster.Frozen = SAPbobsCOM.BoYesNoEnum.tYES Then
            '    oItemMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tYES
            '    oItemMaster_Target.Valid = SAPbobsCOM.BoYesNoEnum.tNO
            'Else
            '    oItemMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
            '    oItemMaster_Target.Valid = SAPbobsCOM.BoYesNoEnum.tYES
            'End If
            'oItemMaster_Target.Frozen = oItemMaster.Frozen
            oItemMaster_Target.FrozenFrom = oItemMaster.FrozenFrom
            oItemMaster_Target.FrozenTo = oItemMaster.FrozenTo
            oItemMaster_Target.ValidFrom = oItemMaster.ValidFrom
            oItemMaster_Target.ValidTo = oItemMaster.ValidTo

            oItemMaster_Target.UserFields.Fields.Item("U_AB_ITEMTYPE").Value = oItemMaster.UserFields.Fields.Item("U_AB_ITEMTYPE").Value
            oItemMaster_Target.UserFields.Fields.Item("U_AB_ITEMSUBGROUP").Value = oItemMaster.UserFields.Fields.Item("U_AB_ITEMSUBGROUP").Value

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        End Sub

        Public Sub BP_Assignment(ByRef oBPMaster As SAPbobsCOM.BusinessPartners, ByRef oBPMaster_Target As SAPbobsCOM.BusinessPartners)

            sFuncName = "BP_Assignment()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oBPMaster_Target.CardName = oBPMaster.CardName
            '' oBPMaster_Target.Series = oBPMaster.Series
            oBPMaster_Target.CardType = oBPMaster.CardType
            oBPMaster_Target.GroupCode = oBPMaster.GroupCode
            oBPMaster_Target.Currency = oBPMaster.Currency
            oBPMaster_Target.FederalTaxID = oBPMaster.FederalTaxID
            'GENERAL TAB
            oBPMaster_Target.Phone1 = oBPMaster.Phone1
            oBPMaster_Target.Phone2 = oBPMaster.Phone2
            oBPMaster_Target.Fax = oBPMaster.Fax
            oBPMaster_Target.EmailAddress = oBPMaster.EmailAddress
            oBPMaster_Target.ContactPerson = oBPMaster.ContactPerson
            'CONTACT PERSON  TAB
            If oBPMaster_Target.ContactEmployees.Count = 0 Then
                oBPMaster_Target.ContactEmployees.Add()
                oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2
            Else
                ' oBPMaster_Target.ContactEmployees.Add()
                oBPMaster_Target.ContactEmployees.SetCurrentLine(0)
                oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2
            End If

            'ADDRESS TAB
            If oBPMaster_Target.Addresses.Count = 0 Then
                For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                    oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                    oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                    oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                    oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                    oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                    oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                    oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                    oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                    oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    oBPMaster_Target.Addresses.Add()
                Next imjs
            Else

                For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                    oBPMaster.Addresses.SetCurrentLine(imjs)
                    If imjs <= oBPMaster_Target.Addresses.Count - 1 Then
                        oBPMaster_Target.Addresses.SetCurrentLine(imjs)

                        If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName) Then
                            oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                        End If
                        If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName2) Then
                            oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                        End If
                        If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName3) Then
                            oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                        End If
                        If Not String.IsNullOrEmpty(oBPMaster.Addresses.Street) Then
                            oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                        End If
                        If Not String.IsNullOrEmpty(oBPMaster.Addresses.Block) Then
                            oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                        End If
                        If Not String.IsNullOrEmpty(oBPMaster.Addresses.City) Then
                            oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                        End If
                        If Not String.IsNullOrEmpty(oBPMaster.Addresses.ZipCode) Then
                            oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                        End If
                        If Not String.IsNullOrEmpty(oBPMaster.Addresses.Country) Then
                            oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                        End If
                        If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressType) Then
                            oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                        End If

                    Else
                        oBPMaster_Target.Addresses.Add()
                        oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                        oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                        oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                        oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                        oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                        oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                        oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                        oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                        oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    End If
                Next imjs
            End If
            'PAYMENT TERMS TAB
            oBPMaster_Target.PayTermsGrpCode = oBPMaster.PayTermsGrpCode

            If oBPMaster_Target.BPBankAccounts.Count > 0 Then
                For imjs As Integer = 0 To oBPMaster_Target.BPBankAccounts.Count - 1
                    oBPMaster_Target.BPBankAccounts.SetCurrentLine(0)
                    oBPMaster_Target.BPBankAccounts.Delete()
                Next
            End If

            For imjs As Integer = 0 To oBPMaster.BPBankAccounts.Count - 1
                oBPMaster.BPBankAccounts.SetCurrentLine(imjs)
                oBPMaster_Target.BPBankAccounts.Country = oBPMaster.BPBankAccounts.Country
                oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo
                oBPMaster_Target.BPBankAccounts.BankCode = oBPMaster.BPBankAccounts.BankCode
                oBPMaster_Target.BPBankAccounts.BPCode = oBPMaster.BPBankAccounts.BPCode
                oBPMaster_Target.BPBankAccounts.Branch = oBPMaster.BPBankAccounts.Branch
                oBPMaster_Target.BPBankAccounts.IBAN = oBPMaster.BPBankAccounts.IBAN
                oBPMaster_Target.BPBankAccounts.AccountName = oBPMaster.BPBankAccounts.AccountName
                oBPMaster_Target.BPBankAccounts.BICSwiftCode = oBPMaster.BPBankAccounts.BICSwiftCode

                oBPMaster_Target.BPBankAccounts.Add()
            Next
            ''  oBPMaster_Target.SaveXML("E:\Abeo-Projects\PWC\SVN\Master Data Synchronization\AE_PWC_AO03\bin\Debug\xml\BP.xml")
            'PAYMENT RUN TAB
            oBPMaster_Target.HouseBankCountry = oBPMaster.HouseBankCountry
            oBPMaster_Target.HouseBank = oBPMaster.HouseBank
            oBPMaster_Target.HouseBankAccount = oBPMaster.HouseBankAccount
            'ACCOUNTING TAB
            If oBPMaster_Target.AccountRecivablePayables.Count = 0 Then
                oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
                oBPMaster_Target.AccountRecivablePayables.Add()
            Else
                oBPMaster_Target.AccountRecivablePayables.SetCurrentLine(0)
                oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
            End If

            oBPMaster_Target.VatLiable = oBPMaster.VatLiable
            ' oBPMaster_Target.WithholdingTaxCertified = oBPMaster.WithholdingTaxCertified
            oBPMaster_Target.VatGroup = oBPMaster.VatGroup
            oBPMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO

            oBPMaster_Target.FreeText = oBPMaster.FreeText
            oBPMaster_Target.Frozen = oBPMaster.Frozen
            If Not String.IsNullOrEmpty(oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value) Then
                oBPMaster_Target.UserFields.Fields.Item("U_AB_WTAXREQ").Value = oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value
            End If
            oBPMaster_Target.PeymentMethodCode = oBPMaster.PeymentMethodCode
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        End Sub

        Public Sub BP_Assignment_Old(ByRef oBPMaster As SAPbobsCOM.BusinessPartners, ByRef oBPMaster_Target As SAPbobsCOM.BusinessPartners)

            sFuncName = "BP_Assignment()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oBPMaster_Target.CardName = oBPMaster.CardName
            oBPMaster_Target.Series = oBPMaster.Series
            oBPMaster_Target.CardType = oBPMaster.CardType
            oBPMaster_Target.GroupCode = oBPMaster.GroupCode
            oBPMaster_Target.Currency = oBPMaster.Currency
            oBPMaster_Target.FederalTaxID = oBPMaster.FederalTaxID
            'GENERAL TAB
            oBPMaster_Target.Phone1 = oBPMaster.Phone1
            oBPMaster_Target.Phone2 = oBPMaster.Phone2
            oBPMaster_Target.Fax = oBPMaster.Fax
            oBPMaster_Target.EmailAddress = oBPMaster.EmailAddress
            oBPMaster_Target.ContactPerson = oBPMaster.ContactPerson
            'CONTACT PERSON  TAB
            If oBPMaster_Target.ContactEmployees.Count = 0 Then
                oBPMaster_Target.ContactEmployees.Add()
                oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2
            Else
                ' oBPMaster_Target.ContactEmployees.Add()
                oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2
            End If

            'ADDRESS TAB
            If oBPMaster_Target.Addresses.Count = 0 Then
                For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                    oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                    oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                    oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                    oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                    oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                    oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                    oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                    oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                    oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    oBPMaster_Target.Addresses.Add()
                Next imjs
            Else
                For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                    oBPMaster.Addresses.SetCurrentLine(imjs)
                    If imjs <= oBPMaster_Target.Addresses.Count - 1 Then
                        oBPMaster_Target.Addresses.SetCurrentLine(imjs)
                        oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                        oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                        oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                        oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                        oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                        oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                        oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                        oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                        oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    Else
                        oBPMaster_Target.Addresses.Add()
                        oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                        oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                        oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                        oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                        oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                        oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                        oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                        oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                        oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    End If
                Next imjs
            End If
            'PAYMENT TERMS TAB
            oBPMaster_Target.PayTermsGrpCode = oBPMaster.PayTermsGrpCode
            oBPMaster_Target.BankCountry = oBPMaster.BankCountry
            oBPMaster_Target.DefaultBankCode = oBPMaster.DefaultBankCode
            MsgBox(oBPMaster.DefaultAccount)
            oBPMaster_Target.DefaultAccount = oBPMaster.DefaultAccount
            oBPMaster_Target.DefaultBranch = oBPMaster.DefaultBranch
            oBPMaster_Target.IBAN = oBPMaster.HouseBankIBAN
            'PAYMENT RUN TAB
            oBPMaster_Target.HouseBankCountry = oBPMaster.HouseBankCountry
            oBPMaster_Target.HouseBank = oBPMaster.HouseBank
            oBPMaster_Target.HouseBankAccount = oBPMaster.HouseBankAccount
            'ACCOUNTING TAB
            ' oBPMaster_Target.AccountRecivablePayables = oBPMaster.AccountRecivablePayables
            oBPMaster_Target.VatLiable = oBPMaster.VatLiable
            ' oBPMaster_Target.WithholdingTaxCertified = oBPMaster.WithholdingTaxCertified
            oBPMaster_Target.VatGroup = oBPMaster.VatGroup
            oBPMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        End Sub

    End Module
End Namespace


