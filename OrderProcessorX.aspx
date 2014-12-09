<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Microsoft.VisualBasic" %>
<script runat="server">

    ' add checks for duplicate column directives (eg 2 components assigned to same column)
    ' $ TEMPLATE directive for WURS/WUIRE specific settings
    ' $ NOTOWN
    ' $ NOPOSTCODE
    ' $ USER <user ID>
    ' allow product keys / user keys?
    ' place ORDER... csv files in sub-directory
    
    ' do we need to allow duplicate order items in gdictOrderItems?
        
    Const DEFAULT_ITEM_ORDER As String = "PQ"
    'Const DEFAULT_ITEM_COLUMN_START As String = "B"    ' CN 24OCT13
    
    Const MAX_ADDRESS_FIELD_LENGTH As Int32 = 50
    Const MAX_EXTERNAL_REFERENCE_LENGTH As Int32 = 50
    Const MAX_CUSTREF_12_FIELD_LENGTH As Int32 = 25
    Const MAX_CUSTREF_34_FIELD_LENGTH As Int32 = 50
    Const COUNTRY_CODE_UK As Integer = 222
    Const HTML_BREAK As String = "<br />"
    
    Const CUSTOMER_COSTA As Int32 = 826
    Const CUSTOMER_FININT As Int32 = 798
        
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private gnCustomerKey As Int32
    Private gnUserKey As Int32
    Private gdtOrderData As DataTable, gdtConflatedOrderData As DataTable
    Private lstConsignmentNos As List(Of Int32)
    
    Dim gsConsignee As String
    Dim gsConsigneeContact As String
    Dim gsAddr1 As String
    Dim gsAddr2 As String
    Dim gsAddr3 As String
    Dim gsTown As String
    Dim gsRegion As String
    Dim gsPostcode As String
    Dim gsCountry As String
    Dim gnCountryCode As Int32
    Dim gsCustRef1 As String
    Dim gsCustRef2 As String
    Dim gsCustRef3 As String
    Dim gsCustRef4 As String
    Dim gsSpecialInstructions As String
    Dim gsPackingNote As String
    Dim gsExternalReference As String
        
    'Private gbHaveLocationCustomer As Boolean = False
    'Private gbHaveLocationCnee As Boolean = False
    'Private gbHaveLocationAddr1 As Boolean = False
    'Private gbHaveLocationTown As Boolean = False     ' new directive NOTOWN
    'Private gbHaveLocationPostCode As Boolean = False ' new directive NOPOSTCODE
    'Private gbHaveLocationCountry As Boolean = False

    Private gsColumnConsignee As String
    Private gsColumnConsigneeContact As String
    Private gsColumnAddr1 As String
    Private gsColumnAddr2 As String
    Private gsColumnAddr3 As String
    Private gsColumnTown As String
    Private gsColumnRegion As String
    Private gsColumnPostcode As String
    Private gsColumnCountry As String

    Private gsColumnCustRef1 As String
    Private gsColumnCustRef2 As String
    Private gsColumnCustRef3 As String
    Private gsColumnCustRef4 As String
    
    Private gsColumnSpecialInstructions As String
    Private gsColumnPackingNote As String
    Private gsColumnExternalReference As String

    Private gsDefaultCustRef1 As String
    Private gsDefaultCustRef2 As String
    Private gsDefaultCustRef3 As String
    Private gsDefaultCustRef4 As String

    Private gsDefaultSpecialInstructions As String
    Private gsDefaultPackingNote As String
    Private gsDefaultExternalReference As String

    Private gsItemOrder As String
    Private gsItemColumnStart As String     ' CN 24OCT13

    Private gbDefaultUK As Boolean = False
    
    Private gbTemplateWU As Boolean = False
    
    Private gbNoTown As Boolean = False
    Private gbNoPostcode As Boolean = False
    
    Private gbOrderFound As Boolean = False
    
    Private gdictOrderItems As Dictionary(Of Int32, Int32) = New Dictionary(Of Int32, Int32)
    Private gdictTotalOrderItems As Dictionary(Of Int32, Int32) = New Dictionary(Of Int32, Int32)
    Private bFoundOrder As Boolean
    
    Private gdictCountryAliases As Dictionary(Of String, String) = New Dictionary(Of String, String)
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call CreateOrderProcessorFolder()
            Call PopulateCustomerDropdown()
            tbOrder.Focus()
        End If
        Call SetTitle()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Order Processor"
    End Sub
    
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sm As New ScriptManager
        sm.ID = "ScriptMgr"
        Try
            PlaceHolderForScriptManager.Controls.Add(sm)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PopulateCustomerDropdown()
        ddlCustomer.Items.Clear()
        Dim sSQL As String = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerStatusID = 'Active' AND ISNULL(AccountHandlerKey, 0) > 0 ORDER BY CustomerAccountCode"
        Dim dtCustomers As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlCustomer.Items.Add(New ListItem("- please select -", 0))
        For Each drCustomer In dtCustomers.Rows
            ddlCustomer.Items.Add(New ListItem(drCustomer("CustomerAccountCode"), drCustomer("CustomerKey")))
        Next
    End Sub

    Private Sub CreateOrderProcessorFolder()
        Dim sPath As String = Server.MapPath("~/")
        If Not Directory.Exists(sPath & "\OrderProcessor") Then
            Directory.CreateDirectory(sPath & "\OrderProcessor")
        End If
    End Sub

    Protected Sub CreateCleanOrderFileFromTextbox()
        psUniqueFilename = Server.MapPath("~/OrderProcessor/") & "\" & "ORDER_" & Format(Now(), "yyyymmddhhmmssff") & ".csv"
        Dim sw As New StreamWriter(psUniqueFilename)
        Dim sLine() As String = tbOrder.Text.Split(Environment.NewLine)
        For Each s As String In sLine
            s = s.Replace("|", vbTab).Trim
            If s.Trim.ToLower.Contains("prodref:") And Not (s.Trim.ToLower.Contains("$comment") Or s.Trim.ToLower.Contains("$ comment")) Then
                s = s.Replace("$ prodref:", "$prodref:")
                While s.StartsWith(".")
                    s = s.Substring(1)
                End While
                Dim nMidPoint As Int32 = s.ToLower.IndexOf("$prodref:")
                nMidPoint += 9
                While IsNumeric(s.Substring(nMidPoint, 1))
                    nMidPoint += 1
                End While
                Dim sRemains As String = s.Substring(nMidPoint, s.Length - nMidPoint)
                While Not (sRemains.StartsWith(vbTab) Or sRemains.StartsWith(Environment.NewLine))
                    sRemains = sRemains.Substring(1)
                End While
                s = s.Substring(0, nMidPoint) & sRemains
            End If
            
            If Not (s.Trim.ToLower.StartsWith("$comment") Or s.Trim.ToLower.StartsWith("$ comment") Or s.Trim.Replace(vbTab, "") = String.Empty) Then
                sw.Write(s & Environment.NewLine)
            End If
        Next
        sw.Close()
    End Sub
    
    Protected Function CheckCleanOrderFile(ByVal bExecuteOrder As Boolean) As Boolean
        Dim bErrorFound As Boolean = False
        Dim sbResult As New StringBuilder
        Dim sr As New StreamReader(psUniqueFilename)
        Dim sLine As String
        Dim nLineNo As Int32 = 1
        Do While sr.Peek >= 0
            sLine = sr.ReadLine()
            If sLine.Trim <> String.Empty Then
                Dim sParseResult As String = TryParse(sLine, bExecuteOrder:=bExecuteOrder)
                If sParseResult.StartsWith("-") Then
                    bErrorFound = True
                    sParseResult = sParseResult.Substring(1, sParseResult.Length - 1)
                    If Not sParseResult.Contains("font color") Then
                        sParseResult = "<font color = 'red'>" & sParseResult & "</font>"
                    End If
                Else
                    sParseResult = sParseResult.Substring(1, sParseResult.Length - 1)    ' ????? Try / Catch ?
                End If
                If sParseResult <> String.Empty Then
                    sbResult.Append("Line " & nLineNo.ToString & ":&nbsp;" & sParseResult)
                    sbResult.Append("<br />")
                End If
            End If
            nLineNo += 1
        Loop
        sr.Close()
    
        sbResult.Append("- all orders scanned -")
        sbResult.Append(HTML_BREAK)
        sbResult.Append(HTML_BREAK)

        If Not bErrorFound Then
            Dim sTotalQtyRequired As String = CheckTotalQtyRequired()
            bErrorFound = sTotalQtyRequired.StartsWith("-")
            sTotalQtyRequired = sTotalQtyRequired.Substring(1, sTotalQtyRequired.Length - 1)
            sbResult.Append(sTotalQtyRequired)
            sbResult.Append("<br />")
        End If

        If Not bErrorFound Then
            sbResult.Append("ORDER DATA VALIDATED SUCCESSFULLY. Click the Process Orders button to place order(s).")
            CheckCleanOrderFile = True
        Else
            sbResult.Append("ONE OR MORE ERRORS FOUND. PLEASE CORRECT THE ORDER DATA.")
            CheckCleanOrderFile = False
        End If
        If bExecuteOrder AndAlso Not bErrorFound Then
            lblResult.Text = "CONSIGNMENTS GENERATED: "
            For Each sConsignmentNo In lstConsignmentNos
                lblResult.Text &= sConsignmentNo.ToString & ", "
            Next
            lblResult.Text = lblResult.Text.Substring(0, lblResult.Text.Length - 2)
        Else
            lblResult.Text = sbResult.ToString
        End If
    End Function
    
    Protected Function CheckTotalQtyRequired() As String
        Dim sbResult As New StringBuilder
        sbResult.Append("TOTAL QUANTITIES REQUIRED SUMMARY" & HTML_BREAK)
        Dim bInsufficientQtyAvailable As Boolean = False
        For Each kv As KeyValuePair(Of Int32, Int32) In gdictTotalOrderItems
            Dim nAvailableQty As Int32 = GetAvailableQty(kv.Key)
            sbResult.Append("Qty Required: ")
            sbResult.Append(kv.Value)
            sbResult.Append(", ")
            sbResult.Append("Qty Available: ")
            sbResult.Append(nAvailableQty)
            sbResult.Append(". PRODUCT: ")
            sbResult.Append(GetProductDetailsFromProductKey(kv.Key))
            If nAvailableQty < kv.Value Then
                bInsufficientQtyAvailable = True
                sbResult.Append(" <font color='red'> (INSUFFICIENT)</font>")
            End If
            sbResult.Append(HTML_BREAK)
        Next
        If bInsufficientQtyAvailable Then
            Return "-" & sbResult.ToString
        Else
            Return "+" & sbResult.ToString
        End If
    End Function
    
    Protected Function TryParse(ByVal sLine As String, ByVal bExecuteOrder As Boolean) As String
        Dim sCol1Qualifier As String = String.Empty
        TryParse = String.Empty
        sLine = sLine.Trim
        If sLine <> String.Empty Then
            Dim arrLine() As String = sLine.Split(vbTab)
            If arrLine(0).StartsWith("$") Then
                ' processing directive
                arrLine(0) = arrLine(0).ToLower

                If arrLine.Count < 2 Then
                    Return "-Not enough data (directive: " & arrLine(0).ToUpper & ")."
                End If
                sCol1Qualifier = arrLine(1).Trim

                If arrLine(0).Contains("customer") Then
                    Dim nCustomerKey As Int32 = GetCustomerKeyFromCustomerCode(sCol1Qualifier)
                    If nCustomerKey <> gnCustomerKey Then
                        gnUserKey = 0
                    End If
                    If nCustomerKey <= 0 Then
                        If sCol1Qualifier = String.Empty Then
                            TryParse = "-Missing customer code."
                        Else
                            TryParse = "-Could not identify user " & sCol1Qualifier & " as an active customer."
                        End If
                    Else
                        gnCustomerKey = nCustomerKey
                        TryParse = "+Set customer " & arrLine(1) & " as order customer."
                    End If
                ElseIf arrLine(0).Contains("user") Then
                    If gnCustomerKey = 0 Then
                        TryParse = "-Cannot process USER directive before CUSTOMER directive."
                    End If
                    Dim nUserKey As Int32 = GetUserKeyFromUserID(sCol1Qualifier)
                    If nUserKey <= 0 Then
                        If sCol1Qualifier = String.Empty Then
                            TryParse = "-Missing UserID."
                        Else
                            TryParse = "-Could not identify " & sCol1Qualifier & " as an active user."
                        End If
                    Else
                        If nUserKey <> GetUserKeyFromUserID(sCol1Qualifier, gnCustomerKey) Then
                            TryParse = "-This user account (" & sCol1Qualifier & ") does not belong to the CUSTOMER you specified."
                        Else
                            gnUserKey = nUserKey
                            TryParse = "+Set User ID " & arrLine(1) & " as placing order."
                        End If
                    End If
                ElseIf arrLine(0).Contains("consignee") And Not arrLine(0).Contains("contact") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "CONSIGNEE", gsColumnConsignee)
                ElseIf arrLine(0).Contains("consigneecontact") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "CONSIGNEECONTACT", gsColumnConsigneeContact)
                ElseIf arrLine(0).Contains("addr1") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "ADDR1", gsColumnAddr1)
                ElseIf arrLine(0).Contains("addr2") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "ADDR2", gsColumnAddr2)
                ElseIf arrLine(0).Contains("addr3") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "ADDR3", gsColumnAddr3)
                ElseIf arrLine(0).Contains("town") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "TOWN", gsColumnTown)
                ElseIf arrLine(0).Contains("region") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "REGION", gsColumnRegion)
                ElseIf arrLine(0).Contains("postcode") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "POSTCODE", gsColumnPostcode)
                ElseIf arrLine(0).Contains("country") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "COUNTRY", gsColumnCountry)
                ElseIf arrLine(0).Contains("custref1") And Not arrLine(0).Contains("default") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "CUSTREF1", gsColumnCustRef1)
                ElseIf arrLine(0).Contains("custref2") And Not arrLine(0).Contains("default") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "CUSTREF2", gsColumnCustRef2)
                ElseIf arrLine(0).Contains("custref3") And Not arrLine(0).Contains("default") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "CUSTREF3", gsColumnCustRef3)
                ElseIf arrLine(0).Contains("custref4") And Not arrLine(0).Contains("default") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "CUSTREF4", gsColumnCustRef4)
                ElseIf arrLine(0).Contains("specialinstructions") And Not arrLine(0).Contains("default") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "SPECIALINSTRUCTIONS", gsColumnSpecialInstructions)
                ElseIf arrLine(0).Contains("packingnote") And Not arrLine(0).Contains("default") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "PACKINGNOTE", gsColumnPackingNote)
                ElseIf arrLine(0).Contains("externalreference") And Not arrLine(0).Contains("default") Then
                    TryParse = ProcessColumnDirective(arrLine(0), sCol1Qualifier, "EXTERNALREFERENCE", gsColumnExternalReference)
                ElseIf arrLine(0).Contains("defaultcustref1") Then
                    TryParse = ProcessDefaultValueDirective(arrLine(0), sCol1Qualifier, "DEFAULTCUSTREF1", gsDefaultCustRef1)
                ElseIf arrLine(0).Contains("defaultcustref2") Then
                    TryParse = ProcessDefaultValueDirective(arrLine(0), sCol1Qualifier, "DEFAULTCUSTREF2", gsDefaultCustRef2)
                ElseIf arrLine(0).Contains("defaultcustref3") Then
                    TryParse = ProcessDefaultValueDirective(arrLine(0), sCol1Qualifier, "DEFAULTCUSTREF3", gsDefaultCustRef3)
                ElseIf arrLine(0).Contains("defaultcustref4") Then
                    TryParse = ProcessDefaultValueDirective(arrLine(0), sCol1Qualifier, "DEFAULTCUSTREF4", gsDefaultCustRef4)
                ElseIf arrLine(0).Contains("defaultspecialinstructions") Then
                    TryParse = ProcessDefaultValueDirective(arrLine(0), sCol1Qualifier, "DEFAULTSPECIALINSTRUCTIONS", gsDefaultSpecialInstructions)
                ElseIf arrLine(0).Contains("defaultpackingnote") Then
                    TryParse = ProcessDefaultValueDirective(arrLine(0), sCol1Qualifier, "DEFAULTPACKINGNOTE", gsDefaultPackingNote)
                ElseIf arrLine(0).Contains("defaultexternalreference") Then
                    TryParse = ProcessDefaultValueDirective(arrLine(0), sCol1Qualifier, "DEFAULTEXTERNALREFERENCE", gsDefaultExternalReference)
                ElseIf arrLine(0).Contains("alias") Then
                    TryParse = ProcessAliasDirective(arrLine(0), sCol1Qualifier, "ALIAS")
                ElseIf arrLine(0).Contains("template") Then
                    TryParse = ProcessTemplateDirective(arrLine(0), sCol1Qualifier, "TEMPLATE")
                ElseIf arrLine(0).Contains("defaultuk") Then
                    gbDefaultUK = True
                    TryParse = "+Set DEFAULTUK" & HTML_BREAK
                ElseIf arrLine(0).Contains("notown") Then
                    gbNoTown = True
                    TryParse = "+Set NOTOWN" & HTML_BREAK
                ElseIf arrLine(0).Contains("nopostcode") Then
                    gbNoPostcode = True
                    TryParse = "+Set NOPOSTCODE" & HTML_BREAK
                ElseIf arrLine(0).Contains("itemorder") Then
                    TryParse = ProcessItemOrderDirective(arrLine(0), sCol1Qualifier, "ITEMORDER", gsItemOrder)
                ElseIf arrLine(0).Contains("item") And Not arrLine(0).Contains("order") Then
                    TryParse = ProcessItemDirective(arrLine, "ITEM")
                ElseIf arrLine(0).Contains("prodref") Then
                    TryParse = ProcessProdRefDirective(arrLine, "PRODREF")
                Else
                    TryParse = "-Unknown directive " & arrLine(0) & " !!"
                End If
            Else
                If gnCustomerKey = 0 Then
                    Return "-Cannot process order line before finding CUSTOMER directive."
                End If
                If gnUserKey = 0 Then
                    Return "-Cannot process order line before finding USER directive. If you already defined the USER but then changed the CUSTOMER, you must redefine the USER."
                End If
                Call ClearAddressFields()
                Dim sParseOrderLineResult As String = ParseOrderLine(arrLine, bExecuteOrder:=bExecuteOrder)
                If sParseOrderLineResult.Substring(0, 1) = "+" Then
                    sParseOrderLineResult = sParseOrderLineResult.Substring(7)
                    If bExecuteOrder Then
                        If gdictOrderItems.Count > 0 Then
                            Dim nConsignmentNo As Int32 = nSubmitConsignment(gdictOrderItems)
                            lstConsignmentNos.Add(nConsignmentNo)
                        End If
                    End If
                    
                End If
                TryParse = "+" & ListOrderItems() & HTML_BREAK & sParseOrderLineResult
            End If
        End If
    End Function

    Protected Function ListOrderItems() As String
        ListOrderItems = String.Empty
        If gdictOrderItems.Count > 0 Then
            ListOrderItems = "BASKET..." & HTML_BREAK
            For Each kv As KeyValuePair(Of Int32, Int32) In gdictOrderItems
                ListOrderItems &= "Product: " & GetProductCodeFromProductKey(kv.Key).ToString & " Qty: " & kv.Value & HTML_BREAK
            Next
        Else
            Return "no order items"
        End If
    End Function
    
    Protected Function ProcessTemplateDirective(ByVal sCol0 As String, ByVal sCol1 As String, ByVal sDirective As String) As String
        If sCol1.Length > 0 Then
            If sCol1.ToLower.Contains("wu") Or sCol1.ToLower.Contains("western") Then
                gbTemplateWU = True
            End If
        Else
            gbTemplateWU = False
        End If
    End Function

    Protected Function ProcessAliasDirective(ByVal sCol0 As String, ByVal sCol1 As String, ByVal sDirective As String) As String
        ProcessAliasDirective = String.Empty
        Dim arrLine() As String = sCol1.Split("=")
        If arrLine.Count < 2 Then
            Return "-Not enough data in ALIAS directive: " & sCol1.ToUpper
        End If
        If arrLine.Count > 2 Then
            Return "-Too much data in ALIAS directive: " & sCol1.ToUpper
        End If

        Dim sAliasCountryName = arrLine(0).Trim.ToUpper
        Dim sAIMSCountryName = arrLine(1).Trim.ToUpper
        
        If sAliasCountryName = String.Empty Or sAIMSCountryName = String.Empty Then
            Return "-Not enough data in ALIAS directive: " & sCol1.ToUpper
        End If
        
        Dim sSQL As String = "SELECT CountryKey FROM Country WHERE CountryName = '" & sAIMSCountryName.Replace("'", "''") & "'"
        Dim dtCountry As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtCountry.Rows.Count = 0 Then
            Return "-Could not match to an AIMS country Name: " & sCol1.ToUpper
        ElseIf dtCountry.Rows.Count > 1 Then
            Return "-Multiple country names matched! How could this happen???? " & sCol1.ToUpper
        End If
        
        If gdictCountryAliases.ContainsKey(sAliasCountryName) Then
            Return "-This alias is already defined: " & sCol1.ToUpper
        End If
        
        gdictCountryAliases.Add(sAliasCountryName, sAIMSCountryName)
        ProcessAliasDirective = "+Country " & sAliasCountryName & " aliased to " & sAIMSCountryName
    End Function

    Protected Function ProcessProdRefDirective(ByVal arrLine() As String, ByVal sDirective As String) As String
        If arrLine.Count < 2 Then
            Return "-Not enough data."
        End If
        Dim nLogisticProductKey As Int32 = -1
        Dim nQty As Int32 = -1
        If gnCustomerKey = 0 Then
            Return "-Cannot process a PRODREF directive before processing CUSTOMER directive."
        End If
        Dim sLogisticProductKey As String = arrLine(0).Substring(9)
        If IsNumeric(sLogisticProductKey) Then
            nLogisticProductKey = CInt(sLogisticProductKey)
            If nLogisticProductKey > 0 Then
                Dim dtProduct As DataTable = ExecuteQueryToDataTable("SELECT CustomerKey, ArchiveFlag, DeletedFlag FROM LogisticProduct WHERE LogisticProductKey = " & nLogisticProductKey)
                If dtProduct.Rows.Count = 1 Then
                    Dim drProduct As DataRow = dtProduct.Rows(0)
                    If drProduct("CustomerKey") = gnCustomerKey Then
                        If drProduct("DeletedFlag") = "N" Then
                            If drProduct("ArchiveFlag") = "Y" Then
                                Return "-The product referenced is marked as ARCHIVED."
                            End If
                        Else
                            Return "-The product referenced is marked as DELETED."
                        End If
                    Else
                        Return "-The product referenced does not belong to the customer named in the CUSTOMER directive."
                    End If
                Else
                    Return "-No product with that PRODREF value exists."
                End If
                '
                ' check here that product is for this customer and exists and is not deleted or archived
                '
            Else
                Return "-PRODREF value must be a positive, non-zero, number."
            End If
        Else
            Return "-Could not extract PRODREF value."
        End If
        If IsNumeric(arrLine(1)) Then
            nQty = CInt(arrLine(1))
            If nQty <= 0 Then
                Return "-Quantity must be greater than 0."
            End If
        Else
            Return "-non-numeric quantity."
        End If
        If gbOrderFound Then
            gdictOrderItems.Clear()
            gbOrderFound = False
        End If
        Dim nAvailableQty As Int32 = GetAvailableQty(nLogisticProductKey)
        If nAvailableQty = 0 Then
            Return "-This product (" & GetProductCodeFromProductKey(nLogisticProductKey) & ") has an available quantity of zero."
        End If
        If nAvailableQty < nQty Then
            Return "-This product (" & GetProductCodeFromProductKey(nLogisticProductKey) & ") has has insufficient available quantity (" & nAvailableQty.ToString & " to fulfil the order requirement (" & nQty & ")."
        End If
        If gdictOrderItems.ContainsKey(nLogisticProductKey) Then
            gdictOrderItems(nLogisticProductKey) = gdictOrderItems(nLogisticProductKey) + nQty
        Else
            gdictOrderItems.Add(nLogisticProductKey, nQty)
        End If
        Dim sMessage As String = "+Added item " & GetProductCodeFromProductKey(nLogisticProductKey)
        sMessage &= " to order contents list."
        Return sMessage
    End Function
    
    Protected Function ProcessItemDirective(ByVal arrLine() As String, ByVal sDirective As String) As String
        Dim sProductCode As String = String.Empty
        Dim sProductDate As String = String.Empty
        Dim nQty As Int32 = -1
        If gnCustomerKey = 0 Then
            Return "-Cannot process an ITEM directive before processing CUSTOMER directive."
        End If

        If gbOrderFound Then
            gdictOrderItems.Clear()
            gbOrderFound = False
        End If

        Dim nItemColumnStart As Int32    ' need to set start position somewhere here, using something similar to GetAddressComponent    CN 24OCT13
        
        Dim nLogisticProductKey As Int32 = 0
        For i As Int32 = 0 To gsItemOrder.Length - 1
            Select Case gsItemOrder.ToLower.Substring(i, 1)
                Case "q"
                    Dim sQty As String = String.Empty
                    Try
                        sQty = arrLine(i + 1)
                    Catch ex As Exception
                        Return "-No item quantity field found."
                    End Try
                    If sQty <> String.Empty Then
                        If IsNumeric(sQty) AndAlso CInt(sQty) > 0 Then
                            nQty = CInt(sQty)
                        Else
                            Return "-Item quantity non-numeric or less than 1."
                        End If
                    Else
                        Return "-Item quantity value not found."
                    End If
                Case "p"
                    Try
                        sProductCode = arrLine(i + 1).Trim
                    Catch ex As Exception
                        Return "-No product code field found."
                    End Try
                    If gsItemOrder.ToLower.Contains("pd") Then
                        Try
                            sProductDate = arrLine(i + 2)
                        Catch ex As Exception
                            'Return "-No product date field found."
                        End Try
                    Else
                        Dim nProductCount As Int32 = GetCountOfProductsWithThisProductCode(sProductCode)
                        If nProductCount > 1 Then
                            Return "-Multiple products qualified by product date found for this account, but no product date column specified."
                        End If
                    End If
                    If sProductCode = String.Empty Then
                        Return "-No product code."
                    End If
                    nLogisticProductKey = GetProductKeyFromProductCode(sProductCode, sProductDate)
                    If nLogisticProductKey < 0 Then
                        Return "-There is more than one product with this product code (" & sProductCode & ")."
                    ElseIf nLogisticProductKey = 0 Then
                        Return "-Cannot match this product (" & sProductCode & ")."
                    End If
            End Select
        Next
        Dim nAvailableQty As Int32 = GetAvailableQty(nLogisticProductKey)
        If nAvailableQty = 0 Then
            Return "-This product (" & sProductCode & "has an available quantity of zero."
        End If
        If nAvailableQty < nQty Then
            Return "-This product (" & sProductCode & "has has insufficient available quantity (" & nAvailableQty.ToString & " to fulfil the order requirement (" & nQty & ")."
        End If
        If gdictOrderItems.ContainsKey(nLogisticProductKey) Then
            gdictOrderItems(nLogisticProductKey) = gdictOrderItems(nLogisticProductKey) + nQty
        Else
            gdictOrderItems.Add(nLogisticProductKey, nQty)
        End If
        Dim sMessage As String = "+Added item " & sProductCode
        If sProductDate <> String.Empty Then
            sMessage &= " / " & sProductDate
        End If
        sMessage &= " to order contents list."
        Return sMessage
    End Function

    Protected Function ProcessItemOrderDirective(ByVal sCol0 As String, ByVal sCol1 As String, ByVal sDirective As String, ByRef sOutputVar As String) As String
        ' WHY ARE WE PASSING IN sCol0 ????
        If sCol1.Length > 0 Then
            If sCol1.ToLower = "qpd" Or sCol1.ToLower = "qp" Or sCol1.ToLower = "pdq" Or sCol1.ToLower = "pq" Then
                sOutputVar = sCol1
                ProcessItemOrderDirective = "+Set " & sDirective & " to " & sCol1 & "."
            Else
                ProcessItemOrderDirective = "-Unrecognised " & sDirective & " directive " & sCol1 & ". Ignored."
            End If
        Else
            If gsItemOrder = DEFAULT_ITEM_ORDER Then
                ProcessItemOrderDirective = "+" & sDirective & " directive found but no value specified (INFORMATIONAL). Current ITEMORDER is " & gsItemOrder & "."
            Else
                ProcessItemOrderDirective = "+" & sDirective & " directive found but no value specified. Resetting ITEMORDER to default (" & DEFAULT_ITEM_ORDER & ")."
            End If
            sOutputVar = DEFAULT_ITEM_ORDER
        End If
    End Function
    
    Protected Function ProcessDefaultValueDirective(ByVal sCol0 As String, ByVal sCol1 As String, ByVal sDirective As String, ByRef sOutputVar As String) As String
        ' WHY ARE WE PASSING IN sCol0 ????
        If sCol1.Length > 0 Then
            sOutputVar = sCol1
            ProcessDefaultValueDirective = "+Set " & sDirective & " to " & sCol1 & "."
        Else
            If sOutputVar <> String.Empty Then
                ProcessDefaultValueDirective = "+" & sDirective & " directive found, clearing previous value."
            Else
                ProcessDefaultValueDirective = "+" & sDirective & " directive found but no value specified (INFORMATIONAL)."
            End If
            sOutputVar = String.Empty
        End If
    End Function
    
    Protected Function ProcessColumnDirective(ByVal sCol0 As String, ByVal sCol1 As String, ByVal sDirective As String, ByRef sOutputVar As String) As String
        ' WHY ARE WE PASSING IN sCol0 ????
        If sCol1.Length > 0 Then
            If sCol1.Length = 1 Then
                If InAtoZ(sCol1) Then
                    If AddressColumnInUse(sCol1) Then
                        ProcessColumnDirective = "-Column " & sCol1 & " already in use for another address field, or already defined for this address field."
                    Else
                        sOutputVar = sCol1
                        ProcessColumnDirective = "+Set " & sDirective & " column to " & sCol1 & "."
                    End If
                Else
                    ProcessColumnDirective = "-Unrecognised " & sDirective & " column" & sCol1 & "."
                End If
            Else
                ProcessColumnDirective = "-Unrecognised " & sDirective & " column" & sCol1 & "."
            End If
        Else
            If sOutputVar <> String.Empty Then
                ProcessColumnDirective = "+" & sDirective & " directive found, clearing previous value."
            Else
                ProcessColumnDirective = "+" & sDirective & " directive found but no column specified (INFORMATIONAL)."
            End If
            sOutputVar = String.Empty
        End If
    End Function
    
    Protected Function AddressColumnInUse(sColumnLetter As String) As Boolean
        AddressColumnInUse = False
        If gsColumnConsignee = sColumnLetter Or gsColumnConsigneeContact = sColumnLetter Or gsColumnAddr1 = sColumnLetter Or gsColumnAddr2 = sColumnLetter Or gsColumnAddr3 = sColumnLetter Or gsColumnTown = sColumnLetter Or gsColumnRegion = sColumnLetter Or gsColumnPostcode = sColumnLetter Or gsColumnCountry = sColumnLetter Then
            AddressColumnInUse = True
        End If
    End Function
    
    Protected Function GetCustomerKeyFromCustomerCode(ByVal sCustomerCode As String) As Int32
        GetCustomerKeyFromCustomerCode = -1
        Dim sSQL As String = "SELECT CustomerKey FROM Customer WHERE DeletedFlag = 'N' AND CustomerStatusId = 'ACTIVE' AND CustomerAccountCode = '" & sCustomerCode.Replace("'", "''") & "'"
        Dim dtCustomer As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtCustomer.Rows.Count = 1 Then
            GetCustomerKeyFromCustomerCode = dtCustomer.Rows(0).Item(0)
        End If
    End Function

    Protected Function GetUserKeyFromUserID(ByVal sUserID As String) As Int32
        GetUserKeyFromUserID = -1
        Dim sSQL As String = "SELECT [key] FROM UserProfile WHERE DeletedFlag = 0 AND Status = 'ACTIVE' AND UserID = '" & sUserID.Replace("'", "''") & "'"
        Dim dtUser As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtUser.Rows.Count = 1 Then
            GetUserKeyFromUserID = dtUser.Rows(0).Item(0)
        End If
    End Function

    Protected Function GetUserKeyFromUserID(ByVal sUserID As String, nCustomerKey As Int32) As Int32
        GetUserKeyFromUserID = -1
        Dim sSQL As String = "SELECT [key] FROM UserProfile WHERE DeletedFlag = 0 AND Status = 'ACTIVE' AND CustomerKey = " & nCustomerKey.ToString & " AND UserID = '" & sUserID.Replace("'", "''") & "'"
        Dim dtUser As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtUser.Rows.Count = 1 Then
            GetUserKeyFromUserID = dtUser.Rows(0).Item(0)
        End If
    End Function

    Protected Function GetCountryKeyFromCountryCode(ByVal sCountryCode As String) As Int32
        GetCountryKeyFromCountryCode = -1
        Dim sSQL As String = "SELECT CountryKey FROM Country WHERE CountryName = " & sCountryCode.Replace("'", "''")
        Dim dtCountry As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtCountry.Rows.Count = 1 Then
            GetCountryKeyFromCountryCode = dtCountry.Rows(0).Item(0)
        End If
    End Function
    
    Protected Function GetProductKeyFromProductCode(ByVal sProductCode As String, ByVal sProductDate As String) As Int32
        GetProductKeyFromProductCode = 0
        Dim sSQL As String = "SELECT LogisticProductKey FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductCode = '" & sProductCode & "' AND ISNULL(ProductDate, '') = '" & sProductDate & "'"
        Dim dtProduct As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtProduct.Rows.Count = 1 Then
            Return dtProduct.Rows(0).Item(0)
        ElseIf dtProduct.Rows.Count > 1 Then
            Return -1
        Else
            Return 0
        End If
    End Function
    
    Protected Function GetCountOfProductsWithThisProductCode(ByVal sProductCode As String) As Int32
        GetCountOfProductsWithThisProductCode = 0
        Dim sSQL As String = "SELECT COUNT (*) FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND DeletedFlag = 'N' AND ProductCode = '" & sProductCode & "'"
        GetCountOfProductsWithThisProductCode = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function
    
    Protected Function GetProductCodeFromProductKey(ByVal nProductKey As Int32) As String
        GetProductCodeFromProductKey = ExecuteQueryToDataTable("SELECT ProductCode FROM LogisticProduct WHERE LogisticProductKey = " & nProductKey).Rows(0).Item(0)
    End Function
    
    Protected Function GetProductDescriptionFromProductCode(ByVal sProductCode As String) As String
        GetProductDescriptionFromProductCode = String.Empty
        Dim sSQL As String = "SELECT TOP 1 ProductDescription FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductCode = '" & sProductCode & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 1 Then
            GetProductDescriptionFromProductCode = dt.Rows(0).Item(0)
        End If
    End Function
    
    Protected Function GetProductDetailsFromProductKey(ByVal nLogisticProductKey As Int32) As String
        GetProductDetailsFromProductKey = String.Empty
        Dim sSQL As String = "SELECT ProductCode, ISNULL(ProductDate, '') 'ProductDate', ProductDescription FROM LogisticProduct WHERE LogisticProductKey = " & nLogisticProductKey
        Dim dr As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        GetProductDetailsFromProductKey = dr("ProductCode")
        If dr("ProductDate") <> String.Empty Then
            GetProductDetailsFromProductKey &= " ~ " & dr("ProductDate")
        End If
        GetProductDetailsFromProductKey &= " " & dr("ProductDescription")
    End Function

    Private Function GetAvailableQty(ByVal nLogisticProductKey As Integer) As Integer
        Dim sSQL As String = "SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ") END"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
            GetAvailableQty = oDataTable.Rows(0)(0)
        Else
            GetAvailableQty = 0
        End If
    End Function

    Protected Function nSubmitConsignment(ByVal dictOrderItems As Dictionary(Of Int32, Int32)) As Integer
        Dim sConn As String = ConfigLib.GetConfigItem_ConnectionString
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(sConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        nSubmitConsignment = 0
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = gnUserKey
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = gnCustomerKey
        oCmdAddBooking.Parameters.Add(param2)

        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "ORDER_PROC_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)

        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = gsCustRef1
        oCmdAddBooking.Parameters.Add(param3)

        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = gsCustRef2
        oCmdAddBooking.Parameters.Add(param4)

        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = gsCustRef3
        oCmdAddBooking.Parameters.Add(param5)

        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = gsCustRef4
        oCmdAddBooking.Parameters.Add(param6)

        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = gsExternalReference
        oCmdAddBooking.Parameters.Add(param6a)

        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = gsSpecialInstructions
        oCmdAddBooking.Parameters.Add(param7)
        
        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = gsPackingNote
        oCmdAddBooking.Parameters.Add(param8)
        
        Dim param9 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
        param9.Value = "STOCK ITEM"
        oCmdAddBooking.Parameters.Add(param9)
        
        Dim param10 As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 4)
        param10.Value = -1
        oCmdAddBooking.Parameters.Add(param10)
        
        Dim param11 As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
        param11.Value = "PRINTED MATTER - FREE DOMICILE"
        oCmdAddBooking.Parameters.Add(param11)
        
        Dim sSQL As String = "SELECT * FROM Customer WHERE CustomerKey = " & gnCustomerKey
        Dim dtCnor As New DataTable
        oConn = New SqlConnection(sConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oAdapter.Fill(dtCnor)
            oConn.Open()
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sSQL & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        
        Dim drCnor As DataRow
        If dtCnor.Rows.Count = 1 Then
            drCnor = dtCnor.Rows(0)
        Else
            WebMsgBox.Show("Couldn't retrieve Consignor details.")
            nSubmitConsignment = -1
            Exit Function
        End If
       
        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        'param13.Value = psCnorCompany
        param13.Value = drCnor("CustomerName")
       
        oCmdAddBooking.Parameters.Add(param13)
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = drCnor("CustomerAddr1")
        
        oCmdAddBooking.Parameters.Add(param14)
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = drCnor("CustomerAddr2")
        
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = drCnor("CustomerAddr3")
        
        oCmdAddBooking.Parameters.Add(param16)
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = drCnor("CustomerTown")
        oCmdAddBooking.Parameters.Add(param17)
        
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = drCnor("CustomerCounty")
        oCmdAddBooking.Parameters.Add(param18)
        
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = drCnor("CustomerPostCode")
        oCmdAddBooking.Parameters.Add(param19)
        
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = drCnor("CustomerCountryKey")
        oCmdAddBooking.Parameters.Add(param20)
        
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = ""
        oCmdAddBooking.Parameters.Add(param21)
        
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = ""
        oCmdAddBooking.Parameters.Add(param22)
        
        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = ""
        oCmdAddBooking.Parameters.Add(param23)
        
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
        
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = gsConsignee
        oCmdAddBooking.Parameters.Add(param25)
        
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = gsAddr1
        oCmdAddBooking.Parameters.Add(param26)
        
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = gsAddr2
        oCmdAddBooking.Parameters.Add(param27)
        
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = gsAddr3
        oCmdAddBooking.Parameters.Add(param28)
        
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = gsTown
        oCmdAddBooking.Parameters.Add(param29)
        
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = gsRegion
        oCmdAddBooking.Parameters.Add(param30)
        
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = gsPostcode
        oCmdAddBooking.Parameters.Add(param31)
        
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = gnCountryCode
        oCmdAddBooking.Parameters.Add(param32)
        
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = gsConsigneeContact
        oCmdAddBooking.Parameters.Add(param33)
        
        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = ""
        oCmdAddBooking.Parameters.Add(param34)
        
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = ""
        oCmdAddBooking.Parameters.Add(param35)
        
        Dim param36 As SqlParameter = New SqlParameter("@CneePreAlertFlag", SqlDbType.Bit)
        param36.Value = 0
        oCmdAddBooking.Parameters.Add(param36)
        
        Dim param37 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
        param37.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param37)
        
        Dim param38 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        param38.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param38)
        
        Try
            BookingFailed = False
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddBooking")
            oCmdAddBooking.Connection = oConn
            oCmdAddBooking.Transaction = oTrans
            oCmdAddBooking.ExecuteNonQuery()
            lBookingKey = CLng(oCmdAddBooking.Parameters("@LogisticBookingKey").Value.ToString)
            lConsignmentKey = CLng(oCmdAddBooking.Parameters("@ConsignmentKey").Value.ToString)
            If lBookingKey > 0 Then
                For Each kvp As KeyValuePair(Of Int32, Int32) In dictOrderItems
                    Dim lProductKey As Int32 = kvp.Key
                    Dim lPickQuantity As Int32 = kvp.Value
                    Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                    oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                    Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                    param51.Value = gnUserKey
                    oCmdAddStockItem.Parameters.Add(param51)
                    Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                    param52.Value = gnCustomerKey
                    oCmdAddStockItem.Parameters.Add(param52)
                    Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                    param53.Value = lBookingKey
                    oCmdAddStockItem.Parameters.Add(param53)
                    Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                    param54.Value = lProductKey
                    oCmdAddStockItem.Parameters.Add(param54)
                    Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                    param55.Value = "PENDING"
                    oCmdAddStockItem.Parameters.Add(param55)
                    Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                    param56.Value = lPickQuantity
                    oCmdAddStockItem.Parameters.Add(param56)
                    Dim param57 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 8)
                    param57.Value = lConsignmentKey
                    oCmdAddStockItem.Parameters.Add(param57)
                    oCmdAddStockItem.Connection = oConn
                    oCmdAddStockItem.Transaction = oTrans
                    oCmdAddStockItem.ExecuteNonQuery()
                Next

                Dim oCmdCompleteBooking As SqlCommand = New SqlCommand("spASPNET_LogisticBooking_Complete", oConn)
                oCmdCompleteBooking.CommandType = CommandType.StoredProcedure
                Dim param71 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                param71.Value = lBookingKey
                oCmdCompleteBooking.Parameters.Add(param71)
                oCmdCompleteBooking.Connection = oConn
                oCmdCompleteBooking.Transaction = oTrans
                oCmdCompleteBooking.ExecuteNonQuery()
            Else
                BookingFailed = True
            End If
            If Not BookingFailed Then
                oTrans.Commit()
                nSubmitConsignment = lConsignmentKey
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            oTrans.Rollback("AddBooking")
        Finally
            oConn.Close()
        End Try
    End Function

    Protected Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oAdapter.Fill(oDataTable)
            oConn.Open()
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function
    
    Property psUniqueFilename() As String
        Get
            Dim o As Object = ViewState("ORDERPROCESSOR_UniqueFilename")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("ORDERPROCESSOR_UniqueFilename") = Value
        End Set
    End Property

    Protected Sub btnReadOrders_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearAllFields()
        Call ReadOrders()
    End Sub
    
    Protected Sub ReadOrders()
        gsItemOrder = "PQ"
        If tbOrder.Text.Trim <> String.Empty Then
            Call CreateCleanOrderFileFromTextbox()
            If CheckCleanOrderFile(bExecuteOrder:=False) Then
                btnProcess.Enabled = True
            Else
                btnProcess.Enabled = False
            End If
        End If
    End Sub
    
    Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearAllFields()
        gsItemOrder = "PQ"
        lstConsignmentNos = New List(Of Int32)
        If tbOrder.Text.Trim <> String.Empty Then
            Call CreateCleanOrderFileFromTextbox()
            If CheckCleanOrderFile(bExecuteOrder:=False) Then
                Call ClearAllFields()
                CheckCleanOrderFile(bExecuteOrder:=True)
            End If
        End If
        btnProcess.Enabled = False
    End Sub

    Protected Function ProcessWesternUnionAgent(arrLine() As String, ByVal bExecuteOrder As Boolean) As String
        Dim sAgentID As String = arrLine(0)
        Dim sSQL As String
        sSQL = "SELECT CustomerKey FROM UserProfile WHERE Status = 'Active' AND DeletedFlag = 0 AND CustomerKey IN (579, 686, 798, 826) AND UserID = '" & sAgentID & "'"
        Dim dtAgent As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtAgent.Rows.Count = 1 Then
            Dim nCustomerKey = dtAgent.Rows(0).Item(0)
            If nCustomerKey = 579 Then
                sSQL = "SELECT * FROM ClientData_WU_Agents WHERE TermID = '" & sAgentID & "'"
                dtAgent = ExecuteQueryToDataTable(sSQL)
                If dtAgent.Rows.Count = 1 Then
                    gsConsignee = dtAgent.Rows(0)("AgentName")
                    gsAddr1 = dtAgent.Rows(0)("Address1")
                    gsAddr2 = dtAgent.Rows(0)("Address2")
                    gsAddr3 = dtAgent.Rows(0)("Address3")
                    gsTown = dtAgent.Rows(0)("City")
                    gsRegion = dtAgent.Rows(0)("State")
                    gsPostcode = dtAgent.Rows(0)("Postcode")
                    gsCountry = "U.K."
                    gnCountryCode = COUNTRY_CODE_UK

                    Dim sbResult As New StringBuilder
                    sbResult.Append("Consignee set to " & gsConsignee & HTML_BREAK)
                    sbResult.Append("No Consignee contact name." & HTML_BREAK)
                    sbResult.Append("Addr1 set to " & gsAddr1 & HTML_BREAK)

                    If gsAddr2 = String.Empty Then
                        sbResult.Append("No Addr2." & HTML_BREAK)
                    Else
                        sbResult.Append("Addr2 set to " & gsAddr2 & HTML_BREAK)
                    End If

                    If gsAddr3 = String.Empty Then
                        sbResult.Append("No Addr3." & HTML_BREAK)
                    Else
                        sbResult.Append("Addr3 set to " & gsAddr3 & HTML_BREAK)
                    End If

                    If gbNoTown Then
                        sbResult.Append("NOTOWN is set." & HTML_BREAK)
                    Else
                        sbResult.Append("Town/City set to " & gsTown & HTML_BREAK)
                    End If

                    If gsRegion = String.Empty Then
                        sbResult.Append("No Region." & HTML_BREAK)
                    Else
                        sbResult.Append("Region set to " & gsRegion & HTML_BREAK)
                    End If

                    If gsCountry <> String.Empty Then
                        sbResult.Append("Country set to " & gsCountry & HTML_BREAK)
                    Else
                        sbResult.Append("Country set to UK" & HTML_BREAK)
                    End If
                    ProcessWesternUnionAgent = "+" & sbResult.ToString
                Else
                    ProcessWesternUnionAgent = "-No address on file for this WURS Agent."
                End If
            ElseIf nCustomerKey = 686 Then
                sSQL = "SELECT * FROM ClientData_WUIRE_Agents WHERE TermID = '" & sAgentID & "'"
                dtAgent = ExecuteQueryToDataTable(sSQL)
                If dtAgent.Rows.Count = 1 Then
                    gsConsignee = dtAgent.Rows(0)("AgentName")
                    gsAddr1 = dtAgent.Rows(0)("Address1")
                    gsAddr2 = dtAgent.Rows(0)("Address2")
                    gsAddr3 = dtAgent.Rows(0)("Address3")
                    gsTown = dtAgent.Rows(0)("City")
                    gsRegion = dtAgent.Rows(0)("State")
                    gsPostcode = dtAgent.Rows(0)("Postcode")

                    Dim sbResult As New StringBuilder
                    sbResult.Append("Consignee set to " & gsConsignee & HTML_BREAK)
                    sbResult.Append("No Consignee contact name." & HTML_BREAK)
                    sbResult.Append("Addr1 set to " & gsAddr1 & HTML_BREAK)

                    If gsAddr2 = String.Empty Then
                        sbResult.Append("No Addr2." & HTML_BREAK)
                    Else
                        sbResult.Append("Addr2 set to " & gsAddr2 & HTML_BREAK)
                    End If

                    If gsAddr3 = String.Empty Then
                        sbResult.Append("No Addr3." & HTML_BREAK)
                    Else
                        sbResult.Append("Addr3 set to " & gsAddr3 & HTML_BREAK)
                    End If

                    If gbNoTown Then
                        sbResult.Append("NOTOWN is set." & HTML_BREAK)
                    Else
                        sbResult.Append("Town/City set to " & gsTown & HTML_BREAK)
                    End If

                    If gsRegion = String.Empty Then
                        sbResult.Append("No Region." & HTML_BREAK)
                    Else
                        sbResult.Append("Region set to " & gsRegion & HTML_BREAK)
                    End If

                    If gsCountry <> String.Empty Then
                        sbResult.Append("Country set to " & gsCountry & HTML_BREAK)
                    Else
                        sbResult.Append("Country set to UK" & HTML_BREAK)
                    End If
                    ProcessWesternUnionAgent = "+" & sbResult.ToString
                Else
                    ProcessWesternUnionAgent = "-No address on file for this WUIRE Agent."
                End If
            ElseIf nCustomerKey = CUSTOMER_FININT Then
                sSQL = "SELECT * FROM ClientData_WU_LegacyNetwork WHERE AgentID = '" & sAgentID & "'"
                dtAgent = ExecuteQueryToDataTable(sSQL)
                If dtAgent.Rows.Count = 1 Then
                    gsConsignee = dtAgent.Rows(0)("LocationName")
                    gsAddr1 = dtAgent.Rows(0)("AddressLine1")
                    gsAddr2 = dtAgent.Rows(0)("AddressLine2")
                    gsTown = dtAgent.Rows(0)("CityName")
                    gsRegion = dtAgent.Rows(0)("Province/County/State")
                    gsPostcode = dtAgent.Rows(0)("PostalCode")
                    gsCountry = "U.K."
                    gnCountryCode = COUNTRY_CODE_UK

                    Dim sbResult As New StringBuilder
                    sbResult.Append("Consignee set to " & gsConsignee & HTML_BREAK)
                    sbResult.Append("No Consignee contact name." & HTML_BREAK)
                    sbResult.Append("Addr1 set to " & gsAddr1 & HTML_BREAK)

                    If gsAddr2 = String.Empty Then
                        sbResult.Append("No Addr2." & HTML_BREAK)
                    Else
                        sbResult.Append("Addr2 set to " & gsAddr2 & HTML_BREAK)
                    End If

                    If gsAddr3 = String.Empty Then
                        sbResult.Append("No Addr3." & HTML_BREAK)
                    Else
                        sbResult.Append("Addr3 set to " & gsAddr3 & HTML_BREAK)
                    End If

                    If gbNoTown Then
                        sbResult.Append("NOTOWN is set." & HTML_BREAK)
                    Else
                        sbResult.Append("Town/City set to " & gsTown & HTML_BREAK)
                    End If

                    If gsRegion = String.Empty Then
                        sbResult.Append("No Region." & HTML_BREAK)
                    Else
                        sbResult.Append("Region set to " & gsRegion & HTML_BREAK)
                    End If

                    If gsCountry <> String.Empty Then
                        sbResult.Append("Country set to " & gsCountry & HTML_BREAK)
                    Else
                        sbResult.Append("Country set to UK" & HTML_BREAK)
                    End If
                    ProcessWesternUnionAgent = "+" & sbResult.ToString
                Else
                    ProcessWesternUnionAgent = "-No address on file for this FININT Agent."
                End If
            ElseIf nCustomerKey = CUSTOMER_COSTA Then
                sSQL = "SELECT * FROM ClientData_WUCOSTA_Agents WHERE TermID = '" & sAgentID & "'"
                dtAgent = ExecuteQueryToDataTable(sSQL)
                If dtAgent.Rows.Count = 1 Then
                    gsConsignee = dtAgent.Rows(0)("LocationName")
                    gsAddr1 = dtAgent.Rows(0)("Address")
                    gsPostcode = dtAgent.Rows(0)("PostCode")
                    gsCountry = "U.K."
                    gnCountryCode = COUNTRY_CODE_UK

                    Dim sbResult As New StringBuilder
                    sbResult.Append("Consignee set to " & gsConsignee & HTML_BREAK)
                    sbResult.Append("No Consignee contact name." & HTML_BREAK)
                    sbResult.Append("Addr1 set to " & gsAddr1 & HTML_BREAK)

                    If gsAddr2 = String.Empty Then
                        sbResult.Append("No Addr2." & HTML_BREAK)
                    Else
                        sbResult.Append("Addr2 set to " & gsAddr2 & HTML_BREAK)
                    End If

                    If gsAddr3 = String.Empty Then
                        sbResult.Append("No Addr3." & HTML_BREAK)
                    Else
                        sbResult.Append("Addr3 set to " & gsAddr3 & HTML_BREAK)
                    End If

                    If gbNoTown Then
                        sbResult.Append("NOTOWN is set." & HTML_BREAK)
                    Else
                        sbResult.Append("Town/City set to " & gsTown & HTML_BREAK)
                    End If

                    If gsRegion = String.Empty Then
                        sbResult.Append("No Region." & HTML_BREAK)
                    Else
                        sbResult.Append("Region set to " & gsRegion & HTML_BREAK)
                    End If

                    If gsCountry <> String.Empty Then
                        sbResult.Append("Country set to " & gsCountry & HTML_BREAK)
                    Else
                        sbResult.Append("Country set to UK" & HTML_BREAK)
                    End If
                    ProcessWesternUnionAgent = "+" & sbResult.ToString
                Else
                    ProcessWesternUnionAgent = "-No address on file for this COSTA Agent."
                End If
            Else
                ProcessWesternUnionAgent = "-Internal error seaching for Agent address."
            End If
        Else
            ProcessWesternUnionAgent = "-Agent account not found."
        End If
    End Function

    Protected Sub ClearAddressFields()
        gsConsignee = String.Empty
        gsConsigneeContact = String.Empty
        gsAddr1 = String.Empty
        gsAddr2 = String.Empty
        gsAddr3 = String.Empty
        gsTown = String.Empty
        gsRegion = String.Empty
        gsPostcode = String.Empty
        gsCountry = String.Empty
        gnCountryCode = -1
    End Sub
    
    Protected Sub ClearAllFields()
        Call ClearAddressFields()
        gsColumnConsignee = String.Empty
        gsColumnConsigneeContact = String.Empty
        gsColumnAddr1 = String.Empty
        gsColumnAddr2 = String.Empty
        gsColumnAddr3 = String.Empty
        gsColumnTown = String.Empty
        gsColumnRegion = String.Empty
        gsColumnPostcode = String.Empty
        gsColumnCountry = String.Empty

        gsColumnCustRef1 = String.Empty
        gsColumnCustRef2 = String.Empty
        gsColumnCustRef3 = String.Empty
        gsColumnCustRef4 = String.Empty
        gsColumnSpecialInstructions = String.Empty
        gsColumnPackingNote = String.Empty
        gsColumnExternalReference = String.Empty

        gsDefaultCustRef1 = String.Empty
        gsDefaultCustRef2 = String.Empty
        gsDefaultCustRef3 = String.Empty
        gsDefaultCustRef4 = String.Empty
        
        'gsItemColumnStart = DEFAULT_ITEM_COLUMN_START   ' CN 24OCT13

        gsDefaultSpecialInstructions = String.Empty
        gsDefaultPackingNote = String.Empty
        gsDefaultExternalReference = String.Empty

        gbOrderFound = False
        
        gdictOrderItems.Clear()
        gdictTotalOrderItems.Clear()
        gdictCountryAliases.Clear()
    End Sub

    Protected Function ParseOrderLine(ByVal arrLine() As String, ByVal bExecuteOrder As Boolean) As String
        gbOrderFound = True
        If gbTemplateWU Then
            Return ProcessWesternUnionAgent(arrLine, bExecuteOrder)
        End If
        
        If gsColumnConsignee = String.Empty Then
            Return "-Cannot process order before finding directive CONSIGNEE"
        End If
        If gsColumnAddr1 = String.Empty Then
            Return "-Cannot process order before finding directive ADDR1"
        End If
        If gsColumnTown = String.Empty Then
            Return "-Cannot process order before finding directive TOWN"
        End If
        If gsColumnPostcode = String.Empty Then
            Return "-Cannot process order before finding directive POSTCODE"
        End If
        If Not gbDefaultUK Then
            If gsColumnCountry = String.Empty Then
                Return "-Cannot process order before finding directive COUNTRY (directive DEFAULTUK not found)"
            End If
        End If
        gsConsignee = GetAddressComponent(arrLine, gsColumnConsignee)
        gsConsigneeContact = GetAddressComponent(arrLine, gsColumnConsigneeContact)
        gsAddr1 = GetAddressComponent(arrLine, gsColumnAddr1)
        gsAddr2 = GetAddressComponent(arrLine, gsColumnAddr2)
        gsAddr3 = GetAddressComponent(arrLine, gsColumnAddr3)
        gsTown = GetAddressComponent(arrLine, gsColumnTown)
        gsRegion = GetAddressComponent(arrLine, gsColumnRegion)
        gsPostcode = GetAddressComponent(arrLine, gsColumnPostcode)
        
        gsCountry = GetAddressComponent(arrLine, gsColumnCountry)
        If gdictCountryAliases.ContainsKey(gsCountry) Then
            gsCountry = gdictCountryAliases(gsCountry)
        End If
        
        gsCustRef1 = GetAddressComponent(arrLine, gsColumnCustRef1)
        gsCustRef2 = GetAddressComponent(arrLine, gsColumnCustRef2)
        gsCustRef3 = GetAddressComponent(arrLine, gsColumnCustRef3)
        gsCustRef4 = GetAddressComponent(arrLine, gsColumnCustRef4)
        gsSpecialInstructions = GetAddressComponent(arrLine, gsColumnSpecialInstructions)
        gsPackingNote = GetAddressComponent(arrLine, gsColumnPackingNote)
        gsExternalReference = GetAddressComponent(arrLine, gsColumnExternalReference)
                
        If gsCustRef1 = String.Empty Then
            gsCustRef1 = gsDefaultCustRef1
        End If

        If gsCustRef1 = String.Empty Then
            gsCustRef2 = gsDefaultCustRef2
        End If

        If gsCustRef3 = String.Empty Then
            gsCustRef3 = gsDefaultCustRef3
        End If

        If gsCustRef4 = String.Empty Then
            gsCustRef4 = gsDefaultCustRef4
        End If

        If gsSpecialInstructions = String.Empty Then
            gsSpecialInstructions = gsDefaultSpecialInstructions
        End If
        
        If gsPackingNote = String.Empty Then
            gsPackingNote = gsDefaultPackingNote
        End If
        
        ' note no check currently done for overlength special instructions or packing note

        If gsExternalReference = String.Empty Then
            gsExternalReference = gsDefaultExternalReference
        End If
        
        If gsExternalReference.Length > MAX_EXTERNAL_REFERENCE_LENGTH Then
            Return "-External Reference overlength field - max length is " & MAX_EXTERNAL_REFERENCE_LENGTH & "."
        End If
        
        If gsCustRef1.Length > MAX_CUSTREF_12_FIELD_LENGTH Then
            Return "-Cust Ref 1 overlength field - max length is " & MAX_CUSTREF_12_FIELD_LENGTH & "."
        End If
        
        If gsCustRef2.Length > MAX_CUSTREF_12_FIELD_LENGTH Then
            Return "-Cust Ref 2 overlength field - max length is " & MAX_CUSTREF_12_FIELD_LENGTH & "."
        End If
        
        If gsCustRef3.Length > MAX_CUSTREF_34_FIELD_LENGTH Then
            Return "-Cust Ref 3 overlength field - max length is " & MAX_CUSTREF_34_FIELD_LENGTH & "."
        End If
        
        If gsCustRef4.Length > MAX_CUSTREF_34_FIELD_LENGTH Then
            Return "-Cust Ref 4 overlength field - max length is " & MAX_CUSTREF_34_FIELD_LENGTH & "."
        End If
        
        If gsConsignee = String.Empty Then
            Return "-Consignee name not found."
        End If
        If gsAddr1 = String.Empty Then
            Return "-Addr1 not found."
        End If
        If gsTown = String.Empty Then
            Return "-Town/City not found."
        End If
        If gsPostcode = String.Empty Then
            Return "-Postcode not found."
        End If
        If gsTown <> String.Empty Then
            If gbNoTown Then
                Return "-Town/City found but NOTOWN is set."
            End If
        End If
        If gsPostcode <> String.Empty Then
            If gbNoPostcode Then
                Return "-Postcode found but NOPOSTCODE is set."
            End If
        End If
        
        If gsConsignee.Length > MAX_ADDRESS_FIELD_LENGTH Then
            Return "-Consignee overlength field - max length is " & MAX_ADDRESS_FIELD_LENGTH & "."
        End If
        If gsConsigneeContact.Length > MAX_ADDRESS_FIELD_LENGTH Then
            Return "-Consignee contact overlength field - max length is " & MAX_ADDRESS_FIELD_LENGTH & "."
        End If
        If gsAddr1.Length > MAX_ADDRESS_FIELD_LENGTH Then
            Return "-Addr1 overlength field - max length is " & MAX_ADDRESS_FIELD_LENGTH & "."
        End If
        If gsAddr2.Length > MAX_ADDRESS_FIELD_LENGTH Then
            Return "-Addr2 overlength field - max length is " & MAX_ADDRESS_FIELD_LENGTH & "."
        End If
        If gsAddr3.Length > MAX_ADDRESS_FIELD_LENGTH Then
            Return "-Addr3 overlength field - max length is " & MAX_ADDRESS_FIELD_LENGTH & "."
        End If
        If gsTown.Length > MAX_ADDRESS_FIELD_LENGTH Then
            Return "-Town/City overlength field - max length is " & MAX_ADDRESS_FIELD_LENGTH & "."
        End If
        If gsRegion.Length > MAX_ADDRESS_FIELD_LENGTH Then
            Return "-Region overlength field - max length is " & MAX_ADDRESS_FIELD_LENGTH & "."
        End If
        If gsPostcode.Length > MAX_ADDRESS_FIELD_LENGTH Then
            Return "-Postcode overlength field - max length is " & MAX_ADDRESS_FIELD_LENGTH & "."
        End If
        
        If gsCountry <> String.Empty Then
            gnCountryCode = GetCountryCodeFromCountryName(gsCountry)
            If gnCountryCode < 0 Then
                Return "-Cannot match " & gsCountry & " to a recognised country."
            End If
        Else
            If Not gbDefaultUK Then
                Return "-No country specified."
            Else
                gsCountry = "U.K."
                gnCountryCode = COUNTRY_CODE_UK
            End If
        End If
        
        For Each kv As KeyValuePair(Of Int32, Int32) In gdictOrderItems
            If gdictTotalOrderItems.ContainsKey(kv.Key) Then
                gdictTotalOrderItems(kv.Key) = gdictTotalOrderItems(kv.Key) + kv.Value
            Else
                gdictTotalOrderItems.Add(kv.Key, kv.Value)
            End If
        Next
        
        Dim sbResult As New StringBuilder
        If bExecuteOrder Then
            sbResult.Append(HTML_BREAK & "PROCESSING ORDER..." & gsConsignee & HTML_BREAK)
        Else
            sbResult.Append(HTML_BREAK & "CHECKING ORDER..." & gsConsignee & HTML_BREAK)
        End If
        sbResult.Append("Consignee set to " & gsConsignee & HTML_BREAK)
        sbResult.Append("No Consignee contact name." & HTML_BREAK)
        sbResult.Append("Addr1 set to " & gsAddr1 & HTML_BREAK)

        If gsAddr2 = String.Empty Then
            sbResult.Append("No Addr2." & HTML_BREAK)
        Else
            sbResult.Append("Addr2 set to " & gsAddr2 & HTML_BREAK)
        End If

        If gsAddr3 = String.Empty Then
            sbResult.Append("No Addr3." & HTML_BREAK)
        Else
            sbResult.Append("Addr3 set to " & gsAddr3 & HTML_BREAK)
        End If

        If gbNoTown Then
            sbResult.Append("NOTOWN is set." & HTML_BREAK)
        Else
            sbResult.Append("Town/City set to " & gsTown & HTML_BREAK)
        End If

        If gsRegion = String.Empty Then
            sbResult.Append("No Region." & HTML_BREAK)
        Else
            sbResult.Append("Region set to " & gsRegion & HTML_BREAK)
        End If

        If gsCountry <> String.Empty Then
            sbResult.Append("Country set to " & gsCountry & HTML_BREAK)
        Else
            sbResult.Append("Country set to UK" & HTML_BREAK)
        End If
        ParseOrderLine = "+" & sbResult.ToString
    End Function
    
    Protected Function GetCountryCodeFromCountryName(ByVal sCountryName As String) As Int32
        GetCountryCodeFromCountryName = -1
        Dim sSQL As String = "SELECT CountryKey FROM Country WHERE CountryName = '" & sCountryName.Replace("'", "''") & "'"
        Dim dtCountry As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtCountry.Rows.Count = 1 Then
            Return dtCountry.Rows(0).Item(0)
        End If
    End Function
    
    Protected Function GetAddressComponent(ByVal arrLine() As String, ByRef sColumnSpecifier As String) As String
        GetAddressComponent = String.Empty
        Dim nArrIndex As Int32 = GetIndexFromColumnLetter(sColumnSpecifier)
        If nArrIndex >= 0 Then
            Try
                GetAddressComponent = arrLine(nArrIndex)
            Catch ex As Exception
            End Try
        End If
    End Function
    
    Protected Function GetIndexFromColumnLetter(ByVal sColumnLetter As String) As Int32
        GetIndexFromColumnLetter = -1
        If sColumnLetter <> String.Empty Then
            GetIndexFromColumnLetter = Strings.Asc(sColumnLetter) - (Strings.Asc("A") + 1)
        End If
    End Function
    
    Protected Function InAtoZ(ByVal sValue As String) As Boolean
        InAtoZ = False
        If sValue.Length = 1 Then
            Dim sRegexPattern As String = "[A-Za-z]"
            Dim r As Regex = New Regex(sRegexPattern, RegexOptions.IgnoreCase)
            If r.Match(sValue).Success Then
                InAtoZ = True
            End If
        End If
    End Function
    
    Protected Sub lnkbtnHide_Click(sender As Object, e As System.EventArgs)
        lblProducts.Text = String.Empty
    End Sub
    
    Protected Sub btnHelpShowHide_Click(sender As Object, e As System.EventArgs)
        If btnHelpShowHide.Text.ToLower.Contains("show") Then
            btnHelpShowHide.Text = "hide help"
            pnlHelp.Visible = True
        Else
            btnHelpShowHide.Text = "show help"
            pnlHelp.Visible = False
        End If
    End Sub
    
    Protected Sub ddlCustomer_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue = 0 Then
            lblProducts.Text = String.Empty
            Call ShowHideProductFinderControls(False)
        Else
            Call ShowHideProductFinderControls(True)
            tbSearchProducts.Text = String.Empty
        End If
        lblProducts.Text = String.Empty
    End Sub
    
    Protected Sub btnAllProducts_Click(sender As Object, e As System.EventArgs)
        tbSearchProducts.Text = String.Empty
        Call ShowProducts(String.Empty)
    End Sub

    Protected Sub btnSearchProducts_Click(sender As Object, e As System.EventArgs)
        If tbSearchProducts.Text.Length < 2 Then
            WebMsgBox.Show("Please enter at least two (2) search characters.")
        End If
        Call ShowProducts(tbSearchProducts.Text)
    End Sub
    
    Protected Sub ShowProducts(sSearchString As String)
        Dim nProductCount As Int32 = ExecuteQueryToDataTable("SELECT COUNT (*) FROM LogisticProduct WHERE DeletedFlag = 'N' AND ArchiveFlag = 'N' AND CustomerKey = " & ddlCustomer.SelectedValue).Rows(0).Item(0)
        If nProductCount > 400 And sSearchString = String.Empty Then
            WebMsgBox.Show("Too many products (" & nProductCount.ToString & ") to display - please use the search facility to filter the product list.")
            Exit Sub
        End If
        Dim sSQL As String
        
        lblProducts.Text = String.Empty
        sSQL = "SELECT ProductCode, ProductDate, ProductDescription, '$PRODREF:' + CAST(LogisticProductKey AS varchar(10)) + ' (' + ProductCode + ' - ' + ProductDescription + ')' 'Shortcut' FROM LogisticProduct WHERE DeletedFlag = 'N' AND ArchiveFlag = 'N' AND CustomerKey = " & ddlCustomer.SelectedValue
        If sSearchString <> String.Empty Then
            sSQL &= " AND ( ProductCode LIKE '%" & sSearchString & "%' OR ProductDate LIKE '%" & sSearchString & "%' OR ProductDescription LIKE '%" & sSearchString & "%' )"
        End If
        sSQL &= " ORDER BY ProductCode"
        Dim dtProducts As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtProducts.Rows.Count > 0 Then
            For Each drProduct As DataRow In dtProducts.Rows
                lblProducts.Text &= "PC: " & drProduct("ProductCode") & " | PD: " & drProduct("ProductDate") & " | DESCR: " & drProduct("ProductDescription") & ".....................<b>" & drProduct("Shortcut") & "</b>....................." & HTML_BREAK
            Next
        Else
            lblProducts.Text = "No products found."
        End If
    End Sub
    
    Protected Sub ShowHideProductFinderControls(bVisibility As Boolean)
        btnAllProducts.Visible = bVisibility
        lblLegendSearch.Visible = bVisibility
        tbSearchProducts.Visible = bVisibility
        btnSearchProducts.Visible = bVisibility
        lnkbtnHide.Visible = bVisibility
    End Sub
    
    Protected Sub btnCountryList_Click(sender As Object, e As System.EventArgs)
        lblCountries.Text = String.Empty
        Dim sSQL As String = "SELECT CountryName FROM Country WHERE ISOPrimary = 1 AND DeletedFlag = 0 ORDER BY CountryName"
        Dim dtCountryList As DataTable = ExecuteQueryToDataTable(sSQL)
        For Each drCountry As DataRow In dtCountryList.Rows
            lblCountries.Text &= drCountry(0) & HTML_BREAK
        Next
    End Sub
    
    Protected Sub lnkbtnHideCountryList_Click(sender As Object, e As System.EventArgs)
        lblCountries.Text = String.Empty
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="frmOrderProcessor" runat="Server">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
    <asp:Label ID="lblLegendOrderProcessor" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        Text="Order Processor - version 20JAN14" Font-Bold="True" />
    <br />
    <asp:Label ID="lblLegendOrder" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
        Text="1.  Copy order from Excel spreadsheet (CTRL+A, CTRL+C) and paste (CTRL+V) into here..." />
    <br />
    <asp:TextBox ID="tbOrder" runat="server" Rows="10" TextMode="MultiLine" Width="100%" />
    <br />
    <br />
    <table width="100%">
        <tr>
            <td width="150px">
                <asp:Label ID="lblLegendOrder1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                    Text="2.  Read the order data:" />
            </td>
            <td>
                <asp:Button ID="btnRead" runat="server" Text="Read Orders" Width="200px" OnClick="btnReadOrders_Click" />
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="lblLegendOrder2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                    Text="3.  Process the order data:" />
            </td>
            <td>
                <asp:Button ID="btnProcess" runat="server" OnClick="btnProcess_Click" Text="Process Orders"
                    Width="200px" Enabled="False" />
            </td>
        </tr>
    </table>
    <br />
    <asp:Label ID="lblResult" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" />
    <br />
    <hr />
    <asp:Label ID="lblLegendOrderProcessor0" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        Text="Product Finder" Font-Bold="True" />
    <br />
    <br />
    <asp:Label ID="lblLegendOrder3" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
        Text="Customer:" />
    <asp:DropDownList ID="ddlCustomer" runat="server" AutoPostBack="True" Font-Names="Verdana"
        Font-Size="XX-Small" Height="16px" OnSelectedIndexChanged="ddlCustomer_SelectedIndexChanged">
    </asp:DropDownList>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:Button ID="btnAllProducts" runat="server" Text="all products" OnClick="btnAllProducts_Click" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:Label ID="lblLegendSearch" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
        Text="Search:" />
    <asp:TextBox ID="tbSearchProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
        Style="margin-bottom: 0px" />
    &nbsp;<asp:Button ID="btnSearchProducts" runat="server" Text="go" OnClick="btnSearchProducts_Click" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:LinkButton ID="lnkbtnHide" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
        OnClick="lnkbtnHide_Click">hide product finder</asp:LinkButton>
    <br />
    <br />
    <asp:Label ID="lblProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
    <br />
    <br />
    <hr />
    <asp:Button ID="btnCountryList" runat="server" OnClick="btnCountryList_Click" Text="AIMS Country List" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:LinkButton ID="lnkbtnHideCountryList" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
        OnClick="lnkbtnHideCountryList_Click">hide country list</asp:LinkButton>
    <br />
    <asp:Label ID="lblCountries" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
    &nbsp;<hr />
    <br />
    <asp:Button ID="btnHelpShowHide" runat="server" Text="show help" OnClick="btnHelpShowHide_Click" />
    <br />
    <br />
    <asp:Panel ID="pnlHelp" runat="server" Font-Names="Verdana" Font-Size="X-Small" Width="100%"
        Visible="false">
        <br />
        <strong>ORDER PROCESSOR INSTRUCTIONS</strong>
        <br />
        <p>
            $ COMMENT This is a sample order file, with documentation. Last updated 30JUL13
            by CN.</p>
        <p>
            $ COMMENT You can download this as an Excel spreadsheet from <a href="http://my.transworld.eu.com/internal/orderprocessor.xlsx">
                http://my.transworld.eu.com/internal/orderprocessor.xlsx</a>. The only differences
            between the this file and the spreadsheet are (a) this line, (b) pipe characters
            are used as column separators in this file instead of tab characters which Excel
            inserts to to separate columns. The processor will accept either.</p>
        <br />
        <p>
            $ CUSTOMER|DEMO</p>
        <p>
            $ USER|geoffdemo</p>
        <br />
        <p>
            $ CONSIGNEE|A</p>
        <p>
            $ CONSIGNEECONTACT</p>
        <p>
            $ ADDR1|B</p>
        <p>
            $ ADDR2|C</p>
        <p>
            $ ADDR3|D</p>
        <p>
            $ TOWN|E</p>
        <p>
            $ REGION|F</p>
        <br />
        <p>
            $ DEFAULTUK</p>
        <p>
            $ POSTCODE|G</p>
        <br />
        <p>
            $ CUSTREF1</p>
        <p>
            $ CUSTREF2</p>
        <p>
            $ CUSTREF3</p>
        <p>
            $ CUSTREF4</p>
        <p>
            $ SPECIALINSTRUCTIONS</p>
        <p>
            $ PACKINGNOTE</p>
        <p>
            $ EXTERNALREFERENCE</p>
        <br />
        <p>
            $ COMMENT Order processing is controlled by DIRECTIVES in the order spreadsheet.</p>
        <p>
            $ COMMENT Directives start with a dollar ($) sign and always appear in column A.
            The directive must be the first word in the column, otherwise it is ignored.</p>
        <p>
            $ COMMENT You can insert optionally a space after the $. Eg $COMMENT is equivalent
            to $ COMMENT.</p>
        <br />
        <p>
            $ COMMENT You can insert a comment anywhere</p>
        <br />
        <p>
            $ COMMENT Here is a list of directives, with notes on their use:</p>
        <br />
        <p>
            $ COMMENT Directive CUSTOMER must have a recognised AIMS / METACS customer account
            code in column B. This is the customer for which the orders following will be processed.</p>
        <p>
            $ COMMENT Directive USER is the UserID of the account placing the order; this field,
            although not mandatory, is strongly recommended</p>
        <p>
            $ COMMENT Directives CONSIGNEE, CONSIGNEECONTACT, ADDR1, ADDR2, ADDR3, TOWN, REGION,
            POSTCODE, COUNTRY should have a column letter in column B (if you omit the column
            letter the directive is ignored). The column letter, eg &#39;F&#39; indicates in
            which column of the address this element of the address is found. Elements CONSIGNEE,
            ADDR1, TOWN, POSTCODE are mandatory.</p>
        <br />
        <p>
            $ COMMENT Directives CUSTREF1/2/3/4, SPECIALINSTRUCTIONS, PACKINGNOTE, EXTERNALREFERENCE
            act like the address directives and are typically used to store order metadata.
            The values, if any, from DEFAULTCUSTREF1/2/3/4, DEFAULTSPECIALINSTRUCTIONS, DEFAULTPACKINGNOTE,
            DEFAULTEXTERNALREFERENCE are used if no explicit value is found.</p>
        <p>
            $ COMMENT Note that ExternalReference is a max 50 character attribute of the order
            that is available to store data with the order. It is visible on the AIMS Desktop
            interface but not on the web interface.</p>
        <br />
        <p>
            $ COMMENT Directive DEFAULTUK indicates that addresses with no COUNTRY (either no
            country field is defined, or the country field is blank) should default the country
            to UK</p>
        $ COMMENT Use directive ALIAS to convert different country name spellings to the
        spelling recognised by AIMS. Example: $ALIAS UNITED KINGDOM = UK<br />
        <p>
            $ COMMENT Directive NOTOWN disables checking that a Town/City is present</p>
        <p>
            $ COMMENT Directive NOPOSTCODE disables checking that a Post Code is present</p>
        <br />
        <p>
            $ COMMENT Directive ITEMORDER takes a value in column B of QPD, QP, PDQ or PQ. Letter
            P represent the Product Code field, letter D represents the Product Date / Value
            field, and letter Q represents Quantity field. This directive allows you to specify
            how the ITEM directives are interpreted. The default is PQ, ie first the Product
            Code field, then the Quantity field, no Value / Date field.</p>
        <p>
            $ COMMENT Directive ITEM takes values in column B &amp; C, and optionally column
            D if Product Date / Value is used. How columns B, C and optionally D are interpreted
            depends on the ITEMORDER directive.</p>
        <p>
            $ COMMENT Alternatively, instead of using the ITEMORDER AND ITEM directives, use
            the PRODREF directive. The format is $ PRODREF:12345 (plus an optional textual description
            of the product, which is ignored).&nbsp; Find the PRODREF values required in the
            <strong>product finder</strong> and paste them into the spreadsheet To make it 
            easier to copy these values, leading and trailing periods are ignored, so if you 
            copy and paste .....$PRODREF:12345 (some text)..... it will work. The PRODREF 
            directive only uses column A.</p>
        <p>
            $ COMMENT The items in an order must be specified before the order address. This 
            list of items is used for each order address found until the processor finds 
            another ITEM or PRODREF directive, upon which it clears any existing item list 
            and builds a new list starting with the product in tthe ITEM directive just 
            found. This is list is then used for any further order addresses, until another 
            ITEM or PRODREF directive is found, and so on.</p>
        <br />
        <p>
            $ COMMENT Any non-blank line that does not start with a directive is assumed to 
            be an order address. The processor builds an order using the products in the 
            list constructed from the ITEM or PRODREF directives found so far.</p>
        <br />
        <p>
            $ COMMENT The system starts processing with row 1 and works down until no more rows
            are found. If a directive is repeated, the most recently encountered directive is
            used. You can therefore include more than one address list in a single spreadsheet
            and the lists can have different formats.</p>
        <br />
        <p>
            $ COMMENT You can enter orders from a simple text file using the pipe character
            (|) to separate fields instead of TABs which Excel inserts between fields when you
            copy and paste from an Excel spreadsheet</p>
        <br />
        <p>
            $ COMMENT Sometimes Excel may automatically reformat a product code or product date
            that you enter. For example if you type a product date such as 06/03 &nbsp;it may
            appear as 06-Jul, which will then not match the value assigned to the product. To
            defeat this feature, enter an apostrophe &#39; before the value, eg &#39;06/03.</p>
        <br />
        <p>
            $ DEFAULTSPECIALINSTRUCTIONS||(no special instructions)</p>
        <p>
            $ DEFAULTPACKINGNOTE||(none)</p>
        <p>
            $ DEFAULTCUSTREF1|(none)</p>
        <p>
            $ DEFAULTCUSTREF2|(none)</p>
        <p>
            $ DEFAULTCUSTREF3|(none)</p>
        <p>
            $ DEFAULTCUSTREF4|(none)</p>
        <br />
        <p>
            $ DEFAULTEXTERNALREFERENCE||(none)</p>
        <br />
        <p>
            $ DEFAULTUK</p>
        <p>
            $ ITEMORDER|QPD</p>
        <br />
        <p>
            $ ITEM|3|123|06/03|||</p>
        <p>
            $ ITEM|1|14790 (HGI)</p>
        <br />
        <p>
            Chris Newport|27 Defoe Avenue|Kew||Richmond-upon-Thames|Surrey|TW9 4DS</p>
        <p>
            Catherine Kelley|Flat 10|Sandways||Richmond-upon-Thames|Surrey|TW9</p>
        <br />
        <br />
        <p>
            $ ITEM|4|14790 (HGI)</p>
        <p>
            John Smith|29 Ruskin Avenue|Kew||Richmond-upon-Thames|Surrey|TW9 4DV|UNITED 
            KINGDOM</p>
        <br />
        <br />
        [end]<br />
    </asp:Panel>
    <br />
    </form>
</body>
</html>
