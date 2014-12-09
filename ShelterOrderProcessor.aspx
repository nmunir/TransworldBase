<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Import Namespace="System.Data.Common" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="FileHelpers" %>
<script runat="server">
    
    'drOrder("CneeAddr1") = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(o.sField4Addr1OrPrice)
    'drOrder("CneeAddr2") = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(o.sField5Addr2)
    'drOrder("CneeAddr3") = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(o.sField6Addr3)
    'drOrder("CneeTown") = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(o.sField7TownCity)


    ' TO DO
    ' dummy user to place orders
    
    Const CUSTOMER_SHELTER As Int32 = 820
    'Const CUSTOMER_SHELTER As Int32 = 16
    Const ORDER_MARKER As String = "Transfer #"
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Dim bHaveOrderHeader As Boolean = False
    Dim gsOrderRef As String = String.Empty
    'Dim gsCneeName As String = String.Empty
    Dim gsCneeCtcName As String = String.Empty
    Dim gsCneeAddr1 As String = String.Empty
    'Dim gsCneeAddr2 As String = String.Empty
    'Dim gsCneeAddr3 As String = String.Empty
    Dim gsCneeTown As String = String.Empty
    Dim gsCneeMoreAddrInfo As String = String.Empty
    Dim gsCneePostcode As String = String.Empty
    Dim gdictProductsHumanReadableList As New SortedDictionary(Of String, Int32)
    Dim gdictProductsToOrder As New SortedDictionary(Of Int32, Int32)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            'Server.Transfer("session_expired.aspx")
        End If

        If Not IsPostBack Then
            Call CreateShelterOrdersFolder()
            Call SetTitle()
            Call PopulateOrderSummariesDropdown()
        End If
    End Sub
    
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sm As New ScriptManager
        sm.ID = "ScriptMgr"
        Try
            PlaceHolderForScriptManager.Controls.Add(sm)
        Catch ex As Exception
        End Try
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Create Shelter Orders"
    End Sub
    
    Protected Sub PopulateOrderSummariesDropdown()
        ddlOrderSummaries.Items.Clear()
        ddlOrderSummaries.Items.Add(New ListItem("- please select -", 0))
        Dim sSQL As String = "SELECT TOP 5 ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), OrderDateTime, 106), ' ', '-') + ' ' + SUBSTRING((CONVERT(VARCHAR(8), OrderDateTime, 108)),1,5) AS varchar(20)),'(never)') 'OrderDateTime', [id] FROM dbo.ClientData_Shelter_OrderSummary ORDER BY OrderDateTime DESC"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        For Each dr As DataRow In dt.Rows
            ddlOrderSummaries.Items.Add(New ListItem(dr("OrderDateTime"), dr("id")))
        Next
    End Sub
    
    Protected Function CaptureOrderHeader(o As ShelterOrderLine) As Boolean
        CaptureOrderHeader = True
        gsOrderRef = o.sField1OrderMarkerOrProductCode.Trim.Trim("""")
        'gsCneeName = o.sField2StockPoint.Trim.Trim("""")
        gsCneeCtcName = o.sField7ContactName.Trim.Trim("""")
        If gsCneeCtcName = String.Empty Then
            CaptureOrderHeader = False
        End If
        gsCneeAddr1 = o.sField8Addr1.Trim.Trim("""")
        If gsCneeAddr1 = String.Empty Then
            CaptureOrderHeader = False
        End If
        'gsCneeAddr2 = o.sField5Addr2.Trim.Trim("""")
        'gsCneeAddr3 = o.sField6Addr3.Trim.Trim("""")
        gsCneeTown = o.sField9TownCity.Trim.Trim("""")
        If gsCneeTown = String.Empty Then
            CaptureOrderHeader = False
        End If
        gsCneeMoreAddrInfo = o.sField10MoreAddressInfo.ToUpper.Trim.Trim("""")
        
        gsCneePostcode = o.sField11Postcode.ToUpper.Trim.Trim("""")
        If gsCneePostcode = String.Empty Then
            CaptureOrderHeader = False
        End If
        gdictProductsHumanReadableList.Clear()
        gdictProductsToOrder.Clear()
    End Function
    
    Protected Sub PopulateOrderHeaderFields(ByRef drOrder As DataRow)
        drOrder("OrderRef") = gsOrderRef
        'drOrder("CneeName") = gsCneeName
        drOrder("CneeCtcName") = gsCneeCtcName
        drOrder("CneeAddr1") = gsCneeAddr1
        drOrder("CneeAddr2") = gsCneeMoreAddrInfo
        'drOrder("CneeAddr3") = gsCneeAddr3
        drOrder("CneeTown") = gsCneeTown
        drOrder("CneePostcode") = gsCneePostcode
    End Sub
                                            
    Protected Sub ReadFile()
        Dim sFilePath As String = Server.MapPath("~/ShelterOrders/" & psFileName)
        Dim ShelterOrderLines As ShelterOrderLine()
        Dim engine As New FileHelperEngine(GetType(ShelterOrderLine))
        Try
            ShelterOrderLines = DirectCast(engine.ReadFile(sFilePath), ShelterOrderLine())
        Catch ex As Exception
            Call WriteToLog("Could not read file: " & ex.Message)
            Exit Sub
        End Try
        
        Dim dtOrders As DataTable = BuildConsignmentTable()
        Dim dictTotalProducts As New SortedDictionary(Of Int32, Int32)

        Dim lstProductsHumanReadableList As New List(Of String)
        Dim dictProducts As New SortedDictionary(Of Int32, Int32)

        Dim nOrderStatusValidated As Int32 = 0
        Dim nOrderStatusValidatedWithErrors As Int32 = 0
        Dim nOrderStatusBad As Int32 = 0
        
        Dim bFoundFirstOrderMarker As Boolean = False
        Dim bFoundOrderMarker As Boolean = False
        Dim nProductsInOrder As Int32 = 0
        Dim bOrderValid As Boolean = False
        
        Call WriteToLog("")
        Call WriteToLog("Processing file...")
        Call WriteToLog("")
        Dim nLineNo As Int32 = 1
        
        For Each o As ShelterOrderLine In ShelterOrderLines
            nLineNo += 1
            If o.sField1OrderMarkerOrProductCode.Contains(ORDER_MARKER) Then
                bFoundFirstOrderMarker = True
                If Not bHaveOrderHeader Then
                    bOrderValid = CaptureOrderHeader(o)
                    If bOrderValid Then
                        Call WriteToLog("LINE " & nLineNo & ": Captured valid order header (" & gsOrderRef & "; " & gsCneeCtcName & "; " & gsCneeAddr1 & "; " & gsCneeTown & "; " & gsCneeMoreAddrInfo & "; " & gsCneePostcode & ").")
                    Else
                        Call WriteToLog("LINE " & nLineNo & ": INVALID ORDER HEADER!")
                    End If
                    bHaveOrderHeader = True
                End If
                If bFoundOrderMarker Then
                    If nProductsInOrder = 0 Then
                        Call WriteToLog("LINE " & nLineNo & ": Empty order detected (or all products in previous order were invalid)!")
                    Else
                        If nProductsInOrder > 0 Then ' end of order, construct order record  ??????????????????????
                            Call WriteToLog("This order contains " & nProductsInOrder & " product(s).")
                            Call WriteToLog("")
                            Dim drOrder As DataRow = dtOrders.NewRow()
                            Call PopulateOrderHeaderFields(drOrder)
                            drOrder("ProductsHumanReadableList") = GetProductsHumanReadableStringFromDictionary(gdictProductsHumanReadableList)
                            drOrder("ProductsToOrder") = GetProductsEncodedStringFromDictionary(gdictProductsToOrder)
                            If Not bOrderValid Then
                                drOrder("StatusExcluded") = -1
                                nOrderStatusBad += 1
                            Else
                                drOrder("StatusExcluded") = 0
                                nOrderStatusValidated += 1
                            End If
                            dtOrders.Rows.Add(drOrder)
                            nProductsInOrder = 0
                            bOrderValid = True
                        Else
                            ' ?
                        End If
                    End If
                    bOrderValid = CaptureOrderHeader(o)
                    If bOrderValid Then
                        Call WriteToLog("LINE " & nLineNo & ": Captured valid order header (" & gsOrderRef & "; " & gsCneeCtcName & "; " & gsCneeAddr1 & "; " & gsCneeTown & "; " & gsCneeMoreAddrInfo & "; " & gsCneePostcode & ").")
                    Else
                        Call WriteToLog("LINE " & nLineNo & ": INVALID ORDER HEADER!")
                    End If
                Else
                    bFoundOrderMarker = True
                End If
            Else
                If bFoundFirstOrderMarker Then
                    If Not bFoundOrderMarker Then
                        Call WriteToLog("LINE " & nLineNo & ": Non start-of-order line detected while looking for start of order!")
                    Else   ' capture order content line
                        Dim bOrderLineValid As Boolean = True
                        Dim sProductCode As String = o.sField1OrderMarkerOrProductCode
                        Dim sQuantity As String = o.sField4ProductQuantity
                        Dim nProductKey As Int32 = 0
                        nProductKey = GetProductKeyFromProductCode(sProductCode)
                        If nProductKey <= 0 Then
                            bOrderLineValid = False
                        End If
                        If Not IsNumeric(sQuantity) Then
                            bOrderLineValid = False
                        Else
                            If CInt(sQuantity) <= 0 Or CInt(sQuantity) <> sQuantity Then
                                bOrderLineValid = False
                            End If
                        End If
                        If bOrderLineValid Then
                            Call WriteToLog("LINE " & nLineNo & ": product: " & sProductCode & "; qty: " & sQuantity & ".")
                            Try
                                gdictProductsHumanReadableList.Add(sProductCode, CInt(sQuantity))
                                gdictProductsToOrder.Add(nProductKey, CInt(sQuantity))
                                nProductsInOrder += 1
                                Call AddProduct(GetProductKeyFromProductCode(sProductCode), CInt(sQuantity), lstProductsHumanReadableList, dictProducts, dictTotalProducts)
                            Catch ex As Exception
                                Call WriteToLog("-- On line " & nLineNo & ": Could not add product " & sProductCode & " to order. This line of the order may be a duplicate.")
                                bOrderValid = False
                            End Try
                        Else
                            Call WriteToLog("LINE " & nLineNo & ": Invalid order product line (" & sProductCode & " - " & sQuantity & ")!")
                            If Not cbIgnoreInvalidProductLines.Checked Then
                                bOrderValid = False
                            End If
                        End If
                    End If
                End If
            End If
        Next

        If bFoundOrderMarker And nProductsInOrder > 0 Then
            Call WriteToLog("This order contains " & nProductsInOrder & " product(s).")
            Call WriteToLog("")
            Dim drOrder As DataRow = dtOrders.NewRow()
            Call PopulateOrderHeaderFields(drOrder)
            drOrder("ProductsHumanReadableList") = GetProductsHumanReadableStringFromDictionary(gdictProductsHumanReadableList)
            drOrder("ProductsToOrder") = GetProductsEncodedStringFromDictionary(gdictProductsToOrder)
            If Not bOrderValid Then
                drOrder("StatusExcluded") = -1
                nOrderStatusBad += 1
            Else
                drOrder("StatusExcluded") = 0
                nOrderStatusValidated += 1
            End If
            dtOrders.Rows.Add(drOrder)
        End If

        '    drOrder("StatusExcluded") = 1 if order already placed

        Dim bAvailableQuantityError As Boolean = False
        
        WriteToLog("")
        For Each kv As KeyValuePair(Of Int32, Int32) In dictTotalProducts
            Dim nAvailableQuantity As Int32 = GetAvailableQty(kv.Key)
            If kv.Value > nAvailableQuantity Then
                WriteToLog("===>>> WARNING: Available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is insufficient to fulfil all orders (required: " & kv.Value & ")")
                bAvailableQuantityError = True
            Else
                WriteToLog("Confirmed required quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & kv.Value & ") is less than or equal to total amount available (" & nAvailableQuantity & ")")
            End If
        Next

        If bAvailableQuantityError Then
            lblInsufficientQuantity.Visible = True
        Else
            lblInsufficientQuantity.Visible = False
        End If
        pdtOrders = dtOrders
        
        gvOrders.DataSource = dtOrders
        gvOrders.DataBind()
        lblLegendOrders.Visible = True
        btnRecheckQuantities.Enabled = True
        btnCreateConsignments.Enabled = True

        Dim sMessage As String = String.Empty
        For Each kv As KeyValuePair(Of Int32, Int32) In dictTotalProducts
            Dim nAvailableQuantity As Int32 = GetAvailableQty(kv.Key)
            Dim sProductCode As String = GetProductCodeFromProductKey(kv.Key)
            If kv.Value <= nAvailableQuantity Then
                'sMessage &= "Sufficient " & GetProductCodeFromProductKey(kv.Key) & " (avail: " & nAvailableQuantity & ", rqd: " & kv.Value & ")" & "\n\n"
                Call WriteToLog("Available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is sufficient to fulfil all orders (required: " & kv.Value & ")")
            Else
                sMessage &= "===>>> WARNING: INSUFFICIENT " & GetProductCodeFromProductKey(kv.Key) & " (avail: " & nAvailableQuantity & ", rqd: " & kv.Value & ")" & "\n\n"
                Call WriteToLog("===>>> WARNING: available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is insufficient to fulfil all orders (required: " & kv.Value & ")")
            End If
        Next
        If nOrderStatusBad > 0 Then
            sMessage &= "\n\nVIEW THE JOURNAL WINDOW FOR MORE DETAILS OF WHICH ORDERS FAILED VALIDATION AND WHY."
        End If
        WebMsgBox.Show("Total orders: " & nOrderStatusValidated + nOrderStatusBad & "\n\nValidated Orders: " & nOrderStatusValidated & "\n\nFailed Validation: " & nOrderStatusBad & "\n\n" & sMessage)
    End Sub
    
    Protected Sub AddProduct(ByVal nProductKey As Int32, ByVal nQuantity As Int32, ByRef lstProductHumanReadableList As List(Of String), ByRef dictProducts As SortedDictionary(Of Int32, Int32), ByRef dictTotalProducts As SortedDictionary(Of Int32, Int32))
        lstProductHumanReadableList.Add(GetProductCodeFromProductKey(nProductKey) & " (" & nQuantity.ToString & ")")
        Try
            dictProducts.Add(nProductKey, nQuantity)
        Catch ex As Exception
            dictProducts(nProductKey) = dictProducts(nProductKey) + nQuantity
        End Try
        Try
            dictTotalProducts.Add(nProductKey, nQuantity)
        Catch ex As Exception
            dictTotalProducts(nProductKey) = dictTotalProducts(nProductKey) + nQuantity
        End Try
    End Sub
    
    Protected Function GetProductKeyFromProductCode(ByVal sProductCode As String) As Int32
        Dim sSQL As String = "SELECT LogisticProductKey FROM LogisticProduct WHERE ProductCode = '" & sProductCode & "' AND DeletedFlag = 'N' AND ArchiveFlag = 'N' AND CustomerKey = " & CUSTOMER_SHELTER
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 1 Then
            GetProductKeyFromProductCode = dt.Rows(0).Item(0)
        Else
            GetProductKeyFromProductCode = 0
        End If
    End Function

    Protected Function GetProductCodeFromProductKey(ByVal sProductKey As Int32) As String
        GetProductCodeFromProductKey = ExecuteQueryToDataTable("SELECT ProductCode FROM LogisticProduct WHERE LogisticProductKey =" & sProductKey).Rows(0).Item(0)
    End Function

    Protected Function IsValidProductAndQuantity(ByVal sProductCode As String, ByVal sQuantity As String) As Int32  ' returns LogisticProductKey if valid, 0 if both fields empty, -1 if invalid
        IsValidProductAndQuantity = 0
        If sProductCode <> String.Empty Then
            If IsNumeric(sQuantity) Then
                Dim nQuantity = CInt(sQuantity)
                If nQuantity > 0 Then
                    Dim nLogisticProductKey As Int32 = GetProductKeyFromProductCode(sProductCode)
                    If nLogisticProductKey > 0 Then
                        IsValidProductAndQuantity = nLogisticProductKey
                    Else
                        Call WriteToLog("Could not match product (" & sProductCode & ").")
                        IsValidProductAndQuantity = -1
                    End If
                Else
                    Call WriteToLog("Quantity must be greater than zero.")
                    IsValidProductAndQuantity = -1
                End If
            Else
                Call WriteToLog("Non-blank product code with non-numeric quantity (" & sQuantity & ").")
                IsValidProductAndQuantity = -1
            End If
        Else
            If sQuantity <> String.Empty Then
                IsValidProductAndQuantity = -1
            End If
        End If
    End Function
    
    Protected Function GetProductsHumanReadableStringFromDictionary(dictHumanReadableProductList As SortedDictionary(Of String, Int32)) As String
        Dim s As String = String.Empty
        For Each kv As KeyValuePair(Of String, Int32) In dictHumanReadableProductList
            s &= kv.Key & " (" & kv.Value.ToString & "), "
        Next
        If s <> String.Empty Then
            s = s.Substring(0, s.Length - 2)
        End If
        GetProductsHumanReadableStringFromDictionary = s
    End Function
    
    Protected Function GetProductsEncodedStringFromDictionary(dictProducts As SortedDictionary(Of Int32, Int32)) As String
        Dim s As String = String.Empty
        For Each kv As KeyValuePair(Of Int32, Int32) In dictProducts
            s &= kv.Key & "," & kv.Value & ","
        Next
        If s <> String.Empty Then
            s = s.Substring(0, s.Length - 1)
        End If
        GetProductsEncodedStringFromDictionary = s
    End Function
    
    Protected Function GetProductsDictionaryFromEncodedString(sProductsToOrder As String) As SortedDictionary(Of Int32, Int32) ' dehydrate from string to dictioary
        Dim dictProducts As New SortedDictionary(Of Int32, Int32)
        Dim arrProductsToOrder() = sProductsToOrder.Split(",")
        For i As Int32 = 0 To arrProductsToOrder.Count - 1 Step 2
            dictProducts.Add(arrProductsToOrder(i), arrProductsToOrder(i + 1))
        Next
        GetProductsDictionaryFromEncodedString = dictProducts
    End Function
    
    Protected Function BuildConsignmentTable() As DataTable
        Dim dtOrders As New DataTable
        'dtOrders.Columns.Add(New DataColumn("CneeName", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("CneeCtcName", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("CneeAddr1", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("CneeAddr2", Type.GetType("System.String")))
        'dtOrders.Columns.Add(New DataColumn("CneeAddr3", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("CneeTown", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("CneePostcode", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("OrderRef", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("ProductsHumanReadableList", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("ProductsToOrder", Type.GetType("System.String")))

        Dim dc As New DataColumn("StatusExcluded")
        dc.DataType = GetType(Integer)
        dc.AllowDBNull = True
        dtOrders.Columns.Add(dc)

        Return dtOrders
    End Function

    '   <DelimitedRecord(",")> _
    '   <IgnoreFirst(1)> _

    <DelimitedRecord(",")> _
    <IgnoreFirst(0)> _
    Public Class ShelterOrderLine
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField1OrderMarkerOrProductCode As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField2StockPoint As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField3ProductDescription As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField4ProductQuantity As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField5ProductPrice As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sBlankField06 As String

        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField7ContactName As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField8Addr1 As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField9TownCity As String

        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField10MoreAddressInfo As String

        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sField11Postcode As String

    End Class

    Protected Sub btnCheckData_Click(sender As Object, e As System.EventArgs)
        Call ReadFile()
    End Sub

    Private Sub CreateShelterOrdersFolder()
        Dim sPath As String = Server.MapPath("~/")
        If Not Directory.Exists(sPath & "\ShelterOrders") Then
            Directory.CreateDirectory(sPath & "\ShelterOrders")
        End If
    End Sub

    Protected Sub btnUpload_Click(ByVal sender As Object, ByVal e As EventArgs)
        If psFileName <> String.Empty Then
            Call ReadFile()
        Else
            Try
                If ruShelterFileUpload.UploadedFiles.Count > 0 Then
                    lblNoResults.Visible = False
                    psFileName = ruShelterFileUpload.UploadedFiles(0).GetName
                    lblFileName.Text = ruShelterFileUpload.UploadedFiles(0).GetName
                    lblLegendFile.Visible = True
                    tbLog.Text = String.Empty
                    lnkbtnClearFile.Visible = True
                    Call WriteToLog("Uploaded " & ruShelterFileUpload.UploadedFiles(0).GetName & " @ " & Format(Date.Now, "d-MMM-yyyy hh:mm:ss"))
                    Call ReadFile()
                Else
                    btnRecheckQuantities.Enabled = False
                    btnCreateConsignments.Enabled = False
                    lnkbtnClearFile.Visible = False
                    lblNoResults.Visible = True
                    divFileInfo.Visible = False
                    Call WriteToLog("Nothing uploaded")
                End If
            Catch ex As Exception
                lblNoResults.Text = ex.Message.ToString()
                Call WriteToLog(ex.Message.ToString())
            End Try
        End If
    End Sub
    
    Private Function GetAvailableQty(ByVal nLogisticProductKey As Integer) As Integer
        Dim sSQL As String = "SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ") END"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
            GetAvailableQty = oDataTable.Rows(0)(0)
        Else
            GetAvailableQty = 0
        End If
    End Function

    Protected Sub WriteToLog(sMessage As String)
        tbLog.Text += sMessage & Environment.NewLine
    End Sub
    
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

    Protected Sub SaveOrderSummary()
        Dim dictTotalProducts As New SortedDictionary(Of Int32, Int32)
        Dim dtOrders As DataTable = pdtOrders
        For Each drOrder As DataRow In dtOrders.Rows
            If drOrder("StatusExcluded") = 0 Then
                Dim dictProductsToOrder As New SortedDictionary(Of Int32, Int32)
                dictProductsToOrder = GetProductsDictionaryFromEncodedString(drOrder("ProductsToOrder"))
                For Each kv As KeyValuePair(Of Int32, Int32) In dictProductsToOrder
                    Try
                        dictTotalProducts.Add(kv.Key, kv.Value)
                    Catch ex As Exception
                        dictTotalProducts(kv.Key) = dictTotalProducts(kv.Key) + kv.Value
                    End Try
                Next
            End If
        Next

        Dim sMessage As String = "SHELTER ORDER SUMMARY - order placed " & DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss") & Environment.NewLine & Environment.NewLine
        For Each kv As KeyValuePair(Of Int32, Int32) In dictTotalProducts
            sMessage &= "Required quantity of product " & GetProductCodeFromProductKey(kv.Key) & ": " & kv.Value & Environment.NewLine
        Next
        Dim sSQL As String = "INSERT INTO ClientData_Shelter_OrderSummary (OrderDateTime, OrderSummary) VALUES (GETDATE(), '" & sMessage & "')"
        Call ExecuteQueryToDataTable(sSQL)
        WebMsgBox.Show(sMessage)
    End Sub
    
    Protected Sub btnRecheckQuantities_Click(sender As Object, e As System.EventArgs)
        Dim dictTotalProducts As New SortedDictionary(Of Int32, Int32)
        Dim dtOrders As DataTable = pdtOrders
        For Each drOrder As DataRow In dtOrders.Rows
            If drOrder("StatusExcluded") = 0 Then
                Dim dictProductsToOrder As New SortedDictionary(Of Int32, Int32)
                dictProductsToOrder = GetProductsDictionaryFromEncodedString(drOrder("ProductsToOrder"))
                For Each kv As KeyValuePair(Of Int32, Int32) In dictProductsToOrder
                    Try
                        dictTotalProducts.Add(kv.Key, kv.Value)
                    Catch ex As Exception
                        dictTotalProducts(kv.Key) = dictTotalProducts(kv.Key) + kv.Value
                    End Try
                Next
            End If
        Next

        Dim sMessage As String = String.Empty
        For Each kv As KeyValuePair(Of Int32, Int32) In dictTotalProducts
            Dim nAvailableQuantity As Int32 = GetAvailableQty(kv.Key)
            If kv.Value <= nAvailableQuantity Then
                'sMessage &= "Available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is sufficient to fulfil all orders (required: " & kv.Value & ")" & "\n\n"
                sMessage &= "Sufficient " & GetProductCodeFromProductKey(kv.Key) & " (avail: " & nAvailableQuantity & ", rqd: " & kv.Value & ")" & "\n\n"

                Call WriteToLog("Available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is sufficient to fulfil all orders (required: " & kv.Value & ")")
            Else
                'sMessage &= "===>>> WARNING: available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is insufficient to fulfil all orders (required: " & kv.Value & ")" & "\n\n"
                sMessage &= "===>>> WARNING: INSUFFICIENT " & GetProductCodeFromProductKey(kv.Key) & " (avail: " & nAvailableQuantity & ", rqd: " & kv.Value & ")" & "\n\n"

                Call WriteToLog("===>>> WARNING: available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is insufficient to fulfil all orders (required: " & kv.Value & ")")
            End If
        Next
        WebMsgBox.Show(sMessage)
    End Sub
    
    Protected Function IsSufficientQuantity() As Boolean
        IsSufficientQuantity = True
        Dim dictTotalProducts As New SortedDictionary(Of Int32, Int32)
        Dim dtOrders As DataTable = pdtOrders
        For Each drOrder As DataRow In dtOrders.Rows
            If drOrder("StatusExcluded") = 0 Then
                Dim dictProductsToOrder As New SortedDictionary(Of Int32, Int32)
                dictProductsToOrder = GetProductsDictionaryFromEncodedString(drOrder("ProductsToOrder"))
                For Each kv As KeyValuePair(Of Int32, Int32) In dictProductsToOrder
                    Try
                        dictTotalProducts.Add(kv.Key, kv.Value)
                    Catch ex As Exception
                        dictTotalProducts(kv.Key) = dictTotalProducts(kv.Key) + kv.Value
                    End Try
                Next
            End If
        Next

        Dim sMessage As String = String.Empty
        For Each kv As KeyValuePair(Of Int32, Int32) In dictTotalProducts
            If kv.Value > GetAvailableQty(kv.Key) Then
                IsSufficientQuantity = False
                Exit For
            End If
        Next
    End Function
    
    Protected Function nSubmitConsignment(ByRef drOrder As DataRow) As Int32
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
        param1.Value = GetGenericUser()
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = CUSTOMER_SHELTER
        oCmdAddBooking.Parameters.Add(param2)
        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)
        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = ""
        oCmdAddBooking.Parameters.Add(param3)
        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = ""
        oCmdAddBooking.Parameters.Add(param4)
        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = drOrder("OrderRef")
        oCmdAddBooking.Parameters.Add(param5)
        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = ""
        oCmdAddBooking.Parameters.Add(param6)
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = Nothing
        oCmdAddBooking.Parameters.Add(param6a)
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = ""
        oCmdAddBooking.Parameters.Add(param7)
        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = ""
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
        
        Dim sSQL As String = "SELECT * FROM Customer WHERE CustomerKey = " & CUSTOMER_SHELTER
        Dim dtCnor As DataTable = ExecuteQueryToDataTable(sSQL)
        
        Dim drCnor As DataRow
        If dtCnor.Rows.Count = 1 Then
            drCnor = dtCnor.Rows(0)
        Else
            WebMsgBox.Show("Couldn't find Consignor details.")
            Exit Function
        End If
       
        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
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
        param25.Value = "SHELTER"
        oCmdAddBooking.Parameters.Add(param25)
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = drOrder("CneeAddr1")
        oCmdAddBooking.Parameters.Add(param26)
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param27)
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param28)
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = drOrder("CneeTown")
        oCmdAddBooking.Parameters.Add(param29)
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = ""
        oCmdAddBooking.Parameters.Add(param30)
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = drOrder("CneePostcode")
        oCmdAddBooking.Parameters.Add(param31)
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = 222
        oCmdAddBooking.Parameters.Add(param32)
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = drOrder("CneeCtcName")
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
                Dim dictProductsToOrder As New SortedDictionary(Of Int32, Int32)
                dictProductsToOrder = GetProductsDictionaryFromEncodedString(drOrder("ProductsToOrder"))
                For Each kv As KeyValuePair(Of Int32, Int32) In dictProductsToOrder
                    Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                    oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                    Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                    param51.Value = GetGenericUser()
                    oCmdAddStockItem.Parameters.Add(param51)
                    Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                    param52.Value = CUSTOMER_SHELTER
                    oCmdAddStockItem.Parameters.Add(param52)
                    Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                    param53.Value = lBookingKey
                    oCmdAddStockItem.Parameters.Add(param53)
                    Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                    param54.Value = kv.Key
                    oCmdAddStockItem.Parameters.Add(param54)
                    Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                    param55.Value = "PENDING"
                    oCmdAddStockItem.Parameters.Add(param55)
                    Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                    param56.Value = kv.Value
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
    
    Protected Function GetGenericUser() As Int32
        GetGenericUser = 0
        Dim sSQL As String = "SELECT [key] FROM UserProfile WHERE UserID = 'shelterGU' AND CustomerKey = " & CUSTOMER_SHELTER
        Dim dtGenericUser As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtGenericUser.Rows.Count = 1 Then
            GetGenericUser = dtGenericUser.Rows(0).Item(0)
        End If
    End Function
    
    Protected Sub btnCreateConsignments_Click(sender As Object, e As System.EventArgs)
        Dim dtOrders As DataTable = pdtOrders
        If IsSufficientQuantity() Then
            Dim nOrderCount As Int32
            Dim bConsignmentFailed As Boolean = False
            For Each drOrder As DataRow In dtOrders.Rows
                If drOrder("StatusExcluded") = 0 Then
                    Dim nConsignmentKey As Int32 = nSubmitConsignment(drOrder)
                    If nConsignmentKey > 0 Then
                        WriteToLog("Order " & drOrder("OrderRef") & " successfully created as consignment " & nConsignmentKey.ToString)
                        nOrderCount += 1
                    Else
                        WriteToLog("WARNING: Could not create consignment for order " & drOrder("OrderRef"))
                        bConsignmentFailed = True
                    End If
                End If
            Next
            If bConsignmentFailed Then
                lblConsignmentFailed.Visible = True
            End If
            WebMsgBox.Show("Created " & nOrderCount & " consignment(s)")
            btnCreateConsignments.Enabled = False
            Call SaveOrderSummary()
            Call PopulateOrderSummariesDropdown()
        Else
            WebMsgBox.Show("One or more products has insufficient quantity to fulfil all orders.\n\nPlease review and adjust the orders and/or quantities\n\nNO CONSIGNMENTS HAVE BEEN CREATED.")
        End If
    End Sub
    
    Protected Sub lnkbtnExclude_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim sOrderRef As String = lnkbtn.CommandArgument
        Dim dtOrders As DataTable = pdtOrders
        If Not lnkbtn.Text.Contains("RESTORE") Then
            For Each drOrder As DataRow In dtOrders.Rows
                If sOrderRef = drOrder("OrderRef") Then
                    drOrder("StatusExcluded") = 1
                    Exit For
                End If
            Next
        Else
            For Each drOrder As DataRow In dtOrders.Rows
                If sOrderRef = drOrder("OrderRef") Then
                    drOrder("StatusExcluded") = 0
                    Exit For
                End If
            Next
        End If
        pdtOrders = dtOrders
        gvOrders.DataSource = dtOrders
        gvOrders.DataBind()
    End Sub
    
    Protected Sub gvOrders_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim hidExcluded As HiddenField
            Dim lnkbtnExclude As LinkButton
            hidExcluded = gvr.Cells(0).FindControl("hidExcluded")
            lnkbtnExclude = gvr.Cells(0).FindControl("lnkbtnExclude")
            If hidExcluded.Value < 0 Then
                gvr.BackColor = Drawing.Color.Red
                lnkbtnExclude.Text = "excluded"
                lnkbtnExclude.Enabled = False
            ElseIf hidExcluded.Value > 0 Then
                gvr.BackColor = Drawing.Color.LightGray
                lnkbtnExclude.Text = "RESTORE"
            End If
        End If
    End Sub
    
    Property psFileName() As String
        Get
            Dim o As Object = ViewState("Shelter_Order_FileName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("Shelter_Order_FileName") = Value
        End Set
    End Property

    Property pdtOrders() As DataTable
        Get
            Dim o As Object = ViewState("ShelterOrders")
            If o Is Nothing Then
                Return Nothing
            End If
            Return o
        End Get
        Set(ByVal Value As DataTable)
            ViewState("ShelterOrders") = Value
        End Set
    End Property
    
    Protected Sub lnkbtnClearFile_Click(sender As Object, e As System.EventArgs)
        psFileName = String.Empty
        lnkbtnClearFile.Visible = False
    End Sub
    
    Protected Sub ddlOrderSummaries_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Dim sSQL As String = "SELECT OrderSummary FROM dbo.ClientData_Shelter_OrderSummary WHERE [id] = " & ddl.SelectedValue
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        tbLog.Text = dt.Rows(0).Item(0)
    End Sub
    
    Protected Sub lnkbtnPlus2_Click(sender As Object, e As System.EventArgs)
        tbLog.Rows = tbLog.Rows + 10
        lnkbtnMinus2.Enabled = True
    End Sub

    Protected Sub lnkbtnMinus2_Click(sender As Object, e As System.EventArgs)
        If tbLog.Rows > 12 Then
            tbLog.Rows = tbLog.Rows - 10
            If tbLog.Rows <= 2 Then
                lnkbtnMinus2.Enabled = False
            End If
        End If
    End Sub
    
</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Create Shelter Orders</title>
    <style type="text/css">
        .label
        {
            margin-left: 1px;
            width: 15%;
            float: left;
            text-align: left;
        }
        .style1
        {
            color: #FF0000;
        }
    </style>
    <script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>
    <%--    <script type="text/javascript">
        $(document).ready(function () {
            $('label').addClass('label');
        });
--%></script>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_addressbook">
            <td style="width: 50%; white-space: nowrap">
                &nbsp;
                    <asp:Label ID="lblTitle" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small">Shelter Order Processor</asp:Label>
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
                &nbsp;
            </td>
        </tr>
    </table>
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server"/>
    <div>
        &nbsp;
        <asp:Label ID="lblLegendSelect" 
            Text="1. Select the Shelter &lt;b&gt;CSV&lt;/b&gt; order file:" runat="server" 
            Font-Bold="False" Font-Names="Verdana" Font-Size="X-Small" />
        <telerik:RadUpload ID="ruShelterFileUpload" Width="100%" TargetFolder="~/ShelterOrders" AllowedFileExtensions=".csv" ToolTip="Select a csv file (.csv)" MaxFileInputsCount="1" OverwriteExistingFiles="true" runat="server" BackColor="#FFE7CE" ControlObjectsVisibility="None" />
        &nbsp;<asp:Label ID="lblLegendFile" Text="File:" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="X-Small" Visible="False" />
        <asp:Label ID="lblFileName" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" />
        &nbsp;&nbsp;
        <asp:LinkButton ID="lnkbtnClearFile" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" onclick="lnkbtnClearFile_Click" Visible="False">clear file</asp:LinkButton>
        <br />
        &nbsp;<asp:Button runat="server" ID="btnUpload" Text="2. Upload &amp; Check Order File" OnClick="btnUpload_Click" Width="200px" />
        &nbsp;&nbsp;
        <asp:CheckBox ID="cbIgnoreInvalidProductLines" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="ignore invalid product lines *" />
        <br />
        <br />
        &nbsp;<asp:Button ID="btnRecheckQuantities" runat="server" OnClick="btnRecheckQuantities_Click" Text="2a. Re-check Quantities" Width="200px" Enabled="False" />
        &nbsp;<asp:Label ID="lblAdvice" Text="(eg after including or excluding one or more orders)" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" />
        <br />
        <br />
        &nbsp;<asp:Button runat="server" ID="btnCreateConsignments" Text="3. Create Consigments" Enabled="false" ValidationGroup="vg" Width="200px" OnClick="btnCreateConsignments_Click" />
    &nbsp;&nbsp;&nbsp;&nbsp; <asp:Label ID="lblLegendViewPreviousOrderSummaries" 
            Text="View previous order summaries:" runat="server" Font-Bold="False" 
            Font-Names="Verdana" Font-Size="XX-Small" />
        <asp:DropDownList ID="ddlOrderSummaries" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" AutoPostBack="True" 
            onselectedindexchanged="ddlOrderSummaries_SelectedIndexChanged">
        </asp:DropDownList>
    </div>
    <br />
    <asp:Label ID="lblNoResults" runat="server" Visible="false" Text="No file uploaded." Font-Names="Verdana" Font-Size="X-Small" Font-Bold="true" />
    &nbsp;<asp:Label ID="lblInsufficientQuantity" runat="server" Visible="False" Text="One or more products has insufficient quantity to fulfil all orders." Font-Bold="True" ForeColor="Red" Font-Names="Verdana" Font-Size="X-Small" />
    <asp:Label ID="lblConsignmentFailed" runat="server" Visible="False" Text="One or more consignments failed." Font-Bold="True" ForeColor="Red" Font-Names="Verdana" Font-Size="X-Small" />
    <br />
    <div id="divFileInfo" runat="server" visible="false">
        &nbsp;
        <br />
    </div>
    <asp:Label ID="lblLegendJournal" Text="Journal:" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="X-Small" />
    &nbsp;
    <asp:LinkButton ID="lnkbtnPlus2" runat="server" Font-Names="Verdana" 
        Font-Size="XX-Small" onclick="lnkbtnPlus2_Click">+10 rows</asp:LinkButton>
&nbsp;<asp:LinkButton ID="lnkbtnMinus2" runat="server" Font-Names="Verdana" 
        Font-Size="XX-Small" onclick="lnkbtnMinus2_Click">-10 rows</asp:LinkButton>
    <asp:TextBox ID="tbLog" runat="server" ReadOnly="True" TextMode="MultiLine" Width="99%" Rows="10" Font-Names="Arial" Font-Size="XX-Small" />
    <br />
    <br />
        <asp:Label ID="lblLegendOrders" Text="Orders:" runat="server" Font-Bold="False" Font-Names="Verdana"
            Font-Size="X-Small" Visible="False" />
    <asp:GridView ID="gvOrders" runat="server" CellPadding="1" Font-Names="Verdana" Font-Size="XX-Small"
        Width="100%" AutoGenerateColumns="False" OnRowDataBound="gvOrders_RowDataBound">
        <Columns>
            <asp:TemplateField>
                <ItemTemplate>
                    <asp:LinkButton ID="lnkbtnExclude" runat="server" CommandArgument='<%# Container.DataItem("OrderRef")%>'
                        OnClick="lnkbtnExclude_Click">exclude</asp:LinkButton>
                    <asp:HiddenField ID="hidExcluded" runat="server" Value='<%# Container.DataItem("StatusExcluded")%>' />
                </ItemTemplate>
                <ItemStyle Width="70px" />
            </asp:TemplateField>
            <asp:BoundField DataField="OrderRef" HeaderText="Order Ref" ReadOnly="True" SortExpression="OrderRef" />
            <%--<asp:BoundField DataField="CneeName" HeaderText="Consignee" ReadOnly="True" SortExpression="CneeName" />--%>
            <asp:BoundField DataField="CneeCtcName" HeaderText="Contact" ReadOnly="True" SortExpression="CneeCtcName" />
            <asp:BoundField DataField="CneeAddr1" HeaderText="Addr 1" ReadOnly="True" SortExpression="CneeAddr1" />
            <%--<asp:BoundField DataField="CneeAddr2" HeaderText="Addr 2" ReadOnly="True" SortExpression="CneeAddr2" />
            <asp:BoundField DataField="CneeAddr3" HeaderText="Addr 3" ReadOnly="True" SortExpression="CneeAddr3" />--%>
            <asp:BoundField DataField="CneeTown" HeaderText="Town / City" ReadOnly="True" SortExpression="CneeTown" />
            <asp:BoundField DataField="CneePostcode" HeaderText="Post code" ReadOnly="True" SortExpression="CneePostcode" />
            <asp:BoundField DataField="ProductsHumanReadableList" HeaderText="Products" ReadOnly="True"
                SortExpression="ProductsHumanReadableList" />
        </Columns>
        <EmptyDataTemplate>
            no orders found
        </EmptyDataTemplate>
    </asp:GridView>
    <asp:Panel ID="pnlHelp" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
        <br />
        <b>INSTRUCTIONS &amp; NOTES</b><br />
        <br />
        1.&nbsp; Download the file of Shelter orders to your PC.<span class="style1"> If this 
        is an Excel Spreadsheet (.xls or .xlsx file) convert it to CSV format.</span><br />
        <br />
        2. Click the <b>Select</b> button. Navigate to the downloaded file.<br />
        <br />
        3.&nbsp; Click the button labelled <b>2. Upload &amp; Check Spreadsheet</b> .
        <br />
        <br />
        The system reads the CSV order file and displays the list of orders for you to view
        before creating consignments,<br />
        <br />
        It validates each order and checks that enough quantity of each requested 
        product exists. Validation errors are listed in the <b>Journal</b> box.&nbsp; Any 
        order that cannot be read is marked is marked &#39;excluded&#39; and highlighted in red. 
        These excluded orders must be investigated and corrected.<br />
        <br />
        * check the <b>ignore invalid product lines</b> check box, then click <b>Upload 
        &amp; Check Order File</b> again, if you want to include orders where validation of 
        one or more of the products failed, eg because the product cannot be matched, or 
        is archived. An order will be created that includes just the product(s) that was 
        validated successfully.<br />
        <br />
        Click <b>exclude</b> to remove any other order you do not want generated (eg 
        because of address problems, insufficient product quantity available to satisfy 
        all orders, etc).<br />
        <br />
        4.&nbsp; After excluding or restoring orders you can re-check the button <b>2a 
        Re-check Quantities</b> to verify that enough of each requested product is in 
        stock.<br />
        <br />
        5.&nbsp; Click the <b>Create Consignments</b> button to generate the orders in the 
        list. The system reports the number of consignments generated.<br />
        <br />
        LAST UPDATE: 15MAR13<br />
        <br />
    </asp:Panel>
    </form>
</body>
</html>
